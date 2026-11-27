/**
 * report-export.helpers.ts
 *
 * Report-agnostic export pipeline used by the report-parity (migration) flow.
 *
 *   exportReportToWorkbook()  embeds the CURRENTLY ACTIVE report (identity set
 *                             via applyReportIdentity), exports every visual on
 *                             every requested page, and writes a stable,
 *                             comparison-friendly .xlsx (one sheet per page).
 *
 *   writeParitySummary()      turns a WorkbookDiffSummary into a readable .xlsx
 *                             summary (overall verdict + per-page diffs + samples).
 *
 * The per-report workbooks contain ONLY data sheets (no run-info sheet) so the
 * two-tier row comparison in excel-compare.helpers stays clean.
 */

import type { Page } from '@playwright/test';
import * as fs   from 'fs';
import ExcelJS    from 'exceljs';
import {
  loadHarnessPage,
  getReportPages,
  setReportPage,
  exportCurrentPageVisuals,
  type ReportPage,
  type VisualExport,
} from './harness.helpers';
import { timed, recordTiming } from './timing.helpers';
import {
  classifyAll, tally, biggestChanges, referenceFacts, rowCapVisuals, sameFiltersAnswer, pageFilterText,
  BUCKET_ORDER, BUCKET_ICON, BUCKET_MEANING, EXPORT_ROW_CAP, type Bucket, type ClassifiedVisual,
} from './summary-format.helpers';
import type { WorkbookDiffSummary, VisualBlockDiff, SheetDiffSummary } from './excel-compare.helpers';

// Visual types that carry no tabular data — never written or compared.
const SKIP_TYPES = new Set(['image', 'shape', 'textbox', 'actionButton']);
const EXCEL_MAX_SHEET_NAME = 31;

// One retry, after a short pause, for the same class of transient hiccup
// readSlicerOptions() in harness.helpers.ts already retries around (observed
// in practice on busy pages) — cheap insurance against a single page silently
// looking like a genuine data difference between source and target.
const NAV_RETRY_WAIT_MS    = 1500;
const EXPORT_RETRY_WAIT_MS = 1500;

/**
 * Errors that cross the page.evaluate() boundary aren't always proper Error
 * instances with a usable .message — some Power BI SDK rejections are plain
 * objects that stringify down to "[object Object]" or just "Object" if you
 * assume .message exists. This pulls out whatever is actually readable.
 */
function describeError(e: unknown): string {
  if (e instanceof Error && e.message) return e.message;
  if (typeof e === 'string') return e;
  try {
    const s = JSON.stringify(e);
    if (s && s !== '{}') return s;
  } catch {
    // fall through
  }
  return String(e);
}

// ── Sheet-name helpers (Excel forbids : \ / ? * [ ] and 31-char names) ────────
function toExcelSafeSheetBase(name: string): string {
  const cleaned = name.replace(/[:\\/?*[\]]/g, ' ').trim();
  return cleaned || 'Sheet';
}

/**
 * Truncates by Unicode CODE POINT, not raw string index. Plain .slice()
 * indexes by UTF-16 code unit - fine for the BMP (covers Latin, accented
 * Latin, Cyrillic, most CJK) but a page name containing an emoji or other
 * astral-plane character (a surrogate pair, 2 code units) could get cut
 * exactly between the two halves of the pair, producing a malformed,
 * unpaired-surrogate string.
 */
export function truncateSafe(text: string, maxLen: number): string {
  return Array.from(text).slice(0, maxLen).join('');
}

function toUniqueSheetName(name: string, usedNames: Set<string>): string {
  const base = toExcelSafeSheetBase(name);
  let candidate = truncateSafe(base, EXCEL_MAX_SHEET_NAME);
  let counter = 2;
  while (usedNames.has(candidate.toLowerCase())) {
    const suffix = `_${counter}`;
    const maxBaseLen = EXCEL_MAX_SHEET_NAME - suffix.length;
    candidate = `${truncateSafe(base, Math.max(1, maxBaseLen))}${suffix}`;
    counter += 1;
  }
  usedNames.add(candidate.toLowerCase());
  return candidate;
}

// ── Public result types ───────────────────────────────────────────────────────
export interface PageExportResult {
  sheetName:        string;
  pageDisplayName:  string;
  pageInternalName: string;
  visualsWithData:  number;
  totalRows:        number;
  /** The visuals written to this sheet, in sheet order (block k of the sheet
   *  = visuals[k]). Used by the mapping-file comparison to find a visual by
   *  its Power BI visual ID; ignored by the normal flow. */
  visuals?:         VisualIndexEntry[];
  /** Visuals on this page that exported NOTHING, with the reason why. */
  notExported?:     NotExportedEntry[];
}

/** A visual that produced no data in the export, and why. */
export interface NotExportedEntry {
  visualName: string;
  title:      string;
  type:       string;
  reason:     string;
}

/** One data visual as written to a sheet (see PageExportResult.visuals). */
export interface VisualIndexEntry {
  /** Power BI visual name (the "visual ID" in a KPI mapping file). */
  visualName: string;
  /** Title exactly as written in the sheet's [VISUAL] marker row. */
  title:      string;
  type:       string;
  rows:       number;
}

/** Sidecar file next to an exported workbook: which visual ID is which block. */
export interface VisualIndexFile {
  version: 1;
  sheets:  Array<{
    sheetName: string; pageDisplayName: string; pageInternalName: string;
    visuals: VisualIndexEntry[];
    /** Visuals that exported nothing, with the reason (added later; optional). */
    notExported?: NotExportedEntry[];
  }>;
}

/** Path of the sidecar index for an exported workbook (expected.xlsx -> expected.visual-index.json). */
export function visualIndexPathFor(xlsxPath: string): string {
  return xlsxPath.replace(/\.xlsx$/i, '') + '.visual-index.json';
}

/** Writes the sidecar index. Best-effort by design: the normal flow never reads it. */
export function writeVisualIndex(result: ReportExportResult): void {
  try {
    const body: VisualIndexFile = {
      version: 1,
      sheets: result.exported.map(e => ({
        sheetName: e.sheetName, pageDisplayName: e.pageDisplayName,
        pageInternalName: e.pageInternalName, visuals: e.visuals ?? [],
        notExported: e.notExported ?? [],
      })),
    };
    fs.writeFileSync(visualIndexPathFor(result.outPath), JSON.stringify(body, null, 2), 'utf8');
  } catch (e) {
    console.warn(`[export] Could not write visual index (non-fatal, only needed for mapping-file mode): ${describeError(e)}`);
  }
}

export interface ReportExportResult {
  pagesFound:  ReportPage[];
  exported:    PageExportResult[];
  outPath:     string;
  /** pagesFilter entries that matched NO page in this report — a strong sign
   *  the two sides' page names have drifted (renamed/removed pages). */
  requestedPagesMissing: string[];
}

/**
 * Canonical form for matching a page/filter display name across two reports.
 * Migrated reports and cross-report filter titles drift in exactly these
 * ways while still meaning the same thing: trailing/leading whitespace
 * ("Vertragsversand " vs "Vertragsversand"), case ("- Cluster" vs
 * "- cluster"), separator style ("Recruiting_Weg" vs "Recruiting Weg"
 * vs "RecruitingWeg" — separators are stripped entirely, not just
 * normalized to one style, so "a_bc" and "abc" compare equal), and accents
 * ("Commodité" vs "Commodite" — observed in production: the exact same page
 * lost its accent during migration and was treated as two different pages).
 * Shared by page-name matching (this file) and cross-report filter-title
 * matching (cross-report-match.helpers.ts) — one rule, used everywhere names
 * from two different reports need to be compared.
 */
// Built from explicit char codes (never a literal combining-mark character in
// this file) - Unicode combining diacritical marks, U+0300 to U+036F. After
// name.normalize('NFD') splits an accented letter into base + mark (e-acute
// -> "e" + U+0301), stripping this range removes the mark and leaves the
// plain base letter.
const COMBINING_DIACRITICS_RE = new RegExp(
  '[' + String.fromCharCode(0x0300) + '-' + String.fromCharCode(0x036f) + ']', 'g',
);

export function pageNameKey(name: string): string {
  return name
    .normalize('NFD').replace(COMBINING_DIACRITICS_RE, '') // strip accents: e-acute -> e, etc.
    .trim().toLowerCase().replace(/[\s_-]+/g, '');
}

/**
 * Pure page-set alignment between a source and target report, for the parity
 * summary. `pageMap` (source name → target name) is applied before matching;
 * names are matched trimmed + case-insensitively.
 */
export function diffPageSets(
  sourcePages: string[],
  targetPages: string[],
  pageMap?: Record<string, string>,
): { inBoth: string[]; onlyInSource: string[]; onlyInTarget: string[] } {
  const mapByKey = new Map(Object.entries(pageMap ?? {}).map(([s, t]) => [pageNameKey(s), t]));
  const targetKeys = new Set(targetPages.map(pageNameKey));

  const inBoth: string[] = [];
  const onlyInSource: string[] = [];
  const matchedTargetKeys = new Set<string>();

  for (const src of sourcePages) {
    const mapped = mapByKey.get(pageNameKey(src)) ?? src;
    const key = pageNameKey(mapped);
    if (targetKeys.has(key)) {
      inBoth.push(src);
      matchedTargetKeys.add(key);
    } else {
      onlyInSource.push(src);
    }
  }

  const onlyInTarget = targetPages.filter(t => !matchedTargetKeys.has(pageNameKey(t)));
  return { inBoth, onlyInSource, onlyInTarget };
}

export interface ExportOptions {
  /** Where to write the .xlsx. */
  outPath: string;
  /** If set, only export pages whose display name is in this list. */
  pagesFilter?: string[];
  /**
   * Map a page's actual display name → the sheet name to write it under.
   * Used to align renamed tabs between source and target. Default: identity.
   */
  sheetNameFor?: (pageDisplayName: string) => string;
  /** Label for console logging, e.g. "source (Import mode)". */
  label?: string;
  /**
   * Optional hook called ONCE, before any page is visited — for filters that
   * apply at the report level (affecting every page in one call), as opposed
   * to applySlicersForPage which runs once per page. If it throws, export
   * continues (a warning is logged) — the per-page hook is the safety net for
   * anything the global hook didn't (or couldn't) cover.
   */
  beforeExport?: (page: Page) => Promise<void>;
  /**
   * Optional hook called AFTER navigating to each page but BEFORE exporting its
   * visuals. Use it to apply slicer selections (filters) for that page. Receives
   * the Playwright page and the page's display name; should apply and settle any
   * filters, then resolve. If it throws, the page is still exported (unfiltered)
   * and a warning is logged, so one bad slicer can't abort the whole run.
   *
   * Kept as a generic callback so this module has no dependency on the slicer
   * scenario config — the spec wires the two together.
   */
  applySlicersForPage?: (page: Page, pageDisplayName: string) => Promise<void>;
}

// ── Styling (light — values are what matter, but readable output helps) ───────
const HDR_FILL: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F3864' } };
const SEC_FILL: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
const BD: ExcelJS.Border = { style: 'thin', color: { argb: 'FFBFBFBF' } };
const CB: Partial<ExcelJS.Borders> = { top: BD, bottom: BD, left: BD, right: BD };

/**
 * Embeds the active report, exports all requested pages' visuals, and writes a
 * comparison-friendly workbook to options.outPath.
 *
 * The report identity MUST already be applied (applyReportIdentity) before this
 * call, and the user token acquired.
 */
export async function exportReportToWorkbook(
  page: Page,
  token: string,
  options: ExportOptions,
): Promise<ReportExportResult> {
  const label = options.label ?? 'report';
  const sheetNameFor = options.sheetNameFor ?? ((n: string) => n);

  // ── Embed ───────────────────────────────────────────────────────────────
  console.log(`\n[export:${label}] Embedding report...`);
  await timed(label, 'embed', () => loadHarnessPage(page, token));

  // ── Discover pages ──────────────────────────────────────────────────────
  const pagesFound = await timed(label, 'getPages', () => getReportPages(page));
  console.log(`[export:${label}] Report has ${pagesFound.length} page(s): ${pagesFound.map(p => p.displayName).join(' | ')}`);

  // Select pages: filtered subset (by display name, trimmed + case-insensitive
  // — migrated reports drift by trailing spaces and case) or all.
  let selected: ReportPage[];
  let requestedPagesMissing: string[] = [];
  if (options.pagesFilter && options.pagesFilter.length > 0) {
    const wanted = options.pagesFilter.map(pageNameKey);
    selected = pagesFound.filter(p => wanted.includes(pageNameKey(p.displayName)));
    requestedPagesMissing = options.pagesFilter.filter(
      w => !pagesFound.some(p => pageNameKey(p.displayName) === pageNameKey(w)),
    );
    if (requestedPagesMissing.length > 0) {
      console.warn(`[export:${label}] Requested pages not found in report: ${requestedPagesMissing.join(', ')}`);
    }
  } else {
    selected = pagesFound;
  }

  if (selected.length === 0) {
    console.warn(`[export:${label}] No pages selected to export.`);
  }

  // ── Build workbook ──────────────────────────────────────────────────────
  // Streaming writer: each row is flushed to disk (and freed from memory) as
  // soon as it's committed, so peak memory stays roughly proportional to ONE
  // page's data instead of the whole report's data — matters once you're
  // exporting many pages and/or large tables.
  const wb = new ExcelJS.stream.xlsx.WorkbookWriter({
    filename: options.outPath,
    useStyles: true,          // keep the fill/font/border formatting
    useSharedStrings: false,  // shared strings trade memory for file size — skip, we're optimising for memory
  });
  wb.creator = 'VériRapport';
  wb.created = new Date();

  const usedSheetNames = new Set<string>();
  const exported: PageExportResult[] = [];

  if (options.beforeExport) {
    try {
      await timed(label, 'beforeExport', () => options.beforeExport!(page));
    } catch (e) {
      console.warn(`[export:${label}]   beforeExport hook failed — continuing without it: ${describeError(e)}`);
    }
  }

  for (const rp of selected) {
    console.log(`[export:${label}]   → page "${rp.displayName}"`);

    let navigated = true;
    try {
      await timed(label, `page:${rp.displayName}:navigate`, () => setReportPage(page, rp.name));
    } catch (e) {
      console.warn(`[export:${label}]     could not navigate to page "${rp.displayName}" — retrying once: ${describeError(e)}`);
      await page.waitForTimeout(NAV_RETRY_WAIT_MS);
      try {
        await setReportPage(page, rp.name);
      } catch (e2) {
        navigated = false;
        console.warn(`[export:${label}]     navigation to "${rp.displayName}" failed again — skipping this page (exporting it now would have captured the WRONG page's data): ${describeError(e2)}`);
      }
    }

    if (!navigated) {
      // Skip slicer application + export entirely — the report is still on
      // whatever page it was on before. Write an explicit marker sheet rather
      // than omitting it, so this reads as "we couldn't read this page" in the
      // comparison output instead of an ambiguous "page missing" diff.
      const desiredSheetName = sheetNameFor(rp.displayName);
      const wsName = toUniqueSheetName(desiredSheetName, usedSheetNames);
      const ws = wb.addWorksheet(wsName);
      const row = ws.addRow([`[PAGE NAVIGATION FAILED — could not reach "${rp.displayName}"; page skipped]`]);
      row.commit();
      ws.commit();
      exported.push({
        sheetName: wsName, pageDisplayName: rp.displayName, pageInternalName: rp.name,
        visualsWithData: 0, totalRows: 0,
      });
      continue;
    }

    // Apply slicer selections for this page (if a hook was provided), after
    // navigation and before export. Isolated so a slicer failure downgrades to
    // an unfiltered export of this page rather than failing the whole run.
    if (options.applySlicersForPage) {
      try {
        await timed(label, `page:${rp.displayName}:filters`, () => options.applySlicersForPage!(page, rp.displayName));
      } catch (e) {
        console.warn(`[export:${label}]     slicer application failed on "${rp.displayName}" — exporting UNFILTERED: ${describeError(e)}`);
      }
    }

    let visuals: VisualExport[] = [];
    try {
      visuals = await timed(label, `page:${rp.displayName}:export`, () => exportCurrentPageVisuals(page));
    } catch (e) {
      console.warn(`[export:${label}]     export failed on "${rp.displayName}" — retrying once: ${describeError(e)}`);
      await page.waitForTimeout(EXPORT_RETRY_WAIT_MS);
      try {
        visuals = await exportCurrentPageVisuals(page);
      } catch (e2) {
        console.warn(`[export:${label}]     export failed again on "${rp.displayName}" — page will be written with no data: ${describeError(e2)}`);
      }
    }

    const desiredSheetName = sheetNameFor(rp.displayName);
    const wsName = toUniqueSheetName(desiredSheetName, usedSheetNames);
    const ws = wb.addWorksheet(wsName);

    {
      const exportedVisuals = visuals.filter(v => v.exportMs !== undefined);
      const sumMs = exportedVisuals.reduce((s, v) => s + (v.exportMs ?? 0), 0);
      const slowest = [...exportedVisuals].sort((a, b) => (b.exportMs ?? 0) - (a.exportMs ?? 0)).slice(0, 3)
        .map(v => `${v.visualTitle || v.visualName} ${((v.exportMs ?? 0) / 1000).toFixed(1)}s`).join(', ');
      console.log(`[export:${label}]     ${exportedVisuals.length} exportData call(s), ${(sumMs / 1000).toFixed(1)}s total` + (slowest ? ` — slowest: ${slowest}` : ''));
      recordTiming(label, `page:${rp.displayName}:exportData`, sumMs);
    }
    const tWrite = Date.now();
    const { visualsWithData, totalRows, index, notExported } = writeVisualsToSheet(ws, visuals);
    ws.commit(); // flush this page's rows to disk now — don't hold them for the rest of the run
    recordTiming(label, `page:${rp.displayName}:writeSheet`, Date.now() - tWrite);

    exported.push({
      sheetName: wsName, pageDisplayName: rp.displayName, pageInternalName: rp.name,
      visualsWithData, totalRows, visuals: index, notExported,
    });
  }

  await timed(label, 'finalizeXlsx', () => wb.commit()); // finalises the xlsx file on disk
  const grandRows = exported.reduce((s, e) => s + e.totalRows, 0);
  console.log(`[export:${label}] Wrote ${exported.length} sheet(s), ${grandRows} data row(s) → ${options.outPath}`);

  return { pagesFound, exported, outPath: options.outPath, requestedPagesMissing };
}

// ── Column width bounds (characters) ───────────────────────────────────────
const MIN_COL_WIDTH = 12;
const MAX_COL_WIDTH = 42;
const WIDTH_PADDING = 2;

/**
 * Writes one page's visuals into an already-created worksheet: a marker row,
 * header row, and data rows per visual (skipping empty/non-data visuals),
 * sized so nothing is truncated.
 *
 * Exported (not just internal) so it can be exercised directly in tests
 * without a live Playwright page — see the parity smoke tests.
 *
 * Works with BOTH the in-memory Workbook and the streaming WorkbookWriter.
 * For the streaming writer specifically, two constraints from ExcelJS's
 * implementation shape this function's structure:
 *   1. Column widths (`ws.columns = [...]`) MUST be set before the FIRST row
 *      is committed — the streaming writer serialises `<cols>` the instant
 *      the first row is flushed, and never revisits it. So this does a
 *      lightweight first pass over the already-in-memory `visuals` data
 *      (just measuring string lengths — no ExcelJS calls) to compute widths
 *      BEFORE writing any row.
 *   2. Each row must be explicitly `.commit()`-ed as it's written, or it just
 *      sits in an internal buffer and is never actually freed — calling
 *      addRow() alone does NOT stream a row to disk.
 *
 * IMPORTANT: only styling (fill/font/border/width/wrap) is applied here.
 * Cell VALUES are written exactly as before — this has zero effect on
 * compareWorkbooks(), which reads only cell.value and ignores all styling.
 */
/** Plain-words reason a visual wrote no data block, or null if it is written. */
export function whyNotExported(v: VisualExport): string | null {
  if (SKIP_TYPES.has(v.visualType)) {
    return `Not a data visual (${v.visualType}) — nothing to export, by design`;
  }
  if (v.errorMessage && !/^skipped/i.test(v.errorMessage)) {
    return `Export failed: ${v.errorMessage}`;
  }
  if (v.errorMessage) return v.errorMessage;
  if (!(v.rowCount > 0)) {
    return 'Returned 0 rows — no data for the filters used, or a hidden/pop-up visual';
  }
  if (!v.headers || v.headers.length === 0) return 'Returned no columns';
  return null;
}

export function writeVisualsToSheet(
  ws: ExcelJS.Worksheet,
  visuals: VisualExport[],
): { visualsWithData: number; totalRows: number; index: VisualIndexEntry[]; notExported: NotExportedEntry[] } {
  // Record why each visual that is NOT written exported nothing.
  const notExported: NotExportedEntry[] = visuals
    .map(v => ({ v, reason: whyNotExported(v) }))
    .filter(x => x.reason !== null)
    .map(x => ({
      visualName: x.v.visualName, title: x.v.visualTitle || x.v.visualName,
      type: x.v.visualType, reason: x.reason as string,
    }));

  // Keep only data-bearing visuals, sorted for stable, aligned output.
  const dataVisuals = visuals
    .filter(v => !SKIP_TYPES.has(v.visualType) && v.rowCount > 0 && v.headers.length > 0)
    .sort((a, b) =>
      (a.visualTitle || a.visualName).localeCompare(b.visualTitle || b.visualName) ||
      a.visualName.localeCompare(b.visualName),
    );

  if (dataVisuals.length === 0) {
    ws.getColumn(1).width = 24;
    const row = ws.addRow(['[NO DATA VISUALS]']);
    row.commit();
    return { visualsWithData: 0, totalRows: 0, index: [], notExported };
  }

  // ── PASS 1: compute per-column widths from data already in memory.
  // No ExcelJS writes yet — must finish this before touching a single row.
  const colWidths: number[] = [];
  const trackWidths = (vals: ReadonlyArray<string>) => {
    vals.forEach((v, i) => {
      const len = (v ?? '').toString().length;
      if (!colWidths[i] || len > colWidths[i]) colWidths[i] = len;
    });
  };
  for (const v of dataVisuals) {
    trackWidths(v.headers);
    for (const row of v.rows) trackWidths(v.headers.map(h => row[h] ?? ''));
  }
  ws.columns = colWidths.map(w => ({
    width: Math.min(Math.max(w + WIDTH_PADDING, MIN_COL_WIDTH), MAX_COL_WIDTH),
  }));

  // ── PASS 2: write + immediately commit every row, so the streaming writer
  // flushes each one to disk and frees it from memory right away instead of
  // holding the whole page (or worse, the whole report) in memory at once.
  let totalRows = 0;
  for (const v of dataVisuals) {
    const nc = v.headers.length;

    // Stable marker row — NO volatile row count, so identical data hashes equal.
    const secRow = ws.addRow([`[VISUAL] ${v.visualTitle || v.visualName} [${v.visualType}]`]);
    if (nc > 1) ws.mergeCells(secRow.number, 1, secRow.number, nc);
    const secCell = secRow.getCell(1);
    secCell.fill = SEC_FILL;
    secCell.font = { bold: true };
    secCell.border = CB;
    secCell.alignment = { vertical: 'middle', wrapText: true };
    secRow.height = 20;
    secRow.commit();

    const hdrRow = ws.addRow(v.headers);
    hdrRow.eachCell(c => {
      c.fill = HDR_FILL;
      c.font = { color: { argb: 'FFFFFFFF' }, bold: true };
      c.border = CB;
      c.alignment = { vertical: 'middle', wrapText: true };
    });
    hdrRow.height = 24;
    hdrRow.commit();

    for (const row of v.rows) {
      const vals = v.headers.map(h => row[h] ?? '');
      const dataRow = ws.addRow(vals);
      dataRow.eachCell(c => {
        c.border = CB;
        c.alignment = { vertical: 'middle' };
      });
      dataRow.commit();
    }
    totalRows += v.rowCount;

    const spacer = ws.addRow([]);
    spacer.commit();
  }

  return {
    visualsWithData: dataVisuals.length, totalRows, notExported,
    index: dataVisuals.map(v => ({
      visualName: v.visualName, title: v.visualTitle || v.visualName, type: v.visualType, rows: v.rowCount,
    })),
  };
}

// ── Parity summary writer ─────────────────────────────────────────────────────

/**
 * One row of the "Filters Applied" sheet: a filter or slicer that was found,
 * cleared, applied or read back on one report.
 */
export interface FilterRecord {
  /** 'Source (old report)' or 'Target (new report)'. */
  report: string;
  /** Page name, or 'All pages' for report-level filters. */
  page:   string;
  level:  'Report' | 'Page' | 'Slicer';
  /** Field as "Table.Column" (and the slicer's title when it has one). */
  field:  string;
  values: string[];
  /** 'Cleared (was saved in the report)' | 'Kept as saved (slicer needs a value)' | 'Applied by you' | 'Active after applying (read-back)' */
  what:   string;
}

export interface ParitySummaryMeta {
  pairName:     string;
  mode:         string;
  sourceLabel:  string;   // e.g. "Import mode"
  targetLabel:  string;   // e.g. "Direct Lake"
  sourceReportId: string;
  targetReportId: string;
  expectedFile: string;
  actualFile:   string;
  /**
   * Problems that make the comparison itself untrustworthy (a filter that
   * applied on one side but failed on the other, requested pages missing on
   * one side, etc). Any entry here FORCES the verdict to a non-pass — a
   * filtered-vs-unfiltered comparison must never present itself as a clean
   * result, pass OR fail.
   */
  comparisonCaveats?: string[];
  /** Plain information lines shown on the summary sheet (never change the verdict). */
  notes?: string[];
  /** Rows for the "Filters Applied" sheet. Omit/empty → the sheet says nothing was recorded. */
  filterRecords?: FilterRecord[];
  /** Mapping-file mode: path/name of the KPI mapping file that drove the comparison. */
  mappingFile?: string;
  /** Full page-set alignment between the two reports (informational). */
  pageAlignment?: { onlyInSource: string[]; onlyInTarget: string[]; inBothCount: number };
}

/**
 * Whether a visual whose header TEXT differs but whose underlying VALUES match
 * (e.g. Import-mode friendly column name "Count of Application ID" vs
 * Direct-Lake raw name "Count of APPLICATION_ID") should fail the run.
 *
 * Default false: the migration is judged on whether the numbers match, not on
 * cosmetic label text. Header differences are still fully visible — every one
 * is listed on the "Visual Comparison" sheet and rolled into the summary
 * counts — they just don't flip the PASS/FAIL verdict on their own. Flip this
 * to true if you want label drift to be a hard failure instead.
 */
const TREAT_HEADER_DIFF_AS_FAILURE = false;

const E_PASS: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00B050' } };
const E_FAIL: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF0000' } };
const E_HDR:  ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F3864' } };
const E_INFO: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE9EFF7' } };
const E_WARN: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC000' } };
const E_MISS: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
const E_BD: ExcelJS.Border = { style: 'thin', color: { argb: 'FFCCCCCC' } };
const E_CB: Partial<ExcelJS.Borders> = { top: E_BD, bottom: E_BD, left: E_BD, right: E_BD };

const VISUAL_STATUS_LABEL: Record<VisualBlockDiff['status'], string> = {
  'identical':         '✅ Identical',
  'header-diff':        '📝 Header text differs (values match)',
  'duplicate-count':    '🔁 Duplicated (values match)',
  'data-diff':          '❌ Values differ',
  'only-in-expected':   '⚠️ Only in Source',
  'only-in-actual':     '🆕 Only in Target',
  'ambiguous':          '❓ Ambiguous — needs manual check',
  'expected-removal':   '🗑️ Removed on purpose (per mapping file)',
  'not-verified':       '🔍 Not verified — please check by eye',
};

const VISUAL_STATUS_FILL: Record<VisualBlockDiff['status'], ExcelJS.Fill> = {
  'identical':        E_PASS,
  'header-diff':       E_INFO,
  'duplicate-count':   E_INFO,
  'data-diff':         E_FAIL,
  'only-in-expected':  E_MISS,
  'only-in-actual':    E_WARN,
  'ambiguous':         E_WARN,
  'expected-removal':  E_INFO,
  'not-verified':      E_WARN,
};

/** True when a visual's status means "the data matches" (part of the relaxed gate). */
function visualDataMatches(status: VisualBlockDiff['status']): boolean {
  return status === 'identical' || status === 'header-diff' || status === 'duplicate-count' || status === 'expected-removal' || status === 'not-verified';
}

// ── Severity tiers ───────────────────────────────────────────────────────────
//
// A value difference and a header/label rename are not the same kind of
// problem — a renamed column is expected, cosmetic noise from a migration; a
// visual with different numbers (or one that vanished entirely) is a real
// finding. VISUAL_STATUS_LABEL/FILL above already say WHICH of 7 fine-grained
// statuses a visual has, but reading 7 different labels to work out "is this
// one bad?" doesn't scale. Severity collapses them to 3 buckets so the sheet
// can be scanned (or filtered) at a glance, and gives the row itself a light
// background tint (full-row, not just the Status cell) so a critical row is
// visually unmistakable even scrolling past it quickly.
export type Severity = 'safe' | 'review' | 'critical';

const VISUAL_SEVERITY: Record<VisualBlockDiff['status'], Severity> = {
  'identical':        'safe',
  'header-diff':      'safe',     // values match — only the label text changed
  'duplicate-count':  'safe',     // values match — only the copy count changed
  'not-verified':     'review',   // mapping file only: could not be checked, not a failure
  'expected-removal': 'safe',     // mapping file says it was dropped on purpose — just listed
  'ambiguous':        'review',   // couldn't be auto-verified either way
  'only-in-expected': 'critical', // a visual genuinely vanished
  'only-in-actual':   'critical', // a visual genuinely appeared
  'data-diff':        'critical', // same visual, different numbers
};

const SEVERITY_LABEL: Record<Severity, string> = {
  safe:     '✅ Safe — cosmetic only',
  review:   '❓ Needs manual review',
  critical: '❌ Critical — real difference',
};

// Pale versions of the status fills, meant for tinting an entire row rather
// than one bold cell — strong enough to scan down a column of rows, light
// enough that black text stays readable on top of it.
const E_PASS_LT:     ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2F3E1' } };
const E_INFO_LT:     ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEAF1FB' } };
const E_WARN_LT:     ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF1D6' } };
const E_FAIL_LT:     ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCE1E1' } };

const SEVERITY_ROW_FILL: Record<Severity, ExcelJS.Fill> = {
  safe:     E_PASS_LT,
  review:   E_WARN_LT,
  critical: E_FAIL_LT,
};
const SEVERITY_CELL_FILL: Record<Severity, ExcelJS.Fill> = {
  safe:     E_PASS,
  review:   E_WARN,
  critical: E_FAIL,
};

// header-diff/duplicate-count share the 'safe' severity but should still tint
// slightly differently from a truly byte-identical visual, so a reviewer can
// still tell "renamed" from "untouched" without reading the Status column —
// this only affects the ROW tint, never the severity bucket itself.
function rowFillFor(status: VisualBlockDiff['status']): ExcelJS.Fill {
  if (status === 'header-diff' || status === 'duplicate-count' || status === 'expected-removal') return E_INFO_LT;
  return SEVERITY_ROW_FILL[VISUAL_SEVERITY[status]];
}

// ── Plain-English analysis ────────────────────────────────────────────────────
//
// The rest of this workbook answers "what differs, exactly, and where" — this
// answers the question a reviewer actually asks first: "do I need to worry
// about this?" It's generated purely from counts already computed elsewhere
// in this function (no new comparison logic), so it can never disagree with
// the detailed sheets — it's a summary OF them, not a second opinion.

export interface AnalysisStats {
  /** Visuals actually inspected in detail (Tier 2) — excludes pages that
   *  matched byte-for-byte at the fast Tier 1 check, which never get a
   *  per-visual breakdown because there's nothing to break down. */
  totalVisuals: number;
  identicalVisuals: number;
  /** header-diff + duplicate-count: values match, only cosmetic. */
  cosmeticOnly: number;
  /** ambiguous: could not be auto-verified either way. */
  needsReview: number;
  /** only-in-expected + only-in-actual + data-diff: a genuine difference. */
  critical: number;
  /** Pages that matched byte-for-byte at the fast check (no visual detail). */
  fastPassPages: number;
  /** Mapping-file mode: visuals the mapping file says were dropped on purpose
   *  (listed, never counted as a difference). Absent/0 in the normal flow. */
  expectedRemovals?: number;
  /** Mapping-file mode: visuals that could not be located in the new report to check. */
  notVerified?: number;
}

export function computeAnalysisStats(diff: WorkbookDiffSummary): AnalysisStats {
  const stats: AnalysisStats = {
    totalVisuals: 0, identicalVisuals: 0, cosmeticOnly: 0, needsReview: 0, critical: 0, fastPassPages: 0,
  };
  for (const s of diff.sheets) {
    if (!s.inExpected || !s.inActual) continue; // whole-page-missing handled separately, via sheetsOnlyIn*
    if (s.visuals.length === 0) {
      if (s.identical) stats.fastPassPages++;
      continue;
    }
    for (const v of s.visuals) {
      stats.totalVisuals++;
      switch (VISUAL_SEVERITY[v.status]) {
        case 'safe':
          if (v.status === 'identical') stats.identicalVisuals++;
          else if (v.status === 'expected-removal') stats.expectedRemovals = (stats.expectedRemovals ?? 0) + 1;
          else stats.cosmeticOnly++;
          break;
        case 'review':
          stats.needsReview++;
          if (v.status === 'not-verified') stats.notVerified = (stats.notVerified ?? 0) + 1;
          break;
        case 'critical': stats.critical++; break;
      }
    }
  }
  return stats;
}

/** Returns 2-4 short sentences, each its own row in the summary. */
export function buildAnalysisNarrative(
  diff: WorkbookDiffSummary,
  stats: AnalysisStats,
  passed: boolean,
  caveats: string[],
): string[] {
  const lines: string[] = [];
  const pagesMissing = diff.sheetsOnlyInExpected.length + diff.sheetsOnlyInActual.length;

  if (caveats.length > 0) {
    lines.push(
      `⚠ This run is NOT directly comparable — ${caveats.length} caveat(s) affected it (see the Caveat rows below). ` +
      `Resolve those before trusting the pass/fail verdict.`,
    );
  }

  const diffVisuals = stats.totalVisuals - stats.identicalVisuals - (stats.expectedRemovals ?? 0);

  if (stats.totalVisuals === 0) {
    lines.push(stats.fastPassPages > 0
      ? `All ${stats.fastPassPages} compared page(s) matched byte-for-byte — no differences of any kind.`
      : 'No pages were compared in detail (nothing to analyze).');
  } else if (diffVisuals === 0) {
    lines.push(`All ${stats.totalVisuals} visual(s) inspected are identical — no differences of any kind.`);
  } else {
    const pctCosmetic = Math.round((stats.cosmeticOnly / diffVisuals) * 100);
    if (stats.critical === 0 && stats.needsReview === 0) {
      lines.push(
        `Good news: every difference found (${stats.cosmeticOnly} of ${diffVisuals}) is cosmetic only — ` +
        `column/visual labels were renamed during migration, but every underlying VALUE still matches. Nothing here needs action.`,
      );
    } else {
      const parts: string[] = [];
      if (stats.cosmeticOnly > 0) parts.push(`${stats.cosmeticOnly} cosmetic label rename(s) (${pctCosmetic}% of all differences — values match, safe to ignore)`);
      if (stats.critical > 0) parts.push(`${stats.critical} critical difference(s) (missing/added visuals or values that genuinely don't match)`);
      if (stats.needsReview > 0) parts.push(`${stats.needsReview} visual(s) needing a manual look (duplicate titles the tool couldn't auto-match)`);
      lines.push(`Of ${diffVisuals} visual(s) that differ: ${parts.join('; ')}.`);
      lines.push(stats.critical > 0
        ? `Bottom line: ${stats.critical} genuine issue(s) need review before sign-off${stats.cosmeticOnly > 0 ? ` — the other ${stats.cosmeticOnly} are just renamed labels and can be ignored` : ''}.`
        : `Bottom line: nothing critical, but ${stats.needsReview} visual(s) need a quick manual check (see "❓ Needs manual review" rows on Visual Comparison).`);
    }
  }

  if ((stats.notVerified ?? 0) > 0) {
    lines.push(`${stats.notVerified} visual(s) could not be verified: the mapping file says they exist in the new report, but their values could not be located there to compare (typically old cards merged into a multi-value card). They are NOT counted as failures. Check them by eye (status "Not verified" on Visual Comparison).`);
  }
  if ((stats.expectedRemovals ?? 0) > 0) {
    lines.push(`${stats.expectedRemovals} visual(s) were removed on purpose according to the mapping file. They are listed on Visual Comparison (status "Removed on purpose") and do not count as differences.`);
  }

  if (pagesMissing > 0) {
    lines.push(`${pagesMissing} entire page(s) exist on only one side of the comparison (see "Pages Only in Source/Target" below) — not counted in the visual stats above.`);
  }

  return lines;
}

// ── Lightweight UI summary (for the desktop app, NOT the xlsx) ───────────────
//
// A trimmed, JSON-serializable view of a comparison result, meant for the
// Electron app to render an in-page summary right after a run finishes,
// instead of the user having to open parity-summary.xlsx just to see
// pass/fail. Deliberately excludes VisualBlockDiff's sampleOnlyInExpected/
// sampleOnlyInActual arrays (up to SAMPLE_SIZE=800 rows each, per visual) —
// those stay exclusive to the xlsx "Differences" sheet. This is a summary
// OF the xlsx, not a replacement for it; fields are picked explicitly
// below, never spread, so a future VisualBlockDiff field never leaks a
// large payload into this JSON by accident.

export interface ParityUiVisual {
  title: string;
  type: string;
  severity: Severity;
  statusLabel: string;
  note?: string;
}

export interface ParityUiPage {
  name: string;
  status: 'identical' | 'header-only' | 'different' | 'only-in-source' | 'only-in-target';
  rowsExpected: number;
  rowsActual: number;
  headerDiffCount: number;
  structuralDiffCount: number;
  visuals: ParityUiVisual[];
}

export interface ParityUiSummary {
  /** Same three-way distinction the xlsx verdict cell shows - a plain
   *  boolean can't tell "genuine fail" from "not a valid comparison". */
  verdict: 'pass' | 'fail' | 'not_comparable';
  passed: boolean;
  differingSheets: number;
  caveats: string[];
  narrative: string[];
  stats: AnalysisStats;
  pagesOnlyInSource: string[];
  pagesOnlyInTarget: string[];
  pages: ParityUiPage[];
}

/** Per-page status - mirrors the exact if/else chain used for the "Page
 *  Comparison" xlsx sheet's Status column, so the two never drift apart. */
function pageUiStatus(s: SheetDiffSummary): ParityUiPage['status'] {
  if (s.inExpected && !s.inActual) return 'only-in-source';
  if (!s.inExpected && s.inActual) return 'only-in-target';
  if (s.identical) return 'identical';
  if (s.dataIdentical) return 'header-only';
  return 'different';
}

/**
 * Builds the lightweight UI summary described above. `passed`/`differingSheets`
 * should be exactly what writeParitySummary() just returned for this same
 * `diff`, so the two representations of the same run can never disagree.
 */
export function summarizeForUi(
  diff: WorkbookDiffSummary,
  caveats: string[],
  passed: boolean,
  differingSheets: number,
): ParityUiSummary {
  const stats = computeAnalysisStats(diff);
  const narrative = buildAnalysisNarrative(diff, stats, passed, caveats);

  const pages: ParityUiPage[] = diff.sheets
    .filter(s => s.inExpected && s.inActual) // sheetsOnlyInExpected/Actual already cover the other two cases below
    .map(s => ({
      name: s.sheet,
      status: pageUiStatus(s),
      rowsExpected: s.rowsExpected,
      rowsActual: s.rowsActual,
      headerDiffCount: s.headerDiffCount,
      structuralDiffCount: s.structuralDiffCount,
      visuals: s.visuals.map(v => ({
        title: v.title,
        type: v.type,
        severity: VISUAL_SEVERITY[v.status],
        statusLabel: VISUAL_STATUS_LABEL[v.status],
        note: v.note,
      })),
    }));

  return {
    verdict: caveats.length > 0 ? 'not_comparable' : (passed ? 'pass' : 'fail'),
    passed,
    differingSheets,
    caveats,
    narrative,
    stats,
    pagesOnlyInSource: diff.sheetsOnlyInExpected,
    pagesOnlyInTarget: diff.sheetsOnlyInActual,
    pages,
  };
}

/**
 * Writes the migration-parity summary workbook from a comparison result.
 * Returns { passed, differingSheets } for the caller's console verdict.
 *
 * PASS/FAIL is gated on `dataIdentical` by default (position-independent and
 * tolerant of header-text-only renames — see TREAT_HEADER_DIFF_AS_FAILURE),
 * not on the strict `identical` flag. Every difference, cosmetic or not, is
 * still written out — the relaxation only affects the verdict, never what's
 * shown.
 */
export function writeFiltersSheet(wb: ExcelJS.Workbook, meta: ParitySummaryMeta): void {
  const ws = wb.addWorksheet('Filters Applied');
  ws.columns = [{ width: 22 }, { width: 34 }, { width: 10 }, { width: 46 }, { width: 50 }, { width: 36 }];
  const recs = meta.filterRecords ?? [];
  const put = (cells: string[], fill?: ExcelJS.Fill, bold = false) => {
    const r = ws.addRow(cells);
    r.eachCell({ includeEmpty: true }, c => {
      c.alignment = { vertical: 'top', wrapText: true }; c.border = E_CB;
      if (fill) c.fill = fill;
      if (bold) c.font = { bold: true };
    });
    return r;
  };
  const line = (text: string, fill?: ExcelJS.Fill, bold = false) => {
    const r = ws.addRow([text]);
    ws.mergeCells(r.number, 1, r.number, 6);
    r.getCell(1).alignment = { vertical: 'top', wrapText: true };
    if (fill) r.getCell(1).fill = fill;
    if (bold) r.getCell(1).font = { bold: true };
  };
  line('Filters and slicers on each report when this comparison was made', E_INFO, true);
  line('Cleared = a filter saved in the report that the tool removed before exporting. Applied by you = a filter you picked in the app. ' +
       'Active after applying = what the report itself says is active after that (read back from Power BI). Where a field has a different name in the ' +
       'new report, each report shows its own field name on its own rows.');
  line(`Run mode: ${meta.mode}.  Source = ${meta.sourceLabel}, Target = ${meta.targetLabel}.`);
  ws.addRow([]);
  if (recs.length === 0) {
    line('Nothing was recorded for this run (for example the report was exported as saved, or the filter read-back was switched off).', E_WARN_LT);
  } else {
    const hdr = put(['Report', 'Page', 'Level', 'Field', 'Value(s)', 'What happened'], E_HDR, true);
    hdr.eachCell(c => { c.font = { bold: true, color: { argb: 'FFFFFFFF' } }; });
    for (const rec of recs) {
      const fill = rec.what.startsWith('Cleared') ? E_WARN_LT : rec.what.startsWith('Applied') || rec.what.startsWith('Kept') ? E_INFO_LT : E_PASS_LT;
      put([rec.report, rec.page, rec.level, rec.field, rec.values.length ? rec.values.join(', ') : '', rec.what], fill);
    }
  }
  ws.addRow([]);
  line('Kept (cannot be cleared by this tool)', E_INFO, true);
  line('• Visual-level filters: part of a visual\'s own definition, so they apply the same way in both reports.');
  line('• Row-level security (RLS): always applies to the signed-in user, so both reports are seen as the same user.');
  ws.views = [{ state: 'frozen', ySplit: 0 }];
}

export async function writeParitySummary(
  diff: WorkbookDiffSummary,
  meta: ParitySummaryMeta,
  outPath: string,
): Promise<{ passed: boolean; differingSheets: number }> {
  const caveats = meta.comparisonCaveats ?? [];
  const dataPassed = TREAT_HEADER_DIFF_AS_FAILURE ? diff.identical : diff.dataIdentical;
  // Caveats override everything: a run where the two sides weren't filtered
  // identically (or pages went missing) is NOT comparable, so neither a green
  // PASS nor a plain FAIL would be honest.
  const passed = dataPassed && caveats.length === 0;
  const differing = diff.sheets.filter(s => TREAT_HEADER_DIFF_AS_FAILURE ? !s.identical : !s.dataIdentical).length;

  const totalHeaderDiffs = diff.sheets.reduce((s, sh) => s + sh.headerDiffCount, 0);
  const totalStructuralDiffs = diff.sheets.reduce((s, sh) => s + sh.structuralDiffCount, 0);

  const wb = new ExcelJS.Workbook();
  wb.creator = 'VériRapport';
  wb.created = new Date();

  // ── Model: every visual goes into ONE plain bucket (see summary-format.helpers.ts)
  const classified = classifyAll(diff);
  const counts = tally(classified);
  const filterRecs = meta.filterRecords ?? [];
  const sameFilters = sameFiltersAnswer(filterRecs, diff.sheets.map(s => s.sheet));
  const facts = referenceFacts(classified);
  const capped = rowCapVisuals(classified);
  const comparedPages = diff.sheets.filter(sh => sh.inExpected && sh.inActual);
  const verdictKind: 'caveat' | 'pass' | 'soft' | 'fail' =
    caveats.length > 0 ? 'caveat'
    : counts.real > 0 ? 'fail'
    : (counts.eye > 0 || counts.expected > 0) ? 'soft'
    : 'pass';
  const verdict =
    verdictKind === 'caveat' ? `⚠️ NOT COMPARABLE — ${caveats.length} problem(s) with how this run was set up (see "Before you trust this" below)`
    : verdictKind === 'fail' ? `❌ ${counts.real} REAL DIFFERENCE(S) found` + (counts.expected ? ` · ${counts.expected} differ as expected` : '') + (counts.eye ? ` · ${counts.eye} to check by eye` : '')
    : verdictKind === 'soft' ? `🟡 NO REAL DIFFERENCES — ${counts.expected} differ as expected (version / dates)` + (counts.eye ? ` · ${counts.eye} to check by eye` : '')
    : '✅ PASS — source and target data match';
  const verdictFill = verdictKind === 'caveat' ? E_WARN : verdictKind === 'fail' ? E_FAIL : verdictKind === 'soft' ? E_WARN : E_PASS;
  const verdictDark = verdictKind === 'fail' || verdictKind === 'pass';

  // Sheets are created in reading order: summary first, details after.
  const glanceWs = wb.addWorksheet('At a glance');
  const whatWs = wb.addWorksheet('What differs');
  const visWs = wb.addWorksheet('All visuals');
  glanceWs.properties.tabColor = { argb: verdictDark ? (verdictKind === 'pass' ? 'FF00B050' : 'FFFF0000') : 'FFFFC000' };

  const styleHeader = (ws: ExcelJS.Worksheet, cols: number) => {
    ws.getRow(1).eachCell(c => {
      c.fill = E_HDR; c.font = { color: { argb: 'FFFFFFFF' }, bold: true };
      c.border = E_CB; c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    });
    ws.getRow(1).height = 30;
    ws.views = [{ state: 'frozen', ySplit: 1 }];          // header stays visible while scrolling
    ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: cols } };
  };
  const link = (cell: ExcelJS.Cell, text: string, sheetName: string, row: number) => {
    cell.value = { text, hyperlink: `#'${sheetName}'!A${row}` };
    cell.font = { color: { argb: 'FF0563C1' }, underline: true };
  };
  const dark = (f: ExcelJS.Fill) => f === E_PASS || f === E_FAIL;
  const num = (v: string): number | null => {
    const t = v.trim().replace(/,/g, '');
    if (!/^-?\d+(\.\d+)?%?$/.test(t)) return null;
    return parseFloat(t);
  };

  const whatRows = new Map<VisualBlockDiff, ExcelJS.Row>();
  const BUCKET_RANK = new Map(BUCKET_ORDER.map((b, i) => [b, i] as const));
  const bucketFill: Record<Bucket, ExcelJS.Fill> = { real: E_FAIL_LT, eye: E_WARN_LT, expected: E_INFO_LT, na: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } }, match: E_PASS_LT };
  const bigNum = (c: ClassifiedVisual): number => {
    let m = 0;
    for (const cr of c.v.changedRows ?? []) for (const ch of cr.changes) {
      const e = num(ch.expected), a = num(ch.actual);
      if (e !== null && a !== null) m = Math.max(m, Math.abs(a - e));
    }
    return m;
  };

  const problemVisuals: Array<{ sheet: string; visual: VisualBlockDiff }> = [];
  for (const c of classified.slice().sort((x, y) => BUCKET_RANK.get(x.bucket)! - BUCKET_RANK.get(y.bucket)!)) {
    if (!visualDataMatches(c.v.status)) problemVisuals.push({ sheet: c.sheet, visual: c.v });
  }
  const detailFirstRow = new Map<VisualBlockDiff, { sheetName: string; row: number }>();

  // ── Sheet 3: All visuals — one row per visual, worst first ────────────────
  visWs.columns = [
    { header: 'Page',                         key: 'sheet',  width: 24 },
    { header: 'Visual',                       key: 'title',  width: 34 },
    { header: 'Type',                         key: 'type',   width: 16 },
    { header: 'Result',                       key: 'result', width: 28 },
    { header: 'What differs',                 key: 'note',   width: 60 },
    { header: 'Rows (old report)',            key: 'rowsE',  width: 11 },
    { header: 'Rows (new report)',            key: 'rowsA',  width: 11 },
    { header: 'How it was paired',            key: 'paired', width: 28 },
    { header: 'Why it came back empty',       key: 'empty',  width: 58 },
    { header: 'Details',                      key: 'details',width: 16 },
  ];
  styleHeader(visWs, visWs.columns.length);
  const visRows = new Map<VisualBlockDiff, ExcelJS.Row>();
  const firstVisRowOfPage = new Map<string, number>();
  const orderedVisuals = classified.slice().sort((x, y) =>
    BUCKET_RANK.get(x.bucket)! - BUCKET_RANK.get(y.bucket)! || x.sheet.localeCompare(y.sheet) || x.v.title.localeCompare(y.v.title));
  for (const c of orderedVisuals) {
    const v = c.v;
    const row = visWs.addRow({
      sheet: c.sheet, title: v.title, type: v.type, result: c.result,
      note: v.note ?? '', rowsE: v.rowsExpected, rowsA: v.rowsActual,
      paired: c.paired, empty: c.emptyReason, details: '',
    });
    visRows.set(v, row);
    if (!firstVisRowOfPage.has(c.sheet)) firstVisRowOfPage.set(c.sheet, row.number);
    row.eachCell(cell => { cell.border = E_CB; cell.alignment = { vertical: 'top', wrapText: true }; cell.fill = bucketFill[c.bucket]; });
    row.getCell('result').font = { bold: true };
    if (v.rowsActual !== v.rowsExpected) row.getCell('rowsA').font = { bold: true, color: { argb: 'FFC00000' } };
  }
  if (classified.length === 0) {
    visWs.addRow(['(No per-visual breakdown — every page matched at the fast whole-sheet check.)']);
  }
  visWs.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: visWs.columns.length } };

  // ── Sheet 2: What differs — real differences only ─────────────────────────
  {
    const W = 8;
    whatWs.columns = [{ width: 24 }, { width: 34 }, { width: 16 }, { width: 60 }, { width: 11 }, { width: 11 }, { width: 14 }, { width: 18 }];
    const wLine = (text: string, fill?: ExcelJS.Fill, bold = false, h?: number) => {
      const r = whatWs.addRow([text]);
      whatWs.mergeCells(r.number, 1, r.number, W);
      const c = r.getCell(1);
      c.alignment = { vertical: 'middle', wrapText: true };
      if (fill) c.fill = fill;
      if (bold) c.font = { bold: true };
      if (h) r.height = h;
      return r;
    };
    const reals = classified.filter(c => c.bucket === 'real');
    wLine('What differs — real differences only', E_HDR, true, 26).getCell(1).font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
    if (reals.length === 0) {
      wLine('No real differences were found. Items that differ as expected, were empty, or need a look by eye are on "All visuals".', E_PASS_LT, true, 30);
    } else {
      wLine(`${reals.length} visual(s) have a real difference. Each shows the old report's value, the new report's value, and how far apart they are.`, undefined, false, 22);
      const big = biggestChanges(classified, 3);
      if (big.length) {
        wLine('The biggest changes', E_INFO, true);
        big.forEach((b, i) => wLine(`${i + 1}.  ${b.sheet} › ${b.visual} › ${b.column}${b.key ? ` (${b.key})` : ''}:  ${b.oldValue}  →  ${b.newValue}   (${b.delta > 0 ? '+' : ''}${Math.round(b.delta * 100) / 100})`));
      }
      whatWs.addRow([]);
      const hdr = whatWs.addRow(['Page', 'Visual', 'Type', 'What differs', 'Rows (old)', 'Rows (new)', 'Values changed', 'Details']);
      hdr.eachCell(c => { c.fill = E_HDR; c.font = { bold: true, color: { argb: 'FFFFFFFF' } }; c.border = E_CB; c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }; });
      for (const c of reals.slice().sort((x, y) => bigNum(y) - bigNum(x))) {
        const v = c.v;
        const r = whatWs.addRow([c.sheet, v.title, v.type, v.note || c.result, v.rowsExpected, v.rowsActual, v.changedRowCount ?? '', '']);
        r.eachCell(cell => { cell.border = E_CB; cell.alignment = { vertical: 'top', wrapText: true }; cell.fill = E_FAIL_LT; });
        whatRows.set(v, r);
      }
    }
  }
  // ── Sheet 4: Changed values — every changed cell, biggest change first ───
  const changedVisuals = problemVisuals.filter(({ visual }) => (visual.changedRows?.length ?? 0) > 0);
  if (changedVisuals.length > 0) {
    const cWs = wb.addWorksheet('Changed values');
    cWs.columns = [
      { header: 'Page',           key: 'sheet',  width: 22 },
      { header: 'Visual',         key: 'visual', width: 30 },
      { header: 'Kind',           key: 'kind',   width: 14 },
      { header: 'Row (key)',      key: 'key',    width: 38 },
      { header: 'Column',         key: 'column', width: 26 },
      { header: 'Old report',     key: 'exp',    width: 16 },
      { header: 'New report',     key: 'act',    width: 16 },
      { header: 'Change',         key: 'delta',  width: 14 },
      { header: 'Change %',       key: 'pct',    width: 10 },
    ];
    styleHeader(cWs, cWs.columns.length);
    type Chg = { sheet: string; label: string; kind: string; key: string; column: string; e: string; a: string };
    const chgs: Chg[] = [];
    for (const { sheet, visual } of changedVisuals) {
      const kind = classified.find(c => c.v === visual)?.bucket === 'expected' ? 'As expected' : 'Real';
      for (const cr of visual.changedRows ?? []) for (const ch of cr.changes) {
        chgs.push({ sheet, label: `${visual.title} [${visual.type}]`, kind, key: cr.key, column: ch.column, e: ch.expected, a: ch.actual });
      }
    }
    const size = (c: Chg) => { const e = num(c.e), a = num(c.a); return e !== null && a !== null ? Math.abs(a - e) : -1; };
    chgs.sort((x, y) => (x.kind === y.kind ? 0 : x.kind === 'Real' ? -1 : 1) || size(y) - size(x));
    for (const c of chgs) {
      const e = num(c.e), a = num(c.a);
      const row = cWs.addRow({ sheet: c.sheet, visual: c.label, kind: c.kind, key: c.key, column: c.column, exp: e ?? c.e, act: a ?? c.a });
      row.eachCell(cell => { cell.border = E_CB; cell.alignment = { vertical: 'middle' }; });
      const pctText = c.e.trim().endsWith('%') && c.a.trim().endsWith('%');
      row.getCell('exp').fill = E_MISS;
      row.getCell('act').fill = c.kind === 'Real' ? E_FAIL_LT : E_INFO_LT;
      row.getCell('act').font = { bold: true };
      if (e !== null && a !== null) {
        const fmt = pctText ? '#,##0.##"%"' : '#,##0.##';
        row.getCell('exp').numFmt = fmt; row.getCell('act').numFmt = fmt;
        row.getCell('delta').value = a - e;
        row.getCell('delta').numFmt = pctText ? '+0.##" pp";-0.##" pp";0' : '+#,##0.##;-#,##0.##;0';
        if (e !== 0) { row.getCell('pct').value = (a - e) / Math.abs(e); row.getCell('pct').numFmt = '+0.0%;-0.0%;0.0%'; }
        else row.getCell('pct').value = 'n/a';
      } else {
        row.getCell('delta').value = 'text';
      }
      row.getCell('delta').alignment = { horizontal: 'right', vertical: 'middle' };
      row.getCell('pct').alignment = { horizontal: 'right', vertical: 'middle' };
    }
  }

  // ── Sheet 5: Filters Applied ──────────────────────────────────────────────
  writeFiltersSheet(wb, meta);

  // ── Sheet 6: Run details — IDs, files, counts, mapping ────────────────────
  {
    const rWs = wb.addWorksheet('Run details');
    rWs.columns = [{ width: 44 }, { width: 80 }];
    const rows: [string, string][] = [
      ['Pair',                meta.pairName],
      ['Run Mode',            meta.mode],
      // Pinned locale + named month so the summary reads the same wherever it was generated.
      ['Run Time',            new Date().toLocaleString('en-US', { dateStyle: 'medium', timeStyle: 'medium' })],
      [`Old report (${meta.sourceLabel})`, meta.sourceReportId],
      [`New report (${meta.targetLabel})`, meta.targetReportId],
      ['Old report file',     meta.expectedFile],
      ['New report file',     meta.actualFile],
      ['Pages compared',      String(comparedPages.length)],
      ['Pages fully identical (strict)', String(diff.sheets.filter(s => s.identical).length)],
      ['Pages with only header-text/order differences', String(diff.sheets.filter(s => !s.identical && s.dataIdentical).length)],
      ['Pages with data/structure differences', String(differing)],
      ['Pages only in old report', diff.sheetsOnlyInExpected.join(', ') || '(none)'],
      ['Pages only in new report', diff.sheetsOnlyInActual.join(', ') || '(none)'],
      ['Visuals with only a header-text difference', String(totalHeaderDiffs)],
      ['Visuals missing / duplicated / ambiguous', String(totalStructuralDiffs)],
      ['Header diffs count as failure?', TREAT_HEADER_DIFF_AS_FAILURE ? 'Yes' : 'No — see TREAT_HEADER_DIFF_AS_FAILURE'],
    ];
    if (meta.mappingFile) {
      rows.push(['Mapping file', meta.mappingFile]);
      const byType = new Map<string, number>();
      for (const sh of diff.sheets) for (const v of sh.visuals) {
        if (v.mapping) byType.set(v.mapping.changeLabel, (byType.get(v.mapping.changeLabel) ?? 0) + 1);
      }
      for (const [label, n] of byType) rows.push([`  Mapped visuals — ${label}`, String(n)]);
    }
    if (meta.pageAlignment) {
      rows.push(['Report pages in both reports', String(meta.pageAlignment.inBothCount)]);
      rows.push(['Report pages only in old report', meta.pageAlignment.onlyInSource.join(', ') || '(none)']);
      rows.push(['Report pages only in new report', meta.pageAlignment.onlyInTarget.join(', ') || '(none)']);
    }
    for (const [label, value] of rows) {
      const r = rWs.addRow([label, value]);
      r.getCell(1).font = { bold: true }; r.getCell(1).fill = E_INFO;
      r.eachCell(c => { c.border = E_CB; c.alignment = { vertical: 'middle', wrapText: true }; });
    }
  }


  const SHOW = 50; // rows shown per kind per visual; the full list is on "Changed Values"
  const colsOf = (v: VisualBlockDiff): string[] =>
    v.diffColumns ?? (v.status === 'only-in-actual' ? v.headerActual : v.headerExpected);

  if (problemVisuals.length > 0) {
    const dWs = wb.addWorksheet('Differences');
    dWs.properties.tabColor = { argb: 'FFC00000' };
    const maxCols = Math.max(3, ...problemVisuals.map(({ visual }) => colsOf(visual).length));
    const lastCol = maxCols + 1;
    dWs.getColumn(1).width = 22;
    for (let c = 2; c <= lastCol; c++) dWs.getColumn(c).width = 20;

    const F = (argb: string): ExcelJS.Fill => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
    const FILL_ORIG = F('FFF2F2F2');
    const FILL_CHG = F('FFE8F1FB');
    const FILL_HIT = F('FFFFB3B3');
    const FILL_ONLY_O = F('FFFFF2CC');
    const FILL_ONLY_C = F('FFE2F3E1');
    const merged = (text: string, opts: { fill?: ExcelJS.Fill; font?: Partial<ExcelJS.Font>; height?: number } = {}) => {
      const r = dWs.addRow([text]);
      dWs.mergeCells(r.number, 1, r.number, lastCol);
      const c = r.getCell(1);
      c.alignment = { vertical: 'middle', wrapText: true };
      if (opts.fill) c.fill = opts.fill;
      if (opts.font) c.font = opts.font;
      if (opts.height) r.height = opts.height;
      return r;
    };

    // Legend
    merged('Differences — original report (source) compared with the changed report (target)', { font: { bold: true, size: 15, color: { argb: 'FF1F3864' } }, height: 28 });
    merged('Each box below is ONE visual that does not match. Read a box from top to bottom:', { font: { italic: true }, height: 20 });
    const legend: Array<[string, ExcelJS.Fill, string]> = [
      ['Original', FILL_ORIG, 'The values in the original report (source, Import mode).'],
      ['Changed', FILL_CHG, 'The values in the changed report (target, Direct Lake). A value that is different is marked in red.'],
      ['Difference', F('FFFFFFFF'), 'How far the value moved: changed minus original (percentage values are shown in percentage points).'],
      ['Only in original', FILL_ONLY_O, 'A row that exists in the original report but is missing from the changed report.'],
      ['Only in changed', FILL_ONLY_C, 'A row that exists only in the changed report.'],
    ];
    for (const [label, fill, text] of legend) {
      const r = dWs.addRow([label, text]);
      dWs.mergeCells(r.number, 2, r.number, lastCol);
      r.getCell(1).fill = fill; r.getCell(1).font = { bold: true };
      r.getCell(1).border = E_CB;
      r.getCell(2).alignment = { vertical: 'middle', wrapText: true };
    }
    const hit = dWs.addRow(['Red cell', 'A value that differs between the two reports.']);
    dWs.mergeCells(hit.number, 2, hit.number, lastCol);
    hit.getCell(1).fill = FILL_HIT; hit.getCell(1).font = { bold: true, color: { argb: 'FF9C0006' } }; hit.getCell(1).border = E_CB;
    dWs.addRow([]);

    const deltaOf = (e: string, a: string): { value: number | string; fmt?: string } => {
      const en = num(e), an = num(a);
      if (en === null || an === null) return { value: 'text changed' };
      const pp = e.trim().endsWith('%') && a.trim().endsWith('%');
      return { value: an - en, fmt: pp ? '+0.##" pp";-0.##" pp";0' : '+#,##0.##;-#,##0.##;0' };
    };
    const padTo = (r: string[], n: number) => Array.from({ length: n }, (_, i) => r[i] ?? '');

    for (const { sheet, visual: v } of problemVisuals) {
      const cols = colsOf(v);
      const n = cols.length;
      const changed = v.changedRows ?? [];
      const unpairedE = Math.max(0, v.onlyInExpectedCount - (v.changedRowCount ?? 0));
      const unpairedA = Math.max(0, v.onlyInActualCount - (v.changedRowCount ?? 0));

      const head = merged(`${sheet}   ›   ${v.title}   (${v.type})`, { fill: E_HDR, font: { bold: true, size: 12, color: { argb: 'FFFFFFFF' } }, height: 24 });
      detailFirstRow.set(v, { sheetName: 'Differences', row: head.number });

      let what: string;
      if (v.status === 'only-in-expected') what = '❌ This visual is in the original report but is MISSING from the changed report.';
      else if (v.status === 'only-in-actual') what = '❌ This visual is NEW: it is in the changed report but not in the original.';
      else if (v.status === 'ambiguous') what = '❓ Several visuals share this title, so the tool could not tell which one matches which. Please compare them by hand.';
      else {
        const parts: string[] = [];
        if (v.changedRowCount) parts.push(`${v.changedRowCount} row(s) have different values`);
        if (unpairedE) parts.push(`${unpairedE} row(s) only in the original`);
        if (unpairedA) parts.push(`${unpairedA} row(s) only in the changed report`);
        const colsDiffer = (v.columnsOnlyInExpected?.length ?? 0) + (v.columnsOnlyInActual?.length ?? 0) > 0;
        what = parts.length > 0
          ? `❌ ${parts.join('  ·  ')}.`
          : colsDiffer ? '❌ The two reports show different columns. The values in the columns they share are identical.'
                       : '❌ The values differ.';
      }
      merged(what, { fill: E_FAIL_LT, font: { bold: true }, height: 22 });
      merged(`Original has ${v.rowsExpected} row(s); the changed report has ${v.rowsActual} row(s).` +
        (v.countExpected !== v.countActual && v.countExpected > 0 && v.countActual > 0 ? ` (Visual appears ${v.countExpected}× in the original and ${v.countActual}× in the changed report.)` : ''),
        { font: { color: { argb: 'FF595959' } } });
      if ((v.columnsOnlyInExpected?.length ?? 0) > 0 || (v.columnsOnlyInActual?.length ?? 0) > 0) {
        const bits: string[] = [];
        if (v.columnsOnlyInExpected?.length) bits.push(`Columns only in the original: ${v.columnsOnlyInExpected.join(', ')}`);
        if (v.columnsOnlyInActual?.length) bits.push(`Columns only in the changed report: ${v.columnsOnlyInActual.join(', ')}`);
        merged(`ℹ ${bits.join('.   ')}. The rows below compare only the columns both reports have.`, { fill: E_WARN_LT });
      } else if (v.note && /renamed|matched by identical columns|different in text/i.test(v.note)) {
        merged(`ℹ ${v.note}`, { fill: E_WARN_LT });
      }
      if (v.status === 'ambiguous' && v.note) merged(v.note, { font: { italic: true, color: { argb: 'FF595959' } } });

      const hasRows = changed.length > 0 || v.sampleOnlyInExpected.length > 0 || v.sampleOnlyInActual.length > 0;
      if (v.status !== 'ambiguous' && hasRows && n > 0) {
        const hdr = dWs.addRow(['', ...cols]);
        hdr.eachCell({ includeEmpty: false }, c => { c.fill = F('FF2F5597'); c.font = { bold: true, color: { argb: 'FFFFFFFF' } }; c.border = E_CB; c.alignment = { vertical: 'middle', wrapText: true }; });

        const line = (label: string, values: string[], fill: ExcelJS.Fill, hitCols?: Set<number>, bold = false) => {
          const r = dWs.addRow([label, ...padTo(values, n)]);
          for (let c = 1; c <= n + 1; c++) {
            const cell = r.getCell(c);
            cell.fill = fill; cell.border = E_CB; cell.alignment = { vertical: 'middle', wrapText: true };
          }
          r.getCell(1).font = { bold: true };
          if (hitCols) for (const i of hitCols) { const cell = r.getCell(i + 2); cell.fill = FILL_HIT; cell.font = { bold: true, color: { argb: 'FF9C0006' } }; }
          if (bold) r.getCell(1).font = { bold: true };
          return r;
        };

        for (const cr of changed.slice(0, SHOW)) {
          const eRow = padTo(cr.expectedRow ?? [], n);
          const aRow = padTo(cr.actualRow ?? [], n);
          const diffIdx = new Set<number>();
          for (let i = 0; i < n; i++) if (eRow[i] !== aRow[i]) diffIdx.add(i);
          line('Original', eRow, FILL_ORIG);
          line('Changed', aRow, FILL_CHG, diffIdx);
          const d = dWs.addRow(['Difference', ...Array(n).fill('')]);
          for (let c = 1; c <= n + 1; c++) { d.getCell(c).border = E_CB; d.getCell(c).alignment = { horizontal: 'right', vertical: 'middle' }; }
          d.getCell(1).font = { bold: true }; d.getCell(1).alignment = { horizontal: 'left', vertical: 'middle' };
          for (const i of diffIdx) {
            const dv = deltaOf(eRow[i], aRow[i]);
            const cell = d.getCell(i + 2);
            cell.value = dv.value;
            if (dv.fmt) cell.numFmt = dv.fmt;
            cell.font = { bold: true, color: { argb: typeof dv.value === 'number' ? (dv.value > 0 ? 'FF1E7B34' : 'FFC00000') : 'FF595959' } };
          }
          dWs.addRow([]).height = 6;
        }
        if (changed.length > SHOW) merged(`… ${v.changedRowCount! - SHOW} more changed row(s) not shown here — see the "Changed Values" sheet for the full list.`, { font: { italic: true, color: { argb: 'FF595959' } } });

        const shownE = v.sampleOnlyInExpected.slice(0, SHOW);
        for (const r of shownE) line('Only in original', r, FILL_ONLY_O);
        if (unpairedE > shownE.length) merged(`… ${unpairedE - shownE.length} more row(s) only in the original are not shown.`, { font: { italic: true, color: { argb: 'FF595959' } } });
        const shownA = v.sampleOnlyInActual.slice(0, SHOW);
        for (const r of shownA) line('Only in changed', r, FILL_ONLY_C);
        if (unpairedA > shownA.length) merged(`… ${unpairedA - shownA.length} more row(s) only in the changed report are not shown.`, { font: { italic: true, color: { argb: 'FF595959' } } });
      }
      dWs.addRow([]);
      dWs.addRow([]);
    }
    dWs.views = [{ state: 'frozen', ySplit: 0, showGridLines: false }];
  }

  // ── Sheet 1: At a glance (filled last, now that every row number is known) ─
  {
    glanceWs.columns = [{ width: 30 }, { width: 36 }, { width: 36 }, { width: 8 }, { width: 8 }, { width: 8 }, { width: 8 }, { width: 8 }, { width: 46 }];
    const COLS = 9;
    const gLine = (text: string, opts: { fill?: ExcelJS.Fill; bold?: boolean; size?: number; height?: number; color?: string; italic?: boolean } = {}) => {
      const r = glanceWs.addRow([text]);
      glanceWs.mergeCells(r.number, 1, r.number, COLS);
      const c = r.getCell(1);
      c.alignment = { vertical: 'middle', wrapText: true };
      if (opts.fill) c.fill = opts.fill;
      c.font = { bold: !!opts.bold, italic: !!opts.italic, size: opts.size, color: opts.color ? { argb: opts.color } : undefined };
      if (opts.height) r.height = opts.height;
      return r;
    };
    const gKV = (label: string, value: string, fill?: ExcelJS.Fill, height?: number) => {
      const r = glanceWs.addRow([label, value]);
      glanceWs.mergeCells(r.number, 2, r.number, COLS);
      r.getCell(1).font = { bold: true }; r.getCell(1).fill = E_INFO;
      r.getCell(1).border = E_CB; r.getCell(2).border = E_CB;
      r.getCell(1).alignment = { vertical: 'middle', wrapText: true };
      r.getCell(2).alignment = { vertical: 'middle', wrapText: true };
      if (fill) r.getCell(2).fill = fill;
      if (height) r.height = height;
      return r;
    };
    const gHead = (text: string) => gLine(text, { fill: E_HDR, bold: true, size: 12, color: 'FFFFFFFF', height: 22 });

    gLine(verdict, { fill: verdictFill, bold: true, size: 16, height: 38, color: verdictDark ? 'FFFFFFFF' : 'FF000000' }).getCell(1).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    glanceWs.addRow([]);

    // Were the same filters used?
    gHead('Were the same filters used on both reports?');
    gKV('Same filters?',
      sameFilters.answer === 'YES' ? '✅ YES' : sameFilters.answer === 'NO' ? '❌ NO' : '❓ Not confirmed',
      sameFilters.answer === 'YES' ? E_PASS_LT : sameFilters.answer === 'NO' ? E_FAIL_LT : E_WARN_LT);
    gKV('', sameFilters.detail, undefined, 32);
    glanceWs.addRow([]);

    // Reasons the data can differ even with the same filters
    if (facts.length > 0 || capped.length > 0) {
      gHead('Expected reasons for differences');
      for (const f of facts) gKV(f.label, `Old report: ${f.oldValue}     →     New report: ${f.newValue}`, E_INFO_LT);
      if (capped.length > 0) {
        gKV('Export row limit hit', `${capped.length} visual(s) exported ${EXPORT_ROW_CAP.toLocaleString('en-US')} rows, the most Power BI will export, so their lists may be cut off at different places: ${capped.slice(0, 5).join('; ')}${capped.length > 5 ? '; …' : ''}`, E_WARN_LT, 44);
      }
      glanceWs.addRow([]);
    }

    // Tally
    gHead(`Visuals compared: ${classified.length}`);
    for (const b of BUCKET_ORDER) {
      const r = glanceWs.addRow([`${BUCKET_ICON[b]}  ${counts[b]}`, BUCKET_MEANING[b]]);
      glanceWs.mergeCells(r.number, 2, r.number, COLS);
      r.getCell(1).font = { bold: true, size: 12 }; r.getCell(1).fill = bucketFill[b];
      r.eachCell(c => { c.border = E_CB; c.alignment = { vertical: 'middle', wrapText: true }; });
      r.height = 22;
    }
    glanceWs.addRow([]);

    // Caveats / notes
    if (caveats.length > 0 || (meta.notes ?? []).length > 0) {
      gHead(caveats.length > 0 ? 'Before you trust this' : 'Notes');
      caveats.forEach(c => gKV('⚠ Problem', c, E_WARN_LT, 44));
      (meta.notes ?? []).forEach(n => gKV('ℹ Note', n, E_INFO_LT, 44));
      glanceWs.addRow([]);
    }

    // Per-page table
    gHead('Page by page');
    const ph = glanceWs.addRow(['Page', 'Filters — old report', 'Filters — new report', '✅', '❌', 'ℹ️', '⚪', '🔍', 'Look at this first']);
    ph.eachCell(c => { c.fill = E_HDR; c.font = { bold: true, color: { argb: 'FFFFFFFF' } }; c.border = E_CB; c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }; });
    const perPage = diff.sheets.map(sh => {
      const mine = classified.filter(c => c.sheet === sh.sheet);
      return { sh, mine, t: tally(mine) };
    }).sort((x, y) => (y.t.real - x.t.real) || (y.t.eye - x.t.eye));
    for (const { sh, mine, t } of perPage) {
      let look = '';
      if (sh.inExpected && !sh.inActual) look = 'Page is only in the old report';
      else if (!sh.inExpected && sh.inActual) look = 'Page is only in the new report';
      else if (t.real > 0) look = mine.filter(c => c.bucket === 'real').sort((x, y) => bigNum(y) - bigNum(x))[0].v.title;
      else if (t.eye > 0) look = `(check by eye) ${mine.find(c => c.bucket === 'eye')!.v.title}`;
      const fast = mine.length === 0 && sh.identical;
      const r = glanceWs.addRow([
        sh.sheet, pageFilterText(filterRecs, sh.sheet, true), pageFilterText(filterRecs, sh.sheet, false),
        fast ? 'all' : t.match, t.real, t.expected, t.na, t.eye, look,
      ]);
      r.eachCell(c => { c.border = E_CB; c.alignment = { vertical: 'top', wrapText: true }; });
      for (let i = 4; i <= 8; i++) r.getCell(i).alignment = { horizontal: 'center', vertical: 'top' };
      if (t.real > 0) r.getCell(5).fill = E_FAIL_LT;
      if (t.eye > 0) r.getCell(8).fill = E_WARN_LT;
      const lr = firstVisRowOfPage.get(sh.sheet);
      if (lr !== undefined) link(r.getCell(1), sh.sheet, 'All visuals', lr);
    }
    glanceWs.addRow([]);
    gHead('Where to look next');
    gKV('What differs', 'Only the real differences, with the 3 biggest changes.');
    gKV('All visuals', 'Every visual, worst first. The last-but-one column says why a visual came back empty.');
    gKV('Changed values', 'Every changed number: old value, new value, change and % change.');
    gKV('Filters Applied', 'Which filters and slicers were found, cleared, applied and really active on each report.');
    gKV('Run details', 'Report IDs, files, page and mapping counts.');
    glanceWs.views = [{ state: 'normal', showGridLines: false }];
  }

  // ── Hyperlinks (set last, once every target row number is known) ──────────
  for (const [v, row] of visRows) {
    const d = detailFirstRow.get(v);
    if (d) link(row.getCell('details'), 'See differences →', d.sheetName, d.row);
  }
  for (const [v, row] of whatRows) {
    const d = detailFirstRow.get(v);
    if (d) link(row.getCell(8), 'See differences →', d.sheetName, d.row);
  }

  wb.eachSheet(ws => { ws.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 }; });
  fs.mkdirSync(require('path').dirname(outPath), { recursive: true });
  await wb.xlsx.writeFile(outPath);
  return { passed, differingSheets: differing };
}
