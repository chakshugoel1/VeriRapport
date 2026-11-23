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
    const { visualsWithData, totalRows } = writeVisualsToSheet(ws, visuals);
    ws.commit(); // flush this page's rows to disk now — don't hold them for the rest of the run
    recordTiming(label, `page:${rp.displayName}:writeSheet`, Date.now() - tWrite);

    exported.push({
      sheetName: wsName, pageDisplayName: rp.displayName, pageInternalName: rp.name,
      visualsWithData, totalRows,
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
export function writeVisualsToSheet(
  ws: ExcelJS.Worksheet,
  visuals: VisualExport[],
): { visualsWithData: number; totalRows: number } {
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
    return { visualsWithData: 0, totalRows: 0 };
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

  return { visualsWithData: dataVisuals.length, totalRows };
}

// ── Parity summary writer ─────────────────────────────────────────────────────

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
};

const VISUAL_STATUS_FILL: Record<VisualBlockDiff['status'], ExcelJS.Fill> = {
  'identical':        E_PASS,
  'header-diff':       E_INFO,
  'duplicate-count':   E_INFO,
  'data-diff':         E_FAIL,
  'only-in-expected':  E_MISS,
  'only-in-actual':    E_WARN,
  'ambiguous':         E_WARN,
};

/** True when a visual's status means "the data matches" (part of the relaxed gate). */
function visualDataMatches(status: VisualBlockDiff['status']): boolean {
  return status === 'identical' || status === 'header-diff' || status === 'duplicate-count';
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
  if (status === 'header-diff' || status === 'duplicate-count') return E_INFO_LT;
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
        case 'safe':     if (v.status === 'identical') stats.identicalVisuals++; else stats.cosmeticOnly++; break;
        case 'review':   stats.needsReview++; break;
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

  const diffVisuals = stats.totalVisuals - stats.identicalVisuals;

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

  // ── Sheet 1: Summary ──────────────────────────────────────────────────────
  const infoWs = wb.addWorksheet('Parity Summary');
  infoWs.columns = [{ width: 44 }, { width: 80 }];

  const verdict = caveats.length > 0
    ? `⚠️ NOT COMPARABLE — ${caveats.length} caveat(s) invalidate this run (see Caveat rows below)`
    : passed
      ? '✅ PASS — source and target data match'
      : `❌ FAIL — ${differing} page(s) have genuine data/structure differences`;
  // Visual-level tallies for the at-a-glance block and the page table.
  const allVisuals = diff.sheets.flatMap(sh => sh.visuals);
  const sevCount = { safe: 0, review: 0, critical: 0 };
  let identicalVisuals = 0;
  for (const v of allVisuals) {
    sevCount[VISUAL_SEVERITY[v.status]]++;
    if (v.status === 'identical') identicalVisuals++;
  }
  const comparedPages = diff.sheets.filter(sh => sh.inExpected && sh.inActual);
  const passingPages = comparedPages.filter(sh => TREAT_HEADER_DIFF_AS_FAILURE ? sh.identical : sh.dataIdentical).length;
  infoWs.properties.tabColor = { argb: caveats.length > 0 ? 'FFFFC000' : (passed ? 'FF00B050' : 'FFFF0000') };

  const banner = infoWs.addRow([verdict, '']);
  infoWs.mergeCells(banner.number, 1, banner.number, 2);
  banner.height = 38;
  banner.getCell(1).fill = caveats.length > 0 ? E_WARN : (passed ? E_PASS : E_FAIL);
  banner.getCell(1).font = { bold: true, size: 16, color: { argb: caveats.length > 0 ? 'FF000000' : 'FFFFFFFF' } };
  banner.getCell(1).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  const glanceRows: Array<[string, string, ExcelJS.Fill | undefined]> = [
    ['Pages matching',            `${passingPages} of ${comparedPages.length}`, passingPages === comparedPages.length ? E_PASS_LT : E_FAIL_LT],
    ['Visuals checked one by one', String(allVisuals.length) + (allVisuals.length === 0 ? ' (every page matched at the fast whole-page check)' : ''), undefined],
    ['  ✅ Identical',             String(identicalVisuals), E_PASS_LT],
    ['  📝 Cosmetic only (label / copy count)', String(sevCount.safe - identicalVisuals), E_INFO_LT],
    ['  ❓ Needs manual review',   String(sevCount.review), sevCount.review > 0 ? E_WARN_LT : undefined],
    ['  ❌ Critical (data differs or visual missing/extra)', String(sevCount.critical), sevCount.critical > 0 ? E_FAIL_LT : E_PASS_LT],
  ];
  for (const [label, value, fill] of glanceRows) {
    const r = infoWs.addRow([label, value]);
    r.getCell(1).font = { bold: true };
    r.eachCell(c => { c.border = E_CB; c.alignment = { vertical: 'middle', wrapText: true }; });
    if (fill) r.getCell(2).fill = fill;
    r.getCell(2).font = { bold: true, size: 12 };
  }
  const guide = infoWs.addRow(['Where to look', 'Page Comparison → which pages fail.  Visual Comparison → which visuals.  Changed Values → which cell changed, from what to what (with % change).  Differences → rows that exist on one side only.']);
  guide.getCell(1).font = { bold: true };
  guide.eachCell(c => { c.border = E_CB; c.alignment = { vertical: 'middle', wrapText: true }; });
  guide.height = 48;
  infoWs.addRow(['', '']);

  // Analysis narrative FIRST — the plain-English "do I need to worry about
  // this" answer a reviewer wants before any raw counts. Purely derived from
  // stats computed below; never a second source of truth.
  const analysisStats = computeAnalysisStats(diff);
  const analysisLines = buildAnalysisNarrative(diff, analysisStats, passed, caveats);
  const analysisHeaderRow = infoWs.addRow(['📋 Analysis', '']);
  infoWs.mergeCells(analysisHeaderRow.number, 1, analysisHeaderRow.number, 2);
  analysisHeaderRow.getCell(1).fill = E_HDR;
  analysisHeaderRow.getCell(1).font = { bold: true, size: 13, color: { argb: 'FFFFFFFF' } };
  analysisHeaderRow.getCell(1).alignment = { vertical: 'middle' };
  analysisHeaderRow.height = 22;
  for (const line of analysisLines) {
    const r = infoWs.addRow([line, '']);
    infoWs.mergeCells(r.number, 1, r.number, 2);
    const c = r.getCell(1);
    c.font = { italic: true };
    c.alignment = { vertical: 'middle', wrapText: true };
    c.fill = E_INFO;
    r.height = 32;
  }
  infoWs.addRow(['', '']); // spacer before the raw stats table below

  const rows: [string, string][] = [
    ['Pair',                meta.pairName],
    ['Run Mode',            meta.mode],
    // Pinned locale + named month: a bare .toLocaleString() renders in
    // whichever OS locale ran the script (dd/mm vs mm/dd is genuinely
    // ambiguous across countries) - this way the summary reads the same
    // regardless of who generated it or where.
    ['Run Time',            new Date().toLocaleString('en-US', { dateStyle: 'medium', timeStyle: 'medium' })],
    [`Source (${meta.sourceLabel})`, meta.sourceReportId],
    [`Target (${meta.targetLabel})`, meta.targetReportId],
    ['Expected File',       meta.expectedFile],
    ['Actual File',         meta.actualFile],
    ['Pages Compared',      String(diff.sheets.filter(s => s.inExpected && s.inActual).length)],
    ['Pages Fully Identical (strict)', String(diff.sheets.filter(s => s.identical).length)],
    ['Pages With Only Header-Text/Order Diffs', String(diff.sheets.filter(s => !s.identical && s.dataIdentical).length)],
    ['Pages With Real Data/Structure Diffs', String(differing)],
    ['Pages Only in Source', diff.sheetsOnlyInExpected.join(', ') || '(none)'],
    ['Pages Only in Target', diff.sheetsOnlyInActual.join(', ') || '(none)'],
    ['Header-Only Visual Diffs (see Visual Comparison)', String(totalHeaderDiffs)],
    ['Structural Visual Diffs — missing/duplicated/ambiguous (see Visual Comparison)', String(totalStructuralDiffs)],
    ['Header diffs count as failure?', TREAT_HEADER_DIFF_AS_FAILURE ? 'Yes' : 'No — see TREAT_HEADER_DIFF_AS_FAILURE'],
  ];
  if (meta.pageAlignment) {
    rows.push(['Report Pages In Both Reports', String(meta.pageAlignment.inBothCount)]);
    rows.push(['Report Pages Only In Source Report', meta.pageAlignment.onlyInSource.join(', ') || '(none)']);
    rows.push(['Report Pages Only In Target Report', meta.pageAlignment.onlyInTarget.join(', ') || '(none)']);
  }
  caveats.forEach((c, i) => rows.push([`⚠ Caveat ${i + 1}`, c]));
  rows.push(['Overall Result', verdict]);
  let verdictRow: ExcelJS.Row | undefined;
  rows.forEach(([label, value]) => {
    const r = infoWs.addRow([label, value]);
    r.getCell(1).font = { bold: true };
    r.getCell(1).fill = label.startsWith('⚠ Caveat') ? E_WARN : E_INFO;
    r.eachCell(c => { c.border = E_CB; c.alignment = { vertical: 'middle', wrapText: true }; });
    if (label === 'Overall Result') verdictRow = r;
  });
  verdictRow!.getCell(2).fill = caveats.length > 0 ? E_WARN : (passed ? E_PASS : E_FAIL);
  verdictRow!.getCell(2).font = { bold: true, color: { argb: caveats.length > 0 ? 'FF000000' : 'FFFFFFFF' } };

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

  // ── Sheet 2: Page Comparison ──────────────────────────────────────────────
  // Worst pages first, so the problems are at the top of the list.
  const pageRank = (sh: typeof diff.sheets[number]) =>
    (!sh.inExpected || !sh.inActual) ? 0 : !sh.dataIdentical ? 1 : !sh.identical ? 2 : 3;
  const pagesOrdered = diff.sheets.slice().sort((x, y) => pageRank(x) - pageRank(y));

  const cmpWs = wb.addWorksheet('Page Comparison');
  cmpWs.columns = [
    { header: 'Page (Sheet)',            key: 'sheet',   width: 30 },
    { header: 'Status',                  key: 'status',  width: 30 },
    { header: 'Visuals checked',         key: 'vis',     width: 11 },
    { header: '✅ Identical',            key: 'ident',   width: 11 },
    { header: '❓ Review',               key: 'review',  width: 10 },
    { header: '❌ Critical',             key: 'crit',    width: 10 },
    { header: 'Visuals matching',        key: 'pct',     width: 12 },
    { header: 'Rows in Source',          key: 'exp',     width: 12 },
    { header: 'Rows in Target',          key: 'act',     width: 12 },
    { header: 'Rows differing (approx.)', key: 'rowsDiff', width: 16 },
  ];
  styleHeader(cmpWs, cmpWs.columns.length);
  const pageCells = new Map<string, ExcelJS.Cell>();
  for (const s of pagesOrdered) {
    let statusText: string;
    let statusFill: ExcelJS.Fill;
    if (s.inExpected && !s.inActual) { statusText = '⚠️ Missing in Target'; statusFill = E_WARN; }
    else if (!s.inExpected && s.inActual) { statusText = '🆕 Only in Target'; statusFill = E_WARN; }
    else if (s.identical) { statusText = '✅ Identical'; statusFill = E_PASS; }
    else if (s.dataIdentical) { statusText = '📝 Data OK — header text/order differs'; statusFill = E_INFO; }
    else { statusText = '❌ Different'; statusFill = E_FAIL; }

    const nVis = s.visuals.length;
    const nIdent = s.visuals.filter(v => v.status === 'identical').length;
    const nRev = s.visuals.filter(v => VISUAL_SEVERITY[v.status] === 'review').length;
    const nCrit = s.visuals.filter(v => VISUAL_SEVERITY[v.status] === 'critical').length;
    const nSafe = s.visuals.filter(v => VISUAL_SEVERITY[v.status] === 'safe').length;
    const fast = nVis === 0 && s.identical;                 // matched at the whole-page hash; no per-visual list
    const row = cmpWs.addRow({
      sheet: s.sheet,
      status: statusText,
      vis: fast ? 'all' : nVis,
      ident: fast ? 'all' : nIdent,
      review: nVis === 0 ? '' : nRev,
      crit: nVis === 0 ? '' : nCrit,
      pct: fast ? 1 : (nVis === 0 ? '' : nSafe / nVis),
      exp: s.rowsExpected,
      act: s.rowsActual,
      rowsDiff: s.rowsOnlyInExpected + s.rowsOnlyInActual,
    });
    row.eachCell(c => { c.border = E_CB; c.alignment = { vertical: 'middle' }; });
    row.getCell('pct').numFmt = '0%';
    const sc = row.getCell('status');
    sc.fill = statusFill;
    sc.font = { bold: true, color: { argb: dark(statusFill) ? 'FFFFFFFF' : 'FF000000' } };
    sc.alignment = { horizontal: 'center', vertical: 'middle' };
    if (nCrit > 0) row.getCell('crit').fill = E_FAIL_LT;
    if (nRev > 0) row.getCell('review').fill = E_WARN_LT;
    pageCells.set(s.sheet, row.getCell('sheet'));
  }
  cmpWs.addConditionalFormatting({
    ref: `G2:G${Math.max(2, cmpWs.rowCount)}`,
    rules: [{ type: 'dataBar', priority: 1, cfvo: [{ type: 'num', value: 0 }, { type: 'num', value: 1 }],
              color: { argb: 'FF63BE7B' }, gradient: false } as any],
  });

  // ── Sheet 3: Visual Comparison (one row per Power BI visual, per page) ────
  // The actionable sheet: which visual on which page has a problem and what
  // kind. Within each page the worst problems are listed first.
  const SEV_RANK: Record<Severity, number> = { critical: 0, review: 1, safe: 2 };
  const visWs = wb.addWorksheet('Visual Comparison');
  visWs.columns = [
    { header: 'Page (Sheet)',   key: 'sheet',   width: 24 },
    { header: 'Visual Title',   key: 'title',   width: 34 },
    { header: 'Type',           key: 'type',    width: 18 },
    { header: 'Severity',       key: 'severity',width: 26 },
    { header: 'Status',         key: 'status',  width: 30 },
    { header: 'Rows Source',    key: 'rowsE',   width: 10 },
    { header: 'Rows Target',    key: 'rowsA',   width: 10 },
    { header: 'Rows Δ',         key: 'rowsD',   width: 9 },
    { header: 'Rows differing', key: 'differ',  width: 11 },
    { header: 'Copies (Src/Tgt)', key: 'counts', width: 11 },
    { header: 'Details',        key: 'details', width: 16 },
    { header: 'What differs',   key: 'note',    width: 70 },
    { header: 'Header (Source)', key: 'hdrExp', width: 34 },
    { header: 'Header (Target)', key: 'hdrAct', width: 34 },
  ];
  styleHeader(visWs, visWs.columns.length);

  const visCells = new Map<VisualBlockDiff, ExcelJS.Row>();
  const firstVisRowOfPage = new Map<string, number>();
  let anyVisualRows = false;
  for (const s of pagesOrdered) {
    const ordered = s.visuals.slice().sort((x, y) => SEV_RANK[VISUAL_SEVERITY[x.status]] - SEV_RANK[VISUAL_SEVERITY[y.status]]);
    for (const v of ordered) {
      anyVisualRows = true;
      const severity = VISUAL_SEVERITY[v.status];
      const row = visWs.addRow({
        sheet: s.sheet,
        title: v.title,
        type: v.type,
        severity: SEVERITY_LABEL[severity],
        status: VISUAL_STATUS_LABEL[v.status],
        rowsE: v.rowsExpected,
        rowsA: v.rowsActual,
        rowsD: v.rowsActual - v.rowsExpected,
        differ: v.onlyInExpectedCount + v.onlyInActualCount,
        counts: `${v.countExpected} / ${v.countActual}`,
        details: '',
        note: v.note ?? '',
        hdrExp: v.headerExpected.join(' | '),
        hdrAct: v.headerActual.join(' | '),
      });
      if (!firstVisRowOfPage.has(s.sheet)) firstVisRowOfPage.set(s.sheet, row.number);
      visCells.set(v, row);
      // Whole-row tint by severity, then the severity/status cells get a bolder fill.
      const rowFill = rowFillFor(v.status);
      row.eachCell(c => { c.border = E_CB; c.alignment = { vertical: 'top', wrapText: true }; c.fill = rowFill; });
      const sevCell = row.getCell('severity');
      sevCell.fill = SEVERITY_CELL_FILL[severity];
      sevCell.font = { bold: true, color: { argb: severity === 'critical' ? 'FFFFFFFF' : 'FF000000' } };
      const sc = row.getCell('status');
      const fill = VISUAL_STATUS_FILL[v.status];
      sc.fill = fill;
      sc.font = { bold: true, color: { argb: (fill === E_WARN || fill === E_INFO || fill === E_MISS) ? 'FF000000' : 'FFFFFFFF' } };
      if (v.rowsActual !== v.rowsExpected) row.getCell('rowsD').font = { bold: true, color: { argb: 'FFC00000' } };
      row.getCell('rowsD').numFmt = '+0;-0;0';
    }
  }
  if (!anyVisualRows) {
    visWs.addRow(['(No per-visual breakdown — every page matched at the fast whole-sheet check.)']);
  }

  // ── Sheet 4: Differences — one box per differing visual ───────────────────
  // Original (source) and Changed (target) are shown as the SAME columns,
  // one above the other, with the changed cells highlighted and a Difference
  // line underneath. Rows that exist on only one side are labelled as such.
  const problemVisuals: Array<{ sheet: string; visual: VisualBlockDiff }> = [];
  for (const s of pagesOrdered) {
    for (const v of s.visuals.slice().sort((x, y) => SEV_RANK[VISUAL_SEVERITY[x.status]] - SEV_RANK[VISUAL_SEVERITY[y.status]])) {
      if (!visualDataMatches(v.status)) problemVisuals.push({ sheet: s.sheet, visual: v });
    }
  }
  const detailFirstRow = new Map<VisualBlockDiff, { sheetName: string; row: number }>();

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

  // ── Sheet 5: Changed Values — flat, filterable list of every changed cell ─
  const changedVisuals = problemVisuals.filter(({ visual }) => (visual.changedRows?.length ?? 0) > 0);
  if (changedVisuals.length > 0) {
    const cWs = wb.addWorksheet('Changed Values');
    cWs.columns = [
      { header: 'Page (Sheet)',   key: 'sheet',  width: 22 },
      { header: 'Visual',         key: 'visual', width: 30 },
      { header: 'Row (key)',      key: 'key',    width: 38 },
      { header: 'Column',         key: 'column', width: 26 },
      { header: 'Original value', key: 'exp',    width: 16 },
      { header: 'Changed value',  key: 'act',    width: 16 },
      { header: 'Change (Δ)',     key: 'delta',  width: 12 },
      { header: 'Change %',       key: 'pct',    width: 10 },
    ];
    styleHeader(cWs, cWs.columns.length);
    for (const { sheet, visual } of changedVisuals) {
      const label = `${visual.title} [${visual.type}]`;
      for (const cr of visual.changedRows ?? []) {
        for (const ch of cr.changes) {
          const e = num(ch.expected), a = num(ch.actual);
          const row = cWs.addRow({ sheet, visual: label, key: cr.key, column: ch.column, exp: ch.expected, act: ch.actual });
          row.eachCell(c => { c.border = E_CB; c.alignment = { vertical: 'middle' }; });
          row.getCell('exp').fill = E_MISS;
          row.getCell('act').fill = E_FAIL_LT;
          row.getCell('act').font = { bold: true };
          if (e !== null && a !== null) {
            row.getCell('delta').value = a - e;
            row.getCell('delta').numFmt = (ch.expected.trim().endsWith('%') && ch.actual.trim().endsWith('%')) ? '+0.##" pp";-0.##" pp";0' : '+0.##;-0.##;0';
            if (e !== 0) { row.getCell('pct').value = (a - e) / Math.abs(e); row.getCell('pct').numFmt = '+0.0%;-0.0%;0.0%'; }
            else row.getCell('pct').value = 'n/a';
          } else {
            row.getCell('delta').value = 'text';
          }
          row.getCell('delta').alignment = { horizontal: 'right', vertical: 'middle' };
          row.getCell('pct').alignment = { horizontal: 'right', vertical: 'middle' };
        }
      }
    }
  }

  // ── Hyperlinks (set last, once every target row number is known) ──────────
  for (const [pageName, cell] of pageCells) {
    const r = firstVisRowOfPage.get(pageName);
    if (r !== undefined) link(cell, pageName, 'Visual Comparison', r);
  }
  for (const [v, row] of visCells) {
    const d = detailFirstRow.get(v);
    if (d) link(row.getCell('details'), 'See differences →', d.sheetName, d.row);
  }

  wb.eachSheet(ws => { ws.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 }; });
  fs.mkdirSync(require('path').dirname(outPath), { recursive: true });
  await wb.xlsx.writeFile(outPath);
  return { passed, differingSheets: differing };
}