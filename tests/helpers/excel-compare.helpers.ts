/**
 * excel-compare.helpers.ts
 *
 * Streaming Excel workbook comparison, in three passes:
 *
 * Tier 1 — per-sheet ordered hash (fast path).
 *   Stream both files' matching sheet, feeding each row into a running SHA-1.
 *   If the two hashes match → sheet is byte-identical → done for that sheet.
 *   A mismatch at Tier 1 proves nothing (could be mere row reordering) so always fall
 *   through to Tier 2 on any mismatch.
 *
 * Tier 2 — visual-block-aware diff (the detailed pass).
 *   report-export.helpers.ts writes each Power BI visual as a self-contained block:
 *   a "[VISUAL] <title> [<type>]" marker row, a header row, its data rows, and a
 *   blank spacer. This tier re-groups the row stream back into those blocks and
 *   compares VISUAL TO VISUAL rather than treating the sheet as one flat bag of
 *   rows. That buys three things a flat row-diff cannot:
 *     - A table that simply moved to a different position/row range in the sheet
 *       is still matched correctly, because matching is by visual identity
 *       (title + type), not row position. A flat row-multiset already tolerated
 *       *row* reordering; this additionally tolerates *visual* reordering and
 *       keeps row matches scoped to the correct visual instead of letting two
 *       coincidentally-identical rows from unrelated visuals cancel out.
 *     - Header-text differences (e.g. an Import-mode friendly column name vs a
 *       Direct-Lake raw column name) are reported separately from genuine value
 *       differences, instead of being blended into one undifferentiated pile of
 *       "only in expected / only in actual" rows.
 *     - The same visual title appearing a different number of times on each side
 *       (duplicate slicers, a visual that only exists on one side, two distinctly
 *       different visuals that happen to share a title) is called out explicitly
 *       instead of silently fragmenting into dozens of unexplained row diffs.
 *
 * `identical` (per sheet and per workbook) stays STRICT: every visual must match
 * on title, type, count, headers AND data. `dataIdentical` is the relaxed sibling:
 * true when every visual's underlying VALUES match, even if headers were renamed,
 * columns were reordered, or a visual is duplicated identically on one side.
 * Callers that want "position/renaming-tolerant" pass/fail should gate on
 * `dataIdentical`; callers that want a byte-for-byte check should use `identical`.
 *
 * Reading strategy: BUFFERED (Workbook.xlsx.readFile), not ExcelJS's streaming
 * reader. That was the original design (a streaming reader touches less
 * memory at once), but ExcelJS's streaming WorkbookReader was found to
 * deadlock indefinitely — never even emitting its first worksheet — on a
 * real, valid export written by this project's own streaming WRITER (~4MB,
 * ~4000 rows across 2 sheets; reproduced deterministically, in isolation,
 * with no pre-seeding and no concurrency involved, so it is not the same
 * zip-entry-order issue readWorkbookSheetInfo works around below - that one
 * corrupts sheet NAMES; this one hangs the reader outright). The buffered
 * reader parses the exact same file correctly in well under a second and,
 * as a bonus, is read ONCE per file and cached for the whole comparison
 * instead of re-opening the zip stream from scratch for every sheet - so
 * this is also strictly faster, not just safer. A 4MB report export fits in
 * memory without concern; this trade-off would need revisiting only for
 * exports far larger than anything this pipeline currently produces.
 */

import * as crypto from 'crypto';
import * as path   from 'path';
import ExcelJS     from 'exceljs';

// Central-directory zip access (Open.file reads only the zip's directory, then
// inflates single entries on demand — no whole-file buffering). Already a
// transitive dependency of exceljs; declared as a direct dependency in
// package.json so an exceljs upgrade can never silently remove it.
// eslint-disable-next-line @typescript-eslint/no-var-requires
const unzipper = require('unzipper');

import { decodeXmlEntities, readSheetRowsRaw, normalizeNumericDrift } from './xlsx-sheet-reader.helpers';

// ── Deterministic sheet-name resolution ──────────────────────────────────────
//
// ExcelJS's streaming READER resolves a worksheet's display name by looking it
// up in workbook.xml — but only if workbook.xml happened to appear EARLIER in
// the zip than the worksheet entry. Files produced by ExcelJS's streaming
// WRITER (i.e. our own exports) put the worksheets FIRST and workbook.xml
// last, and the relative order of workbook.xml vs its .rels file is not even
// deterministic across runs. Depending on that order the reader either:
//   - crashes ("Cannot read properties of undefined (reading 'sheets')" in
//     _parseWorksheet — rels parsed, workbook.xml not yet), or
//   - silently falls back to default names ("Sheet1", "Sheet2", ...), which
//     then destroys sheet pairing between the two workbooks being compared.
// Both were observed in production on a machine validating two real reports.
//
// Fix: read the sheet list ourselves, directly from workbook.xml (+ its rels
// for the sheetN.xml file mapping) via the zip central directory — fully
// order-independent — and PRE-SEED the ExcelJS reader's internal model before
// iterating, so its lookup always succeeds and its crash branch is
// unreachable. Belt-and-braces: sheet matching below also accepts the
// file-number id, and a sheet that is never emitted at all now throws instead
// of silently producing an empty result.

/** One sheet as resolved from workbook.xml: display name + sheetN.xml number. */
export interface SheetInfo {
  name:    string;  // real display name (page name), exactly as written
  fileNo:  string;  // N in xl/worksheets/sheetN.xml
  rId:     string;  // relationship id linking workbook.xml to the file
  sheetId: string;  // workbook.xml sheetId attribute
}

/** Reads the ordered sheet list from an xlsx via its central directory. */
export async function readWorkbookSheetInfo(filePath: string): Promise<SheetInfo[]> {
  const dir = await unzipper.Open.file(filePath);
  const entryFor = (p: string) => dir.files.find((f: any) => f.path === p);

  const wbEntry = entryFor('xl/workbook.xml');
  if (!wbEntry) throw new Error(`Not a valid xlsx (no xl/workbook.xml): ${filePath}`);
  const wbXml = (await wbEntry.buffer()).toString('utf8');

  // rId → sheet file number, from the workbook rels (may be absent in odd files)
  const relTargets = new Map<string, string>();
  const relsEntry = entryFor('xl/_rels/workbook.xml.rels');
  if (relsEntry) {
    const relsXml = (await relsEntry.buffer()).toString('utf8');
    for (const rel of relsXml.match(/<Relationship\b[^>]*>/g) ?? []) {
      const id     = rel.match(/\bId="([^"]+)"/)?.[1];
      const target = rel.match(/\bTarget="[^"]*worksheets\/sheet(\d+)\.xml"/)?.[1];
      if (id && target) relTargets.set(id, target);
    }
  }

  const sheets: SheetInfo[] = [];
  for (const tag of wbXml.match(/<sheet\b[^>]*>/g) ?? []) {
    const name    = tag.match(/\bname="([^"]*)"/)?.[1];
    const rId     = tag.match(/\br:id="([^"]+)"/)?.[1] ?? '';
    const sheetId = tag.match(/\bsheetId="([^"]+)"/)?.[1] ?? '';
    if (name === undefined) continue;
    // Fall back to declaration order when rels are missing — matches how the
    // ExcelJS writer numbers its sheet files.
    const fileNo = (rId && relTargets.get(rId)) || String(sheets.length + 1);
    sheets.push({ name: decodeXmlEntities(name), fileNo, rId, sheetId });
  }
  if (sheets.length === 0) throw new Error(`No sheets declared in workbook.xml: ${filePath}`);
  return sheets;
}

// ── Buffered workbook access (cached per file for the whole comparison) ──────

/**
 * Loads a workbook fully into memory via ExcelJS's buffered reader (NOT the
 * streaming one — see the module comment for why) and caches the result per
 * file path so multiple sheets from the same file only pay the read/parse
 * cost once. `cache` is expected to be fresh per compareWorkbooks() call —
 * never module-level — so results can never leak stale content across
 * independent calls that happen to reuse a file path (e.g. tests writing to
 * the same temp file more than once).
 *
 * This is now the FALLBACK path only — see getSheetRows() below, which tries
 * the lightweight per-sheet reader (xlsx-sheet-reader.helpers.ts) first and
 * only reaches this whole-workbook load when that reader can't confidently
 * handle the file.
 */
interface CompareCache {
  workbooks: Map<string, Promise<ExcelJS.Workbook>>;
  /** Files where the lightweight per-sheet reader hit something it doesn't
   *  recognize — once true for a file, every further sheet from it goes
   *  straight to the buffered path. There's no memory benefit to retrying
   *  the fast path sheet-by-sheet once a file needs buffered mode: the
   *  buffered reader loads every sheet in one shot regardless, so the "whole
   *  workbook resident" cost is already paid the first time it's needed. */
  fastPathFailed: Set<string>;
}

function getBufferedWorkbook(filePath: string, cache: CompareCache): Promise<ExcelJS.Workbook> {
  let p = cache.workbooks.get(filePath);
  if (!p) {
    p = (async () => {
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.readFile(filePath);
      return wb;
    })();
    cache.workbooks.set(filePath, p);
  }
  return p;
}

/**
 * Returns one sheet's rows, normalized exactly like normalizeRow(ws.eachRow
 * row) below — trying the lightweight zip/XML reader first (bounded to
 * roughly one sheet's worth of memory) and falling back to the buffered
 * ExcelJS workbook (the previous, whole-file behavior) only when the fast
 * reader reports it can't confidently handle this file.
 */
async function getSheetRows(filePath: string, target: SheetInfo, cache: CompareCache): Promise<string[][]> {
  if (!cache.fastPathFailed.has(filePath)) {
    const fast = await readSheetRowsRaw(filePath, target);
    if (fast !== null) return fast;
    cache.fastPathFailed.add(filePath);
  }

  const wb = await getBufferedWorkbook(filePath, cache);
  const ws = findWorksheet(wb, target);
  if (!ws) throw sheetNotEmitted(filePath, target);
  const rows: string[][] = [];
  ws.eachRow({ includeEmpty: true }, (row) => { rows.push(normalizeRow(row)); });
  return rows;
}

/**
 * Finds the worksheet described by `target` inside an already-loaded
 * (buffered) workbook. The buffered reader resolves worksheet names
 * directly from workbook.xml with full random access, so — unlike the old
 * streaming path — it never falls back to a default "Sheet1"/"Sheet2" name
 * and this lookup can simply match by name; the sheetId fallback below is
 * belt-and-braces for the (unobserved, in buffered mode) case of a name
 * mismatch, e.g. from a future exceljs version.
 */
function findWorksheet(wb: ExcelJS.Workbook, target: SheetInfo): ExcelJS.Worksheet | undefined {
  return wb.worksheets.find(ws => ws.name === target.name)
    ?? wb.worksheets.find(ws => String(ws.id) === String(target.sheetId) || String(ws.id) === String(target.fileNo));
}

// Built from explicit char codes (never a literal combining-mark character in
// this file) - Unicode combining diacritical marks, U+0300 to U+036F. Mirrors
// the same range used by pageNameKey in report-export.helpers.ts; kept as its
// own local copy rather than imported, to avoid a circular module import
// (that file already imports types from this one).
const COMBINING_DIACRITICS_RE = new RegExp(
  '[' + String.fromCharCode(0x0300) + '-' + String.fromCharCode(0x036f) + ']', 'g',
);

/** Canonical pairing key: sheet names differing only by case/edge-whitespace
 *  (e.g. "…- Cluster" vs "…- cluster" between a source and a migrated report)
 *  or accents (e.g. "Commodité" vs "Commodite" - observed in production,
 *  the exact same page lost its accent during migration) refer to the same
 *  page and must pair up. */
function sheetKey(name: string): string {
  return name.normalize('NFD').replace(COMBINING_DIACRITICS_RE, '').trim().toLowerCase();
}

// ── Public types ─────────────────────────────────────────────────────────────

export type VisualDiffStatus =
  | 'identical'        // headers + data + counts all match
  | 'header-diff'      // data matches; header text and/or column order differs
  | 'duplicate-count'  // headers + data match; this visual appears a different
                        // number of times on each side (all copies mutually identical)
  | 'data-diff'        // matched 1:1 but the underlying values differ
  | 'only-in-expected' // visual (by title+type) exists only in the source export
  | 'only-in-actual'   // visual (by title+type) exists only in the target export
  | 'ambiguous';        // more than one distinct variant on at least one side —
                        // could not safely auto-pair, needs a human look

export interface VisualBlockDiff {
  title:          string;
  type:           string;
  status:         VisualDiffStatus;
  /** How many blocks with this exact (title, type) were found on each side. */
  countExpected:  number;
  countActual:    number;
  rowsExpected:   number;
  rowsActual:     number;
  headerExpected: string[];
  headerActual:   string[];
  /** Up to SAMPLE_SIZE data rows found only in the expected copy of this visual. */
  sampleOnlyInExpected: string[][];
  /** Up to SAMPLE_SIZE data rows found only in the actual copy of this visual. */
  sampleOnlyInActual:   string[][];
  /** True total (not capped by SAMPLE_SIZE) of differing rows, for data-diff visuals. */
  onlyInExpectedCount:  number;
  onlyInActualCount:    number;
  /** Rows paired by their text columns whose VALUES differ, with the columns
   *  that changed (capped at SAMPLE_SIZE). Only set for data-diff visuals. */
  changedRows?:         ChangedRow[];
  /** True number of paired changed rows (not capped). */
  changedRowCount?:     number;
  /** Column names that the row samples and changedRows are laid out in. */
  diffColumns?:         string[];
  /** Columns that exist on only one side (extra / missing / renamed columns). */
  columnsOnlyInExpected?: string[];
  columnsOnlyInActual?:   string[];
  /** Human-readable explanation — why this status, what to go check. */
  note?: string;
}

/** One row that exists on both sides (same text key) but with different values. */
export interface ChangedRow {
  /** The row's text columns, joined — e.g. "Finance | 2026-03". */
  key: string;
  changes: Array<{ column: string; expected: string; actual: string }>;
  /** The full source / target rows (same columns as `VisualBlockDiff.diffColumns`). */
  expectedRow?: string[];
  actualRow?: string[];
}

export interface SheetDiffSummary {
  sheet:         string;
  inExpected:    boolean;
  inActual:      boolean;
  rowsExpected:  number;
  rowsActual:    number;
  /** STRICT: every visual matches on title, type, count, headers AND data. */
  identical:     boolean;
  /** RELAXED: every visual's underlying data matches (headers/order/duplicate
   *  count differences do not count against this). Use this to gate pass/fail
   *  when you only care that the migrated report shows the same numbers. */
  dataIdentical: boolean;
  /** Per-visual breakdown. Empty when the sheet was identical at Tier 1 (no
   *  detailed pass was needed) or when neither export has any visuals. */
  visuals: VisualBlockDiff[];
  /** Visuals whose header text/column order differs but whose data matches. */
  headerDiffCount: number;
  /** Visuals that are missing on one side, duplicated a different number of
   *  times, or otherwise couldn't be unambiguously matched. */
  structuralDiffCount: number;
  /** Approximate, human-scale row-diff counts for at-a-glance display — NOT
   *  used to determine identical/dataIdentical (those are computed from
   *  `visuals` directly). Kept for backward compatibility with existing
   *  consumers of this shape (main-run.spec.ts, excel-compare.spec.ts). */
  rowsOnlyInExpected: number;
  rowsOnlyInActual:   number;
  /** Up to SAMPLE_SIZE rows only in the expected file (human-readable) */
  sampleOnlyInExpected: string[][];
  /** Up to SAMPLE_SIZE rows only in the actual file (human-readable) */
  sampleOnlyInActual:   string[][];
}

export interface WorkbookDiffSummary {
  fileExpected:     string;
  fileActual:       string;
  runAt:            string;
  sheetsOnlyInExpected: string[];
  sheetsOnlyInActual:   string[];
  sheets:           SheetDiffSummary[];
  /** STRICT — every sheet identical, no sheets missing on either side. */
  identical:        boolean;
  /** RELAXED — every sheet's data matches, no sheets missing on either side.
   *  This is what report-parity gates PASS/FAIL on by default. */
  dataIdentical:    boolean;
}

// ── Internal constants ────────────────────────────────────────────────────────

const SAMPLE_SIZE = 800; // max sample rows to keep per differing visual's VALUE diff
// A visual that's entirely missing on one side (or one of several ambiguous
// variants) doesn't need hundreds of rows reproduced in the Differences
// sheet to make the point — the Visual Comparison sheet already has its full
// row count. A short excerpt is plenty and keeps the sheet skimmable.
const MISSING_VISUAL_SAMPLE_CAP = 20;

// ── Cell normalisation ────────────────────────────────────────────────────────

// normalizeNumericDrift lives in xlsx-sheet-reader.helpers.ts now — the fast
// per-sheet reader needs the exact same rounding applied to the raw text it
// extracts as normalizeCell applies below, so ExcelJS-buffered and fast-path
// reads of the same cell always normalize identically. One implementation,
// imported here rather than duplicated.
const LENIENT_TEXT_COMPARE = (process.env['PBI_COMPARE_TEXT_LENIENT'] ?? '').trim() === '1';

/**
 * Canonical form used for equality/hash checks only.
 * When lenient mode is on, case and common separator differences are ignored.
 *
 * Deliberately .toLowerCase(), not .toLocaleLowerCase(): the latter follows
 * the RUNNING MACHINE's default locale, which a pass/fail verdict must
 * never depend on - the same two reports must compare identically whether
 * this runs from a machine set to French, Turkish, or anything else (the
 * Turkish "I" is the textbook case where locale-aware casing diverges from
 * every other locale).
 */
function canonicalizeForCompare(text: string): string {
  if (!LENIENT_TEXT_COMPARE) return text;
  return text
    .toLowerCase()
    .replace(/[_-]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Converts a raw ExcelJS cell value to a consistent string so that
 * formatting-only quirks (extra spaces, number vs string, float drift) don't
 * produce false mismatches.
 */
function normalizeCell(value: ExcelJS.CellValue): string {
  let raw: string;
  if (value === null || value === undefined) {
    raw = '';
  } else if (value instanceof Date) {
    raw = value.toISOString();
  } else if (typeof value === 'object') {
    // RichText, hyperlink, formula result etc.
    const v = value as any;
    if (v.result !== undefined) raw = String(v.result).trim();
    else if (v.text !== undefined) raw = String(v.text).trim();
    else raw = JSON.stringify(value);
  } else {
    raw = String(value).trim();
  }
  return normalizeNumericDrift(raw);
}

/**
 * Normalises a row's cell value array into a plain string array.
 * row.values is 1-indexed in ExcelJS (index 0 is undefined). Exported for
 * direct unit testing of the sparse-row handling (see the comment inside).
 */
export function normalizeRow(row: ExcelJS.Row): string[] {
  const vals = row.values as ExcelJS.CellValue[];
  // Array.from (not a direct .slice/.map on `vals`) is required here:
  // ExcelJS's row.values is column-position-indexed and can be SPARSE - a
  // row with no cell at all in some column leaves a genuine hole, not an
  // explicit undefined. Array.prototype.map SKIPS holes (never calls its
  // callback for them), so normalizeCell never ran for that position and
  // the hole passed straight through as `undefined` instead of becoming
  // '' like every other blank cell - crashing the first thing downstream
  // that assumed every cell is always a string (observed in production on
  // a 24k-row table, the first row large enough to expose the gap).
  // Array.from reads by index up to .length, which materializes an
  // explicit `undefined` at each hole - normalizeCell already handles
  // that correctly, it just never got the chance to run before.
  return Array.from(vals || []).slice(1).map(normalizeCell);
}

// ── Hashing ───────────────────────────────────────────────────────────────────

function hashRow(cells: string[]): string {
  // Trailing blank cells carry no data (a whitespace-only cell on one side
  // only makes that row's array longer), so they must not change the hash.
  let end = cells.length;
  while (end > 0 && canonicalizeForCompare(cells[end - 1]) === '') end--;
  return crypto.createHash('sha1').update(JSON.stringify(cells.slice(0, end).map(canonicalizeForCompare))).digest('hex');
}

function hashRunning(existing: string, rowHash: string): string {
  return crypto.createHash('sha1').update(existing + rowHash).digest('hex');
}

// ── Tier 1: whole-sheet ordered hash ──────────────────────────────────────────

interface SheetHashResult {
  orderedHash: string;
  rowCount:    number;
}

/**
 * One retry for a failed read (transient I/O hiccup on the underlying file).
 * A sheet not being found now THROWS (see sheetNotEmitted) instead of
 * silently returning an empty result that would then be "compared" as if it
 * were real data.
 */
async function withSheetRetry<T>(filePath: string, target: SheetInfo, read: () => Promise<T>): Promise<T> {
  try {
    return await read();
  } catch (e) {
    console.warn(`[excel-compare] Re-reading "${target.name}" from ${path.basename(filePath)} after a failed read: ${(e as Error).message}`);
    return await read();
  }
}

function sheetNotEmitted(filePath: string, target: SheetInfo): Error {
  return new Error(
    `Sheet "${target.name}" (worksheets/sheet${target.fileNo}.xml) exists in ${path.basename(filePath)}'s ` +
    `workbook.xml but was not found when the file was loaded — the file may be corrupted or the read ` +
    `was interrupted. Refusing to treat it as empty.`,
  );
}

async function hashSheetOrdered(
  filePath: string, target: SheetInfo, cache: CompareCache,
): Promise<SheetHashResult> {
  return withSheetRetry(filePath, target, async () => {
    const rows = await getSheetRows(filePath, target, cache);

    let orderedHash = '';
    let rowCount    = 0;
    for (const cells of rows) {
      if (cells.every(c => c === '')) continue; // skip fully-blank spacer rows
      rowCount++;
      orderedHash = hashRunning(orderedHash, hashRow(cells));
    }
    return { orderedHash, rowCount };
  });
}

// ── Tier 2: visual-block parsing ──────────────────────────────────────────────
//
// report-export.helpers.ts (writeVisualsToSheet) writes each visual as:
//   [VISUAL] <title> [<type>]   ← marker row, column A only
//   <header1> <header2> ...     ← header row
//   <data rows...>
//   (blank spacer row)
// This re-groups the flat row stream back into those blocks. The regex is
// greedy on the title so it always splits on the LAST "[...]" in the line,
// which is where writeVisualsToSheet always puts the type — safe even if a
// visual's title itself happens to contain brackets.
const VISUAL_MARKER_RE = /^\[VISUAL\]\s+(.+)\s\[([^[\]]+)\]$/;

interface VisualBlockRaw {
  title:   string;
  type:    string;
  headers: string[];
  rows:    string[][];
}

async function streamSheetBlocks(
  filePath: string, target: SheetInfo, cache: CompareCache,
): Promise<VisualBlockRaw[]> {
  return withSheetRetry(filePath, target, async () => {
    const rows = await getSheetRows(filePath, target, cache);

    const blocks: VisualBlockRaw[] = [];
    let current: VisualBlockRaw | null = null;
    let sawHeaderForCurrent = false;

    for (const cells of rows) {
      const isBlank = cells.every(c => c === '');
      const marker  = !isBlank ? cells[0].match(VISUAL_MARKER_RE) : null;

      if (marker) {
        if (current) blocks.push(current);
        current = { title: marker[1].trim(), type: marker[2].trim(), headers: [], rows: [] };
        sawHeaderForCurrent = false;
        continue;
      }
      if (isBlank) continue;     // spacer row between visuals
      if (!current) continue;    // defensive: content before any marker row

      if (!sawHeaderForCurrent) {
        current.headers = cells;
        sawHeaderForCurrent = true;
      } else {
        current.rows.push(cells);
      }
    }
    if (current) blocks.push(current);
    return blocks;
  });
}

// ── Tier 2: block matching + comparison ───────────────────────────────────────

function headerKey(headers: string[]): string {
  return headers.map(canonicalizeForCompare).join('\u0001');
}

/** Order-independent content signature for a block's data rows. */
function blockDataSignature(rows: string[][]): string {
  const hashes = rows.map(hashRow).sort();
  return crypto.createHash('sha1').update(hashes.join('|')).digest('hex');
}

/**
 * Collapses blocks that are byte-for-byte identical to each other (same
 * headers, same data, any row order) into one representative + a count.
 * This is what lets a slicer that was innocently duplicated on the report
 * canvas — two visual objects, same title, same bound field, same values —
 * be treated as "1 visual, ×2 copies" instead of contaminating the diff.
 */
function dedupeBlocks(blocks: VisualBlockRaw[]): { representative: VisualBlockRaw; count: number }[] {
  const groups: { key: string; representative: VisualBlockRaw; count: number }[] = [];
  for (const b of blocks) {
    const key = headerKey(b.headers) + '::' + blockDataSignature(b.rows);
    const existing = groups.find(g => g.key === key);
    if (existing) existing.count++;
    else groups.push({ key, representative: b, count: 1 });
  }
  return groups.map(({ representative, count }) => ({ representative, count }));
}

/**
 * If `headers` and `targetOrder` contain the exact same column names (just
 * shuffled), remaps every row so column i lines up with targetOrder[i].
 * Returns null when the header sets genuinely differ (renamed/added/removed
 * columns) — in that case realignment isn't meaningful and headers are
 * compared positionally instead.
 */
function realignByHeader(headers: string[], rows: string[][], targetOrder: string[]): string[][] | null {
  if (headers.length !== targetOrder.length) return null;
  const sortedA = headers.map(canonicalizeForCompare).sort().join('\u0001');
  const sortedB = targetOrder.map(canonicalizeForCompare).sort().join('\u0001');
  if (sortedA !== sortedB) return null;

  const indexOf = new Map<string, number>();
  for (let i = 0; i < headers.length; i++) {
    const key = canonicalizeForCompare(headers[i]);
    if (indexOf.has(key)) {
      // Ambiguous (two headers collapse to same canonical key); keep positional compare.
      return null;
    }
    indexOf.set(key, i);
  }
  return rows.map(r => targetOrder.map(h => r[indexOf.get(canonicalizeForCompare(h))!] ?? ''));
}

/** A numeric-looking value (plain number, thousands separators, optional %). */
const NUMERIC_LIKE_RE = /^-?\d[\d,]*(\.\d+)?%?$/;

/**
 * Pairs rows that exist on only one side by their "key" — the text-like
 * columns (names, dates, labels) — and reports which COLUMNS changed. A
 * changed value used to show up as two unrelated rows (one in each pile);
 * pairing them tells the reader "row X: column Y was 10, now 12".
 *
 * Only unambiguous pairs are made: a key must occur exactly once among the
 * unmatched source rows and exactly once among the unmatched target rows.
 * Anything else stays in the plain only-in-source / only-in-target piles.
 */
function pairChangedRows(
  leftExp: string[][], leftAct: string[][], colNames: string[],
): { changed: ChangedRow[]; restExp: string[][]; restAct: string[][] } {
  if (leftExp.length === 0 || leftAct.length === 0) {
    return { changed: [], restExp: leftExp, restAct: leftAct };
  }
  const ncol = Math.max(colNames.length, ...leftExp.map(r => r.length), ...leftAct.map(r => r.length));
  const keyCols: number[] = [];
  for (let c = 0; c < ncol; c++) {
    let nonBlank = 0, numeric = 0;
    for (const rows of [leftExp, leftAct]) {
      for (const r of rows) {
        const v = r[c] ?? '';
        if (v === '') continue;
        nonBlank++;
        if (NUMERIC_LIKE_RE.test(v)) numeric++;
      }
    }
    if (nonBlank > 0 && numeric / nonBlank < 0.5) keyCols.push(c);
  }
  if (keyCols.length === 0) return { changed: [], restExp: leftExp, restAct: leftAct };

  const keyOf = (r: string[]) => keyCols.map(c => canonicalizeForCompare(r[c] ?? '')).join('\u0001');
  const group = (rows: string[][]) => {
    const m = new Map<string, number[]>();
    rows.forEach((r, i) => {
      const k = keyOf(r);
      const l = m.get(k);
      if (l) l.push(i); else m.set(k, [i]);
    });
    return m;
  };
  const expGroups = group(leftExp);
  const actGroups = group(leftAct);

  const pairedExp = new Set<number>();
  const pairedAct = new Set<number>();
  const changed: ChangedRow[] = [];
  for (const [k, eIdx] of expGroups) {
    const aIdx = actGroups.get(k);
    if (!aIdx || eIdx.length !== 1 || aIdx.length !== 1) continue;
    const e = leftExp[eIdx[0]];
    const a = leftAct[aIdx[0]];
    const changes: ChangedRow['changes'] = [];
    for (let c = 0; c < ncol; c++) {
      const ev = e[c] ?? '';
      const av = a[c] ?? '';
      if (canonicalizeForCompare(ev) !== canonicalizeForCompare(av)) {
        changes.push({ column: colNames[c] ?? `Column ${c + 1}`, expected: ev, actual: av });
      }
    }
    if (changes.length === 0) continue;
    pairedExp.add(eIdx[0]);
    pairedAct.add(aIdx[0]);
    changed.push({ key: keyCols.map(c => e[c] ?? '').filter(v => v !== '').join(' | '), changes, expectedRow: e, actualRow: a });
  }
  return {
    changed,
    restExp: leftExp.filter((_, i) => !pairedExp.has(i)),
    restAct: leftAct.filter((_, i) => !pairedAct.has(i)),
  };
}

/** Order-independent row diff between two row sets, with true counts + capped samples. */
function diffRows(rowsExp: string[][], rowsAct: string[][], colNames: string[] = []) {
  const expMap = new Map<string, { count: number; sample: string[] }>();
  for (const r of rowsExp) {
    const h = hashRow(r);
    const e = expMap.get(h);
    if (e) e.count++; else expMap.set(h, { count: 1, sample: r });
  }
  const actMap = new Map<string, { count: number; sample: string[] }>();
  for (const r of rowsAct) {
    const h = hashRow(r);
    const e = actMap.get(h);
    if (e) e.count++; else actMap.set(h, { count: 1, sample: r });
  }

  let onlyInExpectedCount = 0;
  let onlyInActualCount   = 0;
  const leftExp: string[][] = [];
  const leftAct: string[][] = [];

  const allHashes = new Set([...expMap.keys(), ...actMap.keys()]);
  for (const h of allHashes) {
    const ce = expMap.get(h)?.count ?? 0;
    const ca = actMap.get(h)?.count ?? 0;
    const matched = Math.min(ce, ca);
    onlyInExpectedCount += ce - matched;
    onlyInActualCount   += ca - matched;
    for (let i = 0; i < ce - matched; i++) leftExp.push(expMap.get(h)!.sample);
    for (let i = 0; i < ca - matched; i++) leftAct.push(actMap.get(h)!.sample);
  }

  // Pair rows whose text columns match but whose values differ, so a changed
  // number reads as "column X: 10 -> 12" instead of two unrelated rows.
  const paired = pairChangedRows(leftExp, leftAct, colNames);
  const bySortKey = (a: string[], b: string[]) => a.join('\u0001') < b.join('\u0001') ? -1 : 1;
  const sampleOnlyInExpected = paired.restExp.slice().sort(bySortKey).slice(0, SAMPLE_SIZE);
  const sampleOnlyInActual   = paired.restAct.slice().sort(bySortKey).slice(0, SAMPLE_SIZE);

  return {
    onlyInExpectedCount, onlyInActualCount, sampleOnlyInExpected, sampleOnlyInActual,
    changedRows: paired.changed.slice(0, SAMPLE_SIZE),
    changedRowCount: paired.changed.length,
  };
}

/**
 * Columns present on both sides (matched by canonical header text; headers
 * that repeat on a side are left out because they cannot be matched safely).
 */
function sharedColumns(expHeaders: string[], actHeaders: string[]): Array<{ name: string; e: number; a: number }> {
  const uniqueIndex = (hs: string[]) => {
    const seen = new Map<string, number>();
    const dup = new Set<string>();
    hs.forEach((h, i) => {
      const k = canonicalizeForCompare(h);
      if (seen.has(k)) dup.add(k); else seen.set(k, i);
    });
    for (const k of dup) seen.delete(k);
    return seen;
  };
  const eIdx = uniqueIndex(expHeaders);
  const aIdx = uniqueIndex(actHeaders);
  const out: Array<{ name: string; e: number; a: number }> = [];
  for (const [k, e] of eIdx) {
    const a = aIdx.get(k);
    if (a !== undefined) out.push({ name: expHeaders[e], e, a });
  }
  return out.sort((x, y) => x.e - y.e);
}

/** Compares one matched (title, type) pair of representative blocks. */
function compareBlockPair(
  exp: VisualBlockRaw, expCount: number,
  act: VisualBlockRaw, actCount: number,
): VisualBlockDiff {
  const headersEqual = headerKey(exp.headers) === headerKey(act.headers);

  let actRows = act.rows;
  let headerNote: string | undefined;
  if (!headersEqual) {
    const realigned = realignByHeader(act.headers, act.rows, exp.headers);
    if (realigned) {
      actRows = realigned;
      headerNote = 'Same columns on both sides, different order — realigned by column name before comparing values.';
    } else {
      headerNote = 'Column headers differ in text between source and target.';
    }
  }

  const expSig  = blockDataSignature(exp.rows);
  const actSig  = blockDataSignature(actRows);
  const dataEqual = expSig === actSig;

  let onlyInExpectedCount = 0;
  let onlyInActualCount   = 0;
  let sampleOnlyInExpected: string[][] = [];
  let sampleOnlyInActual:   string[][] = [];
  let changedRows: ChangedRow[] | undefined;
  let changedRowCount: number | undefined;
  let columnsOnlyInExpected: string[] | undefined;
  let columnsOnlyInActual:   string[] | undefined;
  let columnNote: string | undefined;
  let diffColumns: string[] | undefined;

  if (!dataEqual) {
    // Positional compare: when a column was renamed, label it with both names.
    const posNames = exp.headers.map((h, i) =>
      actRows === act.rows && act.headers[i] !== undefined && canonicalizeForCompare(act.headers[i]) !== canonicalizeForCompare(h)
        ? `${h} → ${act.headers[i]}` : h);
    let d = diffRows(exp.rows, actRows, posNames);
    diffColumns = posNames;

    // Column-aware pass: when the two sides have DIFFERENT column sets (an
    // extra, missing or renamed column), whole-row hashes differ for every
    // row and the real cause is invisible. Compare on the columns both sides
    // have, and name the columns that are not shared.
    if (!headersEqual && actRows === act.rows) {
      const shared = sharedColumns(exp.headers, act.headers);
      const sharedE = new Set(shared.map(s => s.e));
      const sharedA = new Set(shared.map(s => s.a));
      const onlyE = exp.headers.filter((_, i) => !sharedE.has(i));
      const onlyA = act.headers.filter((_, i) => !sharedA.has(i));
      if (shared.length > 0 && (onlyE.length > 0 || onlyA.length > 0)) {
        columnsOnlyInExpected = onlyE;
        columnsOnlyInActual   = onlyA;
        columnNote =
          (onlyE.length > 0 ? `Column(s) only in source: ${onlyE.join(', ')}. ` : '') +
          (onlyA.length > 0 ? `Column(s) only in target: ${onlyA.join(', ')}. ` : '');
        const projE = exp.rows.map(r => shared.map(s => r[s.e] ?? ''));
        const projA = act.rows.map(r => shared.map(s => r[s.a] ?? ''));
        if (blockDataSignature(projE) === blockDataSignature(projA)) {
          columnNote += `The ${shared.length} shared column(s) hold identical values; the difference is only in the columns listed above.`;
          // The rows themselves match on every shared column: no row-level
          // differences to report, only the column difference above.
          d = { onlyInExpectedCount: 0, onlyInActualCount: 0, sampleOnlyInExpected: [], sampleOnlyInActual: [], changedRows: [], changedRowCount: 0 };
          diffColumns = undefined;
        } else {
          d = diffRows(projE, projA, shared.map(s => s.name));
          diffColumns = shared.map(s => s.name);
          columnNote += `Row differences below are measured on the ${shared.length} shared column(s) only.`;
        }
      }
    }

    onlyInExpectedCount = d.onlyInExpectedCount;
    onlyInActualCount   = d.onlyInActualCount;
    sampleOnlyInExpected = d.sampleOnlyInExpected;
    sampleOnlyInActual   = d.sampleOnlyInActual;
    changedRows     = d.changedRows.length > 0 ? d.changedRows : undefined;
    changedRowCount = d.changedRowCount > 0 ? d.changedRowCount : undefined;
  }

  let status: VisualDiffStatus;
  if (headersEqual && dataEqual && expCount === actCount) status = 'identical';
  else if (headersEqual && dataEqual)                      status = 'duplicate-count';
  else if (dataEqual)                                       status = 'header-diff';
  else                                                       status = 'data-diff';

  const notes: string[] = [];
  if (headerNote) notes.push(headerNote);
  if (columnNote) notes.push(columnNote.trim());
  if (expCount !== actCount) {
    notes.push(`Source had ${expCount} identical cop${expCount === 1 ? 'y' : 'ies'} of this visual; target had ${actCount}.`);
  }

  return {
    title: exp.title, type: exp.type, status,
    countExpected: expCount, countActual: actCount,
    rowsExpected: exp.rows.length, rowsActual: act.rows.length,
    headerExpected: exp.headers, headerActual: act.headers,
    sampleOnlyInExpected, sampleOnlyInActual,
    onlyInExpectedCount, onlyInActualCount,
    changedRows, changedRowCount, diffColumns, columnsOnlyInExpected, columnsOnlyInActual,
    note: notes.length > 0 ? notes.join(' ') : undefined,
  };
}

// ── Pairing of visuals that cannot be matched by exact (title, type) ──────────
//
// Matching used to stop at "same title + same type". A visual renamed during
// migration then showed as one missing plus one added visual, and a title used
// by several different visuals was reported as "ambiguous" and never
// compared. pairByContent() pairs such leftovers by what they contain:
//   1. identical columns AND identical data      -> 'exact'
//   2. identical columns, data differs           -> 'headers' (closest row count)
//   3. only inside one title+type group: closest row count -> 'weak'
// Set PBI_COMPARE_NO_FUZZY_PAIRING=1 to switch this off and get the old
// strict behaviour.
const FUZZY_PAIRING = (process.env['PBI_COMPARE_NO_FUZZY_PAIRING'] ?? '').trim() !== '1';

type DedupItem = { representative: VisualBlockRaw; count: number };
type Cand = { title: string; type: string; item: DedupItem };
type PairStrength = 'exact' | 'headers' | 'weak';

function pairByContent(
  E: Cand[], A: Cand[], allowWeak: boolean,
): { pairs: Array<{ e: Cand; a: Cand; strength: PairStrength }>; restE: Cand[]; restA: Cand[] } {
  const eh = E.map(c => headerKey(c.item.representative.headers));
  const ah = A.map(c => headerKey(c.item.representative.headers));
  const es = E.map(c => blockDataSignature(c.item.representative.rows));
  const as = A.map(c => blockDataSignature(c.item.representative.rows));
  const usedE = new Set<number>();
  const usedA = new Set<number>();
  const pairs: Array<{ e: Cand; a: Cand; strength: PairStrength }> = [];
  const take = (i: number, j: number, strength: PairStrength) => {
    usedE.add(i); usedA.add(j); pairs.push({ e: E[i], a: A[j], strength });
  };
  const hasHeaders = (c: Cand) => c.item.representative.headers.length > 0;
  const sameType = (i: number, j: number) => canonicalizeForCompare(E[i].type) === canonicalizeForCompare(A[j].type);

  // 1. identical columns and data (prefer the same type)
  for (const preferSameType of [true, false]) {
    for (let i = 0; i < E.length; i++) {
      if (usedE.has(i) || !hasHeaders(E[i])) continue;
      for (let j = 0; j < A.length; j++) {
        if (usedA.has(j) || eh[i] !== ah[j] || es[i] !== as[j]) continue;
        if (preferSameType && !sameType(i, j)) continue;
        take(i, j, 'exact');
        break;
      }
    }
  }
  // 2. identical columns, data differs: best-first by how many row values the
  //    two visuals share, then by closest row count (prefer the same type).
  //    Global best-first, so one visual cannot grab another's better partner.
  const rowHashSets = new Map<VisualBlockRaw, Set<string>>();
  const hashesOf = (b: VisualBlockRaw) => {
    let h = rowHashSets.get(b);
    if (!h) { h = new Set(b.rows.map(hashRow)); rowHashSets.set(b, h); }
    return h;
  };
  for (const preferSameType of [true, false]) {
    const cands: Array<{ i: number; j: number; overlap: number; gap: number }> = [];
    for (let i = 0; i < E.length; i++) {
      if (usedE.has(i) || !hasHeaders(E[i])) continue;
      for (let j = 0; j < A.length; j++) {
        if (usedA.has(j) || eh[i] !== ah[j]) continue;
        if (preferSameType && !sameType(i, j)) continue;
        const eb = E[i].item.representative, ab = A[j].item.representative;
        const ha = hashesOf(ab);
        let overlap = 0;
        for (const h of hashesOf(eb)) if (ha.has(h)) overlap++;
        cands.push({ i, j, overlap, gap: Math.abs(eb.rows.length - ab.rows.length) });
      }
    }
    cands.sort((x, y) => y.overlap - x.overlap || x.gap - y.gap || x.i - y.i || x.j - y.j);
    for (const c of cands) {
      if (usedE.has(c.i) || usedA.has(c.j)) continue;
      take(c.i, c.j, 'headers');
    }
  }
  // 3. last resort inside one title+type group: closest row count
  if (allowWeak) {
    for (let i = 0; i < E.length; i++) {
      if (usedE.has(i)) continue;
      let best = -1;
      let bestGap = Infinity;
      for (let j = 0; j < A.length; j++) {
        if (usedA.has(j)) continue;
        const gap = Math.abs(E[i].item.representative.rows.length - A[j].item.representative.rows.length);
        if (gap < bestGap) { best = j; bestGap = gap; }
      }
      if (best >= 0) take(i, best, 'weak');
    }
  }
  return {
    pairs,
    restE: E.filter((_, i) => !usedE.has(i)),
    restA: A.filter((_, i) => !usedA.has(i)),
  };
}

/**
 * Matches every visual block from the expected export against every visual
 * block from the actual export, by (title, type) identity rather than row
 * position, and returns one verdict per visual. Visuals that cannot be
 * matched by title and type are then paired by content (see pairByContent).
 * See the module-level comment for why this is more accurate than a flat row diff.
 */
function matchVisualBlocks(expectedBlocks: VisualBlockRaw[], actualBlocks: VisualBlockRaw[]): VisualBlockDiff[] {
  const results: VisualBlockDiff[] = [];
  const leftE: Cand[] = [];
  const leftA: Cand[] = [];

  // Group by canonical title, then type — computed ONCE per block (the old
  // code re-canonicalised inside nested filters).
  const titleKeyToDisplay = new Map<string, string>();
  const expByTitle = new Map<string, VisualBlockRaw[]>();
  const actByTitle = new Map<string, VisualBlockRaw[]>();
  const addTo = (m: Map<string, VisualBlockRaw[]>, key: string, b: VisualBlockRaw) => {
    const l = m.get(key);
    if (l) l.push(b); else m.set(key, [b]);
    if (!titleKeyToDisplay.has(key)) titleKeyToDisplay.set(key, b.title);
  };
  expectedBlocks.forEach(b => addTo(expByTitle, canonicalizeForCompare(b.title), b));
  actualBlocks.forEach(b => addTo(actByTitle, canonicalizeForCompare(b.title), b));

  for (const [titleKey, title] of titleKeyToDisplay) {
    const expForTitle = expByTitle.get(titleKey) ?? [];
    const actForTitle = actByTitle.get(titleKey) ?? [];

    const typeKeyToDisplay = new Map<string, string>();
    const expByType = new Map<string, VisualBlockRaw[]>();
    const actByType = new Map<string, VisualBlockRaw[]>();
    const addType = (m: Map<string, VisualBlockRaw[]>, b: VisualBlockRaw) => {
      const key = canonicalizeForCompare(b.type);
      const l = m.get(key);
      if (l) l.push(b); else m.set(key, [b]);
      if (!typeKeyToDisplay.has(key)) typeKeyToDisplay.set(key, b.type);
    };
    expForTitle.forEach(b => addType(expByType, b));
    actForTitle.forEach(b => addType(actByType, b));

    for (const [typeKey, type] of typeKeyToDisplay) {
      const expList = expByType.get(typeKey) ?? [];
      const actList = actByType.get(typeKey) ?? [];

      if (expList.length === 0 || actList.length === 0) {
        // Present on one side only — hold back, it may be a renamed visual.
        for (const item of dedupeBlocks(expList)) leftE.push({ title, type, item });
        for (const item of dedupeBlocks(actList)) leftA.push({ title, type, item });
        continue;
      }

      // Both sides have at least one block for this (title, type). Dedupe
      // exact duplicates within each side first — that's what turns "2
      // identical copies vs 1" into a clean 1:1 comparison instead of noise.
      const expDedup = dedupeBlocks(expList);
      const actDedup = dedupeBlocks(actList);

      if (expDedup.length === 1 && actDedup.length === 1) {
        results.push(compareBlockPair(
          expDedup[0].representative, expDedup[0].count,
          actDedup[0].representative, actDedup[0].count,
        ));
        continue;
      }

      if (FUZZY_PAIRING) {
        // Several DISTINCT visuals share this title and type. Pair them by
        // content instead of giving up.
        const candE = expDedup.map(item => ({ title, type, item }));
        const candA = actDedup.map(item => ({ title, type, item }));
        const { pairs, restE, restA } = pairByContent(candE, candA, true);
        for (const p of pairs) {
          const r = compareBlockPair(p.e.item.representative, p.e.item.count, p.a.item.representative, p.a.item.count);
          if (p.strength === 'weak') {
            // Paired only by size: a mismatch could just mean the wrong two
            // visuals were paired, so it needs a human look.
            if (r.status === 'data-diff') {
              r.status = 'ambiguous';
              r.note = `Several visuals share the title "${title}" with different columns; this one was paired by size only, so the mismatch may be a wrong pairing. ${r.note ?? ''}`.trim();
            }
          } else if (p.strength === 'headers' && r.status !== 'identical') {
            r.note = `Several visuals share the title "${title}"; paired with the one that has identical columns. ${r.note ?? ''}`.trim();
          }
          results.push(r);
        }
        leftE.push(...restE);
        leftA.push(...restA);
        continue;
      }

      // Strict mode (PBI_COMPARE_NO_FUZZY_PAIRING=1): report as ambiguous.
      results.push({
        title, type, status: 'ambiguous',
        countExpected: expList.length, countActual: actList.length,
        rowsExpected: expList.reduce((s, b) => s + b.rows.length, 0),
        rowsActual:   actList.reduce((s, b) => s + b.rows.length, 0),
        headerExpected: expDedup[0]?.representative.headers ?? [],
        headerActual:   actDedup[0]?.representative.headers ?? [],
        sampleOnlyInExpected: expDedup.flatMap(d => d.representative.rows.slice(0, 5)),
        sampleOnlyInActual:   actDedup.flatMap(d => d.representative.rows.slice(0, 5)),
        onlyInExpectedCount: expList.reduce((s, b) => s + b.rows.length, 0),
        onlyInActualCount:   actList.reduce((s, b) => s + b.rows.length, 0),
        note: `Found ${expDedup.length} distinct variant(s) of "${title} [${type}]" in source and ${actDedup.length} in target — ` +
              `could not unambiguously match them to each other. Open both files and check this visual manually.`,
      });
    }
  }

  // Leftovers from every title/type group: a visual renamed (or retyped) in
  // the migrated report appears here once on each side. Pair by identical
  // columns, never by size alone.
  let restE = leftE;
  let restA = leftA;
  if (FUZZY_PAIRING && leftE.length > 0 && leftA.length > 0) {
    const { pairs, restE: rE, restA: rA } = pairByContent(leftE, leftA, false);
    restE = rE;
    restA = rA;
    for (const p of pairs) {
      const r = compareBlockPair(p.e.item.representative, p.e.item.count, p.a.item.representative, p.a.item.count);
      const sameTitle = canonicalizeForCompare(p.e.title) === canonicalizeForCompare(p.a.title);
      const sameType  = canonicalizeForCompare(p.e.type)  === canonicalizeForCompare(p.a.type);
      const what = [
        sameTitle ? '' : `title "${p.e.title}" in source vs "${p.a.title}" in target`,
        sameType  ? '' : `type "${p.e.type}" in source vs "${p.a.type}" in target`,
      ].filter(Boolean).join('; ');
      // The visual itself was renamed/retyped — a label-level change, like a
      // renamed column — so identical data is cosmetic, not a pass of a
      // differently-titled visual silently.
      if (r.status === 'identical') r.status = 'header-diff';
      r.note = `Matched by identical columns (${what}). ${p.strength === 'headers' ? 'Columns match but values differ — check these are the same visual. ' : ''}${r.note ?? ''}`.trim();
      results.push(r);
    }
  }

  for (const { title, type, item } of restA) {
    const { representative, count } = item;
    results.push({
      title, type, status: 'only-in-actual',
      countExpected: 0, countActual: count,
      rowsExpected: 0, rowsActual: representative.rows.length,
      headerExpected: [], headerActual: representative.headers,
      sampleOnlyInExpected: [], sampleOnlyInActual: representative.rows.slice(0, MISSING_VISUAL_SAMPLE_CAP),
      onlyInExpectedCount: 0, onlyInActualCount: representative.rows.length * count,
      note: count > 1 ? `${count} identical copies found, only in target.` : 'Only in target — not present in source.',
    });
  }
  for (const { title, type, item } of restE) {
    const { representative, count } = item;
    results.push({
      title, type, status: 'only-in-expected',
      countExpected: count, countActual: 0,
      rowsExpected: representative.rows.length, rowsActual: 0,
      headerExpected: representative.headers, headerActual: [],
      sampleOnlyInExpected: representative.rows.slice(0, MISSING_VISUAL_SAMPLE_CAP), sampleOnlyInActual: [],
      onlyInExpectedCount: representative.rows.length * count, onlyInActualCount: 0,
      note: count > 1 ? `${count} identical copies found, only in source.` : 'Only in source — missing from target.',
    });
  }

  return results;
}

// ── Fallback for sheets with no [VISUAL] markers at all ──────────────────────
//
// Used only when neither side's sheet parsed into any visual blocks — most
// often an arbitrary xlsx passed to the standalone compare:excel tool, not a
// report-export.helpers.ts output. Reads every row directly (same
// blank-row skipping as Tier 1) and diffs them as one plain block, so a
// non-VeriRapport-format sheet still gets a real comparison instead of a vacuous
// "nothing parsed, so call it identical".

async function streamPlainRows(
  filePath: string, target: SheetInfo, cache: CompareCache,
): Promise<string[][]> {
  return withSheetRetry(filePath, target, async () => {
    const rows = await getSheetRows(filePath, target, cache);
    return rows.filter(cells => !cells.every(c => c === ''));
  });
}

async function comparePlainSheet(
  fileExpected: string, fileActual: string,
  expSheet: SheetInfo, actSheet: SheetInfo,
  cache: CompareCache,
): Promise<VisualBlockDiff[]> {
  const [rowsExp, rowsAct] = await Promise.all([
    streamPlainRows(fileExpected, expSheet, cache),
    streamPlainRows(fileActual,   actSheet, cache),
  ]);

  const expSig = blockDataSignature(rowsExp);
  const actSig = blockDataSignature(rowsAct);
  const base = {
    title: '(whole sheet)', type: 'sheet',
    countExpected: 1, countActual: 1,
    rowsExpected: rowsExp.length, rowsActual: rowsAct.length,
    headerExpected: [], headerActual: [],
  };

  if (expSig === actSig) {
    return [{
      ...base, status: 'identical',
      sampleOnlyInExpected: [], sampleOnlyInActual: [],
      onlyInExpectedCount: 0, onlyInActualCount: 0,
    }];
  }

  const d = diffRows(rowsExp, rowsAct);
  return [{
    ...base, status: 'data-diff',
    sampleOnlyInExpected: d.sampleOnlyInExpected, sampleOnlyInActual: d.sampleOnlyInActual,
    onlyInExpectedCount: d.onlyInExpectedCount, onlyInActualCount: d.onlyInActualCount,
    note: 'This sheet has no [VISUAL] marker rows — compared as one plain block instead of per-visual.',
  }];
}

// ── Main public function ──────────────────────────────────────────────────────

/**
 * Compares two Excel workbooks: Tier 1 fast ordered-hash check per sheet,
 * falling through to the visual-block-aware Tier 2 comparison on any
 * mismatch. See the module-level comment for the full design.
 *
 * @param fileExpected  Path to the expected (baseline) .xlsx file
 * @param fileActual    Path to the actual (current run) .xlsx file
 */
export async function compareWorkbooks(
  fileExpected: string,
  fileActual:   string,
): Promise<WorkbookDiffSummary> {
  // Sheet lists come from each file's workbook.xml via the central directory —
  // deterministic and complete, never dependent on zip entry order (see the
  // module comment on readWorkbookSheetInfo).
  const expSheets = await readWorkbookSheetInfo(fileExpected);
  const actSheets = await readWorkbookSheetInfo(fileActual);

  // Fresh per call - see getBufferedWorkbook's doc comment for why this must
  // never be module-level. Each file's buffered Workbook (the fallback path)
  // is loaded/parsed at most ONCE total, no matter how many sheets are
  // compared — but most sheets never reach that path at all; see
  // getSheetRows() for the lightweight per-sheet reader tried first.
  const workbookCache: CompareCache = { workbooks: new Map(), fastPathFailed: new Set() };

  // Pair case-insensitively and whitespace-trimmed: "…- Cluster" vs
  // "…- cluster" (or a trailing space) between a source and a migrated
  // report is the same page, not a missing sheet. Excel forbids two sheets
  // in ONE workbook differing only by case, so keys are unique per side.
  const actByKey = new Map(actSheets.map(s => [sheetKey(s.name), s]));
  const expKeys  = new Set(expSheets.map(s => sheetKey(s.name)));

  const sheetsOnlyInExpected = expSheets.filter(s => !actByKey.has(sheetKey(s.name))).map(s => s.name);
  const sheetsOnlyInActual   = actSheets.filter(s => !expKeys.has(sheetKey(s.name))).map(s => s.name);
  const commonPairs = expSheets
    .filter(s => actByKey.has(sheetKey(s.name)))
    .map(s => ({ exp: s, act: actByKey.get(sheetKey(s.name))! }));

  const sheetResults: SheetDiffSummary[] = [];

  const emptySheetResult = (sheet: string, inExpected: boolean, inActual: boolean): SheetDiffSummary => ({
    sheet, inExpected, inActual,
    rowsExpected: 0, rowsActual: 0,
    identical: false, dataIdentical: false,
    visuals: [], headerDiffCount: 0, structuralDiffCount: 0,
    rowsOnlyInExpected: 0, rowsOnlyInActual: 0,
    sampleOnlyInExpected: [], sampleOnlyInActual: [],
  });

  for (const s of sheetsOnlyInExpected) sheetResults.push(emptySheetResult(s, true, false));
  for (const s of sheetsOnlyInActual)   sheetResults.push(emptySheetResult(s, false, true));

  // Common sheets — Tier 1 then Tier 2. Reported under the EXPECTED side's
  // name when the two sides' names differ only by case/whitespace.
  for (const { exp, act } of commonPairs) {
    const sheetName = exp.name;
    const [expHash, actHash] = await Promise.all([
      hashSheetOrdered(fileExpected, exp, workbookCache),
      hashSheetOrdered(fileActual,   act, workbookCache),
    ]);

    if (expHash.orderedHash === actHash.orderedHash) {
      // Fast path: byte-identical, no need for the detailed pass.
      sheetResults.push({
        sheet: sheetName, inExpected: true, inActual: true,
        rowsExpected: expHash.rowCount, rowsActual: actHash.rowCount,
        identical: true, dataIdentical: true,
        visuals: [], headerDiffCount: 0, structuralDiffCount: 0,
        rowsOnlyInExpected: 0, rowsOnlyInActual: 0,
        sampleOnlyInExpected: [], sampleOnlyInActual: [],
      });
      continue;
    }

    // Tier 2: visual-block-aware detailed comparison.
    const [expBlocks, actBlocks] = await Promise.all([
      streamSheetBlocks(fileExpected, exp, workbookCache),
      streamSheetBlocks(fileActual,   act, workbookCache),
    ]);

    let visuals: VisualBlockDiff[];
    if (expBlocks.length === 0 && actBlocks.length === 0) {
      // Neither side has any [VISUAL]-marker rows — this sheet isn't in the
      // report-export.helpers.ts format (e.g. an arbitrary xlsx passed to the
      // standalone `compare:excel` tool, not a visual export). Tier 1
      // already proved these sheets are NOT byte-identical (that's the only
      // way execution reaches here), so falling through to an empty visuals
      // list would be a vacuous truth — "nothing to compare" is not the same
      // as "everything matches". Compare the raw rows directly instead.
      visuals = await comparePlainSheet(fileExpected, fileActual, exp, act, workbookCache);
    } else {
      visuals = matchVisualBlocks(expBlocks, actBlocks);
    }

    const identical      = visuals.length > 0 ? visuals.every(v => v.status === 'identical') : true;
    const dataIdentical   = visuals.every(v =>
      v.status === 'identical' || v.status === 'header-diff' || v.status === 'duplicate-count',
    );
    const headerDiffCount = visuals.filter(v => v.status === 'header-diff').length;
    const structuralDiffCount = visuals.filter(v =>
      v.status === 'only-in-expected' || v.status === 'only-in-actual' ||
      v.status === 'ambiguous' || v.status === 'duplicate-count',
    ).length;

    // Approximate, display-only aggregate row counts — NOT the source of
    // truth for identical/dataIdentical (see field docs above).
    let rowsOnlyInExpected = 0;
    let rowsOnlyInActual   = 0;
    const sampleOnlyInExpected: string[][] = [];
    const sampleOnlyInActual:   string[][] = [];
    const pushCapped = (arr: string[][], items: string[][]) => {
      for (const r of items) {
        if (arr.length >= SAMPLE_SIZE) break;
        arr.push(r);
      }
    };
    for (const v of visuals) {
      if (v.status === 'identical' || v.status === 'duplicate-count') continue;
      if (v.status === 'header-diff') {
        rowsOnlyInExpected += 1;
        rowsOnlyInActual   += 1;
        pushCapped(sampleOnlyInExpected, [v.headerExpected]);
        pushCapped(sampleOnlyInActual,   [v.headerActual]);
        continue;
      }
      rowsOnlyInExpected += v.onlyInExpectedCount;
      rowsOnlyInActual   += v.onlyInActualCount;
      pushCapped(sampleOnlyInExpected, v.sampleOnlyInExpected);
      pushCapped(sampleOnlyInActual,   v.sampleOnlyInActual);
    }

    sheetResults.push({
      sheet: sheetName, inExpected: true, inActual: true,
      rowsExpected: expHash.rowCount, rowsActual: actHash.rowCount,
      identical, dataIdentical,
      visuals, headerDiffCount, structuralDiffCount,
      rowsOnlyInExpected, rowsOnlyInActual,
      sampleOnlyInExpected, sampleOnlyInActual,
    });
  }

  const noSheetsMissing = sheetsOnlyInExpected.length === 0 && sheetsOnlyInActual.length === 0;
  const identical     = noSheetsMissing && sheetResults.every(s => s.identical);
  const dataIdentical = noSheetsMissing && sheetResults.every(s => s.dataIdentical);

  return {
    fileExpected: path.basename(fileExpected),
    fileActual:   path.basename(fileActual),
    runAt:        new Date().toISOString(),
    sheetsOnlyInExpected,
    sheetsOnlyInActual,
    sheets: sheetResults,
    identical,
    dataIdentical,
  };
}