/**
 * summary-format.helpers.ts
 *
 * Pure (no Excel) logic behind the plain-words parity summary: puts every
 * visual into one of five easy buckets, works out "were the same filters used
 * on both reports?", pulls out the report version / data-refresh facts, and
 * flags pages that hit the export row cap. report-export.helpers.ts turns the
 * results into sheets; keeping this separate makes it unit-testable.
 */

import type { VisualBlockDiff, WorkbookDiffSummary } from './excel-compare.helpers';
import type { FilterRecord } from './report-export.helpers';

/** Power BI export cap: a visual with exactly this many rows was probably cut off. */
export const EXPORT_ROW_CAP = 30000;

export type Bucket = 'real' | 'eye' | 'expected' | 'na' | 'match';

export const BUCKET_ORDER: Bucket[] = ['real', 'eye', 'expected', 'na', 'match'];

export const BUCKET_ICON: Record<Bucket, string> = {
  real: '❌', eye: '🔍', expected: 'ℹ️', na: '⚪', match: '✅',
};

export const BUCKET_MEANING: Record<Bucket, string> = {
  real:     'Real difference — the numbers (or the visual itself) differ and nothing explains it. Look at these.',
  eye:      'Could not be checked by the tool — please compare by eye.',
  expected: 'Differs as expected — report version / data-refresh date / last closed period. Not a problem.',
  na:       'Nothing to compare — empty on both sides, or removed on purpose.',
  match:    'Match — same numbers on both reports.',
};

/** Columns whose values are EXPECTED to differ between two builds / refresh times. */
export const EXPECTED_CHANGE_COLUMNS: RegExp[] = [
  /version/i, /last[ _]?update/i, /last[ _]?closed[ _]?period/i, /refresh/i,
];

export interface ClassifiedVisual {
  sheet: string;
  v: VisualBlockDiff;
  bucket: Bucket;
  /** Short result label for the "Result" column. */
  result: string;
  /** Plain-words reason this visual came back empty (blank when it did not). */
  emptyReason: string;
  /** How the old and new visual were paired. */
  paired: string;
}

function pairedText(v: VisualBlockDiff): string {
  if (!v.mapping) return 'By title + type';
  const same = v.mapping.oldVisualId && v.mapping.oldVisualId === v.mapping.newVisualId;
  return `${v.mapping.changeLabel}${same ? ' (same visual ID)' : ''}`;
}

function isExpectedDiff(v: VisualBlockDiff): boolean {
  if (v.status !== 'data-diff') return false;
  const changed = v.changedRows ?? [];
  if (/^popup/i.test(v.title)) return true;
  if (changed.length === 0) return false;
  // Rows that exist only on one side are not a version/date change.
  const unpaired = Math.max(0, v.onlyInExpectedCount - (v.changedRowCount ?? 0)) +
                   Math.max(0, v.onlyInActualCount - (v.changedRowCount ?? 0));
  if (unpaired > 0) return false;
  return changed.every(cr => cr.changes.every(ch => EXPECTED_CHANGE_COLUMNS.some(re => re.test(ch.column))));
}

export function emptyReasonText(v: VisualBlockDiff): string {
  const o = v.emptyReasonOld, n = v.emptyReasonNew;
  if (o && n) return o === n ? `Both reports: ${o}` : `Old report: ${o}  |  New report: ${n}`;
  if (o) return `Old report: ${o}`;
  if (n) return `New report: ${n}`;
  if (v.emptyBoth) return 'Neither report exported any data for this visual (reason not recorded — re-run to record it).';
  return '';
}

export function classifyVisual(sheet: string, v: VisualBlockDiff): ClassifiedVisual {
  let bucket: Bucket;
  let result: string;
  const bothEmpty = v.emptyBoth || (v.status === 'identical' && v.rowsExpected === 0 && v.rowsActual === 0);
  switch (v.status) {
    case 'identical':
      if (bothEmpty) { bucket = 'na'; result = '⚪ Empty on both reports'; }
      else { bucket = 'match'; result = '✅ Match'; }
      break;
    case 'header-diff': bucket = 'match'; result = '✅ Match (column label differs)'; break;
    case 'duplicate-count': bucket = 'match'; result = '✅ Match (copy count differs)'; break;
    case 'expected-removal': bucket = 'na'; result = '⚪ Removed on purpose'; break;
    case 'not-verified': bucket = 'eye'; result = '🔍 Not verified — check by eye'; break;
    case 'ambiguous': bucket = 'eye'; result = '🔍 Ambiguous — check by eye'; break;
    case 'only-in-actual':
      if (v.emptyReasonOld) { bucket = 'na'; result = '⚪ Old visual was empty'; }
      else { bucket = 'real'; result = '❌ Only in new report'; }
      break;
    case 'only-in-expected': bucket = 'real'; result = '❌ Only in old report'; break;
    case 'data-diff':
    default:
      if (isExpectedDiff(v)) { bucket = 'expected'; result = 'ℹ️ Differs as expected'; }
      else { bucket = 'real'; result = '❌ Values differ'; }
  }
  return { sheet, v, bucket, result, emptyReason: bucket === 'na' || v.emptyReasonOld || v.emptyReasonNew ? emptyReasonText(v) : '', paired: pairedText(v) };
}

export function classifyAll(diff: WorkbookDiffSummary): ClassifiedVisual[] {
  const out: ClassifiedVisual[] = [];
  for (const s of diff.sheets) for (const v of s.visuals) out.push(classifyVisual(s.sheet, v));
  return out;
}

export function tally(rows: ClassifiedVisual[]): Record<Bucket, number> {
  const t: Record<Bucket, number> = { real: 0, eye: 0, expected: 0, na: 0, match: 0 };
  for (const r of rows) t[r.bucket]++;
  return t;
}

// ── Biggest changes ──────────────────────────────────────────────────────────

const toNum = (s: string): number | null => {
  const t = s.trim().replace(/,/g, '');
  return /^-?\d+(\.\d+)?%?$/.test(t) ? parseFloat(t) : null;
};

export interface BigChange { sheet: string; visual: string; column: string; key: string; oldValue: string; newValue: string; delta: number }

/** The largest numeric changes (by absolute size) across the real differences. */
export function biggestChanges(rows: ClassifiedVisual[], n = 3): BigChange[] {
  const all: BigChange[] = [];
  for (const r of rows) {
    if (r.bucket !== 'real') continue;
    for (const cr of r.v.changedRows ?? []) for (const ch of cr.changes) {
      const e = toNum(ch.expected), a = toNum(ch.actual);
      if (e === null || a === null) continue;
      all.push({ sheet: r.sheet, visual: r.v.title, column: ch.column, key: cr.key, oldValue: ch.expected, newValue: ch.actual, delta: a - e });
    }
  }
  return all.sort((x, y) => Math.abs(y.delta) - Math.abs(x.delta)).slice(0, n);
}

// ── Report version / refresh facts ───────────────────────────────────────────

export interface ReferenceFact { label: string; oldValue: string; newValue: string }

const FACT_COLUMNS: Array<[string, RegExp]> = [
  ['Report version', /version/i],
  ['Data last refreshed', /last[ _]?update/i],
  ['Last closed period', /last[ _]?closed[ _]?period/i],
];

/** Version / refresh / closed-period values that differ between the reports (from card visuals). */
export function referenceFacts(rows: ClassifiedVisual[]): ReferenceFact[] {
  const out: ReferenceFact[] = [];
  for (const [label, re] of FACT_COLUMNS) {
    let found: ReferenceFact | null = null;
    for (const r of rows) {
      for (const cr of r.v.changedRows ?? []) for (const ch of cr.changes) {
        if (!found && re.test(ch.column)) found = { label, oldValue: ch.expected, newValue: ch.actual };
      }
      if (found) break;
    }
    if (found) out.push(found);
  }
  return out;
}

// ── Row cap ──────────────────────────────────────────────────────────────────

/** Visuals (page › title) that exported exactly the cap — likely cut off, so their lists may differ. */
export function rowCapVisuals(rows: ClassifiedVisual[], cap = EXPORT_ROW_CAP): string[] {
  return rows
    .filter(r => r.v.rowsExpected >= cap || r.v.rowsActual >= cap)
    .map(r => `${r.sheet} › ${r.v.title}`);
}

// ── Were the same filters used on both reports? ──────────────────────────────

const norm = (s: string) => s.toLowerCase().replace(/\s+/g, ' ').trim();
const isReadBack = (r: FilterRecord) => r.what.startsWith('Active after applying');
const isSource = (r: FilterRecord) => /^source/i.test(r.report);
const hasNoFilterMarker = (r: FilterRecord) => /^\(no /i.test(r.field);

const pageMatches = (rec: FilterRecord, sheet: string) => {
  if (/^all pages$/i.test(rec.page)) return true;
  const a = norm(rec.page), b = norm(sheet);
  return a === b || a.startsWith(b) || b.startsWith(a);
};

/** Filter values active on one side of one page (field names ignored: they can be renamed in a new report). */
function valuesFor(records: FilterRecord[], sheet: string, source: boolean): string[] | null {
  const recs = records.filter(r => isReadBack(r) && isSource(r) === source && pageMatches(r, sheet));
  if (recs.length === 0) return null;
  const vals: string[] = [];
  for (const r of recs) if (!hasNoFilterMarker(r)) for (const v of r.values) vals.push(norm(v));
  return [...new Set(vals)].sort();
}

/** Plain text of what was active on one side of one page, e.g. "PROJECT_NUMBER = 1065110". */
export function pageFilterText(records: FilterRecord[], sheet: string, source: boolean): string {
  const recs = records.filter(r => isReadBack(r) && isSource(r) === source && pageMatches(r, sheet));
  if (recs.length === 0) return '(not recorded)';
  const parts = recs.filter(r => !hasNoFilterMarker(r)).map(r => `${r.field.replace(/^Slicer \((.*)\)$/, '$1')} = ${r.values.join(', ') || '(any)'}`);
  const t = parts.length ? parts.join('; ') : '(no filter)';
  return t.length > 160 ? t.slice(0, 157) + '…' : t;
}

export interface SameFilters { answer: 'YES' | 'NO' | 'UNKNOWN'; detail: string; differingPages: string[] }

export function sameFiltersAnswer(records: FilterRecord[], sheets: string[]): SameFilters {
  const readBack = records.filter(isReadBack);
  if (!readBack.some(isSource) || !readBack.some(r => !isSource(r))) {
    return { answer: 'UNKNOWN', detail: 'The filters that were really active on each report were not read back, so this cannot be confirmed.', differingPages: [] };
  }
  const differing: string[] = [];
  for (const sh of sheets) {
    const a = valuesFor(records, sh, true), b = valuesFor(records, sh, false);
    if (a === null || b === null) continue;
    if (a.join('|') !== b.join('|')) differing.push(sh);
  }
  return differing.length === 0
    ? { answer: 'YES', detail: 'The same filter values were active on both reports (read back from Power BI).', differingPages: [] }
    : { answer: 'NO', detail: `Different filter values were active on both reports for: ${differing.join(', ')}. Differences in the numbers can come from this.`, differingPages: differing };
}
