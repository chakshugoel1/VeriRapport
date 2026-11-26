/**
 * mapping-file.helpers.ts
 *
 * Reads a "KPI mapping" Excel file (old report -> redesigned report) into
 * plain data, WITHOUT depending on column positions or exact header spelling.
 *
 * Why: the file's layout may change between tools/versions, but the
 * information in it stays the same (old/new page + visual IDs, old/new
 * fields, a match type, a reason). So:
 *   - the right sheet is found by looking at HEADERS, not by sheet name;
 *   - each needed column is found by header MEANING (case/space/symbol
 *     insensitive, with a list of accepted synonyms);
 *   - extra or reordered columns are ignored;
 *   - the synonym lists can be extended without code changes through an
 *     optional settings JSON (see MappingSettings).
 *
 * No browser, no Power BI access - pure file reading + logic.
 */

import * as fs from 'fs';
import * as path from 'path';
import ExcelJS from 'exceljs';
import { parseFieldList, normalizeFieldName, splitFieldRef } from './field-name.helpers';

// ── Public types ─────────────────────────────────────────────────────────────

/** How an old visual relates to the new report. */
export type MatchType = 'same' | 'merged' | 'columns-changed' | 'dropped' | 'unknown';

export const MATCH_TYPE_LABEL: Record<MatchType, string> = {
  'same':            'Same KPI',
  'merged':          'Merged into a multi-value card',
  'columns-changed': 'Same table, columns changed',
  'dropped':         'Removed on purpose',
  'unknown':         'Unknown change type',
};

/** One row of the mapping file = one OLD visual and what became of it. */
export interface MappingRow {
  oldPage:      string;
  oldPageId:    string;
  oldVisualId:  string;
  oldVisualType: string;
  oldTitle:     string;
  oldFields:    string[];
  /** The raw Match text from the file (kept for display). */
  matchText:    string;
  matchType:    MatchType;
  newPage:      string;
  newPageId:    string;
  newVisualId:  string;
  newVisualType: string;
  newTitle:     string;
  newFields:    string[];
  fieldsDropped: string;
  fieldsAdded:  string;
  reason:       string;
  /** 1-based row number in the sheet, for error messages. */
  sourceRow:    number;
}

export interface MappingFile {
  filePath:  string;
  sheetName: string;
  /** Report names from a Summary-style sheet, when present (display only). */
  oldReportName?: string;
  newReportName?: string;
  rows:      MappingRow[];
  /** Non-fatal problems found while reading (shown as caveats, never hidden). */
  warnings:  string[];
}

/** Optional settings file: extra header synonyms and match-type words. */
export interface MappingSettings {
  /** logical column name -> extra accepted header texts. */
  columns?:    Partial<Record<LogicalColumn, string[]>>;
  /** match type -> extra words that appear in the Match cell. */
  matchTypes?: Partial<Record<MatchType, string[]>>;
  /** Sheet name to use instead of auto-detection. */
  sheetName?:  string;
}

export type LogicalColumn =
  | 'oldPage' | 'oldPageId' | 'oldVisualId' | 'oldVisualType' | 'oldTitle' | 'oldFields'
  | 'match'
  | 'newPage' | 'newPageId' | 'newVisualId' | 'newVisualType' | 'newTitle' | 'newFields'
  | 'fieldsDropped' | 'fieldsAdded' | 'reason';

// ── Header meaning ───────────────────────────────────────────────────────────

/** Normalises header text: lowercase, letters and digits only. */
function hkey(text: string): string {
  return text.normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, '');
}

/** Built-in accepted header spellings for each logical column (normalised on use). */
const DEFAULT_SYNONYMS: Record<LogicalColumn, string[]> = {
  oldPage:       ['old page', 'source page', 'original page', 'page old'],
  oldPageId:     ['old page id', 'source page id', 'original page id', 'old page name'],
  oldVisualId:   ['old visual id', 'source visual id', 'original visual id', 'old visual', 'old kpi id', 'old visual name'],
  oldVisualType: ['old visual type', 'source visual type', 'old type'],
  oldTitle:      ['old title', 'source title', 'old visual title'],
  oldFields:     ['old fields', 'source fields', 'old columns', 'old field'],
  match:         ['match', 'match type', 'change type', 'mapping type', 'result'],
  newPage:       ['new page', 'target page', 'redesigned page'],
  newPageId:     ['new page id', 'target page id', 'new page name'],
  newVisualId:   ['new visual id', 'target visual id', 'new visual', 'new kpi id', 'new visual name'],
  newVisualType: ['new visual type', 'target visual type', 'new type'],
  newTitle:      ['new title', 'target title', 'new visual title'],
  newFields:     ['new fields', 'target fields', 'new columns', 'new field'],
  fieldsDropped: ['fields dropped', 'fields dropped why', 'dropped fields', 'removed fields', 'fields removed'],
  fieldsAdded:   ['fields added', 'added fields', 'new fields added'],
  reason:        ['reason', 'why', 'why changed', 'comment', 'notes', 'explanation'],
};

/** Columns the loader cannot work without. */
const REQUIRED: LogicalColumn[] = ['oldVisualId', 'match'];

/** Words (normalised) that identify each match type inside the Match cell. */
const DEFAULT_MATCH_WORDS: Record<Exclude<MatchType, 'unknown'>, string[]> = {
  // Checked in this order: the more specific phrases first, "same" last.
  'merged':          ['merged', 'multivalue', 'combined'],
  'columns-changed': ['columnschanged', 'fewercolumns', 'columnsremoved', 'sametablecolumn'],
  'dropped':         ['nosamekpi', 'dropped', 'removed', 'nomatch', 'notinnew', 'deleted'],
  'same':            ['same', 'identical', 'unchanged', 'exact'],
};

export function classifyMatch(text: string, settings?: MappingSettings): MatchType {
  const k = hkey(text);
  if (!k) return 'unknown';
  const words = (t: Exclude<MatchType, 'unknown'>): string[] =>
    [...DEFAULT_MATCH_WORDS[t], ...(settings?.matchTypes?.[t] ?? []).map(hkey)];
  for (const t of ['merged', 'columns-changed', 'dropped', 'same'] as const) {
    if (words(t).some(w => w && k.includes(w))) return t;
  }
  return 'unknown';
}

// ── Cell reading ─────────────────────────────────────────────────────────────

function cellText(v: ExcelJS.CellValue | undefined): string {
  if (v === null || v === undefined) return '';
  if (typeof v === 'object') {
    const o = v as unknown as Record<string, unknown>;
    if ('result' in o) return cellText(o['result'] as ExcelJS.CellValue);           // formula: use the cached result
    if ('richText' in o && Array.isArray(o['richText'])) return (o['richText'] as Array<{ text: string }>).map(r => r.text).join('');
    if ('text' in o) return String(o['text'] ?? '');                                // hyperlink
    if (v instanceof Date) return v.toISOString();
    return '';
  }
  return String(v).trim();
}

interface HeaderHit { row: number; cols: Partial<Record<LogicalColumn, number>> }

function findHeader(ws: ExcelJS.Worksheet, settings?: MappingSettings): HeaderHit | null {
  const maxRow = Math.min(ws.rowCount, 15);
  for (let r = 1; r <= maxRow; r++) {
    const row = ws.getRow(r);
    const byKey = new Map<string, number>();
    row.eachCell({ includeEmpty: false }, (cell, colNo) => {
      const k = hkey(cellText(cell.value));
      if (k && !byKey.has(k)) byKey.set(k, colNo);
    });
    if (byKey.size < 2) continue;
    const cols: Partial<Record<LogicalColumn, number>> = {};
    for (const logical of Object.keys(DEFAULT_SYNONYMS) as LogicalColumn[]) {
      const names = [...DEFAULT_SYNONYMS[logical], ...(settings?.columns?.[logical] ?? [])].map(hkey);
      const hit = names.find(n => byKey.has(n));
      if (hit) cols[logical] = byKey.get(hit);
    }
    if (REQUIRED.every(c => cols[c] !== undefined)) return { row: r, cols };
  }
  return null;
}

// ── Public API ───────────────────────────────────────────────────────────────

/** Loads optional settings JSON; a missing path means "defaults". */
export function loadMappingSettings(settingsPath?: string): MappingSettings | undefined {
  if (!settingsPath || !settingsPath.trim()) return undefined;
  const resolved = path.isAbsolute(settingsPath) ? settingsPath : path.join(process.cwd(), settingsPath);
  if (!fs.existsSync(resolved)) throw new Error(`Mapping settings file not found: ${resolved}`);
  return JSON.parse(fs.readFileSync(resolved, 'utf8')) as MappingSettings;
}

/**
 * Reads a mapping workbook. Throws a plain-English error (naming the columns
 * it did and did not find) if no sheet looks like a KPI mapping.
 */
export async function loadMappingFile(filePath: string, settings?: MappingSettings): Promise<MappingFile> {
  if (!fs.existsSync(filePath)) throw new Error(`Mapping file not found: ${filePath}`);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);

  // Prefer a configured sheet name, then any sheet whose headers fit.
  const sheets = wb.worksheets.slice().sort((a, b) => {
    const pref = (w: ExcelJS.Worksheet) =>
      (settings?.sheetName && hkey(w.name) === hkey(settings.sheetName)) ? 0 : (hkey(w.name).includes('mapping') ? 1 : 2);
    return pref(a) - pref(b);
  });

  let ws: ExcelJS.Worksheet | undefined;
  let hit: HeaderHit | null = null;
  for (const cand of sheets) {
    hit = findHeader(cand, settings);
    if (hit) { ws = cand; break; }
  }
  if (!ws || !hit) {
    const seen = wb.worksheets.map(w => {
      const heads: string[] = [];
      w.getRow(1).eachCell({ includeEmpty: false }, c => heads.push(cellText(c.value)));
      return `"${w.name}": ${heads.slice(0, 8).join(' | ') || '(empty)'}`;
    }).join('; ');
    throw new Error(
      `Could not find a KPI mapping sheet in ${path.basename(filePath)}. ` +
      `I need a sheet with a column meaning "old visual ID" and a column meaning "match type". ` +
      `Headers I saw -> ${seen}. ` +
      `(If your file uses other column names, add them under "columns" in a mapping settings file.)`,
    );
  }

  const { cols } = hit;
  const get = (row: ExcelJS.Row, c: LogicalColumn): string =>
    cols[c] === undefined ? '' : cellText(row.getCell(cols[c]!).value);

  const rows: MappingRow[] = [];
  const warnings: string[] = [];
  for (let r = hit.row + 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const oldVisualId = get(row, 'oldVisualId');
    const matchText = get(row, 'match');
    if (!oldVisualId && !matchText) continue; // blank line
    if (!oldVisualId) { warnings.push(`Mapping row ${r}: no old visual ID - row skipped.`); continue; }
    const matchType = classifyMatch(matchText, settings);
    if (matchType === 'unknown') warnings.push(`Mapping row ${r}: match type "${matchText}" is not recognised - the visual is flagged for manual review.`);
    const mr: MappingRow = {
      oldPage: get(row, 'oldPage'), oldPageId: get(row, 'oldPageId'),
      oldVisualId, oldVisualType: get(row, 'oldVisualType'), oldTitle: get(row, 'oldTitle'),
      oldFields: parseFieldList(get(row, 'oldFields')),
      matchText, matchType,
      newPage: get(row, 'newPage'), newPageId: get(row, 'newPageId'),
      newVisualId: get(row, 'newVisualId'), newVisualType: get(row, 'newVisualType'), newTitle: get(row, 'newTitle'),
      newFields: parseFieldList(get(row, 'newFields')),
      fieldsDropped: get(row, 'fieldsDropped'), fieldsAdded: get(row, 'fieldsAdded'),
      reason: get(row, 'reason'), sourceRow: r,
    };
    if (matchType !== 'dropped' && matchType !== 'unknown' && !mr.newVisualId) {
      warnings.push(`Mapping row ${r}: ${MATCH_TYPE_LABEL[matchType]} but no new visual ID - flagged for manual review.`);
      mr.matchType = 'unknown';
    }
    rows.push(mr);
  }
  if (rows.length === 0) throw new Error(`The mapping sheet "${ws.name}" has headers but no data rows.`);

  // Optional Summary-style sheet: "Old report" / "New report" label rows.
  let oldReportName: string | undefined;
  let newReportName: string | undefined;
  for (const w of wb.worksheets) {
    if (w === ws) continue;
    for (let r = 1; r <= Math.min(w.rowCount, 12); r++) {
      const k = hkey(cellText(w.getRow(r).getCell(1).value));
      if (k === 'oldreport') oldReportName ??= cellText(w.getRow(r).getCell(2).value);
      if (k === 'newreport') newReportName ??= cellText(w.getRow(r).getCell(2).value);
    }
  }

  return { filePath, sheetName: ws.name, oldReportName, newReportName, rows, warnings };
}

// ── Derived helpers ──────────────────────────────────────────────────────────

/** Counts of rows per match type (for the app's status line and the summary). */
export function countByMatchType(m: MappingFile): Record<MatchType, number> {
  const out: Record<MatchType, number> = { same: 0, merged: 0, 'columns-changed': 0, dropped: 0, unknown: 0 };
  for (const r of m.rows) out[r.matchType]++;
  return out;
}

/** Distinct old page names (display) / new page names mentioned in the mapping. */
export function mappedPageNames(m: MappingFile, side: 'old' | 'new'): string[] {
  const set = new Set<string>();
  for (const r of m.rows) {
    const n = side === 'old' ? r.oldPage : r.newPage;
    if (n) set.add(n);
  }
  return [...set];
}

/**
 * Old-field -> new-field translation learned from rows where the two field
 * lists line up by spelling-insensitive column name. Key = "table|column"
 * style old reference (as written in the file), value = the new table/column.
 * Used to translate filters (slicers) from the old report to the new one.
 */
export function buildFieldTranslation(m: MappingFile): Map<string, { table?: string; column: string }> {
  const out = new Map<string, { table?: string; column: string }>();
  for (const r of m.rows) {
    if (r.matchType === 'dropped' || r.matchType === 'unknown') continue;
    for (const oldRef of r.oldFields) {
      const oldParts = splitFieldRef(oldRef);
      const hit = r.newFields.find(n => normalizeFieldName(n) === normalizeFieldName(oldRef));
      if (!hit) continue;
      const nw = splitFieldRef(hit);
      out.set(`${(oldParts.table ?? '').toLowerCase()}|${oldParts.column.toLowerCase()}`, nw);
    }
  }
  return out;
}

// ── Filter (slicer) translation ──────────────────────────────────────────────

import type { SlicerScenario, SlicerSelection } from './slicer-config.helpers';

/** old page display name (lowercase, trimmed) -> new page display name, from the mapping file. */
export function buildPageNameTranslation(m: MappingFile): Map<string, string> {
  const out = new Map<string, string>();
  for (const r of m.rows) {
    if (r.oldPage && r.newPage && !out.has(r.oldPage.trim().toLowerCase())) out.set(r.oldPage.trim().toLowerCase(), r.newPage);
  }
  return out;
}

/**
 * Rewrites a filter scenario (written against the OLD report) so it also works
 * on the redesigned report:
 *   - a flat-field selection whose field is spelled differently in the new
 *     report ("Project Number" -> PROJECT_NUMBER) is split into one selection
 *     tagged for the source side and one tagged for the target side that uses
 *     the new field (the pipeline already supports side-tagged selections);
 *   - selections on a page whose name changed are also listed under the new
 *     page name for the target side.
 * Selections it cannot translate are left exactly as they were (applied to
 * both sides, as today); the existing pre-flight field check still catches a
 * field that does not exist in the new model. Pure: returns a new scenario.
 */
export function translateScenarioForMapping(
  scenario: SlicerScenario,
  fields: Map<string, { table?: string; column: string }>,
  pageNames: Map<string, string>,
): SlicerScenario {
  const pages: Record<string, SlicerSelection[]> = {};
  const add = (page: string, sel: SlicerSelection) => {
    const list = pages[page] ?? (pages[page] = []);
    const key = JSON.stringify([sel.side ?? '', sel.visualName ?? '', sel.targets ?? [], sel.values]);
    if (!list.some(x => JSON.stringify([x.side ?? '', x.visualName ?? '', x.targets ?? [], x.values]) === key)) list.push(sel);
  };

  for (const [pageName, sels] of Object.entries(scenario.pages)) {
    const newPage = pageNames.get(pageName.trim().toLowerCase());
    const pageRenamed = !!newPage && newPage.trim().toLowerCase() !== pageName.trim().toLowerCase();

    for (const sel of sels) {
      const t = sel.targets?.[0];
      const translated = (!sel.side && !sel.isHierarchy && sel.targets?.length === 1 && t?.table && t?.column)
        ? fields.get(`${t.table.toLowerCase()}|${t.column.toLowerCase()}`)
        : undefined;
      const changed = !!translated &&
        (translated.column.toLowerCase() !== t!.column!.toLowerCase() ||
         (!!translated.table && translated.table.toLowerCase() !== t!.table!.toLowerCase()));

      if (changed) {
        const tgtSel: SlicerSelection = {
          ...sel, side: 'target',
          targets: [{ ...t!, table: translated!.table ?? t!.table, column: translated!.column }],
        };
        add(pageName, { ...sel, side: 'source' });
        add(pageName, tgtSel);
        if (pageRenamed) add(newPage!, tgtSel);
      } else {
        add(pageName, sel);
        if (pageRenamed) add(newPage!, sel);
      }
    }
  }
  return { name: scenario.name, pages };
}
