/**
 * mapped-compare.helpers.ts
 *
 * Compares an OLD report export with a REDESIGNED report export using a KPI
 * mapping file (see mapping-file.helpers.ts), instead of the normal
 * title+type pairing in compareWorkbooks(). The normal flow is untouched;
 * this module is used only when a mapping file is supplied.
 *
 * Inputs: two exported workbooks + their sidecar visual indexes (written by
 * writeVisualIndex), which say which Power BI visual ID is which block on
 * which sheet. Output: a WorkbookDiffSummary, so the existing summary writer
 * and UI results work unchanged, with `mapping` info attached to each visual.
 *
 * What happens per mapping row ("Match" column):
 *   same             -> compare the two visuals' values (field names are
 *                       matched ignoring case/spaces/underscores).
 *   columns-changed  -> compare only the columns both visuals still have;
 *                       the removed columns are listed, not failed.
 *   merged           -> look for the old card's column(s) inside the new
 *                       (bigger) card and compare those. If they cannot be
 *                       found, the card is "not verified" - listed for a
 *                       look, never a failure. (Safe by design: this needs
 *                       no sample of a real merged-card export to be correct.)
 *   dropped          -> listed as "removed on purpose", never a failure.
 *   unknown          -> "not verified".
 * Visuals exported on a side but not mentioned in the mapping are reported
 * as only-in-source / only-in-target ("not in the mapping file").
 */

import * as fs from 'fs';
import {
  compareBlockPair, readWorkbookBlocks, sheetKey,
  type VisualBlockRaw, type VisualBlockDiff, type SheetDiffSummary, type WorkbookDiffSummary,
} from './excel-compare.helpers';
import type { VisualIndexFile } from './report-export.helpers';
import { visualIndexPathFor } from './report-export.helpers';
import { MATCH_TYPE_LABEL, type MappingFile, type MappingRow } from './mapping-file.helpers';
import { normalizeFieldName } from './field-name.helpers';
import * as path from 'path';

// ── Loading one side ─────────────────────────────────────────────────────────

interface BlockRef {
  sheetName: string; pageId: string; pageName: string;
  visualId: string; block: VisualBlockRaw;
}

interface Side {
  refs: BlockRef[];
  byPageVisual: Map<string, BlockRef>;
  byVisual: Map<string, BlockRef[]>;
  sheetRows: Map<string, number>;     // sheetName -> total data rows
  sheetByPageId: Map<string, string>; // internal page id -> sheetName
  sheetByPageName: Map<string, string>;
  problems: string[];
}

async function loadSide(xlsxPath: string, label: string): Promise<Side> {
  const side: Side = {
    refs: [], byPageVisual: new Map(), byVisual: new Map(), sheetRows: new Map(),
    sheetByPageId: new Map(), sheetByPageName: new Map(), problems: [],
  };
  const indexPath = visualIndexPathFor(xlsxPath);
  if (!fs.existsSync(indexPath)) {
    throw new Error(
      `Mapping-file mode needs the visual index next to ${path.basename(xlsxPath)} ` +
      `(${path.basename(indexPath)}), which is written when the report is exported. ` +
      `Re-run the ${label} export (it was probably created by an older version).`,
    );
  }
  const index = JSON.parse(fs.readFileSync(indexPath, 'utf8')) as VisualIndexFile;
  const blocksBySheet = await readWorkbookBlocks(xlsxPath);
  const sheetNames = new Map([...blocksBySheet.keys()].map(n => [sheetKey(n), n]));

  for (const sh of index.sheets) {
    const real = sheetNames.get(sheetKey(sh.sheetName));
    const blocks = real ? blocksBySheet.get(real)! : undefined;
    if (!blocks) { side.problems.push(`${label}: sheet "${sh.sheetName}" is in the index but not in ${path.basename(xlsxPath)}.`); continue; }
    if (blocks.length !== sh.visuals.length) {
      side.problems.push(`${label}: sheet "${sh.sheetName}" has ${blocks.length} visual block(s) but the index lists ${sh.visuals.length} - the file and its index are out of step; re-export.`);
      continue;
    }
    let total = 0;
    side.sheetByPageId.set(sh.pageInternalName, sh.sheetName);
    side.sheetByPageName.set(sheetKey(sh.pageDisplayName), sh.sheetName);
    for (let i = 0; i < blocks.length; i++) {
      const v = sh.visuals[i];
      if (blocks[i].title !== v.title.trim() || blocks[i].type !== v.type.trim()) {
        side.problems.push(`${label}: sheet "${sh.sheetName}" block ${i + 1} is "${blocks[i].title} [${blocks[i].type}]" but the index says "${v.title} [${v.type}]" - re-export.`);
        continue;
      }
      const ref: BlockRef = { sheetName: sh.sheetName, pageId: sh.pageInternalName, pageName: sh.pageDisplayName, visualId: v.visualName, block: blocks[i] };
      side.refs.push(ref);
      side.byPageVisual.set(`${sh.pageInternalName}|${v.visualName}`, ref);
      const list = side.byVisual.get(v.visualName) ?? [];
      list.push(ref); side.byVisual.set(v.visualName, list);
      total += blocks[i].rows.length;
    }
    side.sheetRows.set(sh.sheetName, total);
  }
  return side;
}

/** Finds a visual by page id + visual id, falling back to a visual id that is unique in the report. */
function lookup(side: Side, pageId: string, visualId: string): BlockRef | undefined {
  if (!visualId) return undefined;
  const exact = side.byPageVisual.get(`${pageId}|${visualId}`);
  if (exact) return exact;
  const all = side.byVisual.get(visualId);
  return all && all.length === 1 ? all[0] : undefined;
}

// ── Comparing one pair ───────────────────────────────────────────────────────

/** Unique column index by spelling-insensitive name; repeated names are left out (cannot be matched safely). */
function columnIndex(headers: string[]): Map<string, number> {
  const seen = new Map<string, number>();
  const dup = new Set<string>();
  headers.forEach((h, i) => {
    const k = normalizeFieldName(h);
    if (!k) return;
    if (seen.has(k)) dup.add(k); else seen.set(k, i);
  });
  for (const k of dup) seen.delete(k);
  return seen;
}

function emptyDiff(title: string, type: string): VisualBlockDiff {
  return {
    title, type, status: 'identical',
    countExpected: 1, countActual: 1, rowsExpected: 0, rowsActual: 0,
    headerExpected: [], headerActual: [],
    sampleOnlyInExpected: [], sampleOnlyInActual: [],
    onlyInExpectedCount: 0, onlyInActualCount: 0,
  };
}

/**
 * Compares an old visual with its new counterpart on the columns they share.
 * `mode` controls what a column that is missing from the NEW visual means.
 */
export function compareMappedBlocks(
  exp: VisualBlockRaw, act: VisualBlockRaw, mode: 'same' | 'columns-changed' | 'merged',
): VisualBlockDiff {
  const eIdx = columnIndex(exp.headers);
  const aIdx = columnIndex(act.headers);

  const shared: Array<{ e: number; a: number }> = [];
  for (const [k, e] of eIdx) {
    const a = aIdx.get(k);
    if (a !== undefined) shared.push({ e, a });
  }
  shared.sort((x, y) => x.e - y.e);
  const sharedE = new Set(shared.map(s => s.e));
  const sharedA = new Set(shared.map(s => s.a));
  const missingInNew = exp.headers.filter((_, i) => !sharedE.has(i));
  const extraInNew   = act.headers.filter((_, i) => !sharedA.has(i));

  if (shared.length === 0) {
    const d = emptyDiff(exp.title, exp.type);
    d.status = 'not-verified';
    d.rowsExpected = exp.rows.length; d.rowsActual = act.rows.length;
    d.headerExpected = exp.headers; d.headerActual = act.headers;
    d.note = mode === 'merged'
      ? `Could not find this card's value(s) (${exp.headers.join(', ') || 'no columns'}) inside the new combined card (columns: ${act.headers.join(', ') || 'none'}), so it was not compared. Check it by eye.`
      : `None of this visual's columns (${exp.headers.join(', ')}) could be matched by name in the new visual (${act.headers.join(', ')}), so it was not compared. Check it by eye.`;
    return d;
  }

  // Project both blocks onto the shared columns (in the old visual's order),
  // giving the new side the old side's header text so the existing row
  // comparison sees "same columns".
  const expHeaders = shared.map(s => exp.headers[s.e]);
  const expP: VisualBlockRaw = { ...exp, headers: expHeaders, rows: exp.rows.map(r => shared.map(s => r[s.e] ?? '')) };
  const actP: VisualBlockRaw = { ...act, headers: expHeaders, rows: act.rows.map(r => shared.map(s => r[s.a] ?? '')) };
  const d = compareBlockPair(expP, 1, actP, 1);

  // Report the visuals as they really are, not the projections.
  d.title = exp.title; d.type = exp.type;
  d.headerExpected = exp.headers; d.headerActual = act.headers;
  d.rowsExpected = exp.rows.length; d.rowsActual = act.rows.length;

  const notes: string[] = [];
  if (d.note) notes.push(d.note);
  const spellingDiffers = shared.some(s => exp.headers[s.e].trim().toLowerCase() !== act.headers[s.a].trim().toLowerCase());
  if (d.status === 'identical' && spellingDiffers) {
    d.status = 'header-diff';
    notes.push('Same values; field names are spelled differently in the new report (matched ignoring case, spaces and underscores).');
  }
  if (missingInNew.length > 0) {
    d.columnsOnlyInExpected = missingInNew;
    if (mode === 'columns-changed') {
      notes.push(`Compared on the ${shared.length} column(s) both reports have. Removed in the new report: ${missingInNew.join(', ')}.`);
    } else if (mode === 'merged') {
      notes.push(`Only part of this card could be found in the new combined card (${shared.length} of ${exp.headers.length} column(s)); not found: ${missingInNew.join(', ')}.`);
      if (d.status === 'identical' || d.status === 'header-diff') d.status = 'not-verified';
    } else {
      notes.push(`Column(s) not found in the new visual: ${missingInNew.join(', ')}. The ${shared.length} shared column(s) were compared.`);
      if (d.status === 'identical' || d.status === 'header-diff') d.status = 'not-verified';
    }
  }
  if (extraInNew.length > 0 && mode !== 'merged') {
    d.columnsOnlyInActual = extraInNew;
    notes.push(`Extra column(s) in the new visual (not compared): ${extraInNew.join(', ')}.`);
  }
  d.note = notes.join(' ') || undefined;
  return d;
}

function withMapping(d: VisualBlockDiff, row: MappingRow): VisualBlockDiff {
  d.mapping = {
    changeType: row.matchType, changeLabel: row.matchText || MATCH_TYPE_LABEL[row.matchType],
    oldVisualId: row.oldVisualId, newVisualId: row.newVisualId,
    newTitle: row.newTitle || undefined, newType: row.newVisualType || undefined,
    reason: row.reason || undefined,
    fieldsDropped: row.fieldsDropped || undefined,
  };
  return d;
}

// ── Whole-workbook comparison ────────────────────────────────────────────────

export interface MappedCompareResult {
  diff: WorkbookDiffSummary;
  /** Problems that make the comparison itself untrustworthy (shown as caveats). */
  caveats: string[];
}

export async function compareWorkbooksWithMapping(
  fileExpected: string, fileActual: string, mapping: MappingFile,
): Promise<MappedCompareResult> {
  const [old, nw] = await Promise.all([loadSide(fileExpected, 'source'), loadSide(fileActual, 'target')]);
  const caveats = [...old.problems, ...nw.problems, ...mapping.warnings.map(w => `Mapping file: ${w}`)];

  const usedOld = new Set<BlockRef>();
  const usedNew = new Set<BlockRef>();
  const visualsBySheet = new Map<string, VisualBlockDiff[]>();
  const newSheetsForOld = new Map<string, Set<string>>(); // old sheet -> new sheets it maps to
  const push = (sheet: string, d: VisualBlockDiff) => {
    const list = visualsBySheet.get(sheet) ?? [];
    list.push(d); visualsBySheet.set(sheet, list);
  };

  // Rows whose old page was not exported at all (page filter) are skipped.
  const oldSheetOf = (row: MappingRow): string | undefined =>
    old.sheetByPageId.get(row.oldPageId) ?? old.sheetByPageName.get(sheetKey(row.oldPage));

  for (const row of mapping.rows) {
    const sheet = oldSheetOf(row);
    if (!sheet) continue;

    const oldRef = lookup(old, row.oldPageId, row.oldVisualId);
    if (oldRef) usedOld.add(oldRef);
    const newRef = row.newVisualId ? lookup(nw, row.newPageId, row.newVisualId) : undefined;
    if (newRef) {
      usedNew.add(newRef);
      const set = newSheetsForOld.get(sheet) ?? new Set<string>();
      set.add(newRef.sheetName); newSheetsForOld.set(sheet, set);
    }

    const title = oldRef?.block.title ?? (row.oldTitle || '(no title)');
    const type  = oldRef?.block.type  ?? row.oldVisualType;
    let d: VisualBlockDiff;

    if (row.matchType === 'dropped') {
      d = emptyDiff(title, type);
      d.status = 'expected-removal';
      d.rowsExpected = oldRef?.block.rows.length ?? 0;
      d.headerExpected = oldRef?.block.headers ?? [];
      d.countActual = 0;
      d.note = `Removed on purpose according to the mapping file${row.reason ? `: ${row.reason}` : '.'}`;
    } else if (row.matchType === 'unknown') {
      d = emptyDiff(title, type);
      d.status = 'not-verified';
      d.rowsExpected = oldRef?.block.rows.length ?? 0;
      d.note = `The mapping file's change type "${row.matchText}" is not recognised (or the new visual is missing), so this visual was not compared.`;
    } else if (!oldRef && !newRef) {
      // Nothing exported on either side (e.g. a visual with no data). Not a difference.
      d = emptyDiff(title, type);
      d.countExpected = 0; d.countActual = 0;
      d.note = 'No data was exported for this visual on either side, so there is nothing to compare.';
    } else if (oldRef && !newRef) {
      d = emptyDiff(title, type);
      d.status = 'only-in-expected';
      d.countActual = 0;
      d.rowsExpected = oldRef.block.rows.length;
      d.headerExpected = oldRef.block.headers;
      d.sampleOnlyInExpected = oldRef.block.rows.slice(0, 20);
      d.onlyInExpectedCount = oldRef.block.rows.length;
      d.note = `The mapping file says this visual is "${MATCH_TYPE_LABEL[row.matchType]}" (new visual ${row.newVisualId}), but that visual exported no data in the new report.`;
    } else if (!oldRef && newRef) {
      d = emptyDiff(title, type);
      d.status = 'only-in-actual';
      d.countExpected = 0;
      d.rowsActual = newRef.block.rows.length;
      d.headerActual = newRef.block.headers;
      d.sampleOnlyInActual = newRef.block.rows.slice(0, 20);
      d.onlyInActualCount = newRef.block.rows.length;
      d.note = 'The old visual exported no data, but the new visual did.';
    } else {
      d = compareMappedBlocks(oldRef!.block, newRef!.block, row.matchType as 'same' | 'columns-changed' | 'merged');
      d.title = title; d.type = type;
    }
    push(sheet, withMapping(d, row));
  }

  // Exported visuals the mapping file does not mention.
  const sheetOfNewPage = new Map<string, string>(); // new sheet -> an old sheet that maps to it
  for (const [oldSheet, set] of newSheetsForOld) for (const ns of set) if (!sheetOfNewPage.has(ns)) sheetOfNewPage.set(ns, oldSheet);
  const sheetsWithMappedRows = new Set(visualsBySheet.keys());

  for (const ref of old.refs) {
    if (usedOld.has(ref) || !sheetsWithMappedRows.has(ref.sheetName)) continue;
    const d = emptyDiff(ref.block.title, ref.block.type);
    d.status = 'only-in-expected'; d.countActual = 0;
    d.rowsExpected = ref.block.rows.length; d.headerExpected = ref.block.headers;
    d.sampleOnlyInExpected = ref.block.rows.slice(0, 20); d.onlyInExpectedCount = ref.block.rows.length;
    d.note = 'This visual is not listed in the mapping file.';
    push(ref.sheetName, d);
  }
  const strayNewSheets = new Set<string>();
  for (const ref of nw.refs) {
    if (usedNew.has(ref)) continue;
    const home = sheetOfNewPage.get(ref.sheetName);
    if (!home) { strayNewSheets.add(ref.sheetName); continue; }
    const d = emptyDiff(ref.block.title, ref.block.type);
    d.status = 'only-in-actual'; d.countExpected = 0;
    d.rowsActual = ref.block.rows.length; d.headerActual = ref.block.headers;
    d.sampleOnlyInActual = ref.block.rows.slice(0, 20); d.onlyInActualCount = ref.block.rows.length;
    d.note = `New visual on page "${ref.pageName}" that is not listed in the mapping file.`;
    push(home, d);
  }

  // Build per-sheet results (one per OLD page that had mapping rows).
  const sheets: SheetDiffSummary[] = [];
  const dataOk = (v: VisualBlockDiff) =>
    v.status === 'identical' || v.status === 'header-diff' || v.status === 'duplicate-count' ||
    v.status === 'expected-removal' || v.status === 'not-verified';
  for (const [sheet, visuals] of visualsBySheet) {
    let rowsActual = 0;
    for (const ns of newSheetsForOld.get(sheet) ?? []) rowsActual += nw.sheetRows.get(ns) ?? 0;
    let rowsOnlyE = 0, rowsOnlyA = 0;
    const sampleE: string[][] = [], sampleA: string[][] = [];
    for (const v of visuals) {
      if (v.status === 'identical' || v.status === 'duplicate-count' || v.status === 'expected-removal' || v.status === 'not-verified') continue;
      if (v.status === 'header-diff') { rowsOnlyE++; rowsOnlyA++; continue; }
      rowsOnlyE += v.onlyInExpectedCount; rowsOnlyA += v.onlyInActualCount;
      for (const r of v.sampleOnlyInExpected) if (sampleE.length < 800) sampleE.push(r);
      for (const r of v.sampleOnlyInActual) if (sampleA.length < 800) sampleA.push(r);
    }
    sheets.push({
      sheet, inExpected: true, inActual: true,
      rowsExpected: old.sheetRows.get(sheet) ?? 0, rowsActual,
      identical: visuals.every(v => v.status === 'identical'),
      dataIdentical: visuals.every(dataOk),
      visuals,
      headerDiffCount: visuals.filter(v => v.status === 'header-diff').length,
      structuralDiffCount: visuals.filter(v => v.status === 'only-in-expected' || v.status === 'only-in-actual' || v.status === 'ambiguous').length,
      rowsOnlyInExpected: rowsOnlyE, rowsOnlyInActual: rowsOnlyA,
      sampleOnlyInExpected: sampleE, sampleOnlyInActual: sampleA,
    });
  }

  // Pages that exist on only one side after mapping.
  const sheetsOnlyInExpected = [...new Set(old.refs.map(r => r.sheetName))]
    .filter(s => !visualsBySheet.has(s) && (old.sheetRows.get(s) ?? 0) > 0);
  const sheetsOnlyInActual = [...strayNewSheets];
  const empty = (sheet: string, inE: boolean, inA: boolean): SheetDiffSummary => ({
    sheet, inExpected: inE, inActual: inA, rowsExpected: 0, rowsActual: 0,
    identical: false, dataIdentical: false, visuals: [], headerDiffCount: 0, structuralDiffCount: 0,
    rowsOnlyInExpected: 0, rowsOnlyInActual: 0, sampleOnlyInExpected: [], sampleOnlyInActual: [],
  });
  for (const s of sheetsOnlyInExpected) sheets.push(empty(s, true, false));
  for (const s of sheetsOnlyInActual) sheets.push(empty(s, false, true));

  const noMissing = sheetsOnlyInExpected.length === 0 && sheetsOnlyInActual.length === 0;
  const diff: WorkbookDiffSummary = {
    fileExpected: path.basename(fileExpected), fileActual: path.basename(fileActual),
    runAt: new Date().toISOString(),
    sheetsOnlyInExpected, sheetsOnlyInActual, sheets,
    identical: noMissing && sheets.every(s => s.identical),
    dataIdentical: noMissing && sheets.every(s => s.dataIdentical),
  };
  return { diff, caveats };
}
