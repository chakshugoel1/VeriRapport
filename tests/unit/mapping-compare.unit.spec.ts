/**
 * Offline tests for KPI-mapping mode: reading a mapping file by header meaning,
 * field-name matching, and comparing old vs redesigned visuals per match type.
 * Builds small xlsx files in the OS temp dir (no .pbix/.csv, no browser).
 */
import { test, expect } from '@playwright/test';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import ExcelJS from 'exceljs';
import { normalizeFieldName, splitFieldRef, parseFieldList } from '../helpers/field-name.helpers';
import {
  loadMappingFile, classifyMatch, countByMatchType, buildFieldTranslation,
  buildPageNameTranslation, translateScenarioForMapping,
} from '../helpers/mapping-file.helpers';
import { compareMappedBlocks, compareWorkbooksWithMapping } from '../helpers/mapped-compare.helpers';
import { writeVisualsToSheet, writeVisualIndex, type ReportExportResult } from '../helpers/report-export.helpers';
import type { VisualExport } from '../helpers/harness.helpers';
import type { VisualBlockRaw } from '../helpers/excel-compare.helpers';

const tmp: string[] = [];
const tmpPath = (n: string) => { const p = path.join(os.tmpdir(), `vr-map-${process.pid}-${tmp.length}-${n}`); tmp.push(p); return p; };
test.afterAll(() => { for (const f of tmp) for (const g of [f, f.replace(/\.xlsx$/, '.visual-index.json')]) { try { fs.unlinkSync(g); } catch { /* gone */ } } });

test.describe('field names', () => {
  test('spelling styles compare equal', () => {
    expect(normalizeFieldName('Project Number')).toBe('projectnumber');
    expect(normalizeFieldName('PROJECT_NUMBER')).toBe('projectnumber');
    expect(normalizeFieldName('EXT_PROJ_LIST_TBL[Customer Name]')).toBe('customername');
    expect(normalizeFieldName('Sum of Amount')).toBe('amount');
    expect(splitFieldRef("'My Table'[Col X]")).toEqual({ table: 'My Table', column: 'Col X' });
    expect(parseFieldList('A[x]; B[y];')).toEqual(['A[x]', 'B[y]']);
  });
});

async function writeMapping(file: string, headers: string[], rows: string[][], sheetName = 'KPI mapping') {
  const wb = new ExcelJS.Workbook();
  wb.addWorksheet('Summary').addRow(['Old report', 'Old R']);
  const ws = wb.addWorksheet(sheetName);
  ws.addRow(headers);
  rows.forEach(r => ws.addRow(r));
  await wb.xlsx.writeFile(file);
}

const H = ['Old page', 'Old page ID', 'Old visual ID', 'Old visual type', 'Old title', 'Old fields', 'Match',
  'New page', 'New page ID', 'New visual ID', 'New visual type', 'New title', 'New fields', 'Fields dropped (why)', 'Reason'];

test.describe('mapping file reader', () => {
  test('reads rows by header meaning, in any column order, with extra columns', async () => {
    const f = tmpPath('m1.xlsx');
    // reordered + renamed headers + an unrelated column
    await writeMapping(f,
      ['Notes', 'Source visual ID', 'Change type', 'Source page', 'Target visual ID', 'Why'],
      [['x', 'v1', 'Same KPI (same fields)', 'Home', 'v1n', 'card became cardVisual'],
       ['x', 'v2', 'No same KPI in the new report', 'Home', '', 'model lacks the table']]);
    const m = await loadMappingFile(f);
    expect(m.rows.map(r => r.matchType)).toEqual(['same', 'dropped']);
    expect(m.rows[0]).toMatchObject({ oldVisualId: 'v1', newVisualId: 'v1n', oldPage: 'Home', reason: 'card became cardVisual' });
    expect(m.oldReportName).toBe('Old R');
  });

  test('classifies the four match types and flags unknown', () => {
    expect(classifyMatch('Same KPI (same fields)')).toBe('same');
    expect(classifyMatch('Merged into a multi-value card')).toBe('merged');
    expect(classifyMatch('Same table, columns changed')).toBe('columns-changed');
    expect(classifyMatch('No same KPI in the new report')).toBe('dropped');
    expect(classifyMatch('???')).toBe('unknown');
    expect(classifyMatch('Weird', { matchTypes: { merged: ['weird'] } })).toBe('merged');
  });

  test('a file with no recognisable mapping sheet fails with a helpful message', async () => {
    const f = tmpPath('m2.xlsx');
    await writeMapping(f, ['Foo', 'Bar'], [['1', '2']], 'Sheet1');
    await expect(loadMappingFile(f)).rejects.toThrow(/old visual ID.*Foo \| Bar/s);
  });

  test('a non-dropped row with no new visual ID is flagged, not trusted', async () => {
    const f = tmpPath('m3.xlsx');
    await writeMapping(f, H, [['P', 'p1', 'v1', 'card', '', 'T[a]', 'Same KPI (same fields)', 'P', 'p1', '', '', '', '', '', '']]);
    const m = await loadMappingFile(f);
    expect(m.rows[0].matchType).toBe('unknown');
    expect(m.warnings.join(' ')).toMatch(/no new visual ID/);
    expect(countByMatchType(m).unknown).toBe(1);
  });

  test('field + page translation for filters', async () => {
    const f = tmpPath('m4.xlsx');
    await writeMapping(f, H, [['Home', 'p1', 'v1', 'tableEx', '', 'PROJ[Project Number]', 'Same KPI (same fields)', 'Start', 'p9', 'v1', 'tableEx', '', 'PROJ[PROJECT_NUMBER]', '', '']]);
    const m = await loadMappingFile(f);
    const tr = translateScenarioForMapping(
      { name: 's', pages: { Home: [{ values: ['A'], targets: [{ table: 'PROJ', column: 'Project Number' }] }] } },
      buildFieldTranslation(m), buildPageNameTranslation(m));
    expect(tr.pages['Home'].map(s => s.side)).toEqual(['source', 'target']);
    expect(tr.pages['Home'][1].targets![0].column).toBe('PROJECT_NUMBER');
    expect(tr.pages['Start'].map(s => s.side)).toEqual(['target']); // new page name gets the target selection
  });
});

function blk(title: string, type: string, headers: string[], rows: string[][]): VisualBlockRaw { return { title, type, headers, rows }; }

test.describe('compareMappedBlocks', () => {
  test('same values, field names spelled differently -> header-diff (safe)', () => {
    const d = compareMappedBlocks(blk('K', 'card', ['Project Number'], [['5']]), blk('', 'cardVisual', ['PROJECT_NUMBER'], [['5']]), 'same');
    expect(d.status).toBe('header-diff');
  });
  test('same, different values -> data-diff', () => {
    const d = compareMappedBlocks(blk('K', 'card', ['Total'], [['5']]), blk('K', 'cardVisual', ['Total'], [['6']]), 'same');
    expect(d.status).toBe('data-diff');
  });
  test('same, columns in a different order are matched by name', () => {
    const d = compareMappedBlocks(blk('T', 'tableEx', ['A', 'B'], [['1', '2']]), blk('T', 'tableEx', ['B', 'A'], [['2', '1']]), 'same');
    expect(d.status).toMatch(/identical|header-diff/);
  });
  test('columns-changed: removed column is listed, shared columns compared', () => {
    const d = compareMappedBlocks(
      blk('T', 'tableEx', ['Customer Name', 'Project Type', 'Amt'], [['x', 't', '1']]),
      blk('T2', 'tableEx', ['CUSTOMER_NAME', 'AMT'], [['x', '1']]), 'columns-changed');
    expect(d.status).toBe('header-diff');
    expect(d.columnsOnlyInExpected).toEqual(['Project Type']);
  });
  test('columns-changed with a real value change -> data-diff', () => {
    const d = compareMappedBlocks(blk('T', 'tableEx', ['A', 'Gone'], [['1', 'z']]), blk('T', 'tableEx', ['A'], [['2']]), 'columns-changed');
    expect(d.status).toBe('data-diff');
  });
  test('merged: old value found inside the bigger new card and matches', () => {
    const d = compareMappedBlocks(blk('', 'card', ['Budget'], [['10']]), blk('Details', 'cardVisual', ['Name', 'Budget', 'Status'], [['P1', '10', 'Open']]), 'merged');
    expect(['identical', 'header-diff']).toContain(d.status);
  });
  test('merged: old value differs inside the new card -> data-diff', () => {
    const d = compareMappedBlocks(blk('', 'card', ['Budget'], [['10']]), blk('Details', 'cardVisual', ['Name', 'Budget'], [['P1', '11']]), 'merged');
    expect(d.status).toBe('data-diff');
  });
  test('same-KPI card whose label changed is compared by value (position), not skipped', () => {
    const ok = compareMappedBlocks(blk('', 'card', ['M_Last Available Period_EC'], [['10-26']]), blk('', 'cardVisual', ['Last closed period'], [['10-26']]), 'same');
    expect(ok.status).toBe('header-diff');
    expect(ok.note).toMatch(/compared by position/);
    const bad = compareMappedBlocks(blk('', 'card', ['M_Last Available Period_EC'], [['10-26']]), blk('', 'cardVisual', ['Last closed period'], [['09-26']]), 'same');
    expect(bad.status).toBe('data-diff');
  });
  test('multiRowCard -> cardVisual with all labels renamed: paired by position', () => {
    const d = compareMappedBlocks(blk('', 'multiRowCard', ['A1', 'B1'], [['1', '2']]), blk('', 'cardVisual', ['X', 'Y'], [['1', '2']]), 'same');
    expect(d.status).toBe('header-diff');
  });
  test('renamed field paired through the mapping file field lists', () => {
    const d = compareMappedBlocks(
      blk('', 'card', ['Budget EUR'], [['10']]),
      blk('Details', 'cardVisual', ['Name', 'BUDGET_TOTAL'], [['P', '10']]), 'merged',
      { oldFields: ['T[Budget EUR]'], newFields: ['T[Name]', 'T[BUDGET_TOTAL]'] });
    // old field Budget EUR is not in new list by name -> falls through to value search
    expect(d.status).toBe('identical');
    const d2 = compareMappedBlocks(
      blk('', 'card', ['Budget EUR'], [['10']]),
      blk('', 'cardVisual', ['Budget Total'], [['11']]), 'same',
      { oldFields: ['T[Budget EUR]'], newFields: ['T[Budget Total]'] });
    expect(d2.status).toBe('data-diff');
    expect(d2.note).toMatch(/mapping file/);
  });
  test('merged: old value found by value when labels differ', () => {
    const d = compareMappedBlocks(blk('', 'card', ['C_PARENT_BLANK'], [['1,234.50']]), blk('Details', 'cardVisual', ['Name', 'Parent'], [['P1', '1234.5']]), 'merged');
    expect(d.status).toBe('identical');
    expect(d.note).toMatch(/Found by value/);
  });
  test('merged: value cannot be located -> not-verified (never a failure)', () => {
    const d = compareMappedBlocks(blk('', 'card', ['C_PARENT_BLANK'], [['x']]), blk('Details', 'cardVisual', ['Name', 'Status'], [['P1', 'Open']]), 'merged');
    expect(d.status).toBe('not-verified');
  });
});

function vis(id: string, title: string, type: string, headers: string[], rows: string[][]): VisualExport {
  return {
    visualName: id, visualTitle: title, visualType: type, csvData: null, headers,
    rows: rows.map(r => Object.fromEntries(headers.map((h, i) => [h, r[i] ?? '']))), rowCount: rows.length, errorMessage: null,
  };
}

async function exportSide(file: string, pageName: string, pageId: string, visuals: VisualExport[]) {
  const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: file, useStyles: true, useSharedStrings: false });
  const ws = wb.addWorksheet(pageName);
  const { visualsWithData, totalRows, index } = writeVisualsToSheet(ws, visuals);
  ws.commit();
  await wb.commit();
  const res: ReportExportResult = {
    pagesFound: [], outPath: file, requestedPagesMissing: [],
    exported: [{ sheetName: pageName, pageDisplayName: pageName, pageInternalName: pageId, visualsWithData, totalRows, visuals: index }],
  };
  writeVisualIndex(res);
}

test.describe('compareWorkbooksWithMapping (end to end, offline)', () => {
  test('pairs by visual ID across type/title changes and handles every match type', async () => {
    const e = tmpPath('e.xlsx'); const a = tmpPath('a.xlsx'); const mf = tmpPath('map.xlsx');
    await exportSide(e, 'Home', 'p1', [
      vis('o1', '', 'card', ['Project Number'], [['5']]),
      vis('o2', 'Costs', 'tableEx', ['Customer Name', 'Project Type', 'Amt'], [['x', 't', '1']]),
      vis('o3', '', 'card', ['Budget'], [['10']]),
      vis('o4', '', 'card', ['Weird'], [['9']]),
      vis('o5', 'Old broken', 'tableEx', ['Z'], [['1']]),
      vis('o6', 'Not mapped', 'tableEx', ['Q'], [['1']]),
    ]);
    await exportSide(a, 'Home', 'p1', [
      vis('o1', 'KPI', 'cardVisual', ['PROJECT_NUMBER'], [['5']]),
      vis('o2', 'Projects with budget', 'tableEx', ['CUSTOMER_NAME', 'AMT'], [['x', '1']]),
      vis('n3', 'Details', 'cardVisual', ['Name', 'Budget'], [['P1', '10']]),
    ]);
    await writeMapping(mf, H, [
      ['Home', 'p1', 'o1', 'card', '', 'T[Project Number]', 'Same KPI (same fields)', 'Home', 'p1', 'o1', 'cardVisual', 'KPI', 'T[PROJECT_NUMBER]', '', ''],
      ['Home', 'p1', 'o2', 'tableEx', 'Costs', 'T[Customer Name]; T[Project Type]; T[Amt]', 'Same table, columns changed', 'Home', 'p1', 'o2', 'tableEx', '', 'T[CUSTOMER_NAME]; T[AMT]', 'T[Project Type] (decision D-1)', ''],
      ['Home', 'p1', 'o3', 'card', '', 'T[Budget]', 'Merged into a multi-value card', 'Home', 'p1', 'n3', 'cardVisual', 'Details', 'T[Name]; T[Budget]', '', 'auto-merged'],
      ['Home', 'p1', 'o4', 'card', '', 'T[Weird]', 'Merged into a multi-value card', 'Home', 'p1', 'n3', 'cardVisual', 'Details', 'T[Name]; T[Budget]', '', ''],
      ['Home', 'p1', 'o5', 'tableEx', 'Old broken', 'X[Z]', 'No same KPI in the new report', '', '', '', '', '', '', '', 'model lacks table X'],
    ]);
    const m = await loadMappingFile(mf);
    const { diff, caveats } = await compareWorkbooksWithMapping(e, a, m);
    expect(caveats).toEqual([]);
    const v = diff.sheets.find(s => s.sheet === 'Home')!.visuals;
    const byId = (id: string) => v.find(x => x.mapping?.oldVisualId === id)!;
    expect(byId('o1').status).toBe('header-diff');
    expect(byId('o2').status).toBe('header-diff');
    expect(byId('o2').columnsOnlyInExpected).toEqual(['Project Type']);
    expect(byId('o3').status).toMatch(/identical|header-diff/);
    expect(byId('o4').status).toBe('not-verified');
    expect(byId('o5').status).toBe('expected-removal');
    // the exported-but-unlisted visual is reported, not hidden
    expect(v.find(x => x.title === 'Not mapped')!.status).toBe('only-in-expected');
    // not-verified and removed-on-purpose are not failures; the unlisted visual is
    expect(diff.sheets.find(s => s.sheet === 'Home')!.dataIdentical).toBe(false);
  });

  test('without the unlisted visual the page passes', async () => {
    const e = tmpPath('e2.xlsx'); const a = tmpPath('a2.xlsx'); const mf = tmpPath('map2.xlsx');
    await exportSide(e, 'Home', 'p1', [vis('o1', '', 'card', ['A'], [['1']]), vis('o5', 'Gone', 'tableEx', ['Z'], [['1']])]);
    await exportSide(a, 'Start', 'p2', [vis('n1', 'K', 'cardVisual', ['a'], [['1']])]);
    await writeMapping(mf, H, [
      ['Home', 'p1', 'o1', 'card', '', 'T[A]', 'Same KPI (same fields)', 'Start', 'p2', 'n1', 'cardVisual', 'K', 'T[a]', '', ''],
      ['Home', 'p1', 'o5', 'tableEx', '', 'X[Z]', 'No same KPI in the new report', '', '', '', '', '', '', '', 'dropped'],
    ]);
    const { diff } = await compareWorkbooksWithMapping(e, a, await loadMappingFile(mf));
    expect(diff.dataIdentical).toBe(true);
    expect(diff.sheetsOnlyInExpected).toEqual([]);
  });

  test('a missing visual index gives a clear error', async () => {
    const e = tmpPath('e3.xlsx'); const a = tmpPath('a3.xlsx'); const mf = tmpPath('map3.xlsx');
    await exportSide(e, 'Home', 'p1', [vis('o1', '', 'card', ['A'], [['1']])]);
    await exportSide(a, 'Home', 'p1', [vis('o1', '', 'card', ['A'], [['1']])]);
    fs.unlinkSync(a.replace(/\.xlsx$/, '.visual-index.json'));
    await writeMapping(mf, H, [['Home', 'p1', 'o1', 'card', '', 'T[A]', 'Same KPI (same fields)', 'Home', 'p1', 'o1', 'card', '', 'T[A]', '', '']]);
    await expect(compareWorkbooksWithMapping(e, a, await loadMappingFile(mf))).rejects.toThrow(/visual index/);
  });
});
