/**
 * compare-pairing.unit.spec.ts
 *
 * Offline tests for how compareWorkbooks pairs visuals that cannot be matched
 * by exact (title, type): renamed visuals, several visuals sharing a title,
 * extra/missing columns, and how changed values are reported. Uses real xlsx
 * files written by the same streaming writer the export pipeline uses.
 */

import { test, expect } from '@playwright/test';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import ExcelJS from 'exceljs';
import { compareWorkbooks, type WorkbookDiffSummary } from '../helpers/excel-compare.helpers';
import { writeVisualsToSheet, writeParitySummary } from '../helpers/report-export.helpers';
import type { VisualExport } from '../helpers/harness.helpers';

const tmpFiles: string[] = [];
function tmpPath(name: string): string {
  const p = path.join(os.tmpdir(), `verirapport-pair-${process.pid}-${name}`);
  tmpFiles.push(p);
  return p;
}
test.afterAll(() => {
  for (const f of tmpFiles) { try { fs.unlinkSync(f); } catch { /* gone */ } }
});

function vis(title: string, type: string, headers: string[], rows: string[][], name = `v-${title}-${type}`): VisualExport {
  return {
    visualName: name, visualTitle: title, visualType: type, csvData: null,
    headers,
    rows: rows.map(r => Object.fromEntries(headers.map((h, i) => [h, r[i] ?? '']))),
    rowCount: rows.length, errorMessage: null,
  };
}

async function compare(src: VisualExport[], tgt: VisualExport[]): Promise<WorkbookDiffSummary> {
  const write = async (file: string, visuals: VisualExport[]) => {
    const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: file, useStyles: true, useSharedStrings: false });
    const ws = wb.addWorksheet('Page');
    writeVisualsToSheet(ws, visuals);
    ws.commit();
    await wb.commit();
  };
  const e = tmpPath(`e-${tmpFiles.length}.xlsx`);
  const a = tmpPath(`a-${tmpFiles.length}.xlsx`);
  await write(e, src);
  await write(a, tgt);
  return compareWorkbooks(e, a);
}

const H = ['Dept', 'Month', 'Headcount'];
const ROWS = [['Finance', '2026-01', '10'], ['Finance', '2026-02', '12'], ['HR', '2026-01', '5']];

test.describe('visuals renamed during migration', () => {
  test('same columns and data under a new title are matched, not reported missing + added', async () => {
    const d = await compare(
      [vis('Headcount by Dept', 'tableEx', H, ROWS)],
      [vis('Headcount by Department', 'tableEx', H, ROWS)],
    );
    const v = d.sheets[0].visuals;
    expect(v).toHaveLength(1);
    expect(v[0].status).toBe('header-diff');
    expect(v[0].note).toContain('Matched by identical columns');
    expect(d.sheets[0].dataIdentical).toBe(true);
  });

  test('a renamed visual with a changed value is a data-diff that names the column', async () => {
    const changed = ROWS.map(r => [...r]);
    changed[1][2] = '13';
    const d = await compare(
      [vis('Headcount by Dept', 'tableEx', H, ROWS)],
      [vis('Headcount by Department', 'tableEx', H, changed)],
    );
    const v = d.sheets[0].visuals[0];
    expect(v.status).toBe('data-diff');
    expect(v.changedRows?.map(({ key, changes }) => ({ key, changes }))).toEqual([
      { key: 'Finance | 2026-02', changes: [{ column: 'Headcount', expected: '12', actual: '13' }] },
    ]);
    expect(d.sheets[0].dataIdentical).toBe(false);
  });

  test('a visual with a different type but identical columns and data is matched too', async () => {
    const d = await compare(
      [vis('Headcount', 'tableEx', H, ROWS)],
      [vis('Headcount', 'pivotTable', H, ROWS)],
    );
    expect(d.sheets[0].visuals).toHaveLength(1);
    expect(d.sheets[0].visuals[0].status).toBe('header-diff');
    expect(d.sheets[0].visuals[0].note).toContain('type');
  });

  test('a visual that really is missing on one side is still reported as missing', async () => {
    const d = await compare(
      [vis('A', 'tableEx', H, ROWS), vis('Only here', 'tableEx', ['X', 'Y'], [['1', '2']])],
      [vis('A', 'tableEx', H, ROWS)],
    );
    const statuses = d.sheets[0].visuals.map(v => v.status).sort();
    expect(statuses).toEqual(['identical', 'only-in-expected']);
    expect(d.sheets[0].dataIdentical).toBe(false);
  });

  test('different columns are never paired just because both are unmatched', async () => {
    const d = await compare(
      [vis('Old name', 'tableEx', ['A', 'B'], [['x', '1']])],
      [vis('New name', 'tableEx', ['C', 'D'], [['x', '1']])],
    );
    expect(d.sheets[0].visuals.map(v => v.status).sort()).toEqual(['only-in-actual', 'only-in-expected']);
  });
});

test.describe('several different visuals sharing one title', () => {
  const A = ['Name', 'Resigned'];
  const B = ['Name', 'Tenure', 'Grade'];
  test('are paired by their columns and compared, not left ambiguous', async () => {
    const d = await compare(
      [vis('Table', 'tableEx', A, [['a', '1'], ['b', '2']], 'x1'), vis('Table', 'tableEx', B, [['a', '3', 'G1']], 'x2')],
      [vis('Table', 'tableEx', B, [['a', '3', 'G1']], 'y1'), vis('Table', 'tableEx', A, [['a', '1'], ['b', '2']], 'y2')],
    );
    const v = d.sheets[0].visuals;
    expect(v).toHaveLength(2);
    expect(v.every(x => x.status === 'identical')).toBe(true);
    expect(d.sheets[0].dataIdentical).toBe(true);
  });

  test('same columns but different data are still reported as a data difference', async () => {
    const d = await compare(
      [vis('Table', 'tableEx', A, [['a', '1']], 'x1'), vis('Table', 'tableEx', B, [['a', '3', 'G1']], 'x2')],
      [vis('Table', 'tableEx', A, [['a', '9']], 'y1'), vis('Table', 'tableEx', B, [['a', '3', 'G1']], 'y2')],
    );
    const statuses = d.sheets[0].visuals.map(v => v.status).sort();
    expect(statuses).toEqual(['data-diff', 'identical']);
    expect(d.sheets[0].dataIdentical).toBe(false);
  });

  test('pairing by size alone is flagged as ambiguous when the data differs', async () => {
    const d = await compare(
      [vis('Table', 'tableEx', ['P', 'Q'], [['a', '1']], 'x1'), vis('Table', 'tableEx', ['R', 'S'], [['b', '2']], 'x2')],
      [vis('Table', 'tableEx', ['T', 'U'], [['a', '1']], 'y1'), vis('Table', 'tableEx', ['V', 'W'], [['c', '9']], 'y2')],
    );
    const statuses = d.sheets[0].visuals.map(v => v.status);
    expect(statuses).toContain('ambiguous');
  });
});

test.describe('extra, missing or renamed columns', () => {
  test('an extra column in the target is named, and shared columns are checked on their own', async () => {
    const d = await compare(
      [vis('T', 'tableEx', ['Dept', 'Count'], [['HR', '5'], ['IT', '7']])],
      [vis('T', 'tableEx', ['Dept', 'Count', 'Region'], [['HR', '5', 'EU'], ['IT', '7', 'US']])],
    );
    const v = d.sheets[0].visuals[0];
    expect(v.status).toBe('data-diff');
    expect(v.columnsOnlyInActual).toEqual(['Region']);
    expect(v.columnsOnlyInExpected).toEqual([]);
    expect(v.note).toContain('shared column(s) hold identical values');
    expect(d.sheets[0].dataIdentical).toBe(false);
  });

  test('a changed value next to an extra column is measured on shared columns only', async () => {
    const d = await compare(
      [vis('T', 'tableEx', ['Dept', 'Count'], [['HR', '5'], ['IT', '7'], ['Ops', '9']])],
      [vis('T', 'tableEx', ['Dept', 'Count', 'Region'], [['HR', '5', 'EU'], ['IT', '8', 'US'], ['Ops', '9', 'US']])],
    );
    const v = d.sheets[0].visuals[0];
    expect(v.status).toBe('data-diff');
    expect(v.onlyInExpectedCount).toBe(1);
    expect(v.onlyInActualCount).toBe(1);
    expect(v.changedRows?.map(({ key, changes }) => ({ key, changes }))).toEqual([{ key: 'IT', changes: [{ column: 'Count', expected: '7', actual: '8' }] }]);
  });

  test('a purely renamed column with identical values stays a harmless header difference', async () => {
    const d = await compare(
      [vis('T', 'tableEx', ['Dept', 'Count'], [['HR', '5']])],
      [vis('T', 'tableEx', ['Dept', 'Headcount'], [['HR', '5']])],
    );
    expect(d.sheets[0].visuals[0].status).toBe('header-diff');
    expect(d.sheets[0].dataIdentical).toBe(true);
  });
});

test.describe('changed values', () => {
  test('rows are paired by their text columns and each changed cell is listed', async () => {
    const d = await compare(
      [vis('T', 'tableEx', ['Dept', 'Sales', 'Cost'], [['HR', '10', '1'], ['IT', '20', '2'], ['Ops', '30', '3']])],
      [vis('T', 'tableEx', ['Dept', 'Sales', 'Cost'], [['HR', '10', '1'], ['IT', '21', '2'], ['Ops', '30', '4']])],
    );
    const v = d.sheets[0].visuals[0];
    expect(v.changedRowCount).toBe(2);
    const byKey = Object.fromEntries((v.changedRows ?? []).map(r => [r.key, r.changes]));
    expect(byKey['IT']).toEqual([{ column: 'Sales', expected: '20', actual: '21' }]);
    expect(byKey['Ops']).toEqual([{ column: 'Cost', expected: '3', actual: '4' }]);
    // True counts are unchanged by pairing.
    expect(v.onlyInExpectedCount).toBe(2);
    expect(v.onlyInActualCount).toBe(2);
  });

  test('rows with a repeated key are not paired (stay as plain only-in rows)', async () => {
    const d = await compare(
      [vis('T', 'tableEx', ['Dept', 'N'], [['HR', '1'], ['HR', '2']])],
      [vis('T', 'tableEx', ['Dept', 'N'], [['HR', '3'], ['HR', '4']])],
    );
    const v = d.sheets[0].visuals[0];
    expect(v.changedRows).toBeUndefined();
    expect(v.sampleOnlyInExpected).toHaveLength(2);
    expect(v.sampleOnlyInActual).toHaveLength(2);
  });

  test('parity summary gets a Changed Values sheet and one value per cell on Differences', async () => {
    const d = await compare(
      [vis('T', 'tableEx', ['Dept', 'Sales'], [['HR', '10'], ['IT', '20'], ['Legal', '1']])],
      [vis('T', 'tableEx', ['Dept', 'Sales'], [['HR', '10'], ['IT', '21'], ['Fin', '1']])],
    );
    const out = tmpPath('summary.xlsx');
    await writeParitySummary(d, {
      pairName: 'p', mode: 'both', sourceLabel: 's', targetLabel: 't',
      sourceReportId: 'a', targetReportId: 'b', expectedFile: 'e', actualFile: 'a',
    }, out);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(out);
    const names = wb.worksheets.map(w => w.name);
    expect(names).toContain('Changed Values');
    expect(names).toContain('Differences');
    const cv = wb.getWorksheet('Changed Values')!;
    expect(String(cv.getRow(2).getCell(4).value)).toBe('Sales');
    expect(String(cv.getRow(2).getCell(5).value)).toBe('20');
    expect(String(cv.getRow(2).getCell(6).value)).toBe('21');
    const df = wb.getWorksheet('Differences')!;
    const rows = df.getRows(1, df.rowCount)!.map(r => [1, 2, 3].map(c => String(r.getCell(c).value ?? '')));
    // Changed row: Original / Changed lines, then the Difference line with +1.
    const iOrig = rows.findIndex(r => r[0] === 'Original' && r[1] === 'IT');
    expect(iOrig).toBeGreaterThan(-1);
    expect(rows[iOrig + 1]).toEqual(['Changed', 'IT', '21']);
    expect(rows[iOrig + 2][0]).toBe('Difference');
    expect(String(df.getRow(iOrig + 3).getCell(3).value)).toBe('1');
    // Rows on one side only are labelled in plain words, one value per cell.
    expect(rows.filter(r => r[0] === 'Only in original').pop()).toEqual(['Only in original', 'Legal', '1']);
    expect(rows.filter(r => r[0] === 'Only in changed').pop()).toEqual(['Only in changed', 'Fin', '1']);
  });
});

test.describe('correctness audit cases', () => {
  const find = (d: WorkbookDiffSummary, title: string) => d.sheets.flatMap(s => s.visuals).filter(v => v.title === title);

  test('whitespace-only trailing cell on one side does not create a difference', async () => {
    const d = await compare(
      [vis('T', 'tableEx', ['A', 'B', 'C'], [['x', '1', ''], ['y', '2', '']])],
      [vis('T', 'tableEx', ['A', 'B', 'C'], [['x', '1', ' '], ['y', '2', '']])],
    );
    expect(d.sheets[0].dataIdentical).toBe(true);
  });

  test('a missing visual is NOT paired with an unrelated one that has other columns', async () => {
    const d = await compare(
      [vis('Old', 'tableEx', ['A', 'B'], [['x', '1']])],
      [vis('New', 'tableEx', ['C', 'D'], [['x', '1']])],
    );
    const st = d.sheets.flatMap(s => s.visuals).map(v => v.status).sort();
    expect(st).toEqual(['only-in-actual', 'only-in-expected']);
    expect(d.sheets[0].dataIdentical).toBe(false);
  });

  test('same-title variants are paired by shared rows, not by position or size', async () => {
    const hs = ['Dept', 'Count'];
    const s1 = [['A', '1'], ['B', '2'], ['C', '3']];
    const s2 = [['X', '9'], ['Y', '8'], ['Z', '7'], ['W', '6']];
    // target lists them in swapped order and S2 has one changed value
    const t2 = [['X', '9'], ['Y', '8'], ['Z', '7'], ['W', '5']];
    const d = await compare(
      [vis('Card', 'tableEx', hs, s1, 'v1'), vis('Card', 'tableEx', hs, s2, 'v2')],
      [vis('Card', 'tableEx', hs, t2, 'v3'), vis('Card', 'tableEx', hs, s1, 'v4')],
    );
    const vs = find(d, 'Card');
    expect(vs.map(v => v.status).sort()).toEqual(['data-diff', 'identical']);
    const diff = vs.find(v => v.status === 'data-diff')!;
    expect(diff.changedRows?.[0].changes[0]).toMatchObject({ expected: '6', actual: '5' });
  });

  test('a 0-row visual on one side and a populated one on the other is a failure', async () => {
    const d = await compare(
      [vis('E', 'tableEx', ['A', 'B'], [])],
      [vis('E', 'tableEx', ['A', 'B'], [['x', '1']])],
    );
    // the writer drops 0-row visuals, so the populated side shows as extra
    expect(find(d, 'E')[0].status).toBe('only-in-actual');
    expect(d.sheets[0].dataIdentical).toBe(false);
  });

  test('blank cell vs a real value is a difference; duplicate rows are counted', async () => {
    const d = await compare(
      [vis('D', 'tableEx', ['A', 'B'], [['x', ''], ['y', '2'], ['y', '2']])],
      [vis('D', 'tableEx', ['A', 'B'], [['x', '0'], ['y', '2']])],
    );
    const v = find(d, 'D')[0];
    expect(v.status).toBe('data-diff');
    expect(v.onlyInExpectedCount).toBe(2);
    expect(v.onlyInActualCount).toBe(1);
  });
});
