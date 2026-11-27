/**
 * summary-format.unit.spec.ts
 *
 * Offline tests for the plain-words parity summary: visual buckets, the
 * "same filters?" answer, the empty-visual reasons, and the new sheet layout.
 */
import { test, expect } from '@playwright/test';
import ExcelJS from 'exceljs';
import * as os from 'os';
import * as path from 'path';
import { writeParitySummary, whyNotExported, type FilterRecord } from '../helpers/report-export.helpers';
import { classifyVisual, sameFiltersAnswer, rowCapVisuals, classifyAll } from '../helpers/summary-format.helpers';
import { normalizeNumericDrift } from '../helpers/xlsx-sheet-reader.helpers';
import type { VisualBlockDiff, WorkbookDiffSummary } from '../helpers/excel-compare.helpers';

const vd = (o: Partial<VisualBlockDiff>): VisualBlockDiff => ({
  title: 'V', type: 'card', status: 'identical', countExpected: 1, countActual: 1,
  rowsExpected: 1, rowsActual: 1, headerExpected: ['A'], headerActual: ['A'],
  sampleOnlyInExpected: [], sampleOnlyInActual: [], onlyInExpectedCount: 0, onlyInActualCount: 0, ...o,
});

test('whyNotExported gives plain reasons', () => {
  const base = { visualName: 'x', visualTitle: 't', headers: ['A'], rows: [], rowCount: 5, errorMessage: null } as any;
  expect(whyNotExported({ ...base })).toBeNull();
  expect(whyNotExported({ ...base, visualType: 'image', rowCount: 0 })).toMatch(/Not a data visual/);
  expect(whyNotExported({ ...base, visualType: 'card', rowCount: 0 })).toMatch(/0 rows/);
  expect(whyNotExported({ ...base, visualType: 'card', errorMessage: 'boom' })).toMatch(/Export failed: boom/);
  expect(whyNotExported({ ...base, visualType: 'card', headers: [] })).toMatch(/no columns/);
});

test('visuals fall into the five buckets', () => {
  const b = (o: Partial<VisualBlockDiff>) => classifyVisual('P', vd(o)).bucket;
  expect(b({})).toBe('match');
  expect(b({ rowsExpected: 0, rowsActual: 0 })).toBe('na');
  expect(b({ emptyBoth: true })).toBe('na');
  expect(b({ status: 'header-diff' })).toBe('match');
  expect(b({ status: 'expected-removal' })).toBe('na');
  expect(b({ status: 'not-verified' })).toBe('eye');
  expect(b({ status: 'only-in-expected' })).toBe('real');
  expect(b({ status: 'only-in-actual' })).toBe('real');
  expect(b({ status: 'only-in-actual', emptyReasonOld: 'Returned 0 rows' })).toBe('na');
  const real = { key: 'k', changes: [{ column: 'Sales', expected: '1', actual: '2' }] };
  const exp = { key: 'k', changes: [{ column: 'First Version', expected: 'v1', actual: 'v2' }] };
  expect(b({ status: 'data-diff', changedRows: [real], changedRowCount: 1, onlyInExpectedCount: 1, onlyInActualCount: 1 })).toBe('real');
  expect(b({ status: 'data-diff', changedRows: [exp], changedRowCount: 1, onlyInExpectedCount: 1, onlyInActualCount: 1 })).toBe('expected');
});

test('empty reason text covers both sides', () => {
  const c = classifyVisual('P', vd({ emptyBoth: true, rowsExpected: 0, rowsActual: 0, emptyReasonOld: 'Returned 0 rows', emptyReasonNew: 'Returned 0 rows' }));
  expect(c.emptyReason).toContain('Both reports: Returned 0 rows');
  const d = classifyVisual('P', vd({ emptyBoth: true, rowsExpected: 0, rowsActual: 0 }));
  expect(d.emptyReason).toMatch(/not recorded/);
});

const rec = (report: string, page: string, field: string, values: string[]): FilterRecord =>
  ({ report, page, level: 'Page', field, values, what: 'Active after applying (read-back)' });

test('same-filters answer ignores field names but not values', () => {
  const S = 'Source (old report)', T = 'Target (new report)';
  const same = sameFiltersAnswer([rec(S, 'Home', 'A.PROJECT', ['1']), rec(T, 'Home', 'B.PROJECT_NUMBER', ['1'])], ['Home']);
  expect(same.answer).toBe('YES');
  const diff = sameFiltersAnswer([rec(S, 'Home', 'A.PROJECT', ['1']), rec(T, 'Home', 'B.PROJECT_NUMBER', ['2'])], ['Home']);
  expect(diff.answer).toBe('NO');
  expect(diff.differingPages).toEqual(['Home']);
  expect(sameFiltersAnswer([], ['Home']).answer).toBe('UNKNOWN');
});

test('a visual at the export cap is flagged', () => {
  const d = { sheets: [{ sheet: 'P', visuals: [vd({ title: 'Big', rowsExpected: 30000, rowsActual: 29999 })] }] } as unknown as WorkbookDiffSummary;
  expect(rowCapVisuals(classifyAll(d))).toEqual(['P › Big']);
});

test('100% and 100.0% are the same value', () => {
  expect(normalizeNumericDrift('100%')).toBe(normalizeNumericDrift('100.0%'));
  expect(normalizeNumericDrift('12.50%')).toBe('12.5%');
  expect(normalizeNumericDrift('abc%')).toBe('abc%');
});

test('summary workbook has the new sheets and the empty-reason column', async () => {
  const sheet = {
    sheet: 'Home', inExpected: true, inActual: true, rowsExpected: 5, rowsActual: 6,
    identical: false, dataIdentical: false, headerDiffCount: 0, structuralDiffCount: 0,
    rowsOnlyInExpected: 0, rowsOnlyInActual: 0,
    visuals: [
      vd({ title: 'Same', rowsExpected: 3, rowsActual: 3 }),
      vd({ title: 'Empty', rowsExpected: 0, rowsActual: 0, countExpected: 0, countActual: 0, emptyBoth: true,
           emptyReasonOld: 'Returned 0 rows', emptyReasonNew: 'Returned 0 rows' }),
      vd({ title: 'Bad', type: 'tableEx', status: 'data-diff', headerExpected: ['Dept', 'Sales'], headerActual: ['Dept', 'Sales'],
           diffColumns: ['Dept', 'Sales'], onlyInExpectedCount: 1, onlyInActualCount: 1, changedRowCount: 1,
           changedRows: [{ key: 'HR', changes: [{ column: 'Sales', expected: '10', actual: '25' }], expectedRow: ['HR', '10'], actualRow: ['HR', '25'] }],
           sampleOnlyInExpected: [['HR', '10']], sampleOnlyInActual: [['HR', '25']] }),
    ],
  };
  const d = { sheets: [sheet], sheetsOnlyInExpected: [], sheetsOnlyInActual: [], identical: false, dataIdentical: false } as unknown as WorkbookDiffSummary;
  const out = path.join(os.tmpdir(), `summary-fmt-${Date.now()}.xlsx`);
  await writeParitySummary(d, {
    pairName: 'p', mode: 'both', sourceLabel: 's', targetLabel: 't', sourceReportId: 'a', targetReportId: 'b', expectedFile: 'e', actualFile: 'a',
    filterRecords: [rec('Source (old report)', 'Home', 'A.X', ['1']), rec('Target (new report)', 'Home', 'B.X', ['1'])],
  }, out);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(out);
  expect(wb.worksheets.map(w => w.name)).toEqual(['At a glance', 'What differs', 'All visuals', 'Changed values', 'Filters Applied', 'Run details', 'Differences']);
  const glance = wb.getWorksheet('At a glance')!;
  expect(String(glance.getRow(1).getCell(1).value)).toContain('1 REAL DIFFERENCE');
  const vis = wb.getWorksheet('All visuals')!;
  const header = (vis.getRow(1).values as unknown[]).slice(1).map(String);
  expect(header).toContain('Why it came back empty');
  const col = header.indexOf('Why it came back empty') + 1;
  const rows = vis.getRows(2, vis.rowCount - 1)!;
  const emptyRow = rows.find(r => String(r.getCell(2).value) === 'Empty')!;
  expect(String(emptyRow.getCell(col).value)).toContain('Returned 0 rows');
  expect(String(rows[0].getCell(2).value)).toBe('Bad'); // worst first
  const matchRow = rows.find(r => String(r.getCell(2).value) === 'Same')!;
  expect(String(matchRow.getCell(4).value)).toContain('Match');
});
