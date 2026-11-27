/**
 * filters-sheet.unit.spec.ts
 *
 * Offline test for the "Filters Applied" sheet of the parity summary: the
 * records passed in (cleared / applied / read-back) must show up as rows, and
 * an empty record list must say so instead of leaving a silent blank sheet.
 */

import { test, expect } from '@playwright/test';
import ExcelJS from 'exceljs';
import { writeFiltersSheet, type FilterRecord, type ParitySummaryMeta } from '../helpers/report-export.helpers';

const meta = (filterRecords?: FilterRecord[]): ParitySummaryMeta => ({
  pairName: 'p', mode: 'both', sourceLabel: 'Import mode', targetLabel: 'Direct Lake',
  sourceReportId: 's', targetReportId: 't', expectedFile: 'e.xlsx', actualFile: 'a.xlsx', filterRecords,
});

const textOf = (ws: ExcelJS.Worksheet): string[][] => {
  const out: string[][] = [];
  ws.eachRow(r => out.push((r.values as unknown[]).slice(1).map(v => String(v ?? ''))));
  return out;
};

test('Filters Applied sheet lists cleared, applied and read-back filters', () => {
  const wb = new ExcelJS.Workbook();
  writeFiltersSheet(wb, meta([
    { report: 'Source (old report)', page: 'Home Page', level: 'Slicer', field: 'Slicer (T.PROJECT_STATUS_CODE)', values: ['APPROVED'], what: 'Cleared (was saved in the report)' },
    { report: 'Target (new report)', page: 'Home Page', level: 'Page', field: 'T.PROJECT_NUMBER', values: ['1065110', '1065111'], what: 'Applied by you' },
    { report: 'Target (new report)', page: 'Home Page', level: 'Page', field: '(no page filter or slicer selection active)', values: [], what: 'Active after applying (read-back)' },
  ]));
  const ws = wb.getWorksheet('Filters Applied')!;
  expect(ws).toBeTruthy();
  const rows = textOf(ws).map(r => r.join(' | '));
  expect(rows.some(r => r.includes('Cleared') && r.includes('APPROVED'))).toBe(true);
  expect(rows.some(r => r.includes('Applied by you') && r.includes('1065110, 1065111'))).toBe(true);
  expect(rows.some(r => r.includes('read-back'))).toBe(true);
  expect(rows.some(r => r.includes('Row-level security'))).toBe(true);
});

test('Filters Applied sheet says so when nothing was recorded', () => {
  const wb = new ExcelJS.Workbook();
  writeFiltersSheet(wb, meta());
  const rows = textOf(wb.getWorksheet('Filters Applied')!).map(r => r.join(' '));
  expect(rows.some(r => r.includes('Nothing was recorded'))).toBe(true);
});
