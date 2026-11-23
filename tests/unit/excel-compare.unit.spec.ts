/**
 * excel-compare.unit.spec.ts
 *
 * Offline end-to-end regression tests for compareWorkbooks: writes REAL xlsx
 * files with the same streaming writer the export pipeline uses (worksheets
 * placed before workbook.xml in the zip — the exact layout that broke the
 * ExcelJS streaming reader in production), then compares them. No browser, no
 * Power BI.
 *
 * Guards against the production incident where read-back sheet names came out
 * as "Sheet1/Sheet2/Sheet3" (or the reader crashed with "Cannot read
 * properties of undefined (reading 'sheets')"), which destroyed sheet pairing
 * and produced a garbage comparison. Run with: npm run test:unit
 */

import { test, expect } from '@playwright/test';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import ExcelJS from 'exceljs';
import { compareWorkbooks, readWorkbookSheetInfo, normalizeRow } from '../helpers/excel-compare.helpers';
import { readSheetRowsRaw } from '../helpers/xlsx-sheet-reader.helpers';
import { writeVisualsToSheet, diffPageSets, truncateSafe } from '../helpers/report-export.helpers';
import type { VisualExport } from '../helpers/harness.helpers';

const tmpFiles: string[] = [];

function tmpPath(name: string): string {
  const p = path.join(os.tmpdir(), `verirapport-unit-${process.pid}-${name}`);
  tmpFiles.push(p);
  return p;
}

test.afterAll(() => {
  for (const f of tmpFiles) {
    try { fs.unlinkSync(f); } catch { /* already gone */ }
  }
});

function fakeVisual(title: string, rows: Array<[string, string]>): VisualExport {
  return {
    visualName:  `v-${title}`,
    visualTitle: title,
    visualType:  'tableEx',
    csvData:     null,
    headers:     ['Name', 'Value'],
    rows:        rows.map(([a, b]) => ({ Name: a, Value: b })),
    rowCount:    rows.length,
    errorMessage: null,
  };
}

/** Writes an xlsx exactly the way exportReportToWorkbook does: streaming
 *  writer, one committed sheet per page — worksheets land in the zip BEFORE
 *  workbook.xml, which is the trigger for the reader bug this file guards
 *  against. */
async function writeWorkbook(filePath: string, sheets: Array<{ name: string; visuals: VisualExport[] }>): Promise<void> {
  const wb = new ExcelJS.stream.xlsx.WorkbookWriter({
    filename: filePath,
    useStyles: true,
    useSharedStrings: false,
  });
  for (const s of sheets) {
    const ws = wb.addWorksheet(s.name);
    writeVisualsToSheet(ws, s.visuals);
    ws.commit();
  }
  await wb.commit();
}

// German page names with umlauts + a name pair differing only by case, both
// taken from the real report pair that surfaced the bug.
const PAGE_A = 'Dashboard Bewerbungseingänge';
const PAGE_B = 'Dashboard Absagen';
const PAGE_C_SRC = 'Vertragsrücklauf - Cluster';
const PAGE_C_TGT = 'Vertragsrücklauf - cluster';
// Real production case: the exact same page lost its accent during
// migration ("Commodité" -> "Commodite") and was reported as two different
// pages instead of pairing up.
const PAGE_D_SRC = 'Commodit' + String.fromCharCode(0xe9); // "Commodité"
const PAGE_D_TGT = 'Commodite';

test.describe('compareWorkbooks on stream-written files', () => {
  test('reads back REAL page names, never Sheet1/Sheet2 defaults', async () => {
    const file = tmpPath('names.xlsx');
    await writeWorkbook(file, [
      { name: PAGE_A, visuals: [fakeVisual('KPI', [['a', '1']])] },
      { name: PAGE_B, visuals: [fakeVisual('KPI', [['b', '2']])] },
    ]);

    const infos = await readWorkbookSheetInfo(file);
    expect(infos.map(i => i.name)).toEqual([PAGE_A, PAGE_B]);
  });

  test('identical workbooks compare as identical under their real names', async () => {
    const sheets = [
      { name: PAGE_A, visuals: [fakeVisual('Eingänge nach Monat', [['Jan', '10'], ['Feb', '20']])] },
      { name: PAGE_B, visuals: [fakeVisual('Absagegründe', [['x', '5']])] },
    ];
    const fa = tmpPath('ident-a.xlsx');
    const fb = tmpPath('ident-b.xlsx');
    await writeWorkbook(fa, sheets);
    await writeWorkbook(fb, sheets);

    const diff = await compareWorkbooks(fa, fb);
    expect(diff.dataIdentical).toBe(true);
    expect(diff.sheets.map(s => s.sheet).sort()).toEqual([PAGE_A, PAGE_B].sort());
    expect(diff.sheetsOnlyInExpected).toEqual([]);
    expect(diff.sheetsOnlyInActual).toEqual([]);
  });

  test('pairs sheets whose names differ only by case, and flags a real value diff', async () => {
    const fa = tmpPath('case-a.xlsx');
    const fb = tmpPath('case-b.xlsx');
    await writeWorkbook(fa, [
      { name: PAGE_C_SRC, visuals: [fakeVisual('Cycle Time', [['2024', '12']])] },
    ]);
    await writeWorkbook(fb, [
      { name: PAGE_C_TGT, visuals: [fakeVisual('Cycle Time', [['2024', '99']])] }, // changed value
    ]);

    const diff = await compareWorkbooks(fa, fb);
    // Case-different names must PAIR (same page), not appear as missing sheets.
    expect(diff.sheetsOnlyInExpected).toEqual([]);
    expect(diff.sheetsOnlyInActual).toEqual([]);
    expect(diff.sheets).toHaveLength(1);
    expect(diff.sheets[0].dataIdentical).toBe(false); // and the value diff is real
  });

  test('a sheet genuinely on one side only is still reported as such', async () => {
    const fa = tmpPath('only-a.xlsx');
    const fb = tmpPath('only-b.xlsx');
    await writeWorkbook(fa, [
      { name: PAGE_A, visuals: [fakeVisual('KPI', [['a', '1']])] },
      { name: 'PAGES', visuals: [fakeVisual('Nav', [['n', '0']])] },
    ]);
    await writeWorkbook(fb, [
      { name: PAGE_A, visuals: [fakeVisual('KPI', [['a', '1']])] },
    ]);

    const diff = await compareWorkbooks(fa, fb);
    expect(diff.sheetsOnlyInExpected).toEqual(['PAGES']);
    expect(diff.dataIdentical).toBe(false);
  });

  test('pairs sheets whose names differ only by an accent (migration dropped the accent)', async () => {
    const fa = tmpPath('accent-a.xlsx');
    const fb = tmpPath('accent-b.xlsx');
    await writeWorkbook(fa, [
      { name: PAGE_D_SRC, visuals: [fakeVisual('KPI', [['a', '1']])] },
    ]);
    await writeWorkbook(fb, [
      { name: PAGE_D_TGT, visuals: [fakeVisual('KPI', [['a', '1']])] },
    ]);

    const diff = await compareWorkbooks(fa, fb);
    expect(diff.sheetsOnlyInExpected).toEqual([]);
    expect(diff.sheetsOnlyInActual).toEqual([]);
    expect(diff.sheets).toHaveLength(1);
    expect(diff.dataIdentical).toBe(true);
  });
});

// ── Fast per-sheet reader (xlsx-sheet-reader.helpers.ts) ─────────────────────
//
// compareWorkbooks() now tries readSheetRowsRaw() (decompresses + parses just
// one sheetN.xml entry) before falling back to ExcelJS's buffered whole-
// workbook reader — done to bound peak memory to roughly one sheet's worth
// of data instead of the entire workbook, for every sheet compared. These
// tests prove the fast reader's output is byte-for-byte identical to the old
// buffered-reader-only behavior across the cell shapes this project's own
// writer (report-export.helpers.ts) actually produces, and that a file the
// fast reader can't confidently handle correctly falls back rather than
// producing wrong data.

/** Reads a sheet exactly the way this project's OLD (pre-fast-path) code
 *  always did: full buffered Workbook.xlsx.readFile(), then
 *  ws.eachRow({includeEmpty:true}) + normalizeRow(). Used as the ground
 *  truth the fast reader must match. */
async function readSheetOldWay(filePath: string, sheetName: string): Promise<string[][]> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);
  const ws = wb.worksheets.find(w => w.name === sheetName)!;
  const rows: string[][] = [];
  ws.eachRow({ includeEmpty: true }, (row) => rows.push(normalizeRow(row)));
  return rows;
}

test.describe('readSheetRowsRaw matches the buffered reader exactly', () => {
  test('numbers-as-text, decimals, XML-special chars, unicode/emoji, mixed empty cells', async () => {
    const file = tmpPath('fast-equiv.xlsx');
    await writeWorkbook(file, [{
      name: 'Data',
      visuals: [
        fakeVisual('Numbers', [['42', '18544.000001'], ['-7', '0']]),
        fakeVisual('Special chars', [['Alice "A" & <B>', "O'Brien"], ["it's & <ok>", '  padded  ']]),
        fakeVisual('Unicode', [['Commodité', '中文 🚀'], ['emoji 😀X', 'tab\there']]),
      ],
    }]);

    const [info] = await readWorkbookSheetInfo(file);
    const fast = await readSheetRowsRaw(file, info);
    const old  = await readSheetOldWay(file, 'Data');

    expect(fast).not.toBeNull(); // confirms the fast path actually engaged
    expect(fast).toEqual(old);
  });

  test('a data row where every cell is an empty string (edge case) matches the buffered reader', async () => {
    const file = tmpPath('fast-equiv-blank.xlsx');
    // fakeVisual always uses ['Name','Value'] headers - build one row set that
    // includes an all-empty-string row mixed with real data, matching what
    // writeVisualsToSheet would emit for a Power BI visual that legitimately
    // exported a blank row.
    await writeWorkbook(file, [{
      name: 'Data',
      visuals: [fakeVisual('KPI', [['Alice', '1'], ['', ''], ['Bob', '2']])],
    }]);

    const [info] = await readWorkbookSheetInfo(file);
    const fast = await readSheetRowsRaw(file, info);
    const old  = await readSheetOldWay(file, 'Data');

    expect(fast).not.toBeNull();
    expect(fast).toEqual(old);
  });

  test('multi-sheet workbook: each sheet parses independently and correctly', async () => {
    const file = tmpPath('fast-equiv-multi.xlsx');
    await writeWorkbook(file, [
      { name: PAGE_A, visuals: [fakeVisual('KPI', [['a', '1'], ['b', '2']])] },
      { name: PAGE_B, visuals: [fakeVisual('KPI', [['x', '9']])] },
    ]);

    const infos = await readWorkbookSheetInfo(file);
    for (const info of infos) {
      const fast = await readSheetRowsRaw(file, info);
      const old  = await readSheetOldWay(file, info.name);
      expect(fast).not.toBeNull();
      expect(fast).toEqual(old);
    }
  });

  test('a page with zero visuals (just the [NO DATA VISUALS] marker) still matches', async () => {
    const file = tmpPath('fast-equiv-empty-page.xlsx');
    await writeWorkbook(file, [{ name: 'Empty', visuals: [] }]);

    const [info] = await readWorkbookSheetInfo(file);
    const fast = await readSheetRowsRaw(file, info);
    const old  = await readSheetOldWay(file, 'Empty');

    expect(fast).not.toBeNull();
    expect(fast).toEqual(old);
  });
});

test.describe('fast path falls back safely on a shape it does not recognize', () => {
  test('a shared-strings workbook (t="s" cells) is rejected by the fast reader', async () => {
    // useSharedStrings: true forces ExcelJS to emit t="s" cells referencing
    // xl/sharedStrings.xml - a shape report-export.helpers.ts's own writer
    // never produces (it always uses useSharedStrings: false), and one the
    // fast reader deliberately does not implement, so it must return null
    // here rather than silently mis-parsing.
    const file = tmpPath('shared-strings.xlsx');
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Data');
    ws.addRow(['[VISUAL] KPI [tableEx]']);
    ws.addRow(['Name', 'Value']);
    ws.addRow(['Alice', '1']);
    await wb.xlsx.writeFile(file); // buffered writer defaults to shared strings

    const [info] = await readWorkbookSheetInfo(file);
    const fast = await readSheetRowsRaw(file, info);
    expect(fast).toBeNull();
  });

  test('compareWorkbooks still produces a correct result when the fast path bails on one side', async () => {
    // Expected side: normal streaming-writer output (fast path engages).
    const fileExpected = tmpPath('fallback-expected.xlsx');
    await writeWorkbook(fileExpected, [
      { name: 'Data', visuals: [fakeVisual('KPI', [['Alice', '1'], ['Bob', '2']])] },
    ]);

    // Actual side: written with shared strings on (fast path must bail and
    // fall back to the buffered reader), but with a genuine value difference
    // so the comparison result is verifiable, not just "didn't crash".
    const fileActual = tmpPath('fallback-actual.xlsx');
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Data');
    ws.addRow(['[VISUAL] KPI [tableEx]']);
    ws.addRow(['Name', 'Value']);
    ws.addRow(['Alice', '1']);
    ws.addRow(['Bob', '99']); // changed from 2 -> 99
    await wb.xlsx.writeFile(fileActual);

    const diff = await compareWorkbooks(fileExpected, fileActual);
    expect(diff.sheets).toHaveLength(1);
    expect(diff.sheets[0].dataIdentical).toBe(false); // the 2 -> 99 change must still be caught
    expect(diff.sheets[0].visuals[0].status).toBe('data-diff');
  });
});

test.describe('normalizeRow', () => {
  // Regression test for a production crash: ExcelJS's row.values can be a
  // SPARSE array (a row with no cell at all in some column leaves a real
  // hole, not an explicit undefined). Array.prototype.map skips holes, so
  // a naive `.slice(1).map(normalizeCell)` silently let a hole through as
  // `undefined` instead of '' - crashing the first caller that assumed
  // every cell was always a string (`cells[0].match(...)` in
  // streamSheetBlocks, on a 24k-row real table).
  test('a hole in a sparse row.values becomes "", never undefined', () => {
    const sparse: unknown[] = [];
    sparse[0] = undefined; // the always-unused 0-index placeholder
    sparse[2] = 'present'; // column B has a value; column A (index 1) is a genuine hole
    expect(1 in sparse).toBe(false); // confirms this really is a hole, not an explicit undefined

    const fakeRow = { values: sparse } as unknown as ExcelJS.Row;
    const cells = normalizeRow(fakeRow);

    expect(cells).toEqual(['', 'present']);
    expect(cells[0]).not.toBeUndefined();
  });

  test('a normal, fully-populated row is unaffected', () => {
    const fakeRow = { values: [undefined, 'a', 'b', 'c'] } as unknown as ExcelJS.Row;
    expect(normalizeRow(fakeRow)).toEqual(['a', 'b', 'c']);
  });
});

test.describe('diffPageSets', () => {
  test('trims and case-folds before matching (real drift seen in production)', () => {
    const d = diffPageSets(
      ['Vertragsversand', 'Vertragsrücklauf - Cluster', 'Absagen'],
      ['Vertragsversand ', 'Vertragsrücklauf - cluster', 'Absagen', 'PAGES', 'Page 4'],
    );
    expect(d.inBoth).toHaveLength(3);
    expect(d.onlyInSource).toEqual([]);
    expect(d.onlyInTarget).toEqual(['PAGES', 'Page 4']);
  });

  test('applies pageMap (source name -> target name) before matching', () => {
    const d = diffPageSets(['DU Head'], ['DU-Head'], { 'DU Head': 'DU-Head' });
    expect(d.inBoth).toEqual(['DU Head']);
    expect(d.onlyInSource).toEqual([]);
    expect(d.onlyInTarget).toEqual([]);
  });
});

test.describe('truncateSafe', () => {
  // Regression guard for Excel's 31-char sheet-name limit: a page name
  // containing an emoji or other astral-plane character is a surrogate
  // PAIR (2 UTF-16 code units for 1 visible character). Plain .slice()
  // indexes by code unit and could cut exactly between the two halves,
  // producing a malformed, unpaired-surrogate string.
  test('never splits a surrogate pair (emoji) in half', () => {
    const emoji = String.fromCodePoint(0x1f600); // one grinning-face emoji, a real surrogate pair
    const name = 'Dashboard ' + emoji + emoji + emoji + ' Overview'; // well past 31 chars
    const truncated = truncateSafe(name, 12);

    // Every code point in the result round-trips through fromCodePoint - if
    // a surrogate were split, Array.from would already have thrown this off
    // by producing a lone (unpaired) surrogate character instead.
    for (const ch of truncated) {
      const cp = ch.codePointAt(0)!;
      expect(String.fromCodePoint(cp)).toBe(ch);
    }
    expect(Array.from(truncated).length).toBeLessThanOrEqual(12);
  });

  test('behaves like a normal length-limited truncation for plain ASCII/Latin text', () => {
    expect(truncateSafe('Dashboard Overview', 9)).toBe('Dashboard');
    expect(truncateSafe('Bewerbungseingänge', 100)).toBe('Bewerbungseingänge'); // shorter than max: unchanged
  });
});
