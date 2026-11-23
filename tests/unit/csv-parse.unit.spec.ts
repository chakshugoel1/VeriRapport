import { test, expect } from '@playwright/test';
import { parseVisualCsv } from '../helpers/harness.helpers';
import { normalizeNumericDrift } from '../helpers/xlsx-sheet-reader.helpers';

test.describe('parseVisualCsv', () => {
  test('parses a simple CSV with CRLF line endings', () => {
    const r = parseVisualCsv('Region,Sales\r\nEU,10\r\nUS,20\r\n');
    expect(r.headers).toEqual(['Region', 'Sales']);
    expect(r.rows).toEqual([{ Region: 'EU', Sales: '10' }, { Region: 'US', Sales: '20' }]);
  });

  test('keeps a quoted multi-line value in ONE row (was split into bogus rows)', () => {
    const r = parseVisualCsv('Name,Note\n"A","line1\nline2"\n"B","x"');
    expect(r.rows).toEqual([
      { Name: 'A', Note: 'line1\nline2' },
      { Name: 'B', Note: 'x' },
    ]);
  });

  test('handles escaped quotes and commas inside quotes', () => {
    const r = parseVisualCsv('A,B\n"he said ""hi""","x,y"');
    expect(r.rows).toEqual([{ A: 'he said "hi"', B: 'x,y' }]);
  });

  test('duplicate headers keep their own column data', () => {
    const r = parseVisualCsv('Region,Sales,Sales\nEU,1,2');
    expect(r.headers).toEqual(['Region', 'Sales', 'Sales (2)']);
    expect(r.rows).toEqual([{ Region: 'EU', Sales: '1', 'Sales (2)': '2' }]);
  });

  test('duplicate-header suffix never collides with an existing header', () => {
    const r = parseVisualCsv('Sales,Sales,Sales (2)\n1,2,3');
    expect(new Set(r.headers).size).toBe(3);
    expect(r.rows[0][r.headers[0]]).toBe('1');
    expect(r.rows[0][r.headers[1]]).toBe('2');
    expect(r.rows[0][r.headers[2]]).toBe('3');
  });

  test('skips blank lines and returns empty for empty input', () => {
    expect(parseVisualCsv('')).toEqual({ headers: [], rows: [] });
    const r = parseVisualCsv('A\n\n1\n   \n2\n');
    expect(r.rows).toEqual([{ A: '1' }, { A: '2' }]);
  });

  test('short rows are padded with empty strings', () => {
    const r = parseVisualCsv('A,B,C\n1,2');
    expect(r.rows).toEqual([{ A: '1', B: '2', C: '' }]);
  });
});

test.describe('normalizeNumericDrift', () => {
  test('tiny negative and tiny positive both normalise to 0', () => {
    expect(normalizeNumericDrift('-0.00001')).toBe('0');
    expect(normalizeNumericDrift('0.00001')).toBe('0');
  });
  test('still rounds drift and leaves integers/text alone', () => {
    expect(normalizeNumericDrift('18543.999999997')).toBe('18544');
    expect(normalizeNumericDrift('007')).toBe('007');
    expect(normalizeNumericDrift('abc')).toBe('abc');
    expect(normalizeNumericDrift('-1.50')).toBe('-1.5');
  });
});
