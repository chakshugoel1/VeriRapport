/**
 * Filter values must be sent with the column's own type: a text "1065110" on
 * a numeric column silently matches nothing in Power BI.
 */
import { test, expect } from '@playwright/test';
import { coerceFilterValues } from '../helpers/slicer-config.helpers';
import { sampleValueKind } from '../helpers/pbi-api.helpers';

test.describe('coerceFilterValues', () => {
  test('numeric column: plain numbers are sent as numbers', () => {
    expect(coerceFilterValues(['1065110'], 'number')).toEqual([1065110]);
    expect(coerceFilterValues([' 12.5 ', '-3'], 'number')).toEqual([12.5, -3]);
  });
  test('text column or unknown kind: values unchanged', () => {
    expect(coerceFilterValues(['1065110'], 'text')).toEqual(['1065110']);
    expect(coerceFilterValues(['1065110'], undefined)).toEqual(['1065110']);
    expect(coerceFilterValues(['1065110'], 'unknown')).toEqual(['1065110']);
  });
  test('numeric column but a value is not a plain number: left unchanged (no guessing)', () => {
    expect(coerceFilterValues(['1,065,110'], 'number')).toEqual(['1,065,110']);
    expect(coerceFilterValues(['12', 'N/A'], 'number')).toEqual(['12', 'N/A']);
    expect(coerceFilterValues(['007'], 'number')).toEqual([7]);
  });
});

test.describe('sampleValueKind', () => {
  const res = (v: unknown) => ({ tables: [{ rows: [{ '[v]': v }] }] });
  test('reads the column kind from the probe sample', () => {
    expect(sampleValueKind(res(1065110))).toBe('number');
    expect(sampleValueKind(res('1065110'))).toBe('text');
    expect(sampleValueKind(res('2026-11-21T00:00:00'))).toBe('other');
    expect(sampleValueKind(res(true))).toBe('other');
    expect(sampleValueKind(res(null))).toBe('unknown');
    expect(sampleValueKind({ tables: [{ rows: [] }] })).toBe('unknown');
    expect(sampleValueKind(undefined)).toBe('unknown');
  });
});
