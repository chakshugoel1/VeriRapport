/**
 * pbi-api.unit.spec.ts
 *
 * Pure-logic tests for pbi-api.helpers.ts. Covers looksLikePlaceholder — the
 * guard that stops an unfilled pbi-service-principal.json template ("<from
 * IT>" values) from being used verbatim as credentials, which produced a
 * confusing HTTP failure on every embed instead of one clean "not configured"
 * message. Run with: npm run test:unit
 */

import { test, expect } from '@playwright/test';
import { looksLikePlaceholder } from '../helpers/pbi-api.helpers';

test.describe('looksLikePlaceholder', () => {
  test('flags the template boilerplate values', () => {
    expect(looksLikePlaceholder('<from IT>')).toBe(true);
    expect(looksLikePlaceholder('REPLACE_WITH_CLIENT_ID')).toBe(true);
    expect(looksLikePlaceholder('replace_with_secret')).toBe(true);
  });

  test('flags empty and whitespace-only values', () => {
    expect(looksLikePlaceholder('')).toBe(true);
    expect(looksLikePlaceholder('   ')).toBe(true);
    expect(looksLikePlaceholder(undefined)).toBe(true);
    expect(looksLikePlaceholder(null)).toBe(true);
  });

  test('accepts real-looking values', () => {
    expect(looksLikePlaceholder('00000000-0000-0000-0000-000000000000')).toBe(false);
    expect(looksLikePlaceholder('a-real-client-secret~value')).toBe(false);
  });
});
