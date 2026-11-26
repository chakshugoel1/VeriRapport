/**
 * field-name.helpers.ts
 *
 * Pure helpers for treating field names from two reports as "the same field"
 * when they only differ in spelling style: "Project Number" vs PROJECT_NUMBER
 * vs project-number. Used by the KPI-mapping comparison (mapped-compare) and
 * by the filter translation. No I/O.
 */

const COMBINING_DIACRITICS_RE = new RegExp(
  '[' + String.fromCharCode(0x0300) + '-' + String.fromCharCode(0x036f) + ']', 'g',
);

/** Aggregation prefixes Power BI adds to export headers ("Sum of Amount"). */
const AGG_PREFIX_RE = /^(sum|count|distinct count|count of distinct|average|min|max|minimum|maximum|first|last|median|standard deviation|variance)\s+of\s+/i;

/** "Table[Column]" / "'Table'[Column]" -> { table, column }. A bare name has no table. */
export function splitFieldRef(ref: string): { table?: string; column: string } {
  const t = ref.trim();
  const m = t.match(/^(.*)\[(.*)\]$/s);
  if (!m) return { column: t };
  const table = m[1].trim().replace(/^'(.*)'$/, '$1');
  return { table: table || undefined, column: m[2].trim() };
}

/**
 * Spelling-insensitive key for a field's COLUMN name: accents, case, spaces,
 * underscores, dashes and any other symbols are ignored, and Power BI's
 * "Sum of " style prefix is dropped. "Project Number", "PROJECT_NUMBER" and
 * "Sum of project-number" all give "projectnumber".
 */
export function normalizeFieldName(name: string): string {
  const column = splitFieldRef(name).column;
  return column
    .normalize('NFD').replace(COMBINING_DIACRITICS_RE, '')
    .replace(AGG_PREFIX_RE, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '');
}

/** Key for a full "Table[Column]" reference (table and column both normalised). */
export function normalizeFieldRef(ref: string): string {
  const { table, column } = splitFieldRef(ref);
  const t = (table ?? '').normalize('NFD').replace(COMBINING_DIACRITICS_RE, '').toLowerCase().replace(/[^a-z0-9]+/g, '');
  return `${t}|${normalizeFieldName(column)}`;
}

/** Splits a mapping-file field list ("A[x]; B[y]") into trimmed, non-empty entries. */
export function parseFieldList(text: string | undefined | null): string[] {
  if (!text) return [];
  return String(text).split(/[;\n]/).map(s => s.trim()).filter(Boolean);
}
