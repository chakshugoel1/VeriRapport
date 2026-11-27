/**
 * xlsx-sheet-reader.helpers.ts
 *
 * Reads ONE worksheet's rows directly from an .xlsx's zip central directory,
 * bypassing ExcelJS's buffered Workbook.xlsx.readFile() entirely for the
 * common case. This exists purely to bound peak memory: readFile() parses
 * EVERY sheet in the file into ExcelJS's (heavy — every cell is its own JS
 * object) in-memory model, even when only one sheet is needed right now.
 * Decompressing and regex-parsing a single sheetN.xml entry costs roughly
 * "one sheet's worth of XML text", independent of how many other pages the
 * workbook has.
 *
 * Deliberately narrow, not a general OOXML reader: this project's own writer
 * (report-export.helpers.ts, via ExcelJS's streaming WorkbookWriter with
 * useSharedStrings: false) was inspected directly to ground this parser in
 * its ACTUAL output shape — every data/header/marker cell comes out as
 * <c t="str"><v>text</v></c> (plain values, XML-entity-escaped, no shared-
 * strings table). parseSheetXml() targets exactly that shape, plus the
 * standard inline-string/boolean/error/default-numeric variants for
 * resilience against a future ExcelJS version changing its serialization —
 * but the instant it meets a cell/row shape it doesn't recognize, it
 * returns null rather than guessing. The caller (excel-compare.helpers.ts)
 * treats null as "fast path unavailable for this file" and falls back to
 * the existing, fully general ExcelJS buffered reader — so correctness
 * never depends on this parser's coverage being complete, only its
 * detection of its own limits being reliable.
 */

// Central-directory zip access — same technique readWorkbookSheetInfo() in
// excel-compare.helpers.ts already uses for workbook.xml. Already a
// transitive dependency of exceljs; declared direct in package.json.
// eslint-disable-next-line @typescript-eslint/no-var-requires
const unzipper = require('unzipper');

/** Minimal shape needed to locate a sheet's zip entry — matches SheetInfo
 *  from excel-compare.helpers.ts without creating a runtime import of it
 *  (only the type is needed, and TS type-only imports are erased, but
 *  spelling it out locally keeps this module for what it is: a standalone
 *  zip/XML reader with no dependency on the comparison logic). */
export interface SheetLocation {
  name:   string;
  fileNo: string;
}

/** Decodes the small set of entities OOXML actually emits (numeric refs +
 *  the five XML predefined entities). Shared with readWorkbookSheetInfo()'s
 *  workbook.xml parsing in excel-compare.helpers.ts — one implementation,
 *  used everywhere raw OOXML text needs decoding. */
export function decodeXmlEntities(text: string): string {
  return text
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(parseInt(d, 10)))
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
}

// Only plain decimals (e.g. "18544.000001"), never bare integers — touching
// integers risks corrupting IDs like "007" that must keep their exact text.
const PLAIN_DECIMAL_RE = /^-?\d+\.\d+$/;
const NUMERIC_ROUND_DECIMALS = 4;

/**
 * Collapses floating-point noise in decimal-looking text (e.g. a DAX SUM that
 * lands on "18543.999999997" instead of "18544") by rounding to a fixed number
 * of decimals and trimming trailing zeros. Bare integers and non-numeric text
 * are returned untouched, so this never risks reinterpreting an ID or code.
 *
 * Shared between excel-compare.helpers.ts's normalizeCell() (the ExcelJS-
 * object-model path) and parseSheetXml() below (the raw-XML path) — both
 * must apply IDENTICAL rounding to the same underlying text, or a value read
 * via one path could compare unequal to the same value read via the other.
 */
export function normalizeNumericDrift(text: string): string {
  // A percentage written two ways ("100%" vs "100.0%") is the same value.
  const pct = /^(-?\d+(?:\.\d+)?)%$/.exec(text);
  if (pct) {
    const n = Number(pct[1]);
    if (!Number.isFinite(n)) return text;
    const r = n.toFixed(NUMERIC_ROUND_DECIMALS).replace(/0+$/, '').replace(/\.$/, '');
    return (r === '-0' ? '0' : r) + '%';
  }
  if (!PLAIN_DECIMAL_RE.test(text)) return text;
  const n = Number(text);
  if (!Number.isFinite(n)) return text;
  const rounded = n.toFixed(NUMERIC_ROUND_DECIMALS).replace(/0+$/, '').replace(/\.$/, '');
  // A tiny negative (e.g. -0.00001) rounds to "-0"; treat it as plain 0 so it
  // does not compare unequal to the same tiny positive value.
  return rounded === '-0' ? '0' : rounded;
}

/** "A" -> 1, "Z" -> 26, "AA" -> 27, ... (1-based, matches OOXML column refs). */
function columnLettersToIndex(letters: string): number {
  let n = 0;
  for (let i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
  return n;
}

/** Splits a cell ref like "AB123" into its column letters and 1-based row number. */
function splitCellRef(ref: string): { col: number; row: number } | null {
  const m = ref.match(/^([A-Z]+)(\d+)$/);
  if (!m) return null;
  return { col: columnLettersToIndex(m[1]), row: parseInt(m[2], 10) };
}

/**
 * Reads <mergeCells><mergeCell ref="A1:B1"/>...</mergeCells> (a sibling of
 * <sheetData>, not inside it). ExcelJS's own buffered reader propagates a
 * merged range's anchor (top-left) value into every other cell the merge
 * covers when it hands back row.values — verified directly against this
 * project's writer output: writeVisualsToSheet() merges the [VISUAL] marker
 * row across every header column, and the raw XML for the non-anchor cells
 * in that range is a genuinely empty, valueless <c r="B1"/> (no <v> at all).
 * Without replicating this, the fast reader would (correctly, per the raw
 * bytes) treat those cells as holes, while the buffered reader reports them
 * as duplicating the anchor's text — a real, observable divergence for
 * every multi-column visual, not an edge case.
 */
function parseMergeRanges(xml: string): Array<{ startCol: number; startRow: number; endCol: number; endRow: number }> {
  const ranges: Array<{ startCol: number; startRow: number; endCol: number; endRow: number }> = [];
  const mergeCellsMatch = xml.match(/<mergeCells\b[^>]*>([\s\S]*?)<\/mergeCells>/);
  if (!mergeCellsMatch) return ranges;
  const refRe = /<mergeCell\b[^>]*\bref="([^":]+):([^"]+)"/g;
  let m: RegExpExecArray | null;
  while ((m = refRe.exec(mergeCellsMatch[1]))) {
    const start = splitCellRef(m[1]);
    const end   = splitCellRef(m[2]);
    if (!start || !end) continue;
    ranges.push({ startCol: start.col, startRow: start.row, endCol: end.col, endRow: end.row });
  }
  return ranges;
}

/**
 * Parses one sheetN.xml into rows, matching EXACTLY what
 * normalizeRow(ws.eachRow row) produces today for this project's own writer
 * output (verified empirically — see the module comment): a row's returned
 * array length is driven by the highest column that had REAL raw content in
 * the XML (even whitespace-only content counts — only a truly empty <v></v>,
 * or a valueless <c/>, is a hole), and every value is trimmed + numeric-
 * drift-normalized the same way normalizeCell() does. Returns null the
 * instant it meets anything outside the recognized cell-type set, so the
 * caller can fall back safely.
 *
 * One more ExcelJS quirk this replicates: ws.eachRow({includeEmpty:true})
 * synthesizes an empty [] for a row NUMBER that has no <row> element at all
 * (writeVisualsToSheet's spacer rows — ws.addRow([]) with zero cells — are
 * omitted from the XML entirely, verified directly), but only for INTERIOR
 * gaps between rows that do exist; it does not invent trailing rows past the
 * last one actually present. Rows are therefore built keyed by their real
 * `r="N"` row number, and gaps up to the highest such number are filled in
 * afterward — not by pushing sequentially as each <row> is matched.
 */
function parseSheetXml(xml: string): string[][] | null {
  const sdMatch = xml.match(/<sheetData\b[^>]*?(?:\/>|>([\s\S]*?)<\/sheetData>)/);
  if (!sdMatch) return null; // no <sheetData> at all — not a shape we recognize
  const body = sdMatch[1] ?? ''; // self-closing <sheetData/> = zero rows

  const rowsByNumber = new Map<number, string[]>(); // 1-based row number -> parsed row
  let maxRowNumber = 0;

  const rowRe = /<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g;
  const cellRe = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
  const rowNumAttrRe = /\br="(\d+)"/;
  const refAttrRe = /\br="([A-Z]+)\d+"/;
  const typeAttrRe = /\bt="([a-zA-Z]+)"/;
  const vRe = /<v[^>]*>([\s\S]*?)<\/v>/;
  const tRunRe = /<t\b[^>]*>([\s\S]*?)<\/t>/g;

  let rowMatch: RegExpExecArray | null;
  while ((rowMatch = rowRe.exec(body))) {
    const rowAttrs = rowMatch[1];
    const rowInner = rowMatch[2];
    const rowNumStr = rowAttrs.match(rowNumAttrRe)?.[1];
    if (!rowNumStr) return null; // this writer always numbers its rows — unrecognized otherwise
    const rowNum = parseInt(rowNumStr, 10);
    if (rowNum > maxRowNumber) maxRowNumber = rowNum;

    if (rowInner === undefined) { rowsByNumber.set(rowNum, []); continue; } // self-closing <row/>

    const sparse: string[] = [];
    let maxCol = 0; // 1-based highest column with real (possibly whitespace-only) raw content

    let cellMatch: RegExpExecArray | null;
    while ((cellMatch = cellRe.exec(rowInner))) {
      const attrs = cellMatch[1];
      const cellInner = cellMatch[2]; // undefined for a self-closing <c/> (no value at all)

      const refM = attrs.match(refAttrRe);
      if (!refM) return null; // malformed/unexpected — bail out, let the caller fall back
      const colIdx = columnLettersToIndex(refM[1]) - 1; // 0-based

      if (cellInner === undefined) continue; // no <v> at all — hole, doesn't extend the row

      const type = attrs.match(typeAttrRe)?.[1] ?? 'n';
      let rawText: string;
      if (type === 'str' || type === 'n' || type === 'b' || type === 'e') {
        const vM = cellInner.match(vRe);
        rawText = vM ? decodeXmlEntities(vM[1]) : '';
      } else if (type === 'inlineStr') {
        const parts: string[] = [];
        let tM: RegExpExecArray | null;
        while ((tM = tRunRe.exec(cellInner))) parts.push(decodeXmlEntities(tM[1]));
        rawText = parts.join('');
      } else {
        // Shared string (t="s"), date (t="d"), or anything else this writer
        // never produces — not confidently handled here, fall back.
        return null;
      }

      if (rawText === '') continue; // <v></v>, truly empty — a hole, same as no <v> at all

      // Real (possibly whitespace-only) content: this column IS populated —
      // always extends the row, even if trimming below reduces it to ''.
      // Matches ExcelJS: only a literally-empty <v> is a hole.
      if (colIdx + 1 > maxCol) maxCol = colIdx + 1;
      sparse[colIdx] = normalizeNumericDrift(rawText.trim());
    }

    rowsByNumber.set(rowNum, Array.from({ length: maxCol }, (_, i) => sparse[i] ?? ''));
  }

  // Merge-range propagation (see parseMergeRanges' doc comment) — must run
  // before the interior-gap fill below reads row numbers as final, but after
  // every real row is parsed.
  for (const range of parseMergeRanges(xml)) {
    const anchorRow = rowsByNumber.get(range.startRow);
    const anchorValue = anchorRow?.[range.startCol - 1] ?? '';
    if (anchorValue === '') continue; // nothing to propagate

    for (let r = range.startRow; r <= range.endRow; r++) {
      const row = rowsByNumber.get(r);
      if (!row) continue;
      const endCol0 = range.endCol - 1;
      if (row.length <= endCol0) row.length = endCol0 + 1; // extend with holes, filled below
      for (let c = range.startCol - 1; c <= endCol0; c++) {
        if (r === range.startRow && c === range.startCol - 1) continue; // skip the anchor itself
        if (row[c] === undefined || row[c] === '') row[c] = anchorValue;
      }
    }
  }

  // Fill interior gaps (a row number with no <row> element at all, between
  // 1 and the highest row number actually present) with [] — matches
  // ws.eachRow({includeEmpty:true}); never invents rows past maxRowNumber.
  const rows: string[][] = [];
  for (let n = 1; n <= maxRowNumber; n++) rows.push(rowsByNumber.get(n) ?? []);
  return rows;
}

/**
 * Reads one worksheet's rows straight from the zip, without touching
 * ExcelJS's buffered Workbook model. Returns null (never throws for a
 * recognized-but-unhandled shape) when the fast path can't confidently
 * handle this file — the caller is expected to fall back to the buffered
 * reader in that case, not treat null as an error.
 */
export async function readSheetRowsRaw(filePath: string, target: SheetLocation): Promise<string[][] | null> {
  try {
    const dir = await unzipper.Open.file(filePath);
    const entryPath = `xl/worksheets/sheet${target.fileNo}.xml`;
    const entry = dir.files.find((f: any) => f.path === entryPath);
    if (!entry) return null; // let the buffered path's name/sheetId-aware lookup decide
    const xml = (await entry.buffer()).toString('utf8');
    return parseSheetXml(xml);
  } catch {
    return null;
  }
}
