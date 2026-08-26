// Host-independent Excel read/search/verify helpers.
//
// These are deliberately pure so they can be unit tested without an Office
// runtime. `src/office/host.ts` binds them to live Excel.run calls.

export type ExcelReadMode = "compact" | "csv" | "detailed";

export interface ExcelSheetOverview {
  name: string;
  usedRange: string;
  rowCount: number;
  columnCount: number;
  visible?: boolean;
  headers?: unknown[];
  tables?: string[];
  charts?: string[];
}

export interface ExcelWorkbookOverviewInput {
  workbookName?: string;
  sheets: ExcelSheetOverview[];
  namedRanges?: Array<{ name: string; scope?: string; refersTo?: string }>;
  activeSheet?: string;
  selection?: string;
}

export interface ExcelSearchMatch {
  sheet: string;
  address: string;
  value: string;
  formula: string;
}

const MAX_CELL_CHARS = 200;

export function excelColumnLetters(index: number) {
  let column = Math.max(0, Math.trunc(index));
  let name = "";
  do {
    name = String.fromCharCode(65 + (column % 26)) + name;
    column = Math.floor(column / 26) - 1;
  } while (column >= 0);
  return name;
}

export function excelColumnIndex(letters: string) {
  return String(letters || "A")
    .toUpperCase()
    .split("")
    .reduce((sum, char) => sum * 26 + (char.charCodeAt(0) - 64), 0) - 1;
}

export function rangeStartCell(address: unknown) {
  const match = String(address ?? "A1").replace(/\$/g, "").match(/([A-Z]{1,3})(\d{1,7})/i);
  if (!match) return { column: 0, row: 0 };
  return { column: excelColumnIndex(match[1]), row: Number(match[2]) - 1 };
}

export function cellAddress(columnIndex: number, rowIndex: number) {
  return `${excelColumnLetters(columnIndex)}${rowIndex + 1}`;
}

function cellText(value: unknown) {
  if (value === null || value === undefined) return "";
  const text = typeof value === "string" ? value : String(value);
  return text.length > MAX_CELL_CHARS ? `${text.slice(0, MAX_CELL_CHARS)}...` : text;
}

function csvCell(value: unknown) {
  const text = cellText(value);
  return /[",\n\r]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

export function normalizeReadMode(value: unknown): ExcelReadMode {
  const mode = String(value ?? "compact").trim().toLowerCase();
  if (mode === "csv") return "csv";
  if (mode === "detailed" || mode === "full") return "detailed";
  return "compact";
}

/**
 * Render a rectangular block of cells as model-friendly text.
 *
 * compact  - markdown-ish tab table with a leading address column
 * csv      - RFC4180-escaped CSV, no address column
 * detailed - one line per non-empty cell with value + formula + number format
 */
export function formatRangeForModel(
  values: unknown[][],
  options: {
    mode?: unknown;
    baseAddress?: unknown;
    formulas?: unknown[][];
    numberFormats?: unknown[][];
    maxRows?: number;
    maxColumns?: number;
  } = {},
) {
  const mode = normalizeReadMode(options.mode);
  const start = rangeStartCell(options.baseAddress ?? "A1");
  const allRows = Array.isArray(values) ? values : [];
  const maxRows = Math.max(1, Math.trunc(Number(options.maxRows) || 200));
  const maxColumns = Math.max(1, Math.trunc(Number(options.maxColumns) || 50));
  const rows = allRows.slice(0, maxRows);
  const truncatedRows = allRows.length - rows.length;
  const formulas = Array.isArray(options.formulas) ? options.formulas : [];
  const numberFormats = Array.isArray(options.numberFormats) ? options.numberFormats : [];

  let truncatedColumns = 0;
  const lines: string[] = [];

  if (mode === "detailed") {
    rows.forEach((row, rowIndex) => {
      const cells = Array.isArray(row) ? row : [row];
      truncatedColumns = Math.max(truncatedColumns, cells.length - maxColumns);
      cells.slice(0, maxColumns).forEach((value, columnIndex) => {
        const formula = cellText(formulas?.[rowIndex]?.[columnIndex]);
        const text = cellText(value);
        if (!text && !formula) return;
        const address = cellAddress(start.column + columnIndex, start.row + rowIndex);
        const numberFormat = cellText(numberFormats?.[rowIndex]?.[columnIndex]);
        const parts = [`${address}: ${text || "(empty)"}`];
        if (formula && formula !== text) parts.push(`formula=${formula}`);
        if (numberFormat && numberFormat !== "General") parts.push(`format=${numberFormat}`);
        lines.push(parts.join(" | "));
      });
    });
  } else {
    rows.forEach((row, rowIndex) => {
      const cells = Array.isArray(row) ? row : [row];
      truncatedColumns = Math.max(truncatedColumns, cells.length - maxColumns);
      const visible = cells.slice(0, maxColumns);
      if (mode === "csv") {
        lines.push(visible.map(csvCell).join(","));
      } else {
        const address = cellAddress(start.column, start.row + rowIndex);
        lines.push([address, ...visible.map(cellText)].join("\t"));
      }
    });
  }

  const notes: string[] = [];
  if (truncatedRows > 0) notes.push(`${truncatedRows} more rows not shown (increase maxRows to read them).`);
  if (truncatedColumns > 0) notes.push(`${truncatedColumns} more columns not shown (increase maxColumns to read them).`);

  return {
    mode,
    text: lines.join("\n"),
    rowsShown: rows.length,
    truncatedRows: Math.max(0, truncatedRows),
    truncatedColumns: Math.max(0, truncatedColumns),
    notes,
  };
}

/** Build a compact structural blueprint of the workbook for model grounding. */
export function buildWorkbookOverview(input: ExcelWorkbookOverviewInput) {
  const sheets = Array.isArray(input?.sheets) ? input.sheets : [];
  const lines: string[] = [];
  lines.push(`Workbook: ${input?.workbookName || "Untitled workbook"}`);
  if (input?.activeSheet) lines.push(`Active sheet: ${input.activeSheet}`);
  if (input?.selection) lines.push(`Current selection: ${input.selection}`);
  lines.push(`Sheets (${sheets.length}):`);

  for (const sheet of sheets) {
    const detail: string[] = [];
    detail.push(`used ${sheet.usedRange || "empty"}`);
    if (sheet.rowCount || sheet.columnCount) detail.push(`${sheet.rowCount || 0}r x ${sheet.columnCount || 0}c`);
    if (sheet.visible === false) detail.push("hidden");
    if (sheet.tables?.length) detail.push(`tables: ${sheet.tables.join(", ")}`);
    if (sheet.charts?.length) detail.push(`charts: ${sheet.charts.join(", ")}`);
    lines.push(`- ${sheet.name} (${detail.join("; ")})`);
    const headers = (sheet.headers ?? []).map(cellText).filter(Boolean);
    if (headers.length) lines.push(`  headers: ${headers.slice(0, 30).join(" | ")}`);
  }

  const names = Array.isArray(input?.namedRanges) ? input.namedRanges : [];
  if (names.length) {
    lines.push(`Named ranges (${names.length}):`);
    for (const name of names.slice(0, 50)) {
      lines.push(`- ${name.name}${name.scope ? ` [${name.scope}]` : ""}${name.refersTo ? ` -> ${name.refersTo}` : ""}`);
    }
  }

  return {
    text: lines.join("\n"),
    sheetCount: sheets.length,
    namedRangeCount: names.length,
  };
}

function matchesQuery(text: string, query: string, matchCase: boolean, wholeCell: boolean) {
  if (!text) return false;
  const haystack = matchCase ? text : text.toLowerCase();
  const needle = matchCase ? query : query.toLowerCase();
  return wholeCell ? haystack.trim() === needle.trim() : haystack.includes(needle);
}

/** Find cells whose value or formula matches a query, returning real addresses. */
export function searchRangeMatches(
  sheetName: string,
  values: unknown[][],
  formulas: unknown[][] | undefined,
  baseAddress: unknown,
  query: unknown,
  options: { matchCase?: boolean; wholeCell?: boolean; searchFormulas?: boolean; limit?: number } = {},
): ExcelSearchMatch[] {
  const needle = String(query ?? "").trim();
  if (!needle) return [];
  const start = rangeStartCell(baseAddress ?? "A1");
  const rows = Array.isArray(values) ? values : [];
  const formulaRows = Array.isArray(formulas) ? formulas : [];
  const matchCase = options.matchCase === true;
  const wholeCell = options.wholeCell === true;
  const searchFormulas = options.searchFormulas !== false;
  const limit = Math.max(1, Math.trunc(Number(options.limit) || 200));
  const matches: ExcelSearchMatch[] = [];

  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    const row = Array.isArray(rows[rowIndex]) ? (rows[rowIndex] as unknown[]) : [rows[rowIndex]];
    for (let columnIndex = 0; columnIndex < row.length; columnIndex += 1) {
      if (matches.length >= limit) return matches;
      const value = cellText(row[columnIndex]);
      const formula = cellText(formulaRows?.[rowIndex]?.[columnIndex]);
      const hit = matchesQuery(value, needle, matchCase, wholeCell)
        || (searchFormulas && matchesQuery(formula, needle, matchCase, wholeCell));
      if (!hit) continue;
      matches.push({
        sheet: sheetName,
        address: cellAddress(start.column + columnIndex, start.row + rowIndex),
        value,
        formula: formula && formula !== value ? formula : "",
      });
    }
  }
  return matches;
}

export function formatSearchMatches(matches: ExcelSearchMatch[], query: unknown, limit = 200) {
  if (!matches.length) return `No cells matched "${String(query ?? "")}".`;
  const shown = matches.slice(0, limit);
  const lines = shown.map((match) => {
    const suffix = match.formula ? ` (${match.formula})` : "";
    return `${match.sheet}!${match.address}: ${match.value}${suffix}`;
  });
  if (matches.length > shown.length) lines.push(`...and ${matches.length - shown.length} more matches.`);
  return [`${matches.length} match${matches.length === 1 ? "" : "es"} for "${String(query ?? "")}":`, ...lines].join("\n");
}

/**
 * Overwrite protection: report which existing cells a write would destroy.
 * Callers use this to refuse a destructive write unless `overwrite` is set.
 */
export function describeOverwriteConflicts(existingValues: unknown[][] | undefined, baseAddress: unknown) {
  const rows = Array.isArray(existingValues) ? existingValues : [];
  const start = rangeStartCell(baseAddress ?? "A1");
  const occupied: string[] = [];

  rows.forEach((row, rowIndex) => {
    const cells = Array.isArray(row) ? row : [row];
    cells.forEach((value, columnIndex) => {
      if (cellText(value).trim() === "") return;
      occupied.push(cellAddress(start.column + columnIndex, start.row + rowIndex));
    });
  });

  return {
    count: occupied.length,
    addresses: occupied.slice(0, 25),
    summary: occupied.length
      ? `${occupied.length} non-empty cell${occupied.length === 1 ? "" : "s"} would be overwritten (${occupied.slice(0, 10).join(", ")}${occupied.length > 10 ? ", ..." : ""}).`
      : "",
  };
}

/** Post-write verification: compare what we asked Excel to write with what it stored. */
export function verifyWrittenValues(expected: unknown[][], actual: unknown[][] | undefined, baseAddress: unknown) {
  const expectedRows = Array.isArray(expected) ? expected : [];
  const actualRows = Array.isArray(actual) ? actual : [];
  const start = rangeStartCell(baseAddress ?? "A1");
  const mismatches: Array<{ address: string; expected: string; actual: string }> = [];
  let checked = 0;

  expectedRows.forEach((row, rowIndex) => {
    const cells = Array.isArray(row) ? row : [row];
    cells.forEach((value, columnIndex) => {
      const want = cellText(value).trim();
      // Formulas are stored as their computed value, so verification only
      // asserts literal writes. Formula writes are verified separately.
      if (want.startsWith("=")) return;
      checked += 1;
      const got = cellText(actualRows?.[rowIndex]?.[columnIndex]).trim();
      if (want === got) return;
      if (want !== "" && got !== "" && Number(want) === Number(got) && want !== "" && !Number.isNaN(Number(want))) return;
      mismatches.push({ address: cellAddress(start.column + columnIndex, start.row + rowIndex), expected: want, actual: got });
    });
  });

  return {
    verified: mismatches.length === 0,
    checked,
    mismatches: mismatches.slice(0, 20),
    summary: mismatches.length
      ? `Verification found ${mismatches.length} cell${mismatches.length === 1 ? "" : "s"} that did not match the requested values (${mismatches.slice(0, 5).map((item) => item.address).join(", ")}).`
      : `Verified ${checked} written cell${checked === 1 ? "" : "s"}.`,
  };
}