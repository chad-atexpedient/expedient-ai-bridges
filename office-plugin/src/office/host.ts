import { excelInsertionPreview, splitTextForSlides } from "../lib/artifacts";
import { buildWorkbookOverview, describeOverwriteConflicts, formatRangeForModel, formatSearchMatches, normalizeReadMode, searchRangeMatches, verifyWrittenValues } from "./excelRead";
import { findUploadedAsset, imageAssetFromUrl } from "../lib/uploadRegistry";
import { extractTableRows } from "../lib/artifacts";
import type { DocumentContext, OfficeHost, ToolCallRequest, ToolCallResult } from "../lib/types";

declare const Office: any;
declare const Excel: any;
declare const Word: any;
declare const PowerPoint: any;

export function detectHost(): OfficeHost {
  const host = globalThis.Office?.context?.host;
  if (host === globalThis.Office?.HostType?.Excel || host === "Excel") return "excel";
  if (host === globalThis.Office?.HostType?.Word || host === "Word") return "word";
  if (host === globalThis.Office?.HostType?.PowerPoint || host === "PowerPoint") return "powerpoint";
  return "unknown";
}

export async function officeReady(): Promise<OfficeHost> {
  if (!globalThis.Office?.onReady) return "unknown";
  const info = await globalThis.Office.onReady();
  if (info?.host === globalThis.Office?.HostType?.Excel || info?.host === "Excel") return "excel";
  if (info?.host === globalThis.Office?.HostType?.Word || info?.host === "Word") return "word";
  if (info?.host === globalThis.Office?.HostType?.PowerPoint || info?.host === "PowerPoint") return "powerpoint";
  return detectHost();
}

export async function readDocumentContext(host: OfficeHost): Promise<DocumentContext> {
  if (host === "excel" && globalThis.Excel) return readExcelContext();
  if (host === "word" && globalThis.Word) return readWordContext();
  if (host === "powerpoint" && globalThis.PowerPoint) return readPowerPointContext();
  return {
    host,
    title: globalThis.document?.title || "Browser preview",
    selectionLabel: "Preview mode",
    text: "Office.js is not connected. This preview can test settings, branding, and provider calls, but not live document context.",
    metadata: { preview: true },
  };
}

export async function safeReadDocumentContext(host: OfficeHost): Promise<{ context: DocumentContext | null; warning: string }> {
  try {
    return { context: await readDocumentContext(host), warning: "" };
  } catch (error: any) {
    return {
      context: null,
      warning: `Could not read the current ${host} context: ${error?.message || String(error)}`,
    };
  }
}

async function readExcelContext(): Promise<DocumentContext> {
  return globalThis.Excel.run(async (ctx: any) => {
    const workbook = ctx.workbook;
    const worksheets = workbook.worksheets;
    const selected = workbook.getSelectedRange();
    workbook.load("name");
    worksheets.load("items/name");
    selected.load(["address", "values", "formulas", "rowCount", "columnCount"]);
    await ctx.sync();

    const values = selected.values ?? [];
    const formulas = selected.formulas ?? [];
    const lines = values.slice(0, 30).map((row: unknown[], rowIndex: number) =>
      row.map((value, columnIndex) => {
        const formula = formulas?.[rowIndex]?.[columnIndex];
        return formula && formula !== value ? `${value ?? ""} (${formula})` : String(value ?? "");
      }).join("\t")
    );

    return {
      host: "excel",
      title: workbook.name || "Workbook",
      selectionLabel: selected.address,
      text: lines.join("\n"),
      metadata: {
        sheets: worksheets.items.map((sheet: any) => sheet.name),
        rows: selected.rowCount,
        columns: selected.columnCount,
      },
    };
  });
}

async function readWordContext(): Promise<DocumentContext> {
  return globalThis.Word.run(async (ctx: any) => {
    const selection = ctx.document.getSelection();
    selection.load("text");
    await ctx.sync();

    let text = selection.text?.trim();
    let label = "Current selection";
    if (!text) {
      const body = ctx.document.body;
      body.load("text");
      await ctx.sync();
      text = body.text?.slice(0, 6000) || "";
      label = "Document body excerpt";
    }

    return {
      host: "word",
      title: "Word document",
      selectionLabel: label,
      text,
      metadata: { characterCount: text.length },
    };
  });
}

async function readPowerPointContext(): Promise<DocumentContext> {
  return globalThis.PowerPoint.run(async (ctx: any) => {
    const slides = ctx.presentation.slides;
    slides.load("items");
    await ctx.sync();

    const summaries: string[] = [];
    for (let i = 0; i < Math.min(slides.items.length, 12); i += 1) {
      const slide = slides.items[i];
      try {
        slide.shapes.load("items");
        await ctx.sync();
        summaries.push(`Slide ${i + 1}: ${slide.shapes.items.length} shapes`);
      } catch (error: any) {
        summaries.push(`Slide ${i + 1}: shape details unavailable (${error?.message || String(error)})`);
      }
    }

    return {
      host: "powerpoint",
      title: "PowerPoint deck",
      selectionLabel: "Deck summary",
      text: summaries.join("\n") || "No readable slide details were available.",
      metadata: { slideCount: slides.items.length },
    };
  });
}


function toolResult(call: ToolCallRequest, ok: boolean, content: string): ToolCallResult {
  return { id: call.id, name: call.name, ok, content };
}

function asString(value: unknown, fallback = "") {
  return typeof value === "string" ? value : fallback;
}

function asRows(value: unknown): string[][] {
  if (!Array.isArray(value)) return [];
  return value.map((row) => Array.isArray(row) ? row.map((cell) => String(cell ?? "")) : [String(row ?? "")]);
}

function normalizeHeaderKey(value: unknown) {
  return String(value ?? "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");
}

function columnIndexFromReference(reference: unknown, headers: string[], fallback: number | null = null) {
  if (typeof reference === "number" && Number.isFinite(reference)) return Math.max(0, Math.trunc(reference));
  const requested = normalizeHeaderKey(reference);
  if (!requested) return fallback;
  const found = headers.findIndex((header) => normalizeHeaderKey(header) === requested);
  return found >= 0 ? found : fallback;
}

function numericValue(value: unknown) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const cleaned = String(value ?? "").replace(/[$,%]/g, "").trim();
  if (!cleaned) return null;
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : null;
}

export function createSummaryRows(values: unknown[][], groupBy: unknown, valueColumn: unknown, operationValue: unknown = "sum") {
  const rows = Array.isArray(values) ? values : [];
  if (rows.length < 2) throw new Error("Summary source range must include a header row and at least one data row.");
  const headers = rows[0].map((cell) => String(cell ?? "").trim() || "Column");
  const groupIndex = columnIndexFromReference(groupBy, headers, 0);
  const valueIndex = columnIndexFromReference(valueColumn, headers, headers.length > 1 ? 1 : 0);
  if (groupIndex === null || groupIndex < 0 || groupIndex >= headers.length) throw new Error("Summary groupBy column was not found in the source headers.");
  if (valueIndex === null || valueIndex < 0 || valueIndex >= headers.length) throw new Error("Summary valueColumn was not found in the source headers.");
  const operation = normalizeHeaderKey(operationValue || "sum");
  const buckets = new Map<string, { label: string; values: number[]; count: number }>();
  for (const row of rows.slice(1)) {
    const label = String(row?.[groupIndex] ?? "").trim() || "(blank)";
    const numeric = numericValue(row?.[valueIndex]);
    const bucket = buckets.get(label) ?? { label, values: [], count: 0 };
    bucket.count += 1;
    if (numeric !== null) bucket.values.push(numeric);
    buckets.set(label, bucket);
  }
  const aggregate = (bucket: { values: number[]; count: number }) => {
    if (operation === "count") return bucket.count;
    if (!bucket.values.length) return 0;
    if (operation === "average" || operation === "avg") return bucket.values.reduce((sum, value) => sum + value, 0) / bucket.values.length;
    if (operation === "min") return Math.min(...bucket.values);
    if (operation === "max") return Math.max(...bucket.values);
    return bucket.values.reduce((sum, value) => sum + value, 0);
  };
  const operationLabelMap: Record<string, string> = { sum: "Sum", average: "Average", avg: "Average", count: "Count", min: "Min", max: "Max" };
  const operationLabel = operationLabelMap[operation] ?? "Sum";
  const output: Array<Array<string | number>> = [[headers[groupIndex], operation === "count" ? "Count" : operationLabel + " of " + headers[valueIndex], "Rows"]];
  for (const bucket of [...buckets.values()].sort((a, b) => a.label.localeCompare(b.label))) {
    output.push([bucket.label, aggregate(bucket), bucket.count]);
  }
  return output;
}

type ExcelCleanTransformOptions = {
  trimWhitespace?: boolean;
  normalizeSpaces?: boolean;
  removeBlankRows?: boolean;
  removeDuplicateRows?: boolean;
  splitColumn?: unknown;
  delimiter?: unknown;
  hasHeaders?: boolean;
  caseMode?: unknown;
};

type ExcelCombineRangeOptions = ExcelCleanTransformOptions & {
  mode?: unknown;
  matchColumns?: unknown;
  primaryKey?: unknown;
  secondaryKey?: unknown;
  conflictSuffix?: unknown;
  includeSourceColumn?: boolean;
  primaryLabel?: unknown;
  secondaryLabel?: unknown;
};

function normalizeCleanCell(value: unknown, options: ExcelCleanTransformOptions) {
  if (typeof value !== "string") return value ?? "";
  let text = value;
  if (options.trimWhitespace !== false) text = text.trim();
  if (options.normalizeSpaces !== false) text = text.replace(/\s+/g, " ");
  const caseMode = compactKey(options.caseMode || "none");
  if (caseMode === "upper") text = text.toUpperCase();
  else if (caseMode === "lower") text = text.toLowerCase();
  else if (caseMode === "title") text = text.toLowerCase().replace(/\b\p{L}/gu, (match) => match.toUpperCase());
  return text;
}

function isBlankCleanRow(row: unknown[]) {
  return row.every((cell) => String(cell ?? "").trim() === "");
}

function cleanRowKey(row: unknown[]) {
  return row.map((cell) => String(cell ?? "").trim().toLowerCase()).join("\u001f");
}

export function createCleanedRows(values: unknown[][], options: ExcelCleanTransformOptions = {}) {
  const sourceRows = Array.isArray(values) ? values : [];
  if (!sourceRows.length) throw new Error("Clean source range must include at least one row.");
  const width = Math.max(1, ...sourceRows.map((row) => Array.isArray(row) ? row.length : 1));
  const hasHeaders = options.hasHeaders !== false;
  const transforms: string[] = [];
  const normalizedRows: unknown[][] = sourceRows.map((row) => {
    const sourceRow = Array.isArray(row) ? row : [row];
    return Array.from({ length: width }, (_, index) => normalizeCleanCell(sourceRow[index], options));
  });
  if (options.trimWhitespace !== false) transforms.push("trimmed whitespace");
  if (options.normalizeSpaces !== false) transforms.push("normalized repeated spaces");

  const headerRows: unknown[][] = hasHeaders ? normalizedRows.slice(0, 1) : [];
  let dataRows: unknown[][] = hasHeaders ? normalizedRows.slice(1) : normalizedRows.slice();
  const originalDataRowCount = dataRows.length;
  if (options.removeBlankRows !== false) {
    dataRows = dataRows.filter((row) => !isBlankCleanRow(row));
    transforms.push("removed blank rows");
  }

  let splitColumnIndex = -1;
  if (options.splitColumn !== undefined && options.splitColumn !== null && String(options.splitColumn).trim() !== "") {
    const headers = hasHeaders ? headerRows[0].map((cell, index) => String(cell ?? "").trim() || `Column ${index + 1}`) : Array.from({ length: width }, (_, index) => `Column ${index + 1}`);
    splitColumnIndex = columnIndexFromReference(options.splitColumn, headers, null) ?? -1;
    if (splitColumnIndex < 0 || splitColumnIndex >= width) throw new Error(`Split column ${String(options.splitColumn)} was not found in the source range.`);
    const delimiter = asString(options.delimiter, ",") || ",";
    const splitValues = dataRows.map((row) => String(row[splitColumnIndex] ?? "").split(delimiter).map((part) => normalizeCleanCell(part, options)));
    const splitWidth = Math.max(1, ...splitValues.map((parts) => parts.length));
    const expand = (row: unknown[], splitParts: unknown[]) => [
      ...row.slice(0, splitColumnIndex),
      ...Array.from({ length: splitWidth }, (_, index) => splitParts[index] ?? ""),
      ...row.slice(splitColumnIndex + 1),
    ];
    dataRows = dataRows.map((row, index) => expand(row, splitValues[index]));
    if (hasHeaders) {
      const header = headerRows[0];
      const label = String(header[splitColumnIndex] ?? `Column ${splitColumnIndex + 1}`).trim() || `Column ${splitColumnIndex + 1}`;
      headerRows[0] = [
        ...header.slice(0, splitColumnIndex),
        ...Array.from({ length: splitWidth }, (_, index) => splitWidth === 1 ? label : `${label} ${index + 1}`),
        ...header.slice(splitColumnIndex + 1),
      ];
    }
    transforms.push(`split ${headers[splitColumnIndex]} by delimiter`);
  }

  const rowCountAfterBlankRemoval = dataRows.length;
  if (options.removeDuplicateRows === true) {
    const seen = new Set<string>();
    dataRows = dataRows.filter((row) => {
      const key = cleanRowKey(row);
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
    transforms.push("removed duplicate rows");
  }

  const outputRows = [...headerRows, ...dataRows];
  const outputWidth = Math.max(1, ...outputRows.map((row) => row.length));
  return {
    rows: outputRows.map((row) => Array.from({ length: outputWidth }, (_, index) => row[index] ?? "")),
    sourceRows: sourceRows.length,
    sourceDataRows: originalDataRowCount,
    removedBlankRows: originalDataRowCount - rowCountAfterBlankRemoval,
    removedDuplicateRows: rowCountAfterBlankRemoval - dataRows.length,
    outputRows: outputRows.length,
    outputDataRows: dataRows.length,
    outputColumns: outputWidth,
    splitColumnIndex: splitColumnIndex >= 0 ? splitColumnIndex : null,
    transforms,
  };
}

function sourceHeaders(values: unknown[][], options: ExcelCleanTransformOptions, label = "Column") {
  const cleaned = createCleanedRows(values, { ...options, removeBlankRows: false, removeDuplicateRows: false });
  const hasHeaders = options.hasHeaders !== false;
  const width = Math.max(1, cleaned.outputColumns);
  const headers = hasHeaders
    ? cleaned.rows[0].map((cell, index) => String(cell ?? "").trim() || `${label} ${index + 1}`)
    : Array.from({ length: width }, (_, index) => `${label} ${index + 1}`);
  const dataRows = hasHeaders ? cleaned.rows.slice(1) : cleaned.rows;
  return { headers, dataRows };
}

function uniqueHeaderName(baseValue: unknown, existing: Set<string>) {
  const base = String(baseValue ?? "Column").trim() || "Column";
  let candidate = base;
  let suffix = 2;
  while (existing.has(normalizeHeaderKey(candidate))) {
    candidate = `${base} ${suffix}`;
    suffix += 1;
  }
  existing.add(normalizeHeaderKey(candidate));
  return candidate;
}

function rowByHeader(headers: string[], row: unknown[]) {
  const values = new Map<string, unknown>();
  headers.forEach((header, index) => values.set(normalizeHeaderKey(header), row[index] ?? ""));
  return values;
}

export function createCombinedRows(primaryValues: unknown[][], secondaryValues: unknown[][], options: ExcelCombineRangeOptions = {}) {
  const mode = compactKey(options.mode || "append");
  const primary = sourceHeaders(primaryValues, options, "Primary Column");
  const secondary = sourceHeaders(secondaryValues, options, "Secondary Column");
  if (!primary.headers.length || !secondary.headers.length) throw new Error("Combine ranges require two non-empty ranges.");

  if (mode === "merge" || mode === "join") {
    const primaryKeyIndex = columnIndexFromReference(options.primaryKey, primary.headers, null);
    const secondaryKeyIndex = columnIndexFromReference(options.secondaryKey ?? options.primaryKey, secondary.headers, null);
    if (primaryKeyIndex === null || primaryKeyIndex < 0 || primaryKeyIndex >= primary.headers.length) throw new Error("Merge primaryKey was not found in the primary range headers.");
    if (secondaryKeyIndex === null || secondaryKeyIndex < 0 || secondaryKeyIndex >= secondary.headers.length) throw new Error("Merge secondaryKey was not found in the secondary range headers.");
    const existing = new Set(primary.headers.map(normalizeHeaderKey));
    const conflictSuffix = asString(options.conflictSuffix, " from lookup") || " from lookup";
    const secondaryOutput = secondary.headers
      .map((header, index) => ({ header, index }))
      .filter((item) => item.index !== secondaryKeyIndex)
      .map((item) => ({ ...item, outputHeader: uniqueHeaderName(existing.has(normalizeHeaderKey(item.header)) ? `${item.header}${conflictSuffix}` : item.header, existing) }));
    const lookup = new Map<string, unknown[]>();
    for (const row of secondary.dataRows) {
      const key = cleanRowKey([row[secondaryKeyIndex]]);
      if (key && !lookup.has(key)) lookup.set(key, row);
    }
    const output: unknown[][] = [[...primary.headers, ...secondaryOutput.map((item) => item.outputHeader)]];
    let matchedRows = 0;
    for (const row of primary.dataRows) {
      const match = lookup.get(cleanRowKey([row[primaryKeyIndex]]));
      if (match) matchedRows += 1;
      output.push([...row, ...secondaryOutput.map((item) => match?.[item.index] ?? "")]);
    }
    return {
      rows: output,
      mode: "merge",
      primaryRows: primary.dataRows.length,
      secondaryRows: secondary.dataRows.length,
      matchedRows,
      outputRows: output.length,
      outputDataRows: output.length - 1,
      outputColumns: output[0].length,
      transforms: ["left merged ranges by key"],
    };
  }

  const matchMode = compactKey(options.matchColumns || "union");
  const primaryKeys = primary.headers.map(normalizeHeaderKey);
  const secondaryKeys = secondary.headers.map(normalizeHeaderKey);
  const combinedHeaders = matchMode === "intersection"
    ? primary.headers.filter((header) => secondaryKeys.includes(normalizeHeaderKey(header)))
    : (() => {
        const existing = new Set<string>();
        const headers = primary.headers.map((header) => uniqueHeaderName(header, existing));
        for (const header of secondary.headers) {
          if (!existing.has(normalizeHeaderKey(header))) headers.push(uniqueHeaderName(header, existing));
        }
        return headers;
      })();
  if (!combinedHeaders.length) throw new Error("Append ranges had no matching or unionable columns.");
  const includeSource = options.includeSourceColumn === true;
  const finalHeaders = includeSource ? ["Source", ...combinedHeaders] : combinedHeaders;
  const project = (headers: string[], row: unknown[], label: string) => {
    const values = rowByHeader(headers, row);
    const projected = combinedHeaders.map((header) => values.get(normalizeHeaderKey(header)) ?? "");
    return includeSource ? [label, ...projected] : projected;
  };
  const primaryLabel = asString(options.primaryLabel, "Primary") || "Primary";
  const secondaryLabel = asString(options.secondaryLabel, "Secondary") || "Secondary";
  const output: unknown[][] = [finalHeaders, ...primary.dataRows.map((row) => project(primary.headers, row, primaryLabel)), ...secondary.dataRows.map((row) => project(secondary.headers, row, secondaryLabel))];
  return {
    rows: output,
    mode: "append",
    primaryRows: primary.dataRows.length,
    secondaryRows: secondary.dataRows.length,
    matchedRows: null,
    outputRows: output.length,
    outputDataRows: output.length - 1,
    outputColumns: finalHeaders.length,
    transforms: [`appended ranges using ${matchMode === "intersection" ? "matching" : "unioned"} columns`],
  };
}

function excelColumnName(index: number) {
  let column = Math.max(0, index);
  let name = "";
  do {
    name = String.fromCharCode(65 + (column % 26)) + name;
    column = Math.floor(column / 26) - 1;
  } while (column >= 0);
  return name;
}

function formulaShape(formula: string) {
  return formula
    .toUpperCase()
    .replace(/\$?[A-Z]{1,3}\$?\d{1,7}/g, "#REF")
    .replace(/\d+(?:\.\d+)?/g, "#NUM")
    .replace(/\s+/g, "");
}

function formulaFinding(address: string, severity: string, issue: string, formula: string, recommendation: string) {
  return [address, severity, issue, formula.slice(0, 500), recommendation];
}

export function createFormulaAuditRows(formulas: unknown[][], values: unknown[][] = [], baseAddress = "A1", includeLowRisk = false) {
  const rows = Array.isArray(formulas) ? formulas : [];
  const valueRows = Array.isArray(values) ? values : [];
  const startMatch = String(baseAddress || "A1").match(/^([A-Z]{1,3})(\d{1,7})/i);
  const startColumn = startMatch ? startMatch[1].toUpperCase().split("").reduce((sum, char) => sum * 26 + char.charCodeAt(0) - 64, 0) - 1 : 0;
  const startRow = startMatch ? Number(startMatch[2]) - 1 : 0;
  const output: string[][] = [["Cell", "Severity", "Issue", "Formula", "Recommendation"]];
  const columnShapes = new Map<number, Map<string, string[]>>();
  let formulaCount = 0;

  rows.forEach((row, rowIndex) => row.forEach((cell, columnIndex) => {
    const formula = String(cell ?? "").trim();
    if (!formula || !formula.startsWith("=")) return;
    formulaCount += 1;
    const address = `${excelColumnName(startColumn + columnIndex)}${startRow + rowIndex + 1}`;
    const upper = formula.toUpperCase();
    const value = String(valueRows?.[rowIndex]?.[columnIndex] ?? "");
    if (/^#(DIV\/0!|N\/A|VALUE!|REF!|NAME\?|NUM!|NULL!|SPILL!|CALC!)/i.test(value)) output.push(formulaFinding(address, "High", `Formula currently returns ${value}`, formula, "Inspect inputs and dependencies before relying on this result."));
    if (/\[[^\]]+\]|https?:\/\//i.test(formula)) output.push(formulaFinding(address, "High", "External workbook or web reference", formula, "Confirm the external source is trusted, current, and available to recipients."));
    if (/(^|[^A-Z])(OFFSET|INDIRECT|TODAY|NOW|RAND|RANDBETWEEN)\s*\(/i.test(formula)) output.push(formulaFinding(address, "Medium", "Volatile or indirect function", formula, "Volatile formulas can slow workbooks or make results change unexpectedly; document intent or replace where possible."));
    if (/\b(VLOOKUP|HLOOKUP)\s*\(/i.test(formula)) output.push(formulaFinding(address, "Medium", "Legacy lookup function", formula, "Consider XLOOKUP/INDEX-MATCH and verify exact-match behavior."));
    if (/[^A-Z0-9_]\d+(?:\.\d+)?/.test(formula.replace(/"[^"]*"/g, ""))) output.push(formulaFinding(address, "Medium", "Hard-coded numeric constant", formula, "Move assumptions into named input cells or document the constant."));
    if (formula.length > 240) output.push(formulaFinding(address, "Medium", "Long formula", formula, "Break complex logic into helper cells, LET, or named formulas for auditability."));
    if (includeLowRisk && /\bSUM\s*\([^)]*,[^)]*\)/i.test(formula)) output.push(formulaFinding(address, "Low", "Multi-area SUM", formula, "Confirm all intended ranges are included and no range was accidentally skipped."));
    const shape = formulaShape(formula);
    const byShape = columnShapes.get(columnIndex) ?? new Map<string, string[]>();
    byShape.set(shape, [...(byShape.get(shape) ?? []), address]);
    columnShapes.set(columnIndex, byShape);
  }));

  for (const [columnIndex, shapes] of columnShapes) {
    const entries = [...shapes.entries()].filter(([, addresses]) => addresses.length === 1);
    const repeatedShapes = [...shapes.entries()].filter(([, addresses]) => addresses.length >= 3);
    if (repeatedShapes.length && entries.length) {
      for (const [shape, addresses] of entries.slice(0, 25)) {
        output.push(formulaFinding(addresses[0], "Medium", `Formula differs from nearby ${excelColumnName(startColumn + columnIndex)} column pattern`, shape, "Check whether this is an intentional exception or an accidental formula drift."));
      }
    }
  }

  if (formulaCount === 0) output.push(["(none)", "Info", "No formulas found", "", "Audit a larger range or another worksheet if formulas were expected."]);
  else if (output.length === 1) output.push(["Workbook", "Info", `Audited ${formulaCount} formulas with no flagged risks`, "", "Keep this sheet as audit evidence or rerun after model changes."]);
  return output;
}

function uniqueSorted(values: string[]) {
  return [...new Set(values.filter(Boolean))].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
}

function normalizeFormulaReference(reference: string, defaultSheetName = "") {
  const cleaned = reference.replace(/\$/g, "").replace(/^'/, "").replace(/'$/, "");
  const bangIndex = cleaned.lastIndexOf("!");
  if (bangIndex >= 0) {
    const sheet = cleaned.slice(0, bangIndex).replace(/^'/, "").replace(/'$/, "");
    return `${sheet}!${cleaned.slice(bangIndex + 1).toUpperCase()}`;
  }
  return defaultSheetName ? `${defaultSheetName}!${cleaned.toUpperCase()}` : cleaned.toUpperCase();
}

export function formulaReferences(formula: string, defaultSheetName = "") {
  const withoutStrings = formula.replace(/"[^"]*"/g, "");
  const externalReferences = uniqueSorted([
    ...withoutStrings.matchAll(/\[[^\]]+\][A-Za-z0-9_ .'-]+!\$?[A-Z]{1,3}\$?\d{1,7}(?::\$?[A-Z]{1,3}\$?\d{1,7})?/g),
    ...withoutStrings.matchAll(/https?:\/\/[^,)+\s]+/gi),
  ].map((match) => match[0]));
  let localFormulaText = withoutStrings;
  for (const external of externalReferences) localFormulaText = localFormulaText.replace(external, " ");
  const localReferences = uniqueSorted([...localFormulaText.matchAll(/(?:'[^']+'|[A-Za-z0-9_ .]+)!\$?[A-Z]{1,3}\$?\d{1,7}(?::\$?[A-Z]{1,3}\$?\d{1,7})?|\$?[A-Z]{1,3}\$?\d{1,7}(?::\$?[A-Z]{1,3}\$?\d{1,7})?/g)]
    .map((match) => match[0])
    .filter((ref) => !ref.includes("["))
    .filter((ref) => !/^\d/.test(ref))
    .map((ref) => normalizeFormulaReference(ref, defaultSheetName)));
  return { localReferences, externalReferences };
}

export function createFormulaDependencyRows(formulas: unknown[][], baseAddress = "A1", sheetName = "Sheet", includeSummary = true) {
  const rows = Array.isArray(formulas) ? formulas : [];
  const startMatch = String(baseAddress || "A1").match(/^([A-Z]{1,3})(\d{1,7})/i);
  const startColumn = startMatch ? startMatch[1].toUpperCase().split("").reduce((sum, char) => sum * 26 + char.charCodeAt(0) - 64, 0) - 1 : 0;
  const startRow = startMatch ? Number(startMatch[2]) - 1 : 0;
  const output: string[][] = [["Formula Cell", "Precedents", "External References", "Formula"]];
  const dependents = new Map<string, string[]>();
  let formulaCount = 0;

  rows.forEach((row, rowIndex) => row.forEach((cell, columnIndex) => {
    const formula = String(cell ?? "").trim();
    if (!formula || !formula.startsWith("=")) return;
    formulaCount += 1;
    const cellAddress = `${sheetName}!${excelColumnName(startColumn + columnIndex)}${startRow + rowIndex + 1}`;
    const { localReferences, externalReferences } = formulaReferences(formula, sheetName);
    output.push([cellAddress, localReferences.join(", ") || "(none detected)", externalReferences.join(", "), formula.slice(0, 500)]);
    for (const reference of localReferences) dependents.set(reference, [...(dependents.get(reference) ?? []), cellAddress]);
  }));

  if (formulaCount === 0) output.push(["(none)", "No formulas found", "", ""]);
  if (includeSummary && dependents.size) {
    output.push(["", "", "", ""]);
    output.push(["Precedent Cell/Range", "Dependent Formula Cells", "Dependent Count", ""]);
    for (const [reference, dependentCells] of [...dependents.entries()].sort((a, b) => b[1].length - a[1].length || a[0].localeCompare(b[0]))) {
      output.push([reference, uniqueSorted(dependentCells).join(", "), String(uniqueSorted(dependentCells).length), ""]);
    }
  }
  return output;
}

export function normalizeAddress(value: unknown) {
  let address = asString(value, "").trim();
  if (!address) return null;
  address = address.startsWith("=") ? address.slice(1).trim() : address;
  const bangIndex = address.lastIndexOf("!");
  if (bangIndex >= 0) address = address.slice(bangIndex + 1).trim();
  address = address.split("$").join("");
  if ((address.startsWith("'") && address.endsWith("'")) || (address.startsWith("\"") && address.endsWith("\""))) address = address.slice(1, -1);
  const a1Match = address.match(/(?:^|[^A-Z])([A-Z]{1,3}[0-9]{1,7}(?::[A-Z]{1,3}[0-9]{1,7})?)(?:$|[^A-Z0-9])/i);
  if (a1Match) return a1Match[1].toUpperCase();
  const compact = Array.from(address).filter((char) => char.trim()).join("");
  const compactMatch = compact.match(/^[A-Z]{1,3}[0-9]{1,7}(?::[A-Z]{1,3}[0-9]{1,7})?$/i);
  if (compactMatch) return compactMatch[0].toUpperCase();
  return compact || null;
}

function compactKey(value: unknown) {
  return asString(value, "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

function parseNumberLike(value: unknown) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value === "string") {
    const match = value.match(/-?\d+(?:\.\d+)?/);
    if (match) {
      const parsed = Number(match[0]);
      if (Number.isFinite(parsed)) return parsed;
    }
  }
  return null;
}

function firstDefined(...values: unknown[]) {
  return values.find((value) => value !== undefined && value !== null && String(value).trim?.() !== "");
}

export function hasOfficeSlideReference(args: Record<string, unknown>) {
  return firstDefined(args.slideNumber, args.slide, args.index, args.slideIndex) !== undefined;
}

export function normalizeOfficeSlideIndex(argsOrValue: Record<string, unknown> | number | null | undefined, slideCount = 0, fallback = 0) {
  if (typeof argsOrValue === "number" || argsOrValue === null || argsOrValue === undefined) {
    const direct = typeof argsOrValue === "number" && Number.isFinite(argsOrValue) ? Math.trunc(argsOrValue) : fallback;
    if (slideCount > 0 && direct >= slideCount && direct - 1 >= 0 && direct - 1 < slideCount) return direct - 1;
    return Math.max(0, direct);
  }
  const requestedPosition = String(firstDefined(argsOrValue.slidePosition, argsOrValue.position) ?? "").trim().toLowerCase();
  if (["last", "latest", "newest", "end"].includes(requestedPosition)) return Math.max(0, slideCount - 1);
  const explicitSlideNumber = firstDefined(argsOrValue.slideNumber, argsOrValue.slide);
  const raw = explicitSlideNumber ?? firstDefined(argsOrValue.slideIndex, argsOrValue.index);
  const parsed = parseNumberLike(raw);
  if (parsed === null) return Math.max(0, fallback);
  let index = Math.trunc(parsed);
  if (explicitSlideNumber !== undefined) index -= 1;
  else if (slideCount > 0 && index >= slideCount && index - 1 >= 0 && index - 1 < slideCount) index -= 1;
  return Math.max(0, index);
}

export function normalizeOfficeUrl(value: unknown) {
  let url = asString(value, "").trim();
  if (!url) return "";
  if (/^[\w.-]+\.[a-z]{2,}(?:[/:?#].*)?$/i.test(url)) url = "https://" + url;
  if (!/^(https?:|mailto:)/i.test(url)) throw new Error("Unsupported hyperlink URL \"" + url + "\". Use an http(s) or mailto URL.");
  return url;
}

export function safeExcelName(value: unknown) {
  let name = asString(value, "").trim().replace(/[^A-Za-z0-9_.]/g, "_");
  if (!name) return "AI_Range";
  if (/^[0-9]/.test(name)) name = "_" + name;
  if (/^[A-Za-z]{1,3}[0-9]{1,7}$/i.test(name)) name = "_" + name;
  return name.slice(0, 255);
}

function safeSheetName(value: unknown, fallback = "Imported Context") {
  const cleaned = asString(value, fallback).trim().replace(/[\\/?*\[\]:]/g, " ").replace(/\s+/g, " ").slice(0, 31).trim();
  return cleaned || fallback;
}

function sanitizeExcelImportedCell(value: unknown, preserveFormulas = false) {
  const text = String(value ?? "").replace(/\r\n/g, "\n").trim();
  if (!preserveFormulas && /^[=+\-@]/.test(text)) return "'" + text;
  return text;
}

function normalizeImportedContextRows(rows: unknown[][], options: { preserveFormulas?: boolean; includeSourceColumn?: boolean; sourceLabel?: unknown; maxRows?: unknown } = {}) {
  const maxRows = Math.max(1, Math.min(5000, asOptionalNumber(options.maxRows) ?? 500));
  const safeRows = rows.slice(0, maxRows).map((row) => row.map((cell) => sanitizeExcelImportedCell(cell, Boolean(options.preserveFormulas))));
  const width = Math.max(1, ...safeRows.map((row) => row.length));
  const normalized = safeRows.map((row) => Array.from({ length: width }, (_, index) => row[index] ?? ""));
  if (options.includeSourceColumn) {
    const source = sanitizeExcelImportedCell(options.sourceLabel || "Context", false);
    return normalized.map((row, index) => [index === 0 ? "Source" : source, ...row]);
  }
  return normalized;
}

function keyValueRowsFromText(text: string) {
  const rows: string[][] = [["Field", "Value"]];
  for (const line of text.split(/\r?\n/)) {
    const cleaned = line.trim().replace(/^[-*•]\s+/, "");
    if (!cleaned) continue;
    const match = cleaned.match(/^([^:=-][^:=]{1,80})\s*[:=]\s*(.+)$/);
    if (match) rows.push([match[1].trim(), match[2].trim()]);
  }
  return rows.length > 1 ? rows : null;
}

function bulletRowsFromText(text: string) {
  const rows: string[][] = [["Item", "Detail"]];
  for (const line of text.split(/\r?\n/)) {
    const match = line.trim().match(/^[-*•]\s+(.+)$/);
    if (!match) continue;
    const item = match[1].trim();
    const parts = item.split(/\s+[–-]\s+/);
    rows.push([parts[0]?.trim() || item, parts.slice(1).join(" - ").trim()]);
  }
  return rows.length > 1 ? rows : null;
}

export function createExcelContextImportRows(textValue: unknown, options: { mode?: unknown; preserveFormulas?: boolean; includeSourceColumn?: boolean; sourceLabel?: unknown; maxRows?: unknown } = {}) {
  const text = asString(textValue, "").trim();
  if (!text) throw new Error("Context import requires text to import.");
  const mode = compactKey(options.mode || "auto");
  const parsed = mode === "keyvalue" ? keyValueRowsFromText(text)
    : mode === "bullets" ? bulletRowsFromText(text)
    : mode === "table" ? extractTableRows(text)
    : extractTableRows(text) ?? keyValueRowsFromText(text) ?? bulletRowsFromText(text);
  const rows = parsed ?? [["Context"], ...text.split(/\r?\n/).map((line) => [line]).filter((row) => row[0].trim())];
  const normalized = normalizeImportedContextRows(rows, options);
  return { rows: normalized, detectedMode: parsed ? (mode === "auto" ? (parsed[0]?.length === 2 && parsed[0]?.[0] === "Field" ? "keyValue" : parsed[0]?.length === 2 && parsed[0]?.[0] === "Item" ? "bullets" : "table") : mode) : "lines", truncated: rows.length > normalized.length };
}


function excelPivotFieldName(reference: unknown, headers: string[]) {
  const index = columnIndexFromReference(reference, headers, null);
  if (index === null || index < 0 || index >= headers.length) throw new Error(`PivotTable field ${String(reference)} was not found in the source headers.`);
  return headers[index];
}

function excelPivotAggregation(value: unknown) {
  const aggregation = compactKey(value || "sum");
  const functions = globalThis.Excel?.AggregationFunction ?? {};
  const map: Record<string, any> = {
    sum: functions.sum ?? "Sum",
    count: functions.count ?? "Count",
    average: functions.average ?? functions.avg ?? "Average",
    avg: functions.average ?? functions.avg ?? "Average",
    max: functions.max ?? "Max",
    min: functions.min ?? "Min",
  };
  return map[aggregation] ?? map.sum;
}

function excelPivotLayout(value: unknown) {
  const layout = compactKey(value || "");
  const layouts = globalThis.Excel?.PivotLayoutType ?? {};
  const map: Record<string, any> = { compact: layouts.compact ?? "Compact", outline: layouts.outline ?? "Outline", tabular: layouts.tabular ?? "Tabular" };
  return map[layout] ?? null;
}

function asNumber(value: unknown, fallback: number) {
  return typeof value === "number" && Number.isFinite(value) ? value : fallback;
}

function asOptionalNumber(value: unknown) {
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}

async function resolveImageAsset(args: Record<string, unknown>) {
  const assetId = asString(args.assetId, "").trim();
  const assetName = asString(args.assetName, "").trim();
  const imageUrl = asString(args.imageUrl, "").trim();
  const uploaded = findUploadedAsset(assetId || assetName);
  if (uploaded) return uploaded;
  if (imageUrl) return imageAssetFromUrl(imageUrl);
  throw new Error("Image tools require assetId, assetName, or imageUrl. Attach an image or provide a public image URL.");
}

function applyShapeFrame(shape: any, args: Record<string, unknown>, defaults: { left: number; top: number; width: number; height: number }) {
  const left = asOptionalNumber(args.left) ?? defaults.left;
  const top = asOptionalNumber(args.top) ?? defaults.top;
  const width = asOptionalNumber(args.width) ?? defaults.width;
  const height = asOptionalNumber(args.height) ?? defaults.height;
  try { shape.left = left; } catch {}
  try { shape.top = top; } catch {}
  try { shape.width = width; } catch {}
  try { shape.height = height; } catch {}
}

function applyAltText(shape: any, altText: string) {
  if (!altText) return;
  try { shape.altTextDescription = altText; } catch {}
  try { shape.altTextTitle = altText; } catch {}
}

function activeOrNamedWorksheet(ctx: any, sheetNameValue: unknown) {
  const sheetName = asString(sheetNameValue, "").trim();
  return sheetName ? ctx.workbook.worksheets.getItem(sheetName) : ctx.workbook.worksheets.getActiveWorksheet();
}

function boolValue(value: unknown, fallback: boolean) {
  return typeof value === "boolean" ? value : fallback;
}

function styleTextRange(shape: any, args: Record<string, unknown>) {
  const fontSize = asOptionalNumber(args.fontSize);
  const bold = typeof args.bold === "boolean" ? args.bold : null;
  const fontColor = asString(args.fontColor, "").trim();
  try {
    if (fontSize !== null) shape.textFrame.textRange.font.size = fontSize;
    if (bold !== null) shape.textFrame.textRange.font.bold = bold;
    if (fontColor) shape.textFrame.textRange.font.color = fontColor;
  } catch {}
}

function normalizeShapeType(value: unknown) {
  const requested = compactKey(value || "rectangle");
  const shapeType = globalThis.PowerPoint?.ShapeType ?? {};
  const map: Record<string, any> = {
    rectangle: shapeType.rectangle ?? shapeType.rect ?? "Rectangle",
    rect: shapeType.rectangle ?? shapeType.rect ?? "Rectangle",
    square: shapeType.rectangle ?? shapeType.rect ?? "Rectangle",
    roundedrectangle: shapeType.roundRect ?? shapeType.roundedRectangle ?? "RoundRect",
    roundrect: shapeType.roundRect ?? shapeType.roundedRectangle ?? "RoundRect",
    oval: shapeType.ellipse ?? shapeType.oval ?? "Ellipse",
    ellipse: shapeType.ellipse ?? shapeType.oval ?? "Ellipse",
    circle: shapeType.ellipse ?? shapeType.oval ?? "Ellipse",
    triangle: shapeType.triangle ?? "Triangle",
    line: shapeType.line ?? "Line",
    arrow: shapeType.line ?? "Line",
  };
  return map[requested] ?? map.rectangle;
}

function htmlEscape(value: unknown) {
  return asString(value, "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
function wordInsertLocation(location: unknown) {
  const requested = asString(location, "replace");
  return requested === "end" ? globalThis.Word.InsertLocation.end : requested === "start" ? globalThis.Word.InsertLocation.start : globalThis.Word.InsertLocation.replace;
}

function wordTarget(ctx: any, location: unknown) {
  const requested = asString(location, "replace");
  return requested === "end" || requested === "start" ? ctx.document.body : ctx.document.getSelection();
}

function getPowerPointSlideOrThrow(slides: any, argsOrIndex: Record<string, unknown> | number | null | undefined) {
  const index = normalizeOfficeSlideIndex(argsOrIndex, slides.items.length, 0);
  const slide = slides.items[index];
  if (!slide) throw new Error(`Slide ${index + 1} was not found. Use slideNumber for human 1-based slide numbers or slideIndex for zero-based indexes.`);
  return slide;
}

function safeToolArgsForError(args: Record<string, unknown>) {
  const redacted = JSON.parse(JSON.stringify(args ?? {}));
  for (const key of Object.keys(redacted)) {
    if (/apikey|authorization|password|token|base64/i.test(key)) redacted[key] = "[redacted]";
    else if (typeof redacted[key] === "string" && redacted[key].length > 240) redacted[key] = redacted[key].slice(0, 240) + "...";
  }
  return JSON.stringify(redacted);
}

export function officeErrorMessage(error: any, fallback = "Office operation failed", call?: ToolCallRequest) {
  const raw = [
    error?.message,
    error?.debugInfo?.errorLocation ? `Location: ${error.debugInfo.errorLocation}` : "",
    error?.debugInfo?.code ? `Code: ${error.debugInfo.code}` : "",
    error?.debugInfo?.message && error.debugInfo.message !== error?.message ? `Detail: ${error.debugInfo.message}` : "",
  ].filter(Boolean).join(" | ") || fallback;
  const retryable = /invalid|argument|format|missing|not available|does not expose|unsupported|api|property|load method|context\.sync/i.test(raw);
  const retryHint = retryable
    ? "Recovery: retry at most once with simpler normalized inputs: local A1 ranges like A1:C12, slideNumber for human slide numbers, standard chart/shape names, and positive dimensions."
    : "";
  const fallbackHint = call ? officeFallbackHint(call.name) : "";
  const context = call ? `Tool: ${call.name} | Arguments: ${safeToolArgsForError(call.arguments)}` : "";
  return [raw, retryHint, fallbackHint, context].filter(Boolean).join(" | ");
}

function officeFallbackHint(toolName: string) {
  if (toolName.startsWith("powerpoint_")) {
    if (toolName === "powerpoint_generate_deck_file") return "Fallback: report the generated-deck error and summarize what could not be packaged; do not retry indefinitely.";
    return "Fallback: if the simpler retry fails or this PowerPoint runtime does not expose the needed API, use powerpoint_generate_deck_file to produce a real .pptx artifact with slides, images, tables, charts, hyperlinks, notes, and template layout placement instead of giving manual instructions.";
  }
  if (toolName.startsWith("excel_")) {
    if (toolName === "excel_generate_workbook_file") return "Fallback: report the workbook-generation error and summarize what could not be packaged; do not retry indefinitely.";
    return "Fallback: if the simpler retry fails or this Excel runtime does not expose the needed API, use excel_generate_workbook_file for a real .xlsx artifact with sheets, tables, formulas, validations, conditional formats, charts, and images instead of giving manual instructions.";
  }
  if (toolName.startsWith("word_")) {
    if (toolName === "word_generate_document_file") return "Fallback: report the document-generation error and summarize what could not be packaged; do not retry indefinitely.";
    return "Fallback: if the simpler retry fails or this Word runtime does not expose the needed API, use word_generate_document_file for a real .docx artifact with sections, tables, headers, footers, comments, footnotes, hyperlinks, table of contents, and images instead of giving manual instructions.";
  }
  return "";
}

function excelChartType(value: unknown) {
  const requested = compactKey(value || "lineMarkers");
  const chartTypes = globalThis.Excel?.ChartType ?? {};
  const map: Record<string, any> = {
    line: chartTypes.line ?? "Line",
    linechart: chartTypes.line ?? "Line",
    linemarkers: chartTypes.lineMarkers ?? chartTypes.lineMarkersStacked ?? "LineMarkers",
    linewithmarkers: chartTypes.lineMarkers ?? chartTypes.lineMarkersStacked ?? "LineMarkers",
    column: chartTypes.columnClustered ?? "ColumnClustered",
    clusteredcolumn: chartTypes.columnClustered ?? "ColumnClustered",
    columnclustered: chartTypes.columnClustered ?? "ColumnClustered",
    bar: chartTypes.barClustered ?? "BarClustered",
    clusteredbar: chartTypes.barClustered ?? "BarClustered",
    barclustered: chartTypes.barClustered ?? "BarClustered",
    pie: chartTypes.pie ?? "Pie",
    piechart: chartTypes.pie ?? "Pie",
    scatter: chartTypes.xyscatter ?? chartTypes.xyScatter ?? "XYScatter",
    scatterplot: chartTypes.xyscatter ?? chartTypes.xyScatter ?? "XYScatter",
    area: chartTypes.area ?? "Area",
    areachart: chartTypes.area ?? "Area",
  };
  return map[requested] ?? map.linewithmarkers;
}

export function createPivotChartReportRows(values: unknown[][], groupBy: unknown, valueColumn: unknown, operationValue: unknown = "sum", options: { top?: unknown; sortBy?: unknown } = {}) {
  const rows = createSummaryRows(values, groupBy, valueColumn, operationValue);
  const header = rows[0];
  let dataRows = rows.slice(1);
  const sortBy = compactKey(options.sortBy || "label");
  if (sortBy === "value" || sortBy === "valuedesc") dataRows = dataRows.sort((a, b) => numericValue(b[1])! - numericValue(a[1])! || String(a[0]).localeCompare(String(b[0])));
  else if (sortBy === "valueasc") dataRows = dataRows.sort((a, b) => numericValue(a[1])! - numericValue(b[1])! || String(a[0]).localeCompare(String(b[0])));
  else dataRows = dataRows.sort((a, b) => String(a[0]).localeCompare(String(b[0]), undefined, { numeric: true }));
  const top = asOptionalNumber(options.top);
  if (top !== null && top > 0) dataRows = dataRows.slice(0, Math.trunc(top));
  return [header, ...dataRows];
}

function sentenceCount(text: string) {
  return Math.max(1, text.split(/[.!?]+\s+/).filter((part) => part.trim().length > 12).length);
}

export function createWordAuditRows(textValue: unknown, metadata: { headingCount?: number; tableCount?: number; imageCount?: number; hyperlinkCount?: number } = {}) {
  const text = String(textValue ?? "").replace(/\r/g, "");
  const paragraphs = text.split(/\n+/).map((part) => part.trim()).filter(Boolean);
  const rows: string[][] = [["Area", "Severity", "Finding", "Recommendation"]];
  const words = text.match(/\b[\p{L}\p{N}'-]+\b/gu) ?? [];
  const headingLike = paragraphs.filter((paragraph) => paragraph.length <= 90 && !/[.!?]$/.test(paragraph) && /[A-Za-z]/.test(paragraph));
  const headingCount = metadata.headingCount ?? headingLike.length;
  const tableCount = metadata.tableCount ?? 0;
  const imageCount = metadata.imageCount ?? 0;
  const hyperlinkCount = metadata.hyperlinkCount ?? 0;

  if (words.length < 50) rows.push(["Content", "Medium", "Document body is very short or mostly unreadable to the Word runtime.", "Confirm the document has enough body text for review, or attach/upload the source DOCX for deeper package extraction."]);
  if (headingCount === 0 && words.length >= 120) rows.push(["Structure", "High", "No heading-like structure was detected in the readable body text.", "Use Heading 1/2/3 styles so navigation, screen readers, and generated tables of contents work reliably."]);
  if (headingCount === 1 && words.length >= 500) rows.push(["Structure", "Medium", "Only one heading-like section was detected in a longer document.", "Break the document into scannable sections with real Word heading styles."]);

  paragraphs.forEach((paragraph, index) => {
    const paragraphWords = paragraph.match(/\b[\p{L}\p{N}'-]+\b/gu)?.length ?? 0;
    if (paragraphWords > 160) rows.push(["Readability", "Medium", `Paragraph ${index + 1} is ${paragraphWords} words long.`, "Split long paragraphs into shorter paragraphs, bullets, or a table so busy readers and screen readers can navigate the point."]);
    const averageSentenceWords = paragraphWords / sentenceCount(paragraph);
    if (paragraphWords > 60 && averageSentenceWords > 35) rows.push(["Readability", "Low", `Paragraph ${index + 1} has long average sentence length.`, "Shorten sentences or add connective headings to improve executive readability."]);
  });

  const allCapsLong = paragraphs.filter((paragraph) => paragraph.length > 30 && paragraph === paragraph.toUpperCase() && /[A-Z]/.test(paragraph));
  if (allCapsLong.length) rows.push(["Accessibility", "Low", `${allCapsLong.length} long all-caps paragraph${allCapsLong.length === 1 ? "" : "s"} detected.`, "Avoid long all-caps text because it is harder to read and may be announced awkwardly by assistive technology."]);
  if (imageCount > 0) rows.push(["Images", "Info", `${imageCount} image object${imageCount === 1 ? "" : "s"} detected where the runtime exposed image counts.`, "Confirm every meaningful image has descriptive alt text; decorative images should be marked decorative when the target Office version supports it."]);
  else rows.push(["Images", "Info", "No image objects were detected by the live runtime.", "If the document contains images, upload the DOCX for package-level media inspection or check alt text in Word Accessibility Checker."]);
  if (tableCount > 0) rows.push(["Tables", "Medium", `${tableCount} table${tableCount === 1 ? "" : "s"} detected where the runtime exposed table counts.`, "Confirm each table has a header row, simple grid structure, readable labels, and no layout-only tables."]);
  if (hyperlinkCount === 0 && /https?:\/\//i.test(text)) rows.push(["Links", "Medium", "URL text appears in the document body, but no live hyperlink count was available.", "Convert pasted URLs into descriptive Word hyperlinks such as 'View source' rather than exposing raw URLs."]);
  if (/\b(click here|read more|here)\b/i.test(text)) rows.push(["Links", "Low", "Generic link text may be present.", "Use descriptive link text that explains the destination or action."]);
  if (!/[.!?]/.test(text) && words.length > 80) rows.push(["Readability", "Medium", "Readable text has very little sentence punctuation.", "Add punctuation and sectioning so the document is understandable to readers and assistive tools."]);
  if (rows.length === 1) rows.push(["Overall", "Info", "No common live-audit issues were detected from the readable Word body text.", "For production accessibility evidence, still run Word Accessibility Checker and package-level DOCX inspection."]);
  return rows;
}

function firstSentences(text: string, maxSentences = 3, maxChars = 650) {
  const cleaned = text.replace(/\s+/g, " ").trim();
  if (!cleaned) return "";
  const sentences = cleaned.match(/[^.!?]+[.!?]+|[^.!?]+$/g) ?? [cleaned];
  const selected: string[] = [];
  for (const sentence of sentences) {
    const next = [...selected, sentence.trim()].join(" ");
    if (selected.length >= maxSentences || next.length > maxChars) break;
    selected.push(sentence.trim());
  }
  return (selected.join(" ") || cleaned.slice(0, maxChars)).trim();
}

function contextBulletPoints(text: string, tableRows: string[][] | null, maxPoints = 8) {
  const bulletRows = bulletRowsFromText(text);
  if (bulletRows && bulletRows.length > 1) return bulletRows.slice(1, maxPoints + 1).map((row) => [row[0], row[1]].filter(Boolean).join(" - "));
  const keyValues = keyValueRowsFromText(text);
  if (keyValues && keyValues.length > 1) return keyValues.slice(1, maxPoints + 1).map((row) => `${row[0]}: ${row[1]}`);
  if (tableRows && tableRows.length > 1) {
    const headers = tableRows[0] ?? [];
    return tableRows.slice(1, maxPoints + 1).map((row) => row.map((cell, index) => headers[index] ? `${headers[index]}: ${cell}` : cell).filter(Boolean).join("; "));
  }
  return text.split(/\r?\n/).map((line) => line.trim().replace(/^[-*•]\s+/, "")).filter((line) => line.length > 0 && line.length < 220).slice(0, maxPoints);
}

export function createWordContextBrief(textValue: unknown, options: { title?: unknown; sourceLabel?: unknown; audience?: unknown; includeSourceTable?: boolean; maxPoints?: unknown } = {}) {
  const text = asString(textValue, "").trim();
  if (!text) throw new Error("Word context brief requires source text.");
  const title = asString(options.title, "Context Brief").trim() || "Context Brief";
  const sourceLabel = asString(options.sourceLabel, "Provided context").trim() || "Provided context";
  const audience = asString(options.audience, "professional").trim() || "professional";
  const tableRows = extractTableRows(text) ?? keyValueRowsFromText(text);
  const maxPoints = Math.max(3, Math.min(15, asOptionalNumber(options.maxPoints) ?? 8));
  const summary = firstSentences(text, audience === "executive" ? 2 : 3, audience === "executive" ? 500 : 700);
  const keyPoints = contextBulletPoints(text, tableRows, maxPoints);
  return {
    title,
    sourceLabel,
    audience,
    summary,
    keyPoints: keyPoints.length ? keyPoints : [summary],
    tableRows: options.includeSourceTable === false ? null : tableRows,
  };
}

type PowerPointAuditSlide = { slideNumber?: number; shapeCount?: number; text?: string; textBoxCount?: number; imageCount?: number; tableCount?: number };
type PowerPointDiagramStep = { label: string; detail?: string; owner?: string; status?: string };
export type PowerPointDiagramFrame = { kind: "title" | "card" | "connector"; text: string; left: number; top: number; width: number; height: number; index?: number };

export function createPowerPointAuditRows(slidesValue: PowerPointAuditSlide[] = [], options: { audience?: unknown } = {}) {
  const slides = Array.isArray(slidesValue) ? slidesValue : [];
  const audience = asString(options.audience, "executive").trim() || "executive";
  const rows: string[][] = [["Area", "Severity", "Finding", "Recommendation"]];
  if (!slides.length) {
    rows.push(["Deck", "High", "No slides were readable through the live PowerPoint runtime.", "Create or open a deck, or upload the PPTX for package-level review."]);
    return rows;
  }

  const totalShapes = slides.reduce((sum, slide) => sum + Math.max(0, slide.shapeCount ?? 0), 0);
  const emptySlides = slides.filter((slide) => (slide.shapeCount ?? 0) === 0);
  const denseSlides = slides.filter((slide) => (slide.shapeCount ?? 0) > 14);
  const textHeavySlides = slides.filter((slide) => String(slide.text ?? "").length > 750);
  const lowTextSlides = slides.filter((slide) => (slide.shapeCount ?? 0) > 0 && String(slide.text ?? "").trim().length < 12);

  if (slides.length < 3) rows.push(["Narrative", "Medium", `Deck has only ${slides.length} slide${slides.length === 1 ? "" : "s"}.`, `For a ${audience} deck, confirm the story has context, insight, recommendation, and next-step coverage.`]);
  if (slides.length > 25) rows.push(["Narrative", "Medium", `Deck has ${slides.length} slides.`, "Consider an executive summary section, appendix split, or section dividers so the main story stays crisp."]);
  if (emptySlides.length) rows.push(["Completeness", "High", `${emptySlides.length} blank slide${emptySlides.length === 1 ? "" : "s"} detected.`, "Remove blank slides or turn them into intentional section dividers with titles."]);
  if (denseSlides.length) rows.push(["Design", "Medium", `${denseSlides.length} slide${denseSlides.length === 1 ? "" : "s"} have more than 14 objects.`, "Simplify crowded slides, group related elements, or split dense content across multiple slides."]);
  if (textHeavySlides.length) rows.push(["Readability", "Medium", `${textHeavySlides.length} slide${textHeavySlides.length === 1 ? "" : "s"} appear text-heavy.`, "Convert dense prose into concise headlines, bullets, charts, or speaker notes."]);
  if (lowTextSlides.length === slides.length && totalShapes > 0) rows.push(["Accessibility", "Medium", "Slides contain objects but little readable text was exposed to the task-pane runtime.", "Confirm each meaningful image/chart has accessible title or alt-text, and upload the PPTX for package-level media/theme review if needed."]);
  if (totalShapes / slides.length < 2) rows.push(["Visual Story", "Low", "Average object count per slide is low.", "Confirm slides are not overly plain; add charts, icons, tables, or visual hierarchy where they clarify the message."]);
  if (!slides.some((slide) => /agenda|summary|recommendation|next steps?|decision/i.test(String(slide.text ?? "")))) rows.push(["Executive Flow", "Low", "No agenda, summary, recommendation, decision, or next-steps language was detected in readable slide text.", "Add an executive orientation slide or clear decision/next-step slide when the deck is for leadership."]);
  rows.push(["Runtime", "Info", "Live PowerPoint review is based on slides and shape/text properties exposed by Office.js.", "For notes, template/theme, embedded media, chart-data, and speaker-note validation, use uploaded PPTX context or generated-deck QA."]);
  if (rows.length === 2) rows.push(["Overall", "Info", "No common deck review issues were detected from the live runtime snapshot.", "Still run visual QA across target PowerPoint clients before final delivery."]);
  return rows;
}

function normalizeDiagramSteps(value: unknown): PowerPointDiagramStep[] {
  const raw = Array.isArray(value) ? value : [];
  return raw.map((item) => {
    if (typeof item === "string" || typeof item === "number") return { label: String(item).trim() };
    if (item && typeof item === "object") {
      const record = item as Record<string, unknown>;
      return {
        label: asString(record.label ?? record.title ?? record.name, "").trim(),
        detail: asString(record.detail ?? record.description, "").trim() || undefined,
        owner: asString(record.owner, "").trim() || undefined,
        status: asString(record.status, "").trim() || undefined,
      };
    }
    return { label: "" };
  }).filter((step) => step.label).slice(0, 12);
}

function diagramCardText(step: PowerPointDiagramStep, index: number, numbered: boolean) {
  const lines = [`${numbered ? `${index + 1}. ` : ""}${step.label}`];
  if (step.detail) lines.push(step.detail);
  const meta = [step.owner, step.status].filter(Boolean).join(" - ");
  if (meta) lines.push(meta);
  return lines.join("\n");
}

export function createPowerPointDiagramFrames(stepsValue: unknown, typeValue: unknown = "process", boundsValue: Record<string, unknown> = {}): PowerPointDiagramFrame[] {
  const steps = normalizeDiagramSteps(stepsValue);
  if (!steps.length) return [];
  const type = compactKey(typeValue || "process");
  const left = asOptionalNumber(boundsValue.left) ?? 60;
  const top = asOptionalNumber(boundsValue.top) ?? 120;
  const width = asOptionalNumber(boundsValue.width) ?? 600;
  const height = asOptionalNumber(boundsValue.height) ?? 230;
  const title = asString(boundsValue.title, "").trim();
  const numbered = boolValue(boundsValue.numbered, type !== "cycle");
  const frames: PowerPointDiagramFrame[] = [];
  const diagramTop = title ? top + 54 : top;
  const diagramHeight = title ? Math.max(120, height - 54) : height;
  if (title) frames.push({ kind: "title", text: title, left, top, width, height: 38 });

  if (["cycle", "loop", "circular"].includes(type)) {
    const count = steps.length;
    const cardWidth = Math.min(150, Math.max(105, width / Math.max(2.8, Math.sqrt(count) + 1)));
    const cardHeight = Math.min(78, Math.max(54, diagramHeight / 3.2));
    const centerX = left + width / 2;
    const centerY = diagramTop + diagramHeight / 2;
    const radiusX = Math.max(80, width / 2 - cardWidth / 2 - 12);
    const radiusY = Math.max(54, diagramHeight / 2 - cardHeight / 2 - 6);
    steps.forEach((step, index) => {
      const angle = -Math.PI / 2 + (index / count) * Math.PI * 2;
      frames.push({ kind: "card", index, text: diagramCardText(step, index, numbered), left: centerX + Math.cos(angle) * radiusX - cardWidth / 2, top: centerY + Math.sin(angle) * radiusY - cardHeight / 2, width: cardWidth, height: cardHeight });
    });
    steps.forEach((_, index) => {
      const current = frames.find((frame) => frame.kind === "card" && frame.index === index);
      const next = frames.find((frame) => frame.kind === "card" && frame.index === (index + 1) % steps.length);
      if (current && next && steps.length > 1) frames.push({ kind: "connector", index, text: "→", left: (current.left + next.left) / 2 + cardWidth / 2 - 10, top: (current.top + next.top) / 2 + cardHeight / 2 - 13, width: 24, height: 24 });
    });
    return frames;
  }

  const columns = ["roadmap", "timeline"].includes(type) && steps.length > 4 ? Math.min(4, steps.length) : steps.length;
  const rows = Math.ceil(steps.length / Math.max(1, columns));
  const gap = 16;
  const cardWidth = (width - gap * (columns - 1)) / columns;
  const cardHeight = Math.min(92, (diagramHeight - gap * (rows - 1)) / rows);
  steps.forEach((step, index) => {
    const column = index % columns;
    const row = Math.floor(index / columns);
    const cardLeft = left + column * (cardWidth + gap);
    const cardTop = diagramTop + row * (cardHeight + gap);
    frames.push({ kind: "card", index, text: diagramCardText(step, index, numbered), left: cardLeft, top: cardTop, width: cardWidth, height: cardHeight });
    const isRowEnd = column === columns - 1 || index === steps.length - 1;
    if (!isRowEnd) frames.push({ kind: "connector", index, text: "→", left: cardLeft + cardWidth + 2, top: cardTop + cardHeight / 2 - 13, width: Math.max(18, gap - 4), height: 24 });
  });
  return frames;
}

export async function executeOfficeTool(host: OfficeHost, call: ToolCallRequest): Promise<ToolCallResult> {
  try {
    if (call.name === "office_read_context") {
      const context = await readDocumentContext(host);
      return toolResult(call, true, JSON.stringify(context));
    }
    if (host === "excel" && globalThis.Excel) return executeExcelTool(call);
    if (host === "word" && globalThis.Word) return executeWordTool(call);
    if (host === "powerpoint" && globalThis.PowerPoint) return executePowerPointTool(call);
    return toolResult(call, false, `Tool ${call.name} is not available in ${host}.`);
  } catch (error: any) {
    return toolResult(call, false, officeErrorMessage(error, String(error), call));
  }
}

async function executeExcelTool(call: ToolCallRequest): Promise<ToolCallResult> {
  if (call.name === "excel_get_workbook_overview") {
    return globalThis.Excel.run(async (ctx: any) => {
      const includeHeaders = boolValue(call.arguments.includeHeaders, true);
      const maxSheets = Math.max(1, Math.min(100, Math.trunc(asOptionalNumber(call.arguments.maxSheets) ?? 25)));
      const workbook = ctx.workbook;
      const worksheets = workbook.worksheets;
      workbook.load("name");
      worksheets.load("items/name,items/visibility");
      const selected = workbook.getSelectedRange();
      selected.load("address");
      const activeSheet = worksheets.getActiveWorksheet();
      activeSheet.load("name");
      let workbookNames: any = null;
      try { workbookNames = workbook.names; workbookNames.load("items/name,items/formula"); } catch { workbookNames = null; }
      await ctx.sync();

      const sheetItems = (worksheets.items ?? []).slice(0, maxSheets);
      const loaded = sheetItems.map((sheet: any) => {
        const used = sheet.getUsedRangeOrNullObject?.() ?? sheet.getUsedRange?.();
        try { used?.load?.("address,rowCount,columnCount,isNullObject"); } catch {}
        let tables: any = null;
        let charts: any = null;
        try { tables = sheet.tables; tables.load("items/name"); } catch { tables = null; }
        try { charts = sheet.charts; charts.load("items/name"); } catch { charts = null; }
        return { sheet, used, tables, charts };
      });
      await ctx.sync();

      const headerRanges = includeHeaders ? loaded.map((entry: any) => {
        if (!entry.used || entry.used.isNullObject) return null;
        try {
          const header = entry.used.getRow?.(0);
          header?.load?.("values");
          return header;
        } catch { return null; }
      }) : loaded.map(() => null);
      await ctx.sync();

      const sheets = loaded.map((entry: any, index: number) => ({
        name: entry.sheet.name,
        usedRange: entry.used && !entry.used.isNullObject ? String(entry.used.address ?? "") : "",
        rowCount: entry.used && !entry.used.isNullObject ? Number(entry.used.rowCount ?? 0) : 0,
        columnCount: entry.used && !entry.used.isNullObject ? Number(entry.used.columnCount ?? 0) : 0,
        visible: entry.sheet.visibility ? String(entry.sheet.visibility).toLowerCase() === "visible" : true,
        headers: headerRanges[index]?.values?.[0] ?? [],
        tables: (entry.tables?.items ?? []).map((table: any) => String(table.name ?? "")).filter(Boolean),
        charts: (entry.charts?.items ?? []).map((chart: any) => String(chart.name ?? "")).filter(Boolean),
      }));

      const namedRanges = (workbookNames?.items ?? []).map((item: any) => ({
        name: String(item.name ?? ""),
        scope: "workbook",
        refersTo: String(item.formula ?? ""),
      })).filter((item: any) => item.name);

      const overview = buildWorkbookOverview({
        workbookName: workbook.name || "Workbook",
        activeSheet: activeSheet?.name || "",
        selection: selected?.address || "",
        sheets,
        namedRanges,
      });
      const omitted = (worksheets.items ?? []).length - sheets.length;
      const suffix = omitted > 0 ? `${"\n"}(${omitted} more worksheets not described; raise maxSheets to include them.)` : "";
      return toolResult(call, true, overview.text + suffix);
    });
  }

  if (call.name === "excel_read_range") {
    return globalThis.Excel.run(async (ctx: any) => {
      const mode = normalizeReadMode(call.arguments.mode);
      const requested = normalizeAddress(call.arguments.address);
      const sheetNameArg = asString(call.arguments.sheetName, "").trim();
      let range: any;
      if (requested) {
        range = activeOrNamedWorksheet(ctx, call.arguments.sheetName).getRange(requested);
      } else if (sheetNameArg) {
        const sheet = ctx.workbook.worksheets.getItem(sheetNameArg);
        range = sheet.getUsedRangeOrNullObject?.() ?? sheet.getUsedRange?.();
      } else {
        range = ctx.workbook.getSelectedRange();
      }
      if (!range) throw new Error("excel_read_range could not resolve a range to read.");
      const load = ["address", "values", "rowCount", "columnCount"];
      if (mode === "detailed") load.push("formulas", "numberFormat");
      try { range.load(load.join(",")); } catch { range.load("address,values"); }
      await ctx.sync();

      const formatted = formatRangeForModel(range.values ?? [], {
        mode,
        baseAddress: range.address,
        formulas: mode === "detailed" ? (range.formulas ?? []) : [],
        numberFormats: mode === "detailed" ? (range.numberFormat ?? []) : [],
        maxRows: asOptionalNumber(call.arguments.maxRows) ?? 200,
        maxColumns: asOptionalNumber(call.arguments.maxColumns) ?? 50,
      });
      const header = `Range ${range.address} (${mode} mode):`;
      const body = [header, formatted.text, ...formatted.notes].filter(Boolean).join("\n");
      return toolResult(call, true, body);
    });
  }

  if (call.name === "excel_search_workbook") {
    return globalThis.Excel.run(async (ctx: any) => {
      const query = asString(call.arguments.query, "").trim();
      if (!query) throw new Error("excel_search_workbook requires query.");
      const limit = Math.max(1, Math.min(1000, Math.trunc(asOptionalNumber(call.arguments.limit) ?? 200)));
      const targetSheet = asString(call.arguments.sheetName, "").trim();
      const worksheets = ctx.workbook.worksheets;
      worksheets.load("items/name");
      await ctx.sync();

      const sheets = (worksheets.items ?? []).filter((sheet: any) => !targetSheet || String(sheet.name) === targetSheet);
      if (targetSheet && !sheets.length) throw new Error(`Worksheet ${targetSheet} was not found.`);
      const used = sheets.map((sheet: any) => {
        const range = sheet.getUsedRangeOrNullObject?.() ?? sheet.getUsedRange?.();
        try { range?.load?.("address,values,formulas,isNullObject"); } catch {}
        return { sheet, range };
      });
      await ctx.sync();

      const matches = used.flatMap((entry: any) => {
        if (!entry.range || entry.range.isNullObject) return [];
        return searchRangeMatches(String(entry.sheet.name ?? ""), entry.range.values ?? [], entry.range.formulas ?? [], entry.range.address, query, {
          matchCase: boolValue(call.arguments.matchCase, false),
          wholeCell: boolValue(call.arguments.wholeCell, false),
          searchFormulas: boolValue(call.arguments.searchFormulas, true),
          limit,
        });
      }).slice(0, limit);

      return toolResult(call, true, formatSearchMatches(matches, query, limit));
    });
  }
  if (call.name === "excel_add_worksheet") {
    return globalThis.Excel.run(async (ctx: any) => {
      const requestedName = asString(call.arguments.name, "AI Output").slice(0, 31);
      const sheet = ctx.workbook.worksheets.add(requestedName);
      sheet.activate();
      await ctx.sync();
      return toolResult(call, true, `Created worksheet ${requestedName}.`);
    });
  }

  if (call.name === "excel_rename_worksheet") {
    return globalThis.Excel.run(async (ctx: any) => {
      const newName = asString(call.arguments.newName, "").trim().slice(0, 31);
      if (!newName) throw new Error("excel_rename_worksheet requires newName.");
      const sheet = activeOrNamedWorksheet(ctx, call.arguments.sheetName);
      sheet.load("name");
      await ctx.sync();
      const oldName = sheet.name || "active sheet";
      sheet.name = newName;
      await ctx.sync();
      return toolResult(call, true, `Renamed worksheet ${oldName} to ${newName}.`);
    });
  }

  if (call.name === "excel_delete_worksheet") {
    return globalThis.Excel.run(async (ctx: any) => {
      const sheetName = asString(call.arguments.sheetName, "").trim();
      if (!sheetName) throw new Error("excel_delete_worksheet requires sheetName.");
      const sheet = ctx.workbook.worksheets.getItem(sheetName);
      if (!sheet.delete) throw new Error("This Excel runtime does not expose worksheet deletion APIs.");
      sheet.delete();
      await ctx.sync();
      return toolResult(call, true, `Deleted worksheet ${sheetName}.`);
    });
  }

  if (call.name === "excel_set_worksheet_visibility") {
    return globalThis.Excel.run(async (ctx: any) => {
      const requested = asString(call.arguments.visibility, "visible");
      const sheet = activeOrNamedWorksheet(ctx, call.arguments.sheetName);
      const visibilityMap: Record<string, any> = {
        visible: globalThis.Excel.SheetVisibility?.visible ?? "Visible",
        hidden: globalThis.Excel.SheetVisibility?.hidden ?? "Hidden",
        veryHidden: globalThis.Excel.SheetVisibility?.veryHidden ?? "VeryHidden",
      };
      sheet.visibility = visibilityMap[requested] ?? visibilityMap.visible;
      await ctx.sync();
      return toolResult(call, true, `Set worksheet visibility to ${requested}.`);
    });
  }

  if (call.name === "excel_clear_range") {
    return globalThis.Excel.run(async (ctx: any) => {
      const address = normalizeAddress(call.arguments.address);
      if (!address) throw new Error("excel_clear_range requires address.");
      const sheet = activeOrNamedWorksheet(ctx, call.arguments.sheetName);
      const range = sheet.getRange(address);
      const applyTo = asString(call.arguments.applyTo, "contents");
      const clearApplyTo = globalThis.Excel.ClearApplyTo ?? {};
      const map: Record<string, any> = { contents: clearApplyTo.contents ?? "Contents", formats: clearApplyTo.formats ?? "Formats", all: clearApplyTo.all ?? "All" };
      range.clear(map[applyTo] ?? map.contents);
      await ctx.sync();
      return toolResult(call, true, `Cleared ${applyTo} from ${address}.`);
    });
  }

  if (call.name === "excel_write_range") {
    return globalThis.Excel.run(async (ctx: any) => {
      const rows = asRows(call.arguments.values);
      if (!rows.length) throw new Error("excel_write_range requires non-empty values.");
      const columnCount = Math.max(1, ...rows.map((row) => row.length));
      const normalized = rows.map((row) => Array.from({ length: columnCount }, (_, index) => row[index] ?? ""));
      const sheetName = asString(call.arguments.sheetName, "").trim();
      const address = normalizeAddress(call.arguments.address) || "A1";
      const sheet = activeOrNamedWorksheet(ctx, call.arguments.sheetName);
      const start = sheet.getRange(address);
      const target = start.getResizedRange(normalized.length - 1, columnCount - 1);

      // Overwrite protection: never silently destroy existing content.
      const allowOverwrite = boolValue(call.arguments.overwrite, false);
      let conflictSummary = "";
      if (!allowOverwrite) {
        try {
          target.load("values,address");
          await ctx.sync();
          const conflicts = describeOverwriteConflicts(target.values ?? [], target.address ?? address);
          if (conflicts.count > 0) {
            return toolResult(call, false, [
              `Refused to write to ${sheetName || "active sheet"}!${address} because it is not empty.`,
              conflicts.summary,
              "Call excel_read_range to inspect the current values, then retry with overwrite=true if replacing them is intended.",
            ].join(" "));
          }
        } catch {
          conflictSummary = "";
        }
      }

      target.values = normalized;
      target.format.autofitColumns();
      target.format.autofitRows();
      await ctx.sync();

      // Auto-verification: confirm Excel stored what we asked it to store.
      let verification = "";
      try {
        target.load("values,address");
        await ctx.sync();
        const check = verifyWrittenValues(normalized, target.values ?? [], target.address ?? address);
        verification = check.verified ? "" : ` ${check.summary}`;
      } catch {
        verification = "";
      }

      return toolResult(call, true, `Wrote ${normalized.length} rows and ${columnCount} columns to ${sheetName || "active sheet"}!${address}.${verification}${conflictSummary}`);
    });
  }

  if (call.name === "excel_import_context_table") {
    return globalThis.Excel.run(async (ctx: any) => {
      const imported = createExcelContextImportRows(call.arguments.text, {
        mode: call.arguments.mode,
        preserveFormulas: boolValue(call.arguments.preserveFormulas, false),
        includeSourceColumn: boolValue(call.arguments.includeSourceColumn, true),
        sourceLabel: call.arguments.sourceLabel,
        maxRows: call.arguments.maxRows,
      });
      let rows = imported.rows;
      if (boolValue(call.arguments.clean, true) && rows.length) {
        rows = createCleanedRows(rows, {
          hasHeaders: true,
          trimWhitespace: true,
          normalizeSpaces: true,
          removeBlankRows: true,
          removeDuplicateRows: boolValue(call.arguments.removeDuplicateRows, false),
        }).rows.map((row) => row.map((cell) => String(cell ?? "")));
      }
      if (!rows.length) throw new Error("Context import produced no rows to write.");
      const outputSheetName = safeSheetName(call.arguments.outputSheetName, "Imported Context");
      const outputAddress = normalizeAddress(call.arguments.outputAddress) || "A1";
      const columnCount = Math.max(1, ...rows.map((row) => row.length));
      const normalized = rows.map((row) => Array.from({ length: columnCount }, (_, index) => row[index] ?? ""));
      let outputSheet: any;
      try {
        outputSheet = ctx.workbook.worksheets.getItem(outputSheetName);
        outputSheet.load("name");
        await ctx.sync();
      } catch {
        outputSheet = ctx.workbook.worksheets.add(outputSheetName);
        await ctx.sync();
      }
      const used = outputSheet.getUsedRangeOrNullObject?.() ?? outputSheet.getUsedRange?.();
      if (used) {
        try { used.load?.("isNullObject"); await ctx.sync(); if (!used.isNullObject) used.clear(globalThis.Excel.ClearApplyTo?.all ?? "All"); } catch { try { used.clear?.(); } catch {} }
      }
      const target = outputSheet.getRange(outputAddress).getResizedRange(normalized.length - 1, columnCount - 1);
      target.values = normalized;
      try {
        const header = outputSheet.getRange(outputAddress).getResizedRange(0, columnCount - 1);
        header.format.fill.color = "#1f4e79";
        header.format.font.color = "#ffffff";
        header.format.font.bold = true;
      } catch {}
      if (boolValue(call.arguments.createTable, true) && normalized.length >= 2) {
        try {
          const table = outputSheet.tables.add(target, true);
          table.name = safeExcelName(call.arguments.tableName || outputSheetName + " Table");
          table.style = asString(call.arguments.style, "TableStyleMedium2") || "TableStyleMedium2";
        } catch {}
      }
      try { target.format.autofitColumns(); target.format.autofitRows(); } catch {}
      outputSheet.activate();
      await ctx.sync();
      return toolResult(call, true, `Imported ${normalized.length} row${normalized.length === 1 ? "" : "s"} and ${columnCount} column${columnCount === 1 ? "" : "s"} of ${imported.detectedMode} context into ${outputSheetName}!${outputAddress}.`);
    });
  }

  if (call.name === "excel_set_formula") {
    return globalThis.Excel.run(async (ctx: any) => {
      const address = normalizeAddress(call.arguments.address);
      const formula = asString(call.arguments.formula, "");
      if (!address || !formula) throw new Error("excel_set_formula requires address and formula.");
      const sheetName = asString(call.arguments.sheetName, "").trim();
      const sheet = activeOrNamedWorksheet(ctx, call.arguments.sheetName);
      sheet.getRange(address).formulas = [[formula]];
      await ctx.sync();
      return toolResult(call, true, `Set formula in ${sheetName || "active sheet"}!${address}.`);
    });
  }

  if (call.name === "excel_clean_transform_range") {
    return globalThis.Excel.run(async (ctx: any) => {
      const sourceSheet = activeOrNamedWorksheet(ctx, call.arguments.sheetName);
      sourceSheet.load("name");
      const requestedAddress = normalizeAddress(call.arguments.sourceAddress || call.arguments.address);
      const source = requestedAddress ? sourceSheet.getRange(requestedAddress) : sourceSheet.getUsedRange();
      if (!source) throw new Error("No source range was available to clean.");
      source.load(["address", "values", "rowCount", "columnCount"]);
      await ctx.sync();
      const cleaned = createCleanedRows(source.values ?? [], {
        trimWhitespace: boolValue(call.arguments.trimWhitespace, true),
        normalizeSpaces: boolValue(call.arguments.normalizeSpaces, true),
        removeBlankRows: boolValue(call.arguments.removeBlankRows, true),
        removeDuplicateRows: boolValue(call.arguments.removeDuplicateRows, false),
        splitColumn: call.arguments.splitColumn,
        delimiter: call.arguments.delimiter,
        hasHeaders: boolValue(call.arguments.hasHeaders, true),
        caseMode: call.arguments.caseMode,
      });
      if (!cleaned.rows.length) throw new Error("Cleaning produced no rows to write.");
      const outputSheetName = asString(call.arguments.outputSheetName, "Cleaned Data").trim().slice(0, 31) || "Cleaned Data";
      const outputAddress = normalizeAddress(call.arguments.outputAddress) || "A1";
      let outputSheet: any;
      try {
        outputSheet = ctx.workbook.worksheets.getItem(outputSheetName);
        outputSheet.load("name");
        await ctx.sync();
      } catch {
        outputSheet = ctx.workbook.worksheets.add(outputSheetName);
        await ctx.sync();
      }
      const used = outputSheet.getUsedRangeOrNullObject?.() ?? outputSheet.getUsedRange?.();
      if (used) {
        try { used.load?.("isNullObject"); await ctx.sync(); if (!used.isNullObject) used.clear(globalThis.Excel.ClearApplyTo?.all ?? "All"); } catch { try { used.clear?.(); } catch {} }
      }
      const start = outputSheet.getRange(outputAddress);
      const target = start.getResizedRange(cleaned.rows.length - 1, cleaned.outputColumns - 1);
      target.values = cleaned.rows;
      try {
        const header = outputSheet.getRange(outputAddress).getResizedRange(0, cleaned.outputColumns - 1);
        header.format.fill.color = "#1f4e79";
        header.format.font.color = "#ffffff";
        header.format.font.bold = true;
      } catch {}
      if (boolValue(call.arguments.createTable, true)) {
        try {
          const table = outputSheet.tables.add(target, boolValue(call.arguments.hasHeaders, true));
          table.name = safeExcelName(call.arguments.tableName || outputSheetName + " Table");
          table.style = asString(call.arguments.style, "TableStyleMedium2") || "TableStyleMedium2";
        } catch {}
      }
      try { target.format.autofitColumns(); target.format.autofitRows(); } catch {}
      outputSheet.activate();
      await ctx.sync();
      const summary = [
        `Created cleaned Excel table on ${outputSheetName}!${outputAddress}`,
        `from ${sourceSheet.name || "source sheet"}!${source.address || requestedAddress || "used range"}`,
        `with ${cleaned.outputDataRows} data rows and ${cleaned.outputColumns} columns`,
        cleaned.removedBlankRows ? `removed ${cleaned.removedBlankRows} blank row${cleaned.removedBlankRows === 1 ? "" : "s"}` : "",
        cleaned.removedDuplicateRows ? `removed ${cleaned.removedDuplicateRows} duplicate row${cleaned.removedDuplicateRows === 1 ? "" : "s"}` : "",
        cleaned.transforms.length ? `transforms: ${cleaned.transforms.join(", ")}` : "",
      ].filter(Boolean).join("; ") + ".";
      return toolResult(call, true, summary);
    });
  }

  if (call.name === "excel_combine_ranges") {
    return globalThis.Excel.run(async (ctx: any) => {
      const primaryAddress = normalizeAddress(call.arguments.primaryAddress);
      const secondaryAddress = normalizeAddress(call.arguments.secondaryAddress);
      if (!primaryAddress || !secondaryAddress) throw new Error("excel_combine_ranges requires primaryAddress and secondaryAddress.");
      const primarySheet = activeOrNamedWorksheet(ctx, call.arguments.primarySheetName || call.arguments.sheetName);
      const secondarySheet = activeOrNamedWorksheet(ctx, call.arguments.secondarySheetName || call.arguments.sheetName);
      primarySheet.load("name");
      secondarySheet.load("name");
      const primaryRange = primarySheet.getRange(primaryAddress);
      const secondaryRange = secondarySheet.getRange(secondaryAddress);
      primaryRange.load(["address", "values", "rowCount", "columnCount"]);
      secondaryRange.load(["address", "values", "rowCount", "columnCount"]);
      await ctx.sync();
      const combined = createCombinedRows(primaryRange.values ?? [], secondaryRange.values ?? [], {
        mode: call.arguments.mode,
        matchColumns: call.arguments.matchColumns,
        primaryKey: call.arguments.primaryKey,
        secondaryKey: call.arguments.secondaryKey,
        conflictSuffix: call.arguments.conflictSuffix,
        includeSourceColumn: boolValue(call.arguments.includeSourceColumn, false),
        primaryLabel: call.arguments.primaryLabel || primarySheet.name,
        secondaryLabel: call.arguments.secondaryLabel || secondarySheet.name,
        trimWhitespace: boolValue(call.arguments.trimWhitespace, true),
        normalizeSpaces: boolValue(call.arguments.normalizeSpaces, true),
        removeBlankRows: boolValue(call.arguments.removeBlankRows, true),
        removeDuplicateRows: boolValue(call.arguments.removeDuplicateRows, false),
        hasHeaders: boolValue(call.arguments.hasHeaders, true),
        caseMode: call.arguments.caseMode,
      });
      const outputSheetName = asString(call.arguments.outputSheetName, combined.mode === "merge" ? "Merged Data" : "Appended Data").trim().slice(0, 31) || (combined.mode === "merge" ? "Merged Data" : "Appended Data");
      const outputAddress = normalizeAddress(call.arguments.outputAddress) || "A1";
      let outputSheet: any;
      try {
        outputSheet = ctx.workbook.worksheets.getItem(outputSheetName);
        outputSheet.load("name");
        await ctx.sync();
      } catch {
        outputSheet = ctx.workbook.worksheets.add(outputSheetName);
        await ctx.sync();
      }
      const used = outputSheet.getUsedRangeOrNullObject?.() ?? outputSheet.getUsedRange?.();
      if (used) {
        try { used.load?.("isNullObject"); await ctx.sync(); if (!used.isNullObject) used.clear(globalThis.Excel.ClearApplyTo?.all ?? "All"); } catch { try { used.clear?.(); } catch {} }
      }
      const target = outputSheet.getRange(outputAddress).getResizedRange(combined.rows.length - 1, combined.outputColumns - 1);
      target.values = combined.rows;
      try {
        const header = outputSheet.getRange(outputAddress).getResizedRange(0, combined.outputColumns - 1);
        header.format.fill.color = "#1f4e79";
        header.format.font.color = "#ffffff";
        header.format.font.bold = true;
      } catch {}
      if (boolValue(call.arguments.createTable, true)) {
        try {
          const table = outputSheet.tables.add(target, boolValue(call.arguments.hasHeaders, true));
          table.name = safeExcelName(call.arguments.tableName || outputSheetName + " Table");
          table.style = asString(call.arguments.style, "TableStyleMedium4") || "TableStyleMedium4";
        } catch {}
      }
      try { target.format.autofitColumns(); target.format.autofitRows(); } catch {}
      outputSheet.activate();
      await ctx.sync();
      const matchSummary = combined.mode === "merge" ? `; matched ${combined.matchedRows} of ${combined.primaryRows} primary rows` : "";
      return toolResult(call, true, `Created ${combined.mode === "merge" ? "merged" : "appended"} Excel table on ${outputSheetName}!${outputAddress} with ${combined.outputDataRows} data rows and ${combined.outputColumns} columns from ${primarySheet.name || "primary"}!${primaryRange.address || primaryAddress} and ${secondarySheet.name || "secondary"}!${secondaryRange.address || secondaryAddress}${matchSummary}; transforms: ${combined.transforms.join(", ")}.`);
    });
  }

  if (call.name === "excel_audit_formulas") {
    return globalThis.Excel.run(async (ctx: any) => {
      const sourceSheet = activeOrNamedWorksheet(ctx, call.arguments.sheetName);
      sourceSheet.load("name");
      const requestedAddress = normalizeAddress(call.arguments.address);
      const source = requestedAddress ? sourceSheet.getRange(requestedAddress) : sourceSheet.getUsedRange();
      if (!source) throw new Error("No used range was available to audit.");
      source.load(["address", "formulas", "values", "rowCount", "columnCount"]);
      await ctx.sync();
      const auditRows = createFormulaAuditRows(source.formulas ?? [], source.values ?? [], source.address || requestedAddress || "A1", boolValue(call.arguments.includeLowRisk, false));
      const outputSheetName = asString(call.arguments.outputSheetName, "Formula Audit").trim().slice(0, 31) || "Formula Audit";
      let outputSheet: any;
      try {
        outputSheet = ctx.workbook.worksheets.getItem(outputSheetName);
        outputSheet.load("name");
        await ctx.sync();
      } catch {
        outputSheet = ctx.workbook.worksheets.add(outputSheetName);
        await ctx.sync();
      }
      const used = outputSheet.getUsedRangeOrNullObject?.() ?? outputSheet.getUsedRange?.();
      if (used) {
        try { used.load?.("isNullObject"); await ctx.sync(); if (!used.isNullObject) used.clear(globalThis.Excel.ClearApplyTo?.all ?? "All"); } catch { try { used.clear?.(); } catch {} }
      }
      const columnCount = Math.max(1, ...auditRows.map((row) => row.length));
      const normalized = auditRows.map((row) => Array.from({ length: columnCount }, (_, index) => row[index] ?? ""));
      const target = outputSheet.getRange("A1").getResizedRange(normalized.length - 1, columnCount - 1);
      target.values = normalized;
      try {
        const header = outputSheet.getRange("A1:E1");
        header.format.fill.color = "#1f4e79";
        header.format.font.color = "#ffffff";
        header.format.font.bold = true;
      } catch {}
      try {
        const table = outputSheet.tables.add(target, true);
        table.name = safeExcelName(outputSheetName + " Table");
        table.style = "TableStyleMedium2";
      } catch {}
      try { outputSheet.getUsedRange()?.format?.autofitColumns?.(); outputSheet.getUsedRange()?.format?.autofitRows?.(); } catch {}
      outputSheet.activate();
      await ctx.sync();
      const findingCount = Math.max(0, auditRows.length - 1);
      return toolResult(call, true, `Created Excel formula audit on ${outputSheetName} with ${findingCount} finding${findingCount === 1 ? "" : "s"} from ${sourceSheet.name || "source sheet"}!${source.address || requestedAddress || "used range"}.`);
    });
  }

  if (call.name === "excel_map_formula_dependencies") {
    return globalThis.Excel.run(async (ctx: any) => {
      const sourceSheet = activeOrNamedWorksheet(ctx, call.arguments.sheetName);
      sourceSheet.load("name");
      const requestedAddress = normalizeAddress(call.arguments.address);
      const source = requestedAddress ? sourceSheet.getRange(requestedAddress) : sourceSheet.getUsedRange();
      if (!source) throw new Error("No used range was available to map.");
      source.load(["address", "formulas", "rowCount", "columnCount"]);
      await ctx.sync();
      const outputSheetName = asString(call.arguments.outputSheetName, "Dependency Map").trim().slice(0, 31) || "Dependency Map";
      const dependencyRows = createFormulaDependencyRows(source.formulas ?? [], source.address || requestedAddress || "A1", sourceSheet.name || "Sheet", boolValue(call.arguments.includeSummary, true));
      let outputSheet: any;
      try {
        outputSheet = ctx.workbook.worksheets.getItem(outputSheetName);
        outputSheet.load("name");
        await ctx.sync();
      } catch {
        outputSheet = ctx.workbook.worksheets.add(outputSheetName);
        await ctx.sync();
      }
      const used = outputSheet.getUsedRangeOrNullObject?.() ?? outputSheet.getUsedRange?.();
      if (used) {
        try { used.load?.("isNullObject"); await ctx.sync(); if (!used.isNullObject) used.clear(globalThis.Excel.ClearApplyTo?.all ?? "All"); } catch { try { used.clear?.(); } catch {} }
      }
      const columnCount = Math.max(1, ...dependencyRows.map((row) => row.length));
      const normalized = dependencyRows.map((row) => Array.from({ length: columnCount }, (_, index) => row[index] ?? ""));
      const target = outputSheet.getRange("A1").getResizedRange(normalized.length - 1, columnCount - 1);
      target.values = normalized;
      try {
        const header = outputSheet.getRange("A1:D1");
        header.format.fill.color = "#1f4e79";
        header.format.font.color = "#ffffff";
        header.format.font.bold = true;
      } catch {}
      try {
        const table = outputSheet.tables.add(target, true);
        table.name = safeExcelName(outputSheetName + " Table");
        table.style = "TableStyleMedium4";
      } catch {}
      try { outputSheet.getUsedRange()?.format?.autofitColumns?.(); outputSheet.getUsedRange()?.format?.autofitRows?.(); } catch {}
      outputSheet.activate();
      await ctx.sync();
      const formulaCount = dependencyRows.slice(1).filter((row) => row[0] && row[0] !== "" && row[0] !== "Precedent Cell/Range" && row[0] !== "(none)").length;
      return toolResult(call, true, `Created Excel formula dependency map on ${outputSheetName} from ${sourceSheet.name || "source sheet"}!${source.address || requestedAddress || "used range"} with ${formulaCount} mapped row${formulaCount === 1 ? "" : "s"}.`);
    });
  }

  if (call.name === "excel_create_chart") {
    return globalThis.Excel.run(async (ctx: any) => {
      const sourceAddress = normalizeAddress(call.arguments.sourceAddress);
      if (!sourceAddress) throw new Error("excel_create_chart requires sourceAddress.");
      const sheetName = asString(call.arguments.sheetName, "").trim();
      const title = asString(call.arguments.title, "").trim();
      const sheet = activeOrNamedWorksheet(ctx, call.arguments.sheetName);
      const source = sheet.getRange(sourceAddress);
      source.load(["address", "rowCount", "columnCount"]);
      await ctx.sync();
      if (!source.rowCount || !source.columnCount) throw new Error(`Chart source range ${sourceAddress} is empty or invalid.`);

      const chart = sheet.charts.add(excelChartType(call.arguments.chartType), source, globalThis.Excel.ChartSeriesBy.auto);
      if (title) {
        chart.title.text = title;
        chart.title.visible = true;
      }
      const startCell = normalizeAddress(call.arguments.startCell) || "E2";
      const endCell = normalizeAddress(call.arguments.endCell) || "M20";
      try {
        chart.setPosition(sheet.getRange(startCell), sheet.getRange(endCell));
      } catch {
        chart.setPosition("E2", "M20");
      }

      const top = asOptionalNumber(call.arguments.top);
      const left = asOptionalNumber(call.arguments.left);
      const width = asOptionalNumber(call.arguments.width);
      const height = asOptionalNumber(call.arguments.height);
      if (top !== null) chart.top = top;
      if (left !== null) chart.left = left;
      if (width !== null) chart.width = width;
      if (height !== null) chart.height = height;
      await ctx.sync();
      return toolResult(call, true, `Created ${asString(call.arguments.chartType, "lineMarkers")} chart from ${sheetName || "active sheet"}!${source.address || sourceAddress}.`);
    });
  }

  if (call.name === "excel_create_summary_table") {
    return globalThis.Excel.run(async (ctx: any) => {
      const sourceAddress = normalizeAddress(call.arguments.sourceAddress);
      if (!sourceAddress) throw new Error("excel_create_summary_table requires sourceAddress.");
      const outputSheetName = asString(call.arguments.outputSheetName, "AI Summary").trim().slice(0, 31) || "AI Summary";
      const outputAddress = normalizeAddress(call.arguments.outputAddress) || "A1";
      const sheet = activeOrNamedWorksheet(ctx, call.arguments.sheetName);
      const source = sheet.getRange(sourceAddress);
      source.load(["values", "address", "rowCount", "columnCount"]);
      await ctx.sync();
      const summaryRows = createSummaryRows(source.values ?? [], call.arguments.groupBy, call.arguments.valueColumn, call.arguments.operation || "sum");
      let outputSheet: any;
      try {
        outputSheet = ctx.workbook.worksheets.getItem(outputSheetName);
        outputSheet.load("name");
        await ctx.sync();
      } catch {
        outputSheet = ctx.workbook.worksheets.add(outputSheetName);
        await ctx.sync();
      }
      const start = outputSheet.getRange(outputAddress);
      const width = Math.max(1, ...summaryRows.map((row) => row.length));
      const target = start.getResizedRange(summaryRows.length - 1, width - 1);
      target.values = summaryRows.map((row) => Array.from({ length: width }, (_, index) => row[index] ?? ""));
      target.format.autofitColumns();
      target.format.autofitRows();
      try {
        const table = outputSheet.tables.add(target, true);
        const tableName = safeExcelName(call.arguments.tableName || "AI_Summary_Table");
        table.name = tableName;
        table.style = asString(call.arguments.style, "TableStyleMedium4") || "TableStyleMedium4";
      } catch {}
      outputSheet.activate();
      await ctx.sync();
      return toolResult(call, true, `Created Excel summary table with ${summaryRows.length - 1} groups on ${outputSheetName}!${outputAddress}.`);
    });
  }

  if (call.name === "excel_create_pivot_chart_report") {
    return globalThis.Excel.run(async (ctx: any) => {
      const sourceAddress = normalizeAddress(call.arguments.sourceAddress);
      if (!sourceAddress) throw new Error("excel_create_pivot_chart_report requires sourceAddress.");
      const outputSheetName = asString(call.arguments.outputSheetName, "Pivot Chart Report").trim().slice(0, 31) || "Pivot Chart Report";
      const outputAddress = normalizeAddress(call.arguments.outputAddress) || "A1";
      const chartTitle = asString(call.arguments.title, "Pivot-style summary").trim() || "Pivot-style summary";
      const sheet = activeOrNamedWorksheet(ctx, call.arguments.sheetName);
      const source = sheet.getRange(sourceAddress);
      source.load(["values", "address", "rowCount", "columnCount"]);
      await ctx.sync();
      const summaryRows = createPivotChartReportRows(source.values ?? [], call.arguments.groupBy, call.arguments.valueColumn, call.arguments.operation || "sum", { top: call.arguments.top, sortBy: call.arguments.sortBy });
      let outputSheet: any;
      try {
        outputSheet = ctx.workbook.worksheets.getItem(outputSheetName);
        outputSheet.load("name");
        await ctx.sync();
      } catch {
        outputSheet = ctx.workbook.worksheets.add(outputSheetName);
        await ctx.sync();
      }
      const used = outputSheet.getUsedRangeOrNullObject?.() ?? outputSheet.getUsedRange?.();
      if (used) {
        try { used.load?.("isNullObject"); await ctx.sync(); if (!used.isNullObject) used.clear(globalThis.Excel.ClearApplyTo?.all ?? "All"); } catch { try { used.clear?.(); } catch {} }
      }
      const width = Math.max(1, ...summaryRows.map((row) => row.length));
      const normalized = summaryRows.map((row) => Array.from({ length: width }, (_, index) => row[index] ?? ""));
      const target = outputSheet.getRange(outputAddress).getResizedRange(normalized.length - 1, width - 1);
      target.values = normalized;
      target.format.autofitColumns();
      target.format.autofitRows();
      try {
        const table = outputSheet.tables.add(target, true);
        table.name = safeExcelName(call.arguments.tableName || outputSheetName + " Table");
        table.style = asString(call.arguments.style, "TableStyleMedium4") || "TableStyleMedium4";
      } catch {}
      let chartCreated = false;
      if (globalThis.Excel?.ChartType && outputSheet.charts?.add) {
        try {
          const chart = outputSheet.charts.add(excelChartType(call.arguments.chartType || "columnClustered"), target, globalThis.Excel.ChartSeriesBy.auto);
          chart.title.text = chartTitle;
          chart.title.visible = true;
          const startCell = normalizeAddress(call.arguments.chartStartCell) || "E2";
          const endCell = normalizeAddress(call.arguments.chartEndCell) || "M20";
          chart.setPosition(outputSheet.getRange(startCell), outputSheet.getRange(endCell));
          chartCreated = true;
        } catch {}
      }
      outputSheet.activate();
      await ctx.sync();
      return toolResult(call, true, `Created Excel pivot-style chart report on ${outputSheetName}!${outputAddress} with ${summaryRows.length - 1} grouped row${summaryRows.length === 2 ? "" : "s"}${chartCreated ? " and a chart" : "; chart APIs were unavailable, so the summary table was created without a live chart"} from ${source.address || sourceAddress}.`);
    });
  }

  if (call.name === "excel_create_pivottable") {
    return globalThis.Excel.run(async (ctx: any) => {
      const sourceAddress = normalizeAddress(call.arguments.sourceAddress);
      if (!sourceAddress) throw new Error("excel_create_pivottable requires sourceAddress.");
      const sourceSheet = activeOrNamedWorksheet(ctx, call.arguments.sheetName);
      const source = sourceSheet.getRange(sourceAddress);
      source.load(["values", "address", "rowCount", "columnCount"]);
      await ctx.sync();
      const rows = Array.isArray(source.values) ? source.values : [];
      if (rows.length < 2) throw new Error("PivotTable source range must include a header row and at least one data row.");
      const headers = rows[0].map((cell: unknown) => String(cell ?? "").trim()).map((header: string, index: number) => header || `Column ${index + 1}`);
      const destinationSheetName = asString(call.arguments.destinationSheetName, "Pivot Summary").trim().slice(0, 31) || "Pivot Summary";
      const destinationAddress = normalizeAddress(call.arguments.destinationAddress) || "A3";
      let destinationSheet: any;
      try {
        destinationSheet = ctx.workbook.worksheets.getItem(destinationSheetName);
        destinationSheet.load("name");
        await ctx.sync();
      } catch {
        destinationSheet = ctx.workbook.worksheets.add(destinationSheetName);
        await ctx.sync();
      }
      const destination = destinationSheet.getRange(destinationAddress);
      const pivotName = safeExcelName(call.arguments.name || "AI_PivotTable").slice(0, 255);
      const pivotTables = ctx.workbook.pivotTables ?? destinationSheet.pivotTables;
      if (!pivotTables?.add) throw new Error("This Excel runtime does not expose native PivotTable creation APIs. Use excel_create_summary_table or excel_generate_workbook_file as the fallback.");
      const pivot = pivotTables.add(pivotName, source, destination);
      const rowFields = Array.isArray(call.arguments.rows) ? call.arguments.rows : [];
      const columnFields = Array.isArray(call.arguments.columns) ? call.arguments.columns : [];
      const filterFields = Array.isArray(call.arguments.filters) ? call.arguments.filters : [];
      const valueFields = Array.isArray(call.arguments.values) ? call.arguments.values : [];
      const addHierarchy = (collection: any, field: unknown) => {
        const fieldName = excelPivotFieldName(field, headers);
        const hierarchy = pivot.hierarchies?.getItemOrNullObject ? pivot.hierarchies.getItemOrNullObject(fieldName) : pivot.hierarchies?.getItem?.(fieldName);
        if (!hierarchy) throw new Error(`PivotTable hierarchy ${fieldName} was not available.`);
        collection.add(hierarchy);
      };
      for (const field of rowFields) addHierarchy(pivot.rowHierarchies, field);
      for (const field of columnFields) addHierarchy(pivot.columnHierarchies, field);
      for (const field of filterFields) addHierarchy(pivot.filterHierarchies, field);
      for (const value of valueFields) {
        const valueSpec: Record<string, unknown> = value && typeof value === "object" ? value as Record<string, unknown> : { field: value };
        const fieldName = excelPivotFieldName(valueSpec.field, headers);
        const hierarchy = pivot.hierarchies?.getItemOrNullObject ? pivot.hierarchies.getItemOrNullObject(fieldName) : pivot.hierarchies?.getItem?.(fieldName);
        if (!hierarchy) throw new Error(`PivotTable value hierarchy ${fieldName} was not available.`);
        const dataHierarchy = pivot.dataHierarchies.add(hierarchy);
        if (valueSpec.name) try { dataHierarchy.name = asString(valueSpec.name, fieldName); } catch {}
        try { dataHierarchy.summarizeBy = excelPivotAggregation(valueSpec.summarizeBy); } catch {}
      }
      if (!rowFields.length && !columnFields.length && headers.length) addHierarchy(pivot.rowHierarchies, headers[0]);
      if (!valueFields.length && headers.length > 1) {
        const hierarchy = pivot.hierarchies?.getItemOrNullObject ? pivot.hierarchies.getItemOrNullObject(headers[1]) : pivot.hierarchies?.getItem?.(headers[1]);
        if (hierarchy) try { pivot.dataHierarchies.add(hierarchy).summarizeBy = excelPivotAggregation("sum"); } catch {}
      }
      const layout = excelPivotLayout(call.arguments.layout);
      if (layout && pivot.layout?.layoutType !== undefined) try { pivot.layout.layoutType = layout; } catch {}
      try { destinationSheet.getUsedRange()?.format?.autofitColumns?.(); } catch {}
      destinationSheet.activate();
      await ctx.sync();
      return toolResult(call, true, `Created Excel PivotTable ${pivotName} on ${destinationSheetName}!${destinationAddress} from ${source.address || sourceAddress}.`);
    });
  }

  if (call.name === "excel_create_table") {
    return globalThis.Excel.run(async (ctx: any) => {
      const sourceAddress = normalizeAddress(call.arguments.sourceAddress);
      if (!sourceAddress) throw new Error("excel_create_table requires sourceAddress.");
      const sheet = activeOrNamedWorksheet(ctx, call.arguments.sheetName);
      const table = sheet.tables.add(sheet.getRange(sourceAddress), boolValue(call.arguments.hasHeaders, true));
      const tableName = asString(call.arguments.name, "").trim();
      const style = asString(call.arguments.style, "TableStyleMedium2").trim();
      if (tableName) table.name = tableName;
      if (style) table.style = style;
      await ctx.sync();
      return toolResult(call, true, `Created Excel table${tableName ? ` ${tableName}` : ""} from ${sourceAddress}.`);
    });
  }

  if (call.name === "excel_format_range") {
    return globalThis.Excel.run(async (ctx: any) => {
      const address = normalizeAddress(call.arguments.address);
      if (!address) throw new Error("excel_format_range requires address.");
      const sheet = activeOrNamedWorksheet(ctx, call.arguments.sheetName);
      const range = sheet.getRange(address);
      const numberFormat = asString(call.arguments.numberFormat, "").trim();
      if (numberFormat) range.numberFormat = [[numberFormat]];
      const format = range.format;
      const font = format.font;
      const fillColor = asString(call.arguments.fillColor, "").trim();
      const fontColor = asString(call.arguments.fontColor, "").trim();
      const fontSize = asOptionalNumber(call.arguments.fontSize);
      const alignment = asString(call.arguments.horizontalAlignment, "").trim();
      if (typeof call.arguments.bold === "boolean") font.bold = call.arguments.bold;
      if (fillColor) format.fill.color = fillColor;
      if (fontColor) font.color = fontColor;
      if (fontSize !== null) font.size = fontSize;
      if (alignment) format.horizontalAlignment = alignment;
      if (typeof call.arguments.wrapText === "boolean") format.wrapText = call.arguments.wrapText;
      if (boolValue(call.arguments.autofit, true)) {
        format.autofitColumns();
        format.autofitRows();
      }
      await ctx.sync();
      return toolResult(call, true, `Formatted range ${address}.`);
    });
  }

  if (call.name === "excel_sort_range") {
    return globalThis.Excel.run(async (ctx: any) => {
      const address = normalizeAddress(call.arguments.address);
      const keyColumnIndex = asOptionalNumber(call.arguments.keyColumnIndex);
      if (!address || keyColumnIndex === null) throw new Error("excel_sort_range requires address and keyColumnIndex.");
      const sheet = activeOrNamedWorksheet(ctx, call.arguments.sheetName);
      const range = sheet.getRange(address);
      range.sort.apply([{ key: Math.trunc(keyColumnIndex), ascending: boolValue(call.arguments.ascending, true) }], false, boolValue(call.arguments.hasHeaders, true));
      await ctx.sync();
      return toolResult(call, true, `Sorted range ${address}.`);
    });
  }

  if (call.name === "excel_filter_range") {
    return globalThis.Excel.run(async (ctx: any) => {
      const sheet = activeOrNamedWorksheet(ctx, call.arguments.sheetName);
      const tableName = asString(call.arguments.tableName, "").trim();
      const address = normalizeAddress(call.arguments.address);
      const clear = boolValue(call.arguments.clear, false);
      const columnIndex = asOptionalNumber(call.arguments.columnIndex);
      const values = Array.isArray(call.arguments.values) ? call.arguments.values.map((value) => String(value)) : [];

      if (tableName) {
        const table = ctx.workbook.tables.getItem(tableName);
        if (clear) {
          table.clearFilters();
        } else {
          if (columnIndex === null || !values.length) throw new Error("excel_filter_range requires columnIndex and values unless clear is true.");
          table.columns.getItemAt(Math.trunc(columnIndex)).filter.applyValuesFilter(values);
        }
      } else {
        if (!address) throw new Error("excel_filter_range requires tableName or address.");
        const range = sheet.getRange(address);
        if (!sheet.autoFilter) throw new Error("This Excel runtime does not expose worksheet.autoFilter.");
        if (clear) sheet.autoFilter.clearCriteria();
        else {
          if (columnIndex === null || !values.length) throw new Error("excel_filter_range requires columnIndex and values unless clear is true.");
          sheet.autoFilter.apply(range, Math.trunc(columnIndex), { filterOn: globalThis.Excel.FilterOn.values, values });
        }
      }
      await ctx.sync();
      return toolResult(call, true, clear ? "Cleared Excel filters." : "Applied Excel filter.");
    });
  }

  if (call.name === "excel_freeze_panes") {
    return globalThis.Excel.run(async (ctx: any) => {
      const sheet = activeOrNamedWorksheet(ctx, call.arguments.sheetName);
      const panes = sheet.freezePanes;
      if (!panes) throw new Error("This Excel runtime does not expose freeze panes APIs.");
      if (boolValue(call.arguments.clear, false)) panes.unfreeze();
      else {
        const cell = normalizeAddress(call.arguments.cell);
        const rows = asOptionalNumber(call.arguments.rows);
        const columns = asOptionalNumber(call.arguments.columns);
        if (cell && panes.freezeAt) panes.freezeAt(sheet.getRange(cell));
        else if (rows !== null && columns !== null && panes.freezeAt) panes.freezeAt(sheet.getRangeByIndexes(Math.trunc(rows), Math.trunc(columns), 1, 1));
        else if (rows !== null && panes.freezeRows) panes.freezeRows(Math.trunc(rows));
        else if (columns !== null && panes.freezeColumns) panes.freezeColumns(Math.trunc(columns));
        else throw new Error("excel_freeze_panes requires rows, columns, or cell.");
      }
      await ctx.sync();
      return toolResult(call, true, boolValue(call.arguments.clear, false) ? "Unfroze Excel panes." : "Updated Excel freeze panes.");
    });
  }

  if (call.name === "excel_add_data_validation") {
    return globalThis.Excel.run(async (ctx: any) => {
      const address = normalizeAddress(call.arguments.address);
      const listValues = Array.isArray(call.arguments.listValues) ? call.arguments.listValues.map((value) => String(value)) : [];
      if (!address || !listValues.length) throw new Error("excel_add_data_validation requires address and listValues.");
      const sheet = activeOrNamedWorksheet(ctx, call.arguments.sheetName);
      const range = sheet.getRange(address);
      if (!range.dataValidation) throw new Error("This Excel runtime does not expose data validation APIs.");
      range.dataValidation.rule = { list: { inCellDropDown: true, source: listValues.join(",") } };
      const promptTitle = asString(call.arguments.promptTitle, "").trim();
      const promptMessage = asString(call.arguments.promptMessage, "").trim();
      const errorTitle = asString(call.arguments.errorTitle, "").trim();
      const errorMessage = asString(call.arguments.errorMessage, "").trim();
      if (promptTitle || promptMessage) range.dataValidation.prompt = { showPrompt: true, title: promptTitle, message: promptMessage };
      if (errorTitle || errorMessage) range.dataValidation.errorAlert = { showAlert: true, style: "Stop", title: errorTitle, message: errorMessage };
      await ctx.sync();
      return toolResult(call, true, `Added dropdown validation to ${address}.`);
    });
  }

  if (call.name === "excel_apply_conditional_format") {
    return globalThis.Excel.run(async (ctx: any) => {
      const address = normalizeAddress(call.arguments.address);
      const ruleType = asString(call.arguments.ruleType, "");
      if (!address || !ruleType) throw new Error("excel_apply_conditional_format requires address and ruleType.");
      const sheet = activeOrNamedWorksheet(ctx, call.arguments.sheetName);
      const range = sheet.getRange(address);
      if (!range.conditionalFormats?.add) throw new Error("This Excel runtime does not expose conditional formatting APIs.");
      const cfType = globalThis.Excel.ConditionalFormatType ?? {};
      if (ruleType === "colorScale") {
        range.conditionalFormats.add(cfType.colorScale ?? "ColorScale");
      } else {
        const format = range.conditionalFormats.add(cfType.cellValue ?? "CellValue");
        const operatorMap: Record<string, string> = { greaterThan: "GreaterThan", lessThan: "LessThan", equalTo: "EqualTo" };
        if (ruleType === "containsText") {
          format.textComparison.format.fill.color = asString(call.arguments.fillColor, "#fff2cc");
          format.textComparison.rule = { operator: "Contains", text: String(call.arguments.value ?? "") };
        } else {
          format.cellValue.rule = { formula1: String(call.arguments.value ?? ""), operator: operatorMap[ruleType] ?? "GreaterThan" };
          const fillColor = asString(call.arguments.fillColor, "").trim();
          const fontColor = asString(call.arguments.fontColor, "").trim();
          if (fillColor) format.cellValue.format.fill.color = fillColor;
          if (fontColor) format.cellValue.format.font.color = fontColor;
        }
      }
      await ctx.sync();
      return toolResult(call, true, `Applied conditional format to ${address}.`);
    });
  }

  if (call.name === "excel_insert_image") {
    return globalThis.Excel.run(async (ctx: any) => {
      const asset = await resolveImageAsset(call.arguments);
      const sheetName = asString(call.arguments.sheetName, "").trim();
      const sheet = activeOrNamedWorksheet(ctx, call.arguments.sheetName);
      if (!sheet.shapes?.addImage) throw new Error("This Excel runtime does not expose worksheet.shapes.addImage for native image insertion.");
      const image = sheet.shapes.addImage(asset.base64);
      const startCell = normalizeAddress(call.arguments.startCell);
      if (startCell) {
        const anchor = sheet.getRange(startCell);
        anchor.load(["left", "top"]);
        await ctx.sync();
        image.left = anchor.left;
        image.top = anchor.top;
      }
      applyShapeFrame(image, call.arguments, { left: 240, top: 40, width: 320, height: 180 });
      applyAltText(image, asString(call.arguments.altText, asset.name));
      await ctx.sync();
      return toolResult(call, true, `Inserted image ${asset.name} into ${sheetName || "active sheet"}.`);
    });
  }


  if (call.name === "excel_add_comment") {
    return globalThis.Excel.run(async (ctx: any) => {
      const address = normalizeAddress(call.arguments.address);
      const text = asString(call.arguments.text, "");
      if (!address || !text) throw new Error("excel_add_comment requires address and text.");
      const sheet = activeOrNamedWorksheet(ctx, call.arguments.sheetName);
      const range = sheet.getRange(address);
      const comments = ctx.workbook.comments ?? sheet.comments;
      if (!comments?.add) throw new Error("This Excel runtime does not expose comment insertion APIs.");
      comments.add(range, text);
      await ctx.sync();
      return toolResult(call, true, `Added Excel comment to ${address}.`);
    });
  }

  if (call.name === "excel_set_named_range") {
    return globalThis.Excel.run(async (ctx: any) => {
      const requestedName = asString(call.arguments.name, "").trim();
      const name = safeExcelName(requestedName);
      const address = normalizeAddress(call.arguments.address);
      if (!requestedName || !address) throw new Error("excel_set_named_range requires name and address.");
      const sheet = activeOrNamedWorksheet(ctx, call.arguments.sheetName);
      sheet.load("name");
      await ctx.sync();
      const reference = `='${String(sheet.name).replace(/'/g, "''")}'!${address}`;
      const names = ctx.workbook.names;
      if (!names?.add) throw new Error("This Excel runtime does not expose workbook named range APIs.");
      try { names.getItem(name).delete(); } catch {}
      names.add(name, reference);
      await ctx.sync();
      return toolResult(call, true, `Created Excel named range ${name} for ${reference}.`);
    });
  }

  if (call.name === "excel_protect_sheet") {
    return globalThis.Excel.run(async (ctx: any) => {
      const sheet = activeOrNamedWorksheet(ctx, call.arguments.sheetName);
      const protection = sheet.protection;
      if (!protection) throw new Error("This Excel runtime does not expose worksheet protection APIs.");
      const protect = boolValue(call.arguments.protect, true);
      if (protect) {
        const options = {
          allowFormatCells: boolValue(call.arguments.allowFormatCells, false),
          allowSort: boolValue(call.arguments.allowSort, false),
          allowAutoFilter: boolValue(call.arguments.allowAutoFilter, false),
        };
        protection.protect(options, asString(call.arguments.password, "") || undefined);
      } else {
        protection.unprotect(asString(call.arguments.password, "") || undefined);
      }
      await ctx.sync();
      return toolResult(call, true, protect ? "Protected Excel worksheet." : "Unprotected Excel worksheet.");
    });
  }

  if (call.name === "excel_set_page_layout") {
    return globalThis.Excel.run(async (ctx: any) => {
      const sheet = activeOrNamedWorksheet(ctx, call.arguments.sheetName);
      const layout = sheet.pageLayout;
      if (!layout) throw new Error("This Excel runtime does not expose worksheet page layout APIs.");
      const orientation = asString(call.arguments.orientation, "").trim();
      const paperSize = asString(call.arguments.paperSize, "").trim();
      const orientations = globalThis.Excel.PageOrientation ?? {};
      const paperSizes = globalThis.Excel.PaperType ?? {};
      if (orientation) layout.orientation = orientation === "landscape" ? (orientations.landscape ?? "Landscape") : (orientations.portrait ?? "Portrait");
      if (paperSize) {
        const map: Record<string, any> = { letter: paperSizes.letter ?? "Letter", legal: paperSizes.legal ?? "Legal", a4: paperSizes.a4 ?? "A4" };
        layout.paperSize = map[paperSize] ?? map.letter;
      }
      for (const [arg, prop] of [["topMargin", "topMargin"], ["bottomMargin", "bottomMargin"], ["leftMargin", "leftMargin"], ["rightMargin", "rightMargin"], ["fitToPagesWide", "fitToPagesWide"], ["fitToPagesTall", "fitToPagesTall"], ["scale", "zoom"]] as const) {
        const value = asOptionalNumber(call.arguments[arg]);
        if (value !== null) try { layout[prop] = value; } catch {}
      }
      const printArea = asString(call.arguments.printArea, "").trim();
      if (printArea || call.arguments.printArea === "") {
        const normalizedPrintArea = normalizeAddress(printArea);
        try { layout.printArea = normalizedPrintArea || ""; } catch {}
      }
      const repeatRows = asString(call.arguments.repeatRows, "").trim();
      const repeatColumns = asString(call.arguments.repeatColumns, "").trim();
      if (repeatRows) try { layout.setPrintTitleRows?.(repeatRows.replace(/\$/g, "")); } catch { try { layout.printTitleRows = repeatRows; } catch {} }
      if (repeatColumns) try { layout.setPrintTitleColumns?.(repeatColumns.replace(/\$/g, "")); } catch { try { layout.printTitleColumns = repeatColumns; } catch {} }
      for (const [arg, prop] of [["showGridlines", "printGridlines"], ["showHeadings", "printHeadings"], ["centerHorizontally", "centerHorizontally"], ["centerVertically", "centerVertically"], ["blackAndWhite", "blackAndWhite"], ["draftMode", "draftMode"]] as const) {
        if (typeof call.arguments[arg] === "boolean") try { layout[prop] = call.arguments[arg]; } catch {}
      }
      await ctx.sync();
      return toolResult(call, true, "Updated Excel worksheet page layout and print setup.");
    });
  }
  return toolResult(call, false, `Unsupported Excel tool: ${call.name}.`);
}

async function executeWordTool(call: ToolCallRequest): Promise<ToolCallResult> {
  if (call.name === "word_audit_document") {
    return globalThis.Word.run(async (ctx: any) => {
      const body = ctx.document.body;
      body.load("text");
      let tables: any = null;
      let inlinePictures: any = null;
      try { tables = body.tables; tables.load?.("items"); } catch {}
      try { inlinePictures = body.inlinePictures; inlinePictures.load?.("items"); } catch {}
      await ctx.sync();
      const rows = createWordAuditRows(body.text || "", {
        tableCount: Array.isArray(tables?.items) ? tables.items.length : undefined,
        imageCount: Array.isArray(inlinePictures?.items) ? inlinePictures.items.length : undefined,
      });
      const includeHeading = boolValue(call.arguments.includeHeading, true);
      const style = asString(call.arguments.style, "Grid Table 4 - Accent 1").trim();
      if (includeHeading) {
        const heading = body.insertText("\nCTRL Word Document Audit\n", globalThis.Word.InsertLocation.end);
        try { heading.styleBuiltIn = globalThis.Word.Style.heading1 ?? "Heading 1"; } catch { try { heading.style = "Heading 1"; } catch {} }
      }
      if (!body.insertTable) throw new Error("This Word runtime does not expose insertTable for native audit table insertion.");
      const table = body.insertTable(rows.length, rows[0].length, globalThis.Word.InsertLocation.end, rows);
      if (style) table.style = style;
      await ctx.sync();
      const findingCount = Math.max(0, rows.length - 1);
      return toolResult(call, true, `Inserted Word document audit table with ${findingCount} finding${findingCount === 1 ? "" : "s"}.`);
    });
  }

  if (call.name === "word_create_context_brief") {
    return globalThis.Word.run(async (ctx: any) => {
      const brief = createWordContextBrief(call.arguments.text, {
        title: call.arguments.title,
        sourceLabel: call.arguments.sourceLabel,
        audience: call.arguments.audience,
        includeSourceTable: boolValue(call.arguments.includeSourceTable, true),
        maxPoints: call.arguments.maxPoints,
      });
      const body = ctx.document.body;
      const location = asString(call.arguments.location, "end");
      const insertLocation = location === "start" ? globalThis.Word.InsertLocation.start : location === "replace" ? globalThis.Word.InsertLocation.replace : globalThis.Word.InsertLocation.end;
      const target = location === "replace" ? ctx.document.getSelection() : body;
      const heading = target.insertText(`${brief.title}\n`, insertLocation);
      try { heading.styleBuiltIn = globalThis.Word.Style.heading1 ?? "Heading 1"; } catch { try { heading.style = "Heading 1"; } catch {} }
      const source = body.insertText(`Source: ${brief.sourceLabel}\nAudience: ${brief.audience}\n\n`, globalThis.Word.InsertLocation.end);
      try { source.style = "Intense Quote"; } catch {}
      const summaryHeading = body.insertText("Executive summary\n", globalThis.Word.InsertLocation.end);
      try { summaryHeading.styleBuiltIn = globalThis.Word.Style.heading2 ?? "Heading 2"; } catch { try { summaryHeading.style = "Heading 2"; } catch {} }
      body.insertText(`${brief.summary}\n\n`, globalThis.Word.InsertLocation.end);
      const pointsHeading = body.insertText("Key points\n", globalThis.Word.InsertLocation.end);
      try { pointsHeading.styleBuiltIn = globalThis.Word.Style.heading2 ?? "Heading 2"; } catch { try { pointsHeading.style = "Heading 2"; } catch {} }
      const pointRows = [["#", "Point"], ...brief.keyPoints.map((point, index) => [String(index + 1), point])];
      if (body.insertTable) {
        const pointTable = body.insertTable(pointRows.length, 2, globalThis.Word.InsertLocation.end, pointRows);
        try { pointTable.style = asString(call.arguments.style, "Grid Table 4 - Accent 1") || "Grid Table 4 - Accent 1"; } catch {}
      } else {
        body.insertText(pointRows.map((row) => row.join("\t")).join("\n") + "\n", globalThis.Word.InsertLocation.end);
      }
      if (brief.tableRows && brief.tableRows.length >= 2) {
        const evidenceHeading = body.insertText("\nExtracted source table\n", globalThis.Word.InsertLocation.end);
        try { evidenceHeading.styleBuiltIn = globalThis.Word.Style.heading2 ?? "Heading 2"; } catch { try { evidenceHeading.style = "Heading 2"; } catch {} }
        const columnCount = Math.max(1, ...brief.tableRows.map((row) => row.length));
        const tableRows = brief.tableRows.slice(0, Math.max(2, Math.min(50, asOptionalNumber(call.arguments.maxTableRows) ?? 20))).map((row) => Array.from({ length: columnCount }, (_, index) => row[index] ?? ""));
        if (body.insertTable) {
          const evidenceTable = body.insertTable(tableRows.length, columnCount, globalThis.Word.InsertLocation.end, tableRows);
          try { evidenceTable.style = asString(call.arguments.style, "Grid Table 4 - Accent 1") || "Grid Table 4 - Accent 1"; } catch {}
        } else {
          body.insertText(tableRows.map((row) => row.join("\t")).join("\n") + "\n", globalThis.Word.InsertLocation.end);
        }
      }
      await ctx.sync();
      return toolResult(call, true, `Created Word context brief from ${brief.sourceLabel} with ${brief.keyPoints.length} key point${brief.keyPoints.length === 1 ? "" : "s"}${brief.tableRows ? " and an extracted source table" : ""}.`);
    });
  }

  if (call.name === "word_insert_text") {
    return globalThis.Word.run(async (ctx: any) => {
      const text = asString(call.arguments.text, "");
      const location = asString(call.arguments.location, "replace");
      const insertLocation = location === "end" ? globalThis.Word.InsertLocation.end : location === "start" ? globalThis.Word.InsertLocation.start : globalThis.Word.InsertLocation.replace;
      const target = location === "end" || location === "start" ? ctx.document.body : ctx.document.getSelection();
      target.insertText(text, insertLocation);
      await ctx.sync();
      return toolResult(call, true, `Inserted ${text.length} characters into Word.`);
    });
  }

  if (call.name === "word_insert_heading") {
    return globalThis.Word.run(async (ctx: any) => {
      const headingText = asString(call.arguments.text, "");
      if (!headingText) throw new Error("word_insert_heading requires text.");
      const location = asString(call.arguments.location, "replace");
      const insertLocation = location === "end" ? globalThis.Word.InsertLocation.end : location === "start" ? globalThis.Word.InsertLocation.start : globalThis.Word.InsertLocation.replace;
      const target = location === "end" || location === "start" ? ctx.document.body : ctx.document.getSelection();
      const range = target.insertText(headingText, insertLocation);
      const level = Math.min(6, Math.max(1, Math.trunc(asOptionalNumber(call.arguments.level) ?? 1)));
      try { range.styleBuiltIn = globalThis.Word.Style[`heading${level}`] ?? `Heading ${level}`; } catch { range.style = `Heading ${level}`; }
      await ctx.sync();
      return toolResult(call, true, `Inserted Heading ${level} into Word.`);
    });
  }

  if (call.name === "word_insert_table") {
    return globalThis.Word.run(async (ctx: any) => {
      const rows = asRows(call.arguments.values);
      if (!rows.length) throw new Error("word_insert_table requires non-empty values.");
      const columnCount = Math.max(1, ...rows.map((row) => row.length));
      const normalized = rows.map((row) => Array.from({ length: columnCount }, (_, index) => row[index] ?? ""));
      const location = asString(call.arguments.location, "replace");
      const insertLocation = location === "end" ? globalThis.Word.InsertLocation.end : location === "start" ? globalThis.Word.InsertLocation.start : globalThis.Word.InsertLocation.replace;
      const target = location === "end" || location === "start" ? ctx.document.body : ctx.document.getSelection();
      if (!target.insertTable) throw new Error("This Word runtime does not expose insertTable for native table insertion.");
      const table = target.insertTable(normalized.length, columnCount, insertLocation, normalized);
      const style = asString(call.arguments.style, "").trim();
      if (style) table.style = style;
      await ctx.sync();
      return toolResult(call, true, `Inserted a ${normalized.length} by ${columnCount} Word table.`);
    });
  }

  if (call.name === "word_apply_style") {
    return globalThis.Word.run(async (ctx: any) => {
      const style = asString(call.arguments.style, "").trim();
      if (!style) throw new Error("word_apply_style requires style.");
      const target = asString(call.arguments.target, "selection") === "body" ? ctx.document.body : ctx.document.getSelection();
      target.style = style;
      await ctx.sync();
      return toolResult(call, true, `Applied Word style ${style}.`);
    });
  }

  if (call.name === "word_insert_page_break") {
    return globalThis.Word.run(async (ctx: any) => {
      const target = wordTarget(ctx, call.arguments.location);
      target.insertBreak(globalThis.Word.BreakType.page, wordInsertLocation(call.arguments.location));
      await ctx.sync();
      return toolResult(call, true, "Inserted Word page break.");
    });
  }

  if (call.name === "word_insert_section_break") {
    return globalThis.Word.run(async (ctx: any) => {
      const target = wordTarget(ctx, call.arguments.location);
      const requested = asString(call.arguments.breakType, "nextPage");
      const breakTypes = globalThis.Word.BreakType ?? {};
      const sectionBreak = requested === "continuous"
        ? (breakTypes.sectionContinuous ?? breakTypes.sectionBreakContinuous ?? "SectionContinuous")
        : (breakTypes.sectionNext ?? breakTypes.sectionNextPage ?? "SectionNext");
      if (!target.insertBreak) throw new Error("This Word runtime does not expose break insertion APIs.");
      target.insertBreak(sectionBreak, wordInsertLocation(call.arguments.location));
      await ctx.sync();
      return toolResult(call, true, `Inserted Word ${requested} section break.`);
    });
  }

  if (call.name === "word_set_header_footer") {
    return globalThis.Word.run(async (ctx: any) => {
      const header = asString(call.arguments.header, "");
      const footer = asString(call.arguments.footer, "");
      if (!header && !footer) throw new Error("word_set_header_footer requires header or footer text.");
      const sections = ctx.document.sections;
      if (!sections?.getFirst) throw new Error("This Word runtime does not expose section header/footer APIs.");
      const section = sections.getFirst();
      const headerFooterType = globalThis.Word.HeaderFooterType?.primary ?? "Primary";
      if (header) {
        const headerBody = section.getHeader(headerFooterType);
        headerBody.insertText(header, globalThis.Word.InsertLocation.replace);
      }
      if (footer) {
        const footerBody = section.getFooter(headerFooterType);
        footerBody.insertText(footer, globalThis.Word.InsertLocation.replace);
      }
      await ctx.sync();
      return toolResult(call, true, `Updated Word ${header && footer ? "header and footer" : header ? "header" : "footer"}.`);
    });
  }

  if (call.name === "word_find_replace") {
    return globalThis.Word.run(async (ctx: any) => {
      const find = asString(call.arguments.find, "");
      const replace = asString(call.arguments.replace, "");
      if (!find) throw new Error("word_find_replace requires find text.");
      const results = ctx.document.body.search(find, { matchCase: boolValue(call.arguments.matchCase, false), matchWholeWord: boolValue(call.arguments.matchWholeWord, false) });
      results.load("items");
      await ctx.sync();
      for (const item of results.items) item.insertText(replace, globalThis.Word.InsertLocation.replace);
      await ctx.sync();
      return toolResult(call, true, `Replaced ${results.items.length} Word occurrence${results.items.length === 1 ? "" : "s"}.`);
    });
  }

  if (call.name === "word_insert_comment") {
    return globalThis.Word.run(async (ctx: any) => {
      const text = asString(call.arguments.text, "");
      if (!text) throw new Error("word_insert_comment requires text.");
      const selection = ctx.document.getSelection();
      if (!selection.insertComment) throw new Error("This Word runtime does not expose comment insertion APIs.");
      selection.insertComment(text);
      await ctx.sync();
      return toolResult(call, true, "Inserted Word comment.");
    });
  }

  if (call.name === "word_insert_image") {
    return globalThis.Word.run(async (ctx: any) => {
      const asset = await resolveImageAsset(call.arguments);
      const location = asString(call.arguments.location, "replace");
      const insertLocation = location === "end" ? globalThis.Word.InsertLocation.end : location === "start" ? globalThis.Word.InsertLocation.start : globalThis.Word.InsertLocation.replace;
      const target = location === "end" || location === "start" ? ctx.document.body : ctx.document.getSelection();
      if (!target.insertInlinePictureFromBase64) throw new Error("This Word runtime does not expose insertInlinePictureFromBase64 for native image insertion.");
      const picture = target.insertInlinePictureFromBase64(asset.base64, insertLocation);
      const width = asOptionalNumber(call.arguments.width);
      const height = asOptionalNumber(call.arguments.height);
      if (width !== null) picture.width = width;
      if (height !== null) picture.height = height;
      applyAltText(picture, asString(call.arguments.altText, asset.name));
      await ctx.sync();
      return toolResult(call, true, `Inserted image ${asset.name} into Word.`);
    });
  }


  if (call.name === "word_insert_hyperlink") {
    return globalThis.Word.run(async (ctx: any) => {
      const text = asString(call.arguments.text, "").trim();
      const url = asString(call.arguments.url, "").trim();
      if (!text || !url) throw new Error("word_insert_hyperlink requires text and url.");
      const location = asString(call.arguments.location, "replace");
      const insertLocation = location === "end" ? globalThis.Word.InsertLocation.end : location === "start" ? globalThis.Word.InsertLocation.start : globalThis.Word.InsertLocation.replace;
      const target = location === "end" || location === "start" ? ctx.document.body : ctx.document.getSelection();
      if (!target.insertHtml) throw new Error("This Word runtime does not expose HTML insertion APIs needed for hyperlinks.");
      target.insertHtml(`<a href="${url.replace(/"/g, "&quot;")}">${text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")}</a>`, insertLocation);
      await ctx.sync();
      return toolResult(call, true, `Inserted Word hyperlink to ${url}.`);
    });
  }

  if (call.name === "word_format_selection") {
    return globalThis.Word.run(async (ctx: any) => {
      const target = asString(call.arguments.target, "selection") === "body" ? ctx.document.body : ctx.document.getSelection();
      const font = target.font;
      if (!font) throw new Error("This Word runtime does not expose selection/body font formatting APIs.");
      const fontSize = asOptionalNumber(call.arguments.fontSize);
      const fontColor = asString(call.arguments.fontColor, "").trim();
      const highlightColor = asString(call.arguments.highlightColor, "").trim();
      if (typeof call.arguments.bold === "boolean") font.bold = call.arguments.bold;
      if (typeof call.arguments.italic === "boolean") font.italic = call.arguments.italic;
      if (typeof call.arguments.underline === "boolean") font.underline = call.arguments.underline ? (globalThis.Word.UnderlineType?.single ?? "Single") : (globalThis.Word.UnderlineType?.none ?? "None");
      if (fontSize !== null) font.size = fontSize;
      if (fontColor) font.color = fontColor;
      if (highlightColor) font.highlightColor = highlightColor;
      await ctx.sync();
      return toolResult(call, true, "Formatted Word selection.");
    });
  }
  if (call.name === "word_insert_content_control") {
    return globalThis.Word.run(async (ctx: any) => {
      const title = asString(call.arguments.title, "").trim();
      if (!title) throw new Error("word_insert_content_control requires title.");
      const tag = asString(call.arguments.tag, title).trim();
      const placeholderText = asString(call.arguments.placeholderText, title).trim();
      const requestedType = compactKey(call.arguments.type || "richText");
      const target = wordTarget(ctx, call.arguments.location || "replace");
      const insertLocation = wordInsertLocation(call.arguments.location || "replace");
      let contentControl: any = null;

      if (requestedType === "checkbox" && typeof target.insertCheckBoxContentControl === "function") {
        contentControl = target.insertCheckBoxContentControl(insertLocation);
      } else if (["dropdown", "dropdownlist", "combo"].includes(requestedType) && typeof target.insertDropdownListContentControl === "function") {
        contentControl = target.insertDropdownListContentControl(insertLocation);
      } else if (["date", "datepicker"].includes(requestedType) && typeof target.insertDatePickerContentControl === "function") {
        contentControl = target.insertDatePickerContentControl(insertLocation);
      } else if (["plaintext", "text"].includes(requestedType) && typeof target.insertPlainTextContentControl === "function") {
        contentControl = target.insertPlainTextContentControl(insertLocation);
      } else if (typeof target.insertContentControl === "function") {
        contentControl = target.insertContentControl();
      } else {
        throw new Error("This Word runtime does not expose native content-control insertion APIs. Use word_generate_document_file for template-style DOCX generation, or try a newer Word build with content-control APIs enabled.");
      }

      try { contentControl.title = title; } catch {}
      try { contentControl.tag = tag; } catch {}
      if (typeof call.arguments.cannotDelete === "boolean") try { contentControl.cannotDelete = call.arguments.cannotDelete; } catch {}
      if (typeof call.arguments.cannotEdit === "boolean") try { contentControl.cannotEdit = call.arguments.cannotEdit; } catch {}
      if (placeholderText && typeof contentControl.insertText === "function") {
        try { contentControl.insertText(placeholderText, globalThis.Word.InsertLocation.replace); } catch {}
      } else if (placeholderText && contentControl.appearance !== undefined) {
        try { contentControl.placeholderText = placeholderText; } catch {}
      }
      const options = Array.isArray(call.arguments.options) ? call.arguments.options.map((item) => String(item)).filter(Boolean) : [];
      if (options.length) {
        for (const option of options.slice(0, 50)) {
          try { contentControl.addItem?.(option, option); } catch { try { contentControl.items?.add?.(option, option); } catch {} }
        }
      }
      await ctx.sync();
      return toolResult(call, true, "Inserted Word content control " + title + ".");
    });
  }
  return toolResult(call, false, `Unsupported Word tool: ${call.name}.`);
}

async function newestPowerPointSlide(ctx: any) {
  const slides = ctx.presentation.slides;
  slides.load("items");
  await ctx.sync();
  return slides.items[slides.items.length - 1] ?? null;
}

async function powerPointSlideAtOrCreate(ctx: any, argsOrIndex: Record<string, unknown> | number | null | undefined) {
  const slides = ctx.presentation.slides;
  slides.load("items");
  await ctx.sync();
  if (!slides.items.length) {
    slides.add();
    return newestPowerPointSlide(ctx);
  }
  const index = normalizeOfficeSlideIndex(argsOrIndex, slides.items.length, 0);
  if (slides.items[index]) return slides.items[index];
  return slides.items[0];
}

function normalizeShapeIndexes(value: unknown, count: number) {
  const raw = Array.isArray(value) ? value : [];
  const indexes = raw
    .map((item) => typeof item === "number" && Number.isFinite(item) ? Math.trunc(item) - 1 : null)
    .filter((item): item is number => item !== null && item >= 0 && item < count);
  return indexes.length ? [...new Set(indexes)] : Array.from({ length: count }, (_, shapeIndex) => shapeIndex);
}

export function arrangeShapeFrames(shapes: Array<{ left: number; top: number; width: number; height: number }>, actionValue: unknown) {
  const action = compactKey(actionValue || "alignleft");
  if (shapes.length < 1) return 0;
  const left = Math.min(...shapes.map((shape) => shape.left));
  const top = Math.min(...shapes.map((shape) => shape.top));
  const right = Math.max(...shapes.map((shape) => shape.left + shape.width));
  const bottom = Math.max(...shapes.map((shape) => shape.top + shape.height));
  const centerX = (left + right) / 2;
  const centerY = (top + bottom) / 2;
  if (["alignleft", "left"].includes(action)) shapes.forEach((shape) => { shape.left = left; });
  else if (["aligncenter", "center", "horizontalcenter"].includes(action)) shapes.forEach((shape) => { shape.left = centerX - shape.width / 2; });
  else if (["alignright", "right"].includes(action)) shapes.forEach((shape) => { shape.left = right - shape.width; });
  else if (["aligntop", "top"].includes(action)) shapes.forEach((shape) => { shape.top = top; });
  else if (["alignmiddle", "middle", "verticalcenter"].includes(action)) shapes.forEach((shape) => { shape.top = centerY - shape.height / 2; });
  else if (["alignbottom", "bottom"].includes(action)) shapes.forEach((shape) => { shape.top = bottom - shape.height; });
  else if (["distributehorizontal", "distributedhorizontally", "horizontal"].includes(action) && shapes.length > 2) {
    const sorted = [...shapes].sort((a, b) => a.left - b.left);
    const totalWidth = sorted.reduce((sum, shape) => sum + shape.width, 0);
    const gap = (right - left - totalWidth) / (sorted.length - 1);
    let cursor = left;
    sorted.forEach((shape) => { shape.left = cursor; cursor += shape.width + gap; });
  } else if (["distributevertical", "distributedvertically", "vertical"].includes(action) && shapes.length > 2) {
    const sorted = [...shapes].sort((a, b) => a.top - b.top);
    const totalHeight = sorted.reduce((sum, shape) => sum + shape.height, 0);
    const gap = (bottom - top - totalHeight) / (sorted.length - 1);
    let cursor = top;
    sorted.forEach((shape) => { shape.top = cursor; cursor += shape.height + gap; });
  } else {
    return 0;
  }
  return shapes.length;
}

async function executePowerPointTool(call: ToolCallRequest): Promise<ToolCallResult> {
  if (call.name === "powerpoint_audit_deck") {
    return globalThis.PowerPoint.run(async (ctx: any) => {
      const slides = ctx.presentation.slides;
      slides.load("items");
      await ctx.sync();
      const summaries: PowerPointAuditSlide[] = [];
      for (let index = 0; index < slides.items.length; index += 1) {
        const slide = slides.items[index];
        let shapeCount = 0;
        let text = "";
        try {
          slide.shapes.load("items");
          await ctx.sync();
          shapeCount = slide.shapes.items.length;
          const snippets: string[] = [];
          for (const shape of slide.shapes.items.slice(0, 40)) {
            try { shape.textFrame?.textRange?.load?.("text"); } catch {}
          }
          await ctx.sync();
          for (const shape of slide.shapes.items.slice(0, 40)) {
            const shapeText = asString(shape.textFrame?.textRange?.text, "").trim();
            if (shapeText) snippets.push(shapeText);
          }
          text = snippets.join("\n");
        } catch {}
        summaries.push({ slideNumber: index + 1, shapeCount, text });
      }
      const rows = createPowerPointAuditRows(summaries, { audience: call.arguments.audience });
      slides.add();
      const auditSlide = await newestPowerPointSlide(ctx);
      if (!auditSlide?.shapes?.addTextBox) throw new Error("This PowerPoint runtime does not expose text boxes needed to create a live deck audit slide.");
      const title = auditSlide.shapes.addTextBox(asString(call.arguments.title, "CTRL PowerPoint Deck Audit") || "CTRL PowerPoint Deck Audit");
      applyShapeFrame(title, { left: 40, top: 28, width: 640, height: 42 }, { left: 40, top: 28, width: 640, height: 42 });
      styleTextRange(title, { fontSize: 24, bold: true, fontColor: "#1f2937" });
      const subtitle = auditSlide.shapes.addTextBox(`Reviewed ${summaries.length} slide${summaries.length === 1 ? "" : "s"}. Findings are based on live Office.js-visible slide objects and text.`);
      applyShapeFrame(subtitle, { left: 42, top: 72, width: 635, height: 34 }, { left: 42, top: 72, width: 635, height: 34 });
      styleTextRange(subtitle, { fontSize: 11, fontColor: "#475569" });
      const maxRows = Math.min(rows.length, 8);
      const tableRows = rows.slice(0, maxRows);
      const columnWidths = [118, 74, 218, 250];
      const rowHeight = 34;
      const left = 36;
      const top = 122;
      tableRows.forEach((row, rowIndex) => {
        let cursor = left;
        row.forEach((cell, columnIndex) => {
          const box = auditSlide.shapes.addTextBox(String(cell ?? ""));
          applyShapeFrame(box, { left: cursor, top: top + rowIndex * rowHeight, width: columnWidths[columnIndex] ?? 140, height: rowHeight }, { left: cursor, top: top + rowIndex * rowHeight, width: columnWidths[columnIndex] ?? 140, height: rowHeight });
          styleTextRange(box, { fontSize: rowIndex === 0 ? 10 : 8.5, bold: rowIndex === 0, fontColor: rowIndex === 0 ? "#ffffff" : "#111827" });
          try { box.fill.setSolidColor(rowIndex === 0 ? "#1f4e79" : columnIndex === 1 && /high/i.test(String(cell)) ? "#fee2e2" : columnIndex === 1 && /medium/i.test(String(cell)) ? "#fef3c7" : "#f8fafc"); } catch {}
          cursor += columnWidths[columnIndex] ?? 140;
        });
      });
      if (rows.length > maxRows) {
        const more = auditSlide.shapes.addTextBox(`+ ${rows.length - maxRows} more finding${rows.length - maxRows === 1 ? "" : "s"}. Ask CTRL for a generated audit artifact or uploaded-PPTX review for full detail.`);
        applyShapeFrame(more, { left: 42, top: top + maxRows * rowHeight + 14, width: 620, height: 28 }, { left: 42, top: top + maxRows * rowHeight + 14, width: 620, height: 28 });
        styleTextRange(more, { fontSize: 10, fontColor: "#475569" });
      }
      await ctx.sync();
      const findingCount = Math.max(0, rows.length - 1);
      return toolResult(call, true, `Created PowerPoint deck audit slide with ${findingCount} finding${findingCount === 1 ? "" : "s"} from ${summaries.length} slide${summaries.length === 1 ? "" : "s"}.`);
    });
  }

  if (call.name === "powerpoint_create_slides") {
    return globalThis.PowerPoint.run(async (ctx: any) => {
      const slidesArg = Array.isArray(call.arguments.slides) ? call.arguments.slides : [];
      if (!slidesArg.length) throw new Error("powerpoint_create_slides requires at least one slide.");
      for (const item of slidesArg.slice(0, 20)) {
        ctx.presentation.slides.add();
        const slide = await newestPowerPointSlide(ctx);
        if (!slide?.shapes?.addTextBox) throw new Error("This PowerPoint runtime can create slides, but its shape/text-box API is unavailable. Try updating Office or use Insert last as plain text.");
        const title = typeof item === "object" && item ? asString((item as any).title, "") : "";
        const body = typeof item === "object" && item ? asString((item as any).body, String(item ?? "")) : String(item ?? "");
        const titleBox = slide.shapes.addTextBox(title || "Untitled");
        titleBox.left = 36;
        titleBox.top = 30;
        titleBox.width = 612;
        titleBox.height = 48;
        const bodyBox = slide.shapes.addTextBox(body);
        bodyBox.left = 36;
        bodyBox.top = 96;
        bodyBox.width = 612;
        bodyBox.height = 300;
      }
      await ctx.sync();
      return toolResult(call, true, `Created ${Math.min(slidesArg.length, 20)} PowerPoint slides.`);
    });
  }

  if (call.name === "powerpoint_add_textbox") {
    return globalThis.PowerPoint.run(async (ctx: any) => {
      const text = asString(call.arguments.text, "");
      if (!text) throw new Error("powerpoint_add_textbox requires text.");
      const slide = await powerPointSlideAtOrCreate(ctx, call.arguments);
      if (!slide?.shapes?.addTextBox) throw new Error("This PowerPoint runtime does not expose addTextBox.");
      const shape = slide.shapes.addTextBox(text);
      applyShapeFrame(shape, call.arguments, { left: 72, top: 72, width: 576, height: 120 });
      styleTextRange(shape, call.arguments);
      await ctx.sync();
      return toolResult(call, true, "Added PowerPoint text box.");
    });
  }

  if (call.name === "powerpoint_add_table") {
    return globalThis.PowerPoint.run(async (ctx: any) => {
      const rows = asRows(call.arguments.values);
      if (!rows.length) throw new Error("powerpoint_add_table requires non-empty values.");
      const columnCount = Math.max(1, ...rows.map((row) => row.length));
      const normalized = rows.map((row) => Array.from({ length: columnCount }, (_, index) => row[index] ?? ""));
      const slide = await powerPointSlideAtOrCreate(ctx, call.arguments);
      const left = asOptionalNumber(call.arguments.left) ?? 48;
      const top = asOptionalNumber(call.arguments.top) ?? 120;
      const width = asOptionalNumber(call.arguments.width) ?? 624;
      const height = asOptionalNumber(call.arguments.height) ?? Math.max(120, normalized.length * 34);
      const fontSize = asOptionalNumber(call.arguments.fontSize) ?? 12;
      const cellWidth = width / columnCount;
      const cellHeight = height / normalized.length;
      const headerFillColor = asString(call.arguments.headerFillColor, "#1f4e79");
      if (slide?.shapes?.addTable) {
        try {
          const table = slide.shapes.addTable(normalized.length, columnCount, { left, top, width, height });
          try {
            normalized.forEach((row, rowIndex) => row.forEach((cell, columnIndex) => { table.getCell(rowIndex, columnIndex).text = cell; }));
          } catch {}
          await ctx.sync();
          return toolResult(call, true, `Added native PowerPoint table with ${normalized.length} rows.`);
        } catch {
          // Some PowerPoint builds expose addTable but reject its argument shape. Fall through to the reliable text-box grid.
        }
      }

      if (!slide?.shapes?.addTextBox) throw new Error("This PowerPoint runtime does not expose text boxes needed for table fallback.");
      normalized.forEach((row, rowIndex) => {
        row.forEach((cell, columnIndex) => {
          const box = slide.shapes.addTextBox(cell);
          box.left = left + columnIndex * cellWidth;
          box.top = top + rowIndex * cellHeight;
          box.width = cellWidth;
          box.height = cellHeight;
          try { box.textFrame.textRange.font.size = fontSize; } catch {}
          if (rowIndex === 0) {
            try { box.fill.setSolidColor(headerFillColor); } catch {}
            try { box.textFrame.textRange.font.color = "#ffffff"; } catch {}
            try { box.textFrame.textRange.font.bold = true; } catch {}
          }
        });
      });
      await ctx.sync();
      return toolResult(call, true, `Added PowerPoint table grid with ${normalized.length} rows and ${columnCount} columns.`);
    });
  }

  if (call.name === "powerpoint_add_shape") {
    return globalThis.PowerPoint.run(async (ctx: any) => {
      const slide = await powerPointSlideAtOrCreate(ctx, call.arguments);
      if (!slide?.shapes?.addGeometricShape && !slide?.shapes?.addShape) throw new Error("This PowerPoint runtime does not expose shape insertion APIs.");
      const addShape = slide.shapes.addGeometricShape ?? slide.shapes.addShape;
      const shape = addShape.call(slide.shapes, normalizeShapeType(call.arguments.shapeType));
      applyShapeFrame(shape, call.arguments, { left: 96, top: 140, width: 180, height: 90 });
      const fillColor = asString(call.arguments.fillColor, "").trim();
      const lineColor = asString(call.arguments.lineColor, "").trim();
      const text = asString(call.arguments.text, "").trim();
      try { if (fillColor) shape.fill.setSolidColor(fillColor); } catch {}
      try { if (lineColor) shape.lineFormat.color = lineColor; } catch {}
      try { if (text && shape.textFrame?.textRange) shape.textFrame.textRange.text = text; } catch {}
      await ctx.sync();
      return toolResult(call, true, "Added PowerPoint shape.");
    });
  }

  if (call.name === "powerpoint_create_diagram") {
    return globalThis.PowerPoint.run(async (ctx: any) => {
      const frames = createPowerPointDiagramFrames(call.arguments.steps, call.arguments.diagramType, call.arguments);
      const cardCount = frames.filter((frame) => frame.kind === "card").length;
      if (!cardCount) throw new Error("powerpoint_create_diagram requires at least one step with a label.");
      const slide = await powerPointSlideAtOrCreate(ctx, call.arguments);
      if (!slide?.shapes?.addTextBox) throw new Error("This PowerPoint runtime does not expose text boxes needed for live editable diagrams.");
      const canAddShape = Boolean(slide?.shapes?.addGeometricShape || slide?.shapes?.addShape);
      const addShape = slide?.shapes?.addGeometricShape ?? slide?.shapes?.addShape;
      const fillColor = asString(call.arguments.fillColor, "#dbeafe").trim() || "#dbeafe";
      const accentColor = asString(call.arguments.accentColor, "#2563eb").trim() || "#2563eb";
      const fontColor = asString(call.arguments.fontColor, "#111827").trim() || "#111827";

      for (const frame of frames) {
        if (frame.kind === "title") {
          const title = slide.shapes.addTextBox(frame.text);
          applyShapeFrame(title, frame, frame);
          styleTextRange(title, { fontSize: 22, bold: true, fontColor: accentColor });
          continue;
        }
        if (frame.kind === "connector") {
          const connector = slide.shapes.addTextBox(frame.text);
          applyShapeFrame(connector, frame, frame);
          styleTextRange(connector, { fontSize: 18, bold: true, fontColor: accentColor });
          continue;
        }
        const shape = canAddShape ? addShape.call(slide.shapes, normalizeShapeType("roundedRectangle")) : slide.shapes.addTextBox(frame.text);
        applyShapeFrame(shape, frame, frame);
        try { if (canAddShape) shape.fill.setSolidColor(fillColor); } catch {}
        try { if (canAddShape) shape.lineFormat.color = accentColor; } catch {}
        try {
          if (shape.textFrame?.textRange) shape.textFrame.textRange.text = frame.text;
        } catch {}
        styleTextRange(shape, { fontSize: 11, bold: false, fontColor });
      }
      await ctx.sync();
      return toolResult(call, true, `Created live editable PowerPoint ${asString(call.arguments.diagramType, "process") || "process"} diagram with ${cardCount} step${cardCount === 1 ? "" : "s"}.`);
    });
  }

  if (call.name === "powerpoint_arrange_shapes") {
    return globalThis.PowerPoint.run(async (ctx: any) => {
      const slide = await powerPointSlideAtOrCreate(ctx, call.arguments);
      if (!slide?.shapes) throw new Error("This PowerPoint runtime does not expose slide shapes.");
      slide.shapes.load("items");
      await ctx.sync();
      const indexes = normalizeShapeIndexes(call.arguments.shapeNumbers, slide.shapes.items.length);
      const selected = indexes.map((shapeIndex) => slide.shapes.items[shapeIndex]).filter(Boolean);
      if (!selected.length) throw new Error("No PowerPoint shapes were available to arrange.");
      for (const shape of selected) shape.load?.(["left", "top", "width", "height"]);
      await ctx.sync();
      const order = compactKey(call.arguments.order);
      let changed = 0;
      if (order) {
        for (const shape of selected) {
          try {
            if (["front", "bringtofront", "bringfront"].includes(order)) shape.zOrder?.bringToFront?.();
            else if (["back", "sendtoback", "sendback"].includes(order)) shape.zOrder?.sendToBack?.();
            else if (["forward", "bringforward"].includes(order)) shape.zOrder?.bringForward?.();
            else if (["backward", "sendbackward"].includes(order)) shape.zOrder?.sendBackward?.();
            changed += 1;
          } catch {}
        }
      }
      const action = call.arguments.action;
      if (asString(action, "").trim()) changed += arrangeShapeFrames(selected, action);
      await ctx.sync();
      return toolResult(call, true, `Arranged ${selected.length} PowerPoint shape${selected.length === 1 ? "" : "s"}.`);
    });
  }

  if (call.name === "powerpoint_group_shapes") {
    return globalThis.PowerPoint.run(async (ctx: any) => {
      const action = compactKey(call.arguments.action || "group");
      const slide = await powerPointSlideAtOrCreate(ctx, call.arguments);
      if (!slide?.shapes) throw new Error("This PowerPoint runtime does not expose slide shapes.");
      slide.shapes.load("items");
      await ctx.sync();
      const explicitShapeNumbers = Array.isArray(call.arguments.shapeNumbers) && call.arguments.shapeNumbers.length > 0;
      const indexes = normalizeShapeIndexes(call.arguments.shapeNumbers, slide.shapes.items.length);
      const selected = indexes.map((shapeIndex) => slide.shapes.items[shapeIndex]).filter(Boolean);
      if (!selected.length) throw new Error("No PowerPoint shapes were available for grouping.");

      if (["ungroup", "ungroupshapes", "split"].includes(action)) {
        let ungrouped = 0;
        for (const shape of selected) {
          try {
            if (typeof shape.ungroup === "function") {
              shape.ungroup();
              ungrouped += 1;
            } else if (typeof shape.group?.ungroup === "function") {
              shape.group.ungroup();
              ungrouped += 1;
            }
          } catch {}
        }
        if (!ungrouped) throw new Error("This PowerPoint runtime does not expose a supported shape ungroup API. The add-in detected the request, but this Office build cannot ungroup shapes from a task pane.");
        await ctx.sync();
        return toolResult(call, true, "Ungrouped " + ungrouped + " PowerPoint shape group" + (ungrouped === 1 ? "" : "s") + ".");
      }

      if (selected.length < 2) throw new Error("powerpoint_group_shapes requires at least two shapeNumbers when action is group.");
      const shapesApi = slide.shapes;
      const groupingAttempts: Array<() => unknown> = [];
      if (typeof shapesApi.group === "function") groupingAttempts.push(() => shapesApi.group(selected));
      if (typeof shapesApi.groupShapes === "function") groupingAttempts.push(() => shapesApi.groupShapes(selected));
      if (typeof shapesApi.addGroup === "function") groupingAttempts.push(() => shapesApi.addGroup(selected));
      if (typeof shapesApi.createGroup === "function") groupingAttempts.push(() => shapesApi.createGroup(selected));

      for (const attempt of groupingAttempts) {
        try {
          attempt();
          await ctx.sync();
          return toolResult(call, true, "Grouped " + selected.length + " PowerPoint shapes" + (explicitShapeNumbers ? "" : " on the slide") + ".");
        } catch {}
      }

      throw new Error("This PowerPoint runtime does not expose a supported shape grouping API. The add-in will adopt grouping automatically when Microsoft exposes it to task-pane add-ins; for now use generated decks for deterministic grouped layouts or PowerPoint native ribbon Group command.");
    });
  }
  if (call.name === "powerpoint_delete_slide") {
    return globalThis.PowerPoint.run(async (ctx: any) => {
      if (!hasOfficeSlideReference(call.arguments)) throw new Error("powerpoint_delete_slide requires slideIndex or slideNumber.");
      const slides = ctx.presentation.slides;
      slides.load("items");
      await ctx.sync();
      const index = normalizeOfficeSlideIndex(call.arguments, slides.items.length, 0);
      const slide = getPowerPointSlideOrThrow(slides, index);
      if (!slide.delete) throw new Error("This PowerPoint runtime does not expose slide deletion APIs.");
      slide.delete();
      await ctx.sync();
      return toolResult(call, true, `Deleted PowerPoint slide ${index + 1}.`);
    });
  }

  if (call.name === "powerpoint_clear_slide") {
    return globalThis.PowerPoint.run(async (ctx: any) => {
      const slide = await powerPointSlideAtOrCreate(ctx, call.arguments);
      if (!slide?.shapes) throw new Error("This PowerPoint runtime does not expose slide shapes.");
      slide.shapes.load("items");
      await ctx.sync();
      let removed = 0;
      for (const shape of slide.shapes.items) {
        if (!shape.delete) throw new Error("This PowerPoint runtime can read shapes but does not expose shape deletion APIs.");
        shape.delete();
        removed += 1;
      }
      await ctx.sync();
      return toolResult(call, true, `Removed ${removed} shape${removed === 1 ? "" : "s"} from the PowerPoint slide.`);
    });
  }

  if (call.name === "powerpoint_set_slide_background") {
    return globalThis.PowerPoint.run(async (ctx: any) => {
      const slide = await powerPointSlideAtOrCreate(ctx, call.arguments);
      const color = asString(call.arguments.color, "").trim();
      const hasImage = Boolean(asString(call.arguments.assetId, "").trim() || asString(call.arguments.assetName, "").trim() || asString(call.arguments.imageUrl, "").trim());
      if (hasImage) {
        const asset = await resolveImageAsset(call.arguments);
        const addImage = slide?.shapes?.addImage ?? slide?.shapes?.addPicture;
        if (!addImage) throw new Error("This PowerPoint runtime does not expose native image insertion APIs for slide backgrounds.");
        const image = addImage.call(slide.shapes, asset.base64);
        applyShapeFrame(image, { left: 0, top: 0, width: 720, height: 405 }, { left: 0, top: 0, width: 720, height: 405 });
        try { image.zOrder?.sendToBack?.(); } catch {}
        applyAltText(image, asString(call.arguments.altText, asset.name));
        await ctx.sync();
        return toolResult(call, true, `Added ${asset.name} as a full-slide PowerPoint background image.`);
      }
      if (!color) throw new Error("powerpoint_set_slide_background requires color or image asset/url.");
      if (slide?.background?.fill?.setSolidColor) {
        slide.background.fill.setSolidColor(color);
      } else if (slide?.shapes?.addGeometricShape || slide?.shapes?.addShape) {
        const addShape = slide.shapes.addGeometricShape ?? slide.shapes.addShape;
        const rect = addShape.call(slide.shapes, normalizeShapeType("rectangle"));
        applyShapeFrame(rect, { left: 0, top: 0, width: 720, height: 405 }, { left: 0, top: 0, width: 720, height: 405 });
        try { rect.fill.setSolidColor(color); } catch {}
        try { rect.lineFormat.color = color; } catch {}
        try { rect.zOrder?.sendToBack?.(); } catch {}
      } else {
        throw new Error("This PowerPoint runtime does not expose slide background or shape APIs.");
      }
      await ctx.sync();
      return toolResult(call, true, `Set PowerPoint slide background to ${color}.`);
    });
  }

  if (call.name === "powerpoint_add_speaker_notes") {
    return globalThis.PowerPoint.run(async (ctx: any) => {
      const notes = asString(call.arguments.notes, "");
      if (!notes) throw new Error("powerpoint_add_speaker_notes requires notes.");
      const slide = await powerPointSlideAtOrCreate(ctx, call.arguments);
      const notesPage = slide.notesPage ?? slide.notes;
      if (!notesPage) throw new Error("This PowerPoint runtime does not expose speaker notes APIs. Use the roadmap Open XML deck lane for guaranteed notes support.");
      if (notesPage.insertText) notesPage.insertText(notes);
      else if (notesPage.body?.insertText) notesPage.body.insertText(notes);
      else if (typeof notesPage.text === "string") notesPage.text = notes;
      else throw new Error("This PowerPoint runtime exposes a notes object, but not a supported notes text insertion method.");
      await ctx.sync();
      return toolResult(call, true, "Added PowerPoint speaker notes.");
    });
  }

  if (call.name === "powerpoint_insert_image") {
    return globalThis.PowerPoint.run(async (ctx: any) => {
      const asset = await resolveImageAsset(call.arguments);
      const slide = await powerPointSlideAtOrCreate(ctx, call.arguments);
      const addImage = slide?.shapes?.addImage ?? slide?.shapes?.addPicture;
      if (!addImage) throw new Error("This PowerPoint runtime does not expose a native slide image insertion API. The image is available as context, but this Office build cannot place it as a real picture object from the add-in.");
      const image = addImage.call(slide.shapes, asset.base64);
      applyShapeFrame(image, call.arguments, { left: 72, top: 110, width: 560, height: 315 });
      applyAltText(image, asString(call.arguments.altText, asset.name));
      await ctx.sync();
      return toolResult(call, true, `Inserted image ${asset.name} into PowerPoint.`);
    });
  }


  if (call.name === "powerpoint_duplicate_slide") {
    return globalThis.PowerPoint.run(async (ctx: any) => {
      const slides = ctx.presentation.slides;
      slides.load("items");
      await ctx.sync();
      const index = normalizeOfficeSlideIndex(call.arguments, slides.items.length, 0);
      const slide = getPowerPointSlideOrThrow(slides, index);
      if (!slide.duplicate) throw new Error("This PowerPoint runtime does not expose slide duplication APIs. Use the generated deck file tool for template-like duplicated layouts.");
      slide.duplicate();
      await ctx.sync();
      return toolResult(call, true, `Duplicated PowerPoint slide ${index + 1}.`);
    });
  }

  if (call.name === "powerpoint_add_hyperlink_textbox") {
    return globalThis.PowerPoint.run(async (ctx: any) => {
      const text = asString(call.arguments.text, "").trim();
      const url = normalizeOfficeUrl(call.arguments.url);
      if (!text || !url) throw new Error("powerpoint_add_hyperlink_textbox requires text and url.");
      const slide = await powerPointSlideAtOrCreate(ctx, call.arguments);
      if (!slide?.shapes?.addTextBox) throw new Error("This PowerPoint runtime does not expose text box APIs.");
      const shape = slide.shapes.addTextBox(text);
      applyShapeFrame(shape, call.arguments, { left: 72, top: 340, width: 420, height: 40 });
      styleTextRange(shape, { ...call.arguments, fontColor: asString(call.arguments.fontColor, "#2563eb") });
      try {
        const textRange = shape.textFrame?.textRange;
        if (textRange?.hyperlink) textRange.hyperlink.address = url;
        else if (shape.hyperlink) shape.hyperlink.address = url;
        else if (shape.actionSettings?.mouseClick?.hyperlink) shape.actionSettings.mouseClick.hyperlink.address = url;
        else throw new Error("No hyperlink API");
      } catch {
        throw new Error("This PowerPoint runtime can add the text box, but does not expose hyperlink APIs for add-ins. Use powerpoint_generate_deck_file for hyperlink-bearing generated decks once hyperlink XML support is enabled.");
      }
      await ctx.sync();
      return toolResult(call, true, `Added PowerPoint hyperlink text box to ${url}.`);
    });
  }
  return toolResult(call, false, `Unsupported PowerPoint tool: ${call.name}.`);
}
export async function insertIntoOffice(host: OfficeHost, text: string): Promise<string> {
  if (host === "excel" && globalThis.Excel) return insertIntoExcel(text);
  if (host === "word" && globalThis.Word) return insertIntoWord(text);
  if (host === "powerpoint" && globalThis.PowerPoint) return insertIntoPowerPoint(text);
  await navigator.clipboard?.writeText(text).catch(() => undefined);
  return "Copied to clipboard in browser preview mode.";
}

async function insertIntoExcel(text: string): Promise<string> {
  return globalThis.Excel.run(async (ctx: any) => {
    const selected = ctx.workbook.getSelectedRange();
    selected.load("address");
    await ctx.sync();

    const preview = excelInsertionPreview(text);
    const rowCount = preview.rows.length;
    const columnCount = Math.max(1, ...preview.rows.map((row) => row.length));
    const target = selected.getResizedRange(rowCount - 1, columnCount - 1);
    target.values = preview.rows.map((row) => Array.from({ length: columnCount }, (_, index) => row[index] ?? ""));
    target.format.autofitColumns();
    target.format.autofitRows();
    await ctx.sync();

    if (preview.kind === "table") return `Inserted ${rowCount} rows and ${columnCount} columns starting at ${selected.address}.`;
    return `Inserted response into ${selected.address}.`;
  });
}

async function insertIntoWord(text: string): Promise<string> {
  return globalThis.Word.run(async (ctx: any) => {
    ctx.document.getSelection().insertText(text, globalThis.Word.InsertLocation.replace);
    await ctx.sync();
    return "Inserted into the current Word selection.";
  });
}

async function insertIntoPowerPoint(text: string): Promise<string> {
  return globalThis.PowerPoint.run(async (ctx: any) => {
    const slides = ctx.presentation.slides;
    slides.load("items");
    await ctx.sync();

    const slideTexts = splitTextForSlides(text).slice(0, 20);
    let inserted = 0;
    for (const slideText of slideTexts) {
      let slide = inserted === 0 && slides.items[0] ? slides.items[0] : null;
      if (!slide) {
        slides.add();
        slides.load("items");
        await ctx.sync();
        slide = slides.items[slides.items.length - 1];
      }
      if (!slide?.shapes?.addTextBox) throw new Error("This PowerPoint runtime does not expose the shape/text-box API needed to insert slide text.");
      const shape = slide.shapes.addTextBox(slideText);
      shape.left = 36;
      shape.top = 72;
      shape.width = 612;
      shape.height = 300;
      inserted += 1;
    }

    await ctx.sync();
    return inserted === 1 ? "Inserted a text box on the first PowerPoint slide." : `Inserted text boxes across ${inserted} PowerPoint slides.`;
  });
}







