import { describe, expect, it } from "vitest";
import {
  buildWorkbookOverview,
  cellAddress,
  describeOverwriteConflicts,
  excelColumnIndex,
  excelColumnLetters,
  formatRangeForModel,
  formatSearchMatches,
  normalizeReadMode,
  rangeStartCell,
  searchRangeMatches,
  verifyWrittenValues,
} from "../src/office/excelRead";

describe("Excel address helpers", () => {
  it("round-trips column letters and indexes", () => {
    expect(excelColumnLetters(0)).toBe("A");
    expect(excelColumnLetters(25)).toBe("Z");
    expect(excelColumnLetters(26)).toBe("AA");
    expect(excelColumnIndex("A")).toBe(0);
    expect(excelColumnIndex("AA")).toBe(26);
    expect(excelColumnIndex(excelColumnLetters(731))).toBe(731);
  });

  it("resolves the start cell of absolute and sheet-qualified ranges", () => {
    expect(rangeStartCell("$C$5:$F$20")).toEqual({ column: 2, row: 4 });
    expect(rangeStartCell("B2")).toEqual({ column: 1, row: 1 });
    expect(rangeStartCell(undefined)).toEqual({ column: 0, row: 0 });
  });

  it("builds addresses offset from a base cell", () => {
    expect(cellAddress(2, 4)).toBe("C5");
  });
});

describe("Range rendering for the model", () => {
  const values = [
    ["Region", "Revenue"],
    ["West", 10],
    ["East", 5],
  ];

  it("defaults to compact mode and prefixes real addresses", () => {
    expect(normalizeReadMode(undefined)).toBe("compact");
    const rendered = formatRangeForModel(values, { baseAddress: "B2" });
    expect(rendered.mode).toBe("compact");
    expect(rendered.text.split("\n")).toEqual([
      "B2\tRegion\tRevenue",
      "B3\tWest\t10",
      "B4\tEast\t5",
    ]);
  });

  it("escapes CSV output without address columns", () => {
    const rendered = formatRangeForModel([["a,b", 'say "hi"']], { mode: "csv" });
    expect(rendered.text).toBe('"a,b","say ""hi"""');
  });

  it("includes formulas and number formats in detailed mode", () => {
    const rendered = formatRangeForModel([[10, 25]], {
      mode: "detailed",
      baseAddress: "A1",
      formulas: [["10", "=A1*2.5"]],
      numberFormats: [["General", "0.00"]],
    });
    expect(rendered.text).toContain("A1: 10");
    expect(rendered.text).toContain("B1: 25 | formula==A1*2.5 | format=0.00");
  });

  it("bounds output and reports what was truncated", () => {
    const wide = [Array.from({ length: 12 }, (_, index) => `c${index}`)];
    const tall = Array.from({ length: 8 }, (_, index) => [`r${index}`]);
    const boundedColumns = formatRangeForModel(wide, { maxColumns: 3 });
    expect(boundedColumns.truncatedColumns).toBe(9);
    expect(boundedColumns.notes.join(" ")).toContain("9 more columns");

    const boundedRows = formatRangeForModel(tall, { maxRows: 3 });
    expect(boundedRows.rowsShown).toBe(3);
    expect(boundedRows.truncatedRows).toBe(5);
    expect(boundedRows.notes.join(" ")).toContain("5 more rows");
  });
});

describe("Workbook overview blueprint", () => {
  it("summarizes sheets, headers, objects, and named ranges", () => {
    const overview = buildWorkbookOverview({
      workbookName: "Q3.xlsx",
      activeSheet: "Data",
      selection: "Data!A1:C4",
      sheets: [
        { name: "Data", usedRange: "A1:C40", rowCount: 40, columnCount: 3, headers: ["Region", "Revenue", "Margin"], tables: ["SalesTable"], charts: [] },
        { name: "Scratch", usedRange: "", rowCount: 0, columnCount: 0, visible: false },
      ],
      namedRanges: [{ name: "TaxRate", refersTo: "=Data!$C$1" }],
    });

    expect(overview.sheetCount).toBe(2);
    expect(overview.namedRangeCount).toBe(1);
    expect(overview.text).toContain("Workbook: Q3.xlsx");
    expect(overview.text).toContain("Active sheet: Data");
    expect(overview.text).toContain("- Data (used A1:C40; 40r x 3c; tables: SalesTable)");
    expect(overview.text).toContain("headers: Region | Revenue | Margin");
    expect(overview.text).toContain("hidden");
    expect(overview.text).toContain("- TaxRate -> =Data!$C$1");
  });
});

describe("Workbook search", () => {
  const values = [
    ["Region", "Revenue"],
    ["West", 10],
    ["East", 5],
  ];
  const formulas = [
    ["Region", "Revenue"],
    ["West", "=SUM(Sales!A1:A9)"],
    ["East", "5"],
  ];

  it("returns real addresses for value matches", () => {
    const matches = searchRangeMatches("Data", values, formulas, "A1", "east");
    expect(matches).toEqual([{ sheet: "Data", address: "A3", value: "East", formula: "" }]);
  });

  it("matches formula text and reports the formula", () => {
    const matches = searchRangeMatches("Data", values, formulas, "A1", "Sales!");
    expect(matches).toEqual([{ sheet: "Data", address: "B2", value: "10", formula: "=SUM(Sales!A1:A9)" }]);
  });

  it("honors case sensitivity, whole-cell, and formula opt-out", () => {
    expect(searchRangeMatches("Data", values, formulas, "A1", "east", { matchCase: true })).toEqual([]);
    expect(searchRangeMatches("Data", values, formulas, "A1", "Reg", { wholeCell: true })).toEqual([]);
    expect(searchRangeMatches("Data", values, formulas, "A1", "Sales!", { searchFormulas: false })).toEqual([]);
  });

  it("offsets addresses by the searched range origin and respects the limit", () => {
    const matches = searchRangeMatches("Data", [["hit", "hit"], ["hit", "hit"]], [], "C10", "hit", { limit: 3 });
    expect(matches.map((match) => match.address)).toEqual(["C10", "D10", "C11"]);
  });

  it("formats an empty result set clearly", () => {
    expect(formatSearchMatches([], "missing")).toBe('No cells matched "missing".');
    expect(formatSearchMatches(searchRangeMatches("Data", values, formulas, "A1", "east"), "east")).toContain("Data!A3: East");
  });
});

describe("Write safety", () => {
  it("reports the cells a write would destroy", () => {
    const conflicts = describeOverwriteConflicts([["keep", ""], ["", "also"]], "B2");
    expect(conflicts.count).toBe(2);
    expect(conflicts.addresses).toEqual(["B2", "C3"]);
    expect(conflicts.summary).toContain("2 non-empty cells would be overwritten");
  });

  it("treats an empty target range as safe", () => {
    expect(describeOverwriteConflicts([["", ""], ["", ""]], "A1").count).toBe(0);
    expect(describeOverwriteConflicts(undefined, "A1").count).toBe(0);
  });

  it("verifies literal writes and ignores formula cells", () => {
    const verified = verifyWrittenValues([["Total", "=SUM(A1:A9)"]], [["Total", 42]], "A1");
    expect(verified.verified).toBe(true);
    expect(verified.checked).toBe(1);
    expect(verified.summary).toBe("Verified 1 written cell.");
  });

  it("flags cells Excel did not store as requested", () => {
    const verified = verifyWrittenValues([["Total", "10"]], [["Total", "7"]], "A1");
    expect(verified.verified).toBe(false);
    expect(verified.mismatches).toEqual([{ address: "B1", expected: "10", actual: "7" }]);
    expect(verified.summary).toContain("B1");
  });

  it("accepts numeric round-trips that change representation", () => {
    expect(verifyWrittenValues([["10.0"]], [[10]], "A1").verified).toBe(true);
  });
});