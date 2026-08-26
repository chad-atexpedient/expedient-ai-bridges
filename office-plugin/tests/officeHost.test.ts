import { describe, expect, it } from "vitest";
import { arrangeShapeFrames, createCleanedRows, createCombinedRows, createExcelContextImportRows, createFormulaAuditRows, createFormulaDependencyRows, createPivotChartReportRows, createPowerPointAuditRows, createPowerPointDiagramFrames, createSummaryRows, createWordAuditRows, createWordContextBrief, formulaReferences, normalizeAddress, normalizeOfficeSlideIndex, normalizeOfficeUrl, officeErrorMessage, safeExcelName } from "../src/office/host";

describe("Office tool argument normalization", () => {
  it("normalizes Excel addresses from user/model-friendly forms", () => {
    expect(normalizeAddress("'Gas Prices'!$A$5:$C$17")).toBe("A5:C17");
    expect(normalizeAddress("select A5:C17 on Sheet1")).toBe("A5:C17");
    expect(normalizeAddress(" = b2 ")).toBe("B2");
  });

  it("accepts human slide numbers and zero-based indexes", () => {
    expect(normalizeOfficeSlideIndex({ slideNumber: 2 }, 3)).toBe(1);
    expect(normalizeOfficeSlideIndex({ slideIndex: 2 }, 2)).toBe(1);
    expect(normalizeOfficeSlideIndex({ slideIndex: 0 }, 3)).toBe(0);
    expect(normalizeOfficeSlideIndex({ slidePosition: "last" }, 5)).toBe(4);
  });

  it("normalizes safe links and Excel names", () => {
    expect(normalizeOfficeUrl("example.com/report")).toBe("https://example.com/report");
    expect(() => normalizeOfficeUrl("javascript:alert(1)")).toThrow(/Unsupported hyperlink URL/);
    expect(safeExcelName("2026 gas prices")).toBe("_2026_gas_prices");
  });

  it("creates pivot-style summary rows from worksheet values", () => {
    const rows = createSummaryRows([
      ["Region", "Revenue"],
      ["West", 10],
      ["East", 5],
      ["West", 7],
    ], "Region", "Revenue", "sum");
    expect(rows).toEqual([
      ["Region", "Sum of Revenue", "Rows"],
      ["East", 5, 1],
      ["West", 17, 2],
    ]);
  });

  it("creates sorted top-N pivot-chart report rows", () => {
    const rows = createPivotChartReportRows([
      ["Region", "Revenue"],
      ["West", 10],
      ["East", 5],
      ["West", 7],
      ["North", 20],
    ], "Region", "Revenue", "sum", { sortBy: "valueDesc", top: 2 });

    expect(rows).toEqual([
      ["Region", "Sum of Revenue", "Rows"],
      ["North", 20, 1],
      ["West", 17, 2],
    ]);
  });

  it("creates cleaned Excel output rows for Power Query-style cleanup", () => {
    const cleaned = createCleanedRows([
      ["Name", "Tags", "Amount"],
      [" Acme   Inc ", "red, blue", " 10 "],
      ["", "", ""],
      ["Acme Inc", "red, blue", "10"],
      ["Beta", "green", 5],
    ], {
      splitColumn: "Tags",
      delimiter: ",",
      removeBlankRows: true,
      removeDuplicateRows: true,
      trimWhitespace: true,
      normalizeSpaces: true,
    });

    expect(cleaned.rows).toEqual([
      ["Name", "Tags 1", "Tags 2", "Amount"],
      ["Acme Inc", "red", "blue", "10"],
      ["Beta", "green", "", 5],
    ]);
    expect(cleaned.removedBlankRows).toBe(1);
    expect(cleaned.removedDuplicateRows).toBe(1);
    expect(cleaned.outputColumns).toBe(4);
    expect(cleaned.transforms).toContain("removed duplicate rows");
  });

  it("appends two Excel ranges with unioned columns and source labels", () => {
    const combined = createCombinedRows([
      ["Customer", "Amount"],
      [" Acme ", 10],
    ], [
      ["Customer", "Region"],
      ["Beta", "West"],
    ], { mode: "append", matchColumns: "union", includeSourceColumn: true, primaryLabel: "Orders", secondaryLabel: "Pipeline" });

    expect(combined.rows).toEqual([
      ["Source", "Customer", "Amount", "Region"],
      ["Orders", "Acme", 10, ""],
      ["Pipeline", "Beta", "", "West"],
    ]);
    expect(combined.mode).toBe("append");
    expect(combined.outputColumns).toBe(4);
  });

  it("left-merges lookup columns by key", () => {
    const combined = createCombinedRows([
      ["Customer", "Amount"],
      ["Acme", 10],
      ["Beta", 5],
    ], [
      ["Customer", "Region", "Amount"],
      ["Acme", "East", 99],
    ], { mode: "merge", primaryKey: "Customer", secondaryKey: "Customer" });

    expect(combined.rows).toEqual([
      ["Customer", "Amount", "Region", "Amount from lookup"],
      ["Acme", 10, "East", 99],
      ["Beta", 5, "", ""],
    ]);
    expect(combined.mode).toBe("merge");
    expect(combined.matchedRows).toBe(1);
  });

  it("imports cross-surface markdown context into safe Excel rows", () => {
    const imported = createExcelContextImportRows(`| Task | Owner | Risk |\n| --- | --- | --- |\n| Build model | Finance | =HIGH |\n| Review deck | Sales | Medium |`, { sourceLabel: "Planning Notes.docx", includeSourceColumn: true });
    expect(imported.detectedMode).toBe("table");
    expect(imported.rows).toEqual([
      ["Source", "Task", "Owner", "Risk"],
      ["Planning Notes.docx", "Build model", "Finance", "'=HIGH"],
      ["Planning Notes.docx", "Review deck", "Sales", "Medium"],
    ]);
  });

  it("imports key-value and bullet context for Excel", () => {
    expect(createExcelContextImportRows("Client: Acme\nBudget: $10,000", { mode: "keyValue", includeSourceColumn: false }).rows).toEqual([
      ["Field", "Value"],
      ["Client", "Acme"],
      ["Budget", "$10,000"],
    ]);
    expect(createExcelContextImportRows("- Discover - Gather requirements\n- Build - Create workbook", { mode: "bullets", includeSourceColumn: false }).rows).toEqual([
      ["Item", "Detail"],
      ["Discover", "Gather requirements"],
      ["Build", "Create workbook"],
    ]);
  });

  it("creates Word document audit findings for structure and accessibility review", () => {
    const rows = createWordAuditRows(`Executive summary
This paragraph is intentionally long because it is trying to mimic a dense business report paragraph that keeps going and going without giving the reader a useful break, which makes the document harder to scan, harder to review, and harder for assistive technology users to process efficiently because the paragraph contains too many words before it reaches a natural stopping point for the reader.
CLICK HERE for details: https://example.test/source`, { imageCount: 1, tableCount: 1 });
    const text = rows.flat().join(" | ");
    expect(rows[0]).toEqual(["Area", "Severity", "Finding", "Recommendation"]);
    expect(text).toContain("Paragraph 2");
    expect(text).toContain("Confirm every meaningful image has descriptive alt text");
    expect(text).toContain("Confirm each table has a header row");
    expect(text).toContain("Generic link text");
  });

  it("creates a Word context brief from cross-surface table context", () => {
    const brief = createWordContextBrief(`| Metric | Value | Note |\n| --- | --- | --- |\n| Pipeline | $4.2M | Needs exec review |\n| Risk | Medium | Staffing dependency |`, { title: "Pipeline Brief", sourceLabel: "Sales Pipeline.xlsx", audience: "executive" });
    expect(brief.title).toBe("Pipeline Brief");
    expect(brief.sourceLabel).toBe("Sales Pipeline.xlsx");
    expect(brief.keyPoints[0]).toContain("Metric: Pipeline");
    expect(brief.tableRows?.[0]).toEqual(["Metric", "Value", "Note"]);
  });

  it("creates a Word context brief from key-value notes", () => {
    const brief = createWordContextBrief("Client: Acme\nDecision: Approve pilot\nOwner: Ops", { sourceLabel: "Meeting notes" });
    expect(brief.keyPoints).toEqual(["Client: Acme", "Decision: Approve pilot", "Owner: Ops"]);
    expect(brief.tableRows).toEqual([["Field", "Value"], ["Client", "Acme"], ["Decision", "Approve pilot"], ["Owner", "Ops"]]);
  });

  it("creates PowerPoint deck audit findings from slide snapshots", () => {
    const rows = createPowerPointAuditRows([
      { slideNumber: 1, shapeCount: 0, text: "" },
      { slideNumber: 2, shapeCount: 18, text: "A".repeat(800) },
      { slideNumber: 3, shapeCount: 3, text: "Revenue by region" },
    ], { audience: "board" });
    const text = rows.flat().join(" | ");
    expect(rows[0]).toEqual(["Area", "Severity", "Finding", "Recommendation"]);
    expect(text).toContain("blank slide");
    expect(text).toContain("more than 14 objects");
    expect(text).toContain("text-heavy");
    expect(text).toContain("executive orientation");
  });

  it("creates PowerPoint diagram frames with cards and connectors", () => {
    const frames = createPowerPointDiagramFrames([
      { label: "Discover", detail: "Gather needs" },
      { label: "Build", detail: "Create draft" },
      { label: "Review", detail: "Polish output" },
    ], "process", { title: "Workflow", left: 50, top: 80, width: 600, height: 220 });
    expect(frames.find((frame) => frame.kind === "title")?.text).toBe("Workflow");
    expect(frames.filter((frame) => frame.kind === "card")).toHaveLength(3);
    expect(frames.filter((frame) => frame.kind === "connector")).toHaveLength(2);
    expect(frames.every((frame) => frame.left >= 0 && frame.top >= 0 && frame.width > 0 && frame.height > 0)).toBe(true);
  });

  it("creates cycle diagram frames around a center point", () => {
    const frames = createPowerPointDiagramFrames(["Plan", "Do", "Check", "Act"], "cycle", { left: 60, top: 100, width: 500, height: 240 });
    expect(frames.filter((frame) => frame.kind === "card")).toHaveLength(4);
    expect(frames.filter((frame) => frame.kind === "connector")).toHaveLength(4);
  });

  it("creates professional formula audit findings from formula/value matrices", () => {
    const rows = createFormulaAuditRows([
      ["Revenue", "Margin", "Risk"],
      ["=B1*1.2", "=VLOOKUP(A2,Sheet2!A:B,2,FALSE)", "=TODAY()"],
      ["=B2*1.2", "=B2/[Book2.xlsx]Inputs!A1", "=SUM(A1,A3)"],
      ["=B3*9.99", "=B3/C3", "=A3+B3"],
      ["=B4*1.2", "=B4/C4", "=A4+B4"],
      ["=B5*1.2", "=B5/C5", "=A5+B5"],
    ], [
      ["Revenue", "Margin", "Risk"],
      [120, 4, "2026-01-01"],
      [144, "#DIV/0!", 12],
      [200, 8, 14],
      [240, 9, 16],
      [288, 10, 18],
    ], "A1", true);

    const text = rows.flat().join(" | ");
    expect(rows[0]).toEqual(["Cell", "Severity", "Issue", "Formula", "Recommendation"]);
    expect(text).toContain("Volatile or indirect function");
    expect(text).toContain("Legacy lookup function");
    expect(text).toContain("External workbook or web reference");
    expect(text).toContain("Hard-coded numeric constant");
    expect(text).toContain("Formula differs from nearby");
    expect(text).toContain("Formula currently returns #DIV/0!");
  });

  it("maps formula precedents and dependent summaries", () => {
    const refs = formulaReferences("=SUM(A1:A3)+'Inputs'!$B$2+[Budget.xlsx]Sheet1!C4+https://example.test/source", "Model");
    expect(refs.localReferences).toEqual(["Inputs!B2", "Model!A1:A3"]);
    expect(refs.externalReferences).toEqual(["[Budget.xlsx]Sheet1!C4", "https://example.test/source"]);

    const rows = createFormulaDependencyRows([
      ["Input", "Calc", "Output"],
      ["", "=A2*1.1", "=B2+B3"],
      ["", "=A3*1.1", "=B2*2"],
      ["", "='Inputs'!$B$2+[Budget.xlsx]Plan!C4", "=SUM(B2:B4)"],
    ], "A1", "Model", true);
    const text = rows.flat().join(" | ");
    expect(rows[0]).toEqual(["Formula Cell", "Precedents", "External References", "Formula"]);
    expect(text).toContain("Model!B2");
    expect(text).toContain("Model!A2");
    expect(text).toContain("Inputs!B2");
    expect(text).toContain("[Budget.xlsx]Plan!C4");
    expect(text).toContain("Precedent Cell/Range");
    expect(text).toContain("Dependent Formula Cells");
  });

  it("adds recovery guidance for cryptic live Office API failures", () => {
    const message = officeErrorMessage(
      { message: "The argument is invalid or missing or has an incorrect format." },
      "Office operation failed",
      { id: "call-1", name: "powerpoint_add_shape", arguments: { slideNumber: 2, left: -10, apiKey: "secret" } },
    );
    expect(message).toContain("Recovery: retry at most once");
    expect(message).toContain("powerpoint_generate_deck_file");
    expect(message).toContain("instead of giving manual instructions");
    expect(message).toContain('"apiKey":"[redacted]"');
  });

  it("points Excel and Word live API failures at generated Office artifact fallbacks", () => {
    expect(officeErrorMessage(
      { message: "This Excel runtime does not expose chart APIs." },
      "Office operation failed",
      { id: "call-2", name: "excel_create_chart", arguments: { sourceAddress: "A1:C12" } },
    )).toContain("excel_generate_workbook_file");
    expect(officeErrorMessage(
      { message: "The property 'name' is not available. Before reading the property's value, call the load method on the containing object and call context.sync()." },
      "Office operation failed",
      { id: "call-3", name: "word_insert_content_control", arguments: { title: "Client" } },
    )).toContain("word_generate_document_file");
  });
  it("arranges PowerPoint shape frames for alignment and distribution", () => {
    const shapes = [
      { left: 10, top: 10, width: 20, height: 20 },
      { left: 80, top: 20, width: 20, height: 20 },
      { left: 150, top: 30, width: 20, height: 20 },
    ];
    expect(arrangeShapeFrames(shapes, "alignTop")).toBe(3);
    expect(shapes.map((shape) => shape.top)).toEqual([10, 10, 10]);
    expect(arrangeShapeFrames(shapes, "distributeHorizontal")).toBe(3);
    expect(shapes.map((shape) => shape.left)).toEqual([10, 80, 150]);
  });
});
