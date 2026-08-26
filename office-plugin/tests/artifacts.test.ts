import { describe, expect, it } from "vitest";
import { excelInsertionPreview, extractTableRows, splitTextForSlides } from "../src/lib/artifacts";

describe("insertable artifacts", () => {
  it("extracts CSV into rows for Excel insertion", () => {
    expect(extractTableRows("Month,Price\nJan,3.10\nFeb,3.20")).toEqual([
      ["Month", "Price"],
      ["Jan", "3.10"],
      ["Feb", "3.20"],
    ]);
  });

  it("extracts JSON object arrays into header rows", () => {
    expect(extractTableRows('[{"month":"Jan","price":3.1},{"month":"Feb","price":3.2}]')).toEqual([
      ["month", "price"],
      ["Jan", "3.1"],
      ["Feb", "3.2"],
    ]);
  });

  it("detects Excel table insertion instead of single-cell insertion", () => {
    const preview = excelInsertionPreview("| A | B |\n| --- | --- |\n| 1 | 2 |");
    expect(preview.kind).toBe("table");
    expect(preview.rows).toEqual([["A", "B"], ["1", "2"]]);
  });

  it("splits explicit PowerPoint slide sections", () => {
    expect(splitTextForSlides("Title\n\n--- slide\n\nNext")).toEqual(["Title", "Next"]);
  });
});