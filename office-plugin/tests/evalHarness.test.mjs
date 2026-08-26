import { describe, expect, it } from "vitest";
import { applyEdits, buildWorkbook } from "../evals/lib/fixtures.mjs";
import { readWorkbook } from "../evals/lib/workbook.mjs";
import { formulasEqual, gradeRun, normalizeFormula, valuesEqual } from "../evals/lib/grade.mjs";
import { buildSeed, gradeResultBuffer, listTasks, loadTask, selfTest, seedSheetsForVariant } from "../evals/run.mjs";

const CORRECT_FIXES = {
  "Model!F4": "=(E4-B4*D4)/E4",
  "Model!F6": "=(E6-B6*D6)/E6",
  "Model!E7": "=SUM(E2:E6)",
};

function gradeWithEdits(task, edits, variant) {
  const sheets = seedSheetsForVariant(task, variant);
  return gradeResultBuffer(task, buildWorkbook(applyEdits(sheets, edits)), variant);
}

describe("Workbook reader", () => {
  it("round-trips values, formulas, booleans, and sheet names", () => {
    const workbook = readWorkbook(buildWorkbook([
      { name: "Data", rows: [["Name", "Qty"], ["Widget", 3], ["Total", "=SUM(B2:B2)"]] },
      { name: "Flags", rows: [["On", true]] },
    ]));

    expect(workbook.sheetNames).toEqual(["Data", "Flags"]);
    expect(workbook.sheets.get("Data").get("A1")).toEqual({ value: "Name", formula: "" });
    expect(workbook.sheets.get("Data").get("B2")).toEqual({ value: 3, formula: "" });
    expect(workbook.sheets.get("Data").get("B3")).toEqual({ value: "", formula: "=SUM(B2:B2)" });
    expect(workbook.sheets.get("Flags").get("B1").value).toBe(true);
  });

  it("omits empty cells so diffs stay sparse", () => {
    const workbook = readWorkbook(buildWorkbook([{ name: "S", rows: [["a", "", "c"]] }]));
    expect(workbook.sheets.get("S").has("B1")).toBe(false);
    expect(workbook.sheets.get("S").has("C1")).toBe(true);
  });

  it("rejects packages that are not valid archives", () => {
    expect(() => readWorkbook(Buffer.from("not a workbook"))).toThrow(/not a valid ZIP/);
  });
});

describe("Formula normalization", () => {
  it("ignores case and whitespace outside string literals", () => {
    expect(formulasEqual("=sum( A1 : A9 )", "=SUM(A1:A9)")).toBe(true);
    expect(normalizeFormula("=IF(A1>0, B1, 0)")).toBe("=IF(A1>0,B1,0)");
  });

  it("preserves whitespace and case inside string literals", () => {
    expect(formulasEqual('=IF(A1,"Past Due","ok")', '=IF(A1,"PastDue","ok")')).toBe(false);
    expect(formulasEqual('=A1&" x "', '=A1&" x "')).toBe(true);
  });

  it("handles escaped quotes inside literals", () => {
    expect(formulasEqual('=IF(A1,"a""b","c")', '=IF(A1,"a""b","c")')).toBe(true);
  });
});

describe("Value comparison", () => {
  it("uses relative tolerance for nonzero numbers", () => {
    expect(valuesEqual(1000, 1000.0000001)).toBe(true);
    expect(valuesEqual(1000, 1001)).toBe(false);
  });

  it("uses absolute tolerance at zero", () => {
    expect(valuesEqual(0, 1e-12)).toBe(true);
    expect(valuesEqual(0, 0.5)).toBe(false);
  });

  it("keeps booleans distinct from numbers", () => {
    expect(valuesEqual(true, 1)).toBe(false);
    expect(valuesEqual(true, true)).toBe(true);
  });
});

describe("Two-level grading", () => {
  const task = loadTask("margin-model-doctor");

  it("passes both levels for a run that makes exactly the intended fixes", () => {
    const verdict = gradeWithEdits(task, CORRECT_FIXES);
    expect(verdict.modif).toBe(true);
    expect(verdict.acc).toBe(true);
    expect(verdict.unintendedEditCount).toBe(0);
  });

  it("fails a run that fixed every target bug but flattened healthy formulas", () => {
    const verdict = gradeWithEdits(task, {
      ...CORRECT_FIXES,
      "Model!E2": 2500,
      "Model!E3": 3000,
      "Model!F2": 0.28,
    });

    // This is the whole point of the harness: Modif. looks perfect.
    expect(verdict.modif).toBe(true);
    expect(verdict.acc).toBe(false);
    expect(verdict.unintendedEditCount).toBe(3);

    const destructive = verdict.checks.find((check) => check.name === "unintended_edited_cells");
    expect(destructive.failures.map((failure) => failure.ref)).toEqual(["Model!E2", "Model!E3", "Model!F2"]);
    expect(destructive.failures.every((failure) => failure.reason === "formula-replaced-by-value")).toBe(true);
    expect(verdict.summary).toContain("This is a destructive run.");
  });

  it("fails a run that edited a protected sheet", () => {
    const verdict = gradeWithEdits(task, { ...CORRECT_FIXES, "Inputs!B2": 0.25 });
    expect(verdict.modif).toBe(true);
    expect(verdict.acc).toBe(false);
    expect(verdict.checks.find((check) => check.name === "no_mutation").pass).toBe(false);
  });

  it("fails a run that only fixed some of the target bugs", () => {
    const verdict = gradeWithEdits(task, { "Model!F4": "=(E4-B4*D4)/E4" });
    expect(verdict.modif).toBe(false);
    expect(verdict.acc).toBe(false);
  });

  it("detects a deleted formula even when the used range shrinks", () => {
    const verdict = gradeWithEdits(task, { ...CORRECT_FIXES, "Model!E5": "" });
    expect(verdict.acc).toBe(false);
    const refs = verdict.checks.find((check) => check.name === "unintended_edited_cells").failures.map((failure) => failure.ref);
    expect(refs).toContain("Model!E5");
  });

  it("treats a deleted sheet as a destructive edit", () => {
    const seed = readWorkbook(buildSeed(task));
    const withoutModel = readWorkbook(buildWorkbook(task.seed.sheets.filter((sheet) => sheet.name !== "Model")));
    const verdict = gradeRun(seed, withoutModel, task);
    expect(verdict.acc).toBe(false);
    expect(verdict.checks.find((check) => check.name === "unintended_edited_cells").failures.some((failure) => failure.reason === "sheet-deleted")).toBe(true);
  });

  it("does not penalize a newly added sheet", () => {
    const seed = readWorkbook(buildSeed(task));
    const withExtra = readWorkbook(buildWorkbook([
      ...applyEdits(task.seed.sheets, CORRECT_FIXES),
      { name: "Notes", rows: [["Explanation", "Fixed three formulas."]] },
    ]));
    const verdict = gradeRun(seed, withExtra, task);
    expect(verdict.acc).toBe(true);
  });
});

describe("Hidden variants", () => {
  const task = loadTask("margin-model-doctor");

  it("defines variants that perturb input literals", () => {
    expect(task.variants.map((variant) => variant.id)).toEqual(["a", "b", "c"]);
  });

  it("accepts the same correct formulas across every variant", () => {
    for (const variant of task.variants) {
      const verdict = gradeWithEdits(task, CORRECT_FIXES, variant.id);
      expect(verdict.modif, `variant ${variant.id} modif`).toBe(true);
      expect(verdict.acc, `variant ${variant.id} acc`).toBe(true);
      expect(verdict.variant).toBe(variant.id);
    }
  });

  it("changes the seed data between variants so memorized values cannot pass", () => {
    const baseline = readWorkbook(buildSeed(task, "a")).sheets.get("Model").get("B2").value;
    const perturbed = readWorkbook(buildSeed(task, "b")).sheets.get("Model").get("B2").value;
    expect(perturbed).not.toBe(baseline);
  });

  it("rejects an unknown variant instead of silently grading the baseline", () => {
    expect(() => buildSeed(task, "zz")).toThrow(/Variant zz is not defined/);
  });
});

describe("Additive task: quarterly rollup", () => {
  const task = loadTask("quarterly-rollup");
  const CORRECT = {
    "Rollup!E1": "Q4",
    "Rollup!F1": "QoQ",
    "Rollup!E2": "=('Q4 Data'!B2)",
    "Rollup!E3": "=('Q4 Data'!B3)",
    "Rollup!E4": "=('Q4 Data'!B4)",
    "Rollup!E5": "=('Q4 Data'!B5)",
    "Rollup!E6": "=SUM(E2:E5)",
    "Rollup!F2": "=(E2-D2)/D2",
    "Rollup!F3": "=(E3-D3)/D3",
    "Rollup!F4": "=(E4-D4)/D4",
    "Rollup!F5": "=(E5-D5)/D5",
  };

  it("passes when new columns are added without touching prior quarters", () => {
    for (const variant of task.variants) {
      const verdict = gradeWithEdits(task, CORRECT, variant.id);
      expect(verdict.modif, `variant ${variant.id}`).toBe(true);
      expect(verdict.acc, `variant ${variant.id}`).toBe(true);
    }
  });

  it("treats allowed edits as non-destructive", () => {
    const verdict = gradeWithEdits(task, CORRECT);
    expect(verdict.unintendedEditCount).toBe(0);
  });

  it("flags flattening the existing Total formulas even when the new work is correct", () => {
    const verdict = gradeWithEdits(task, { ...CORRECT, "Rollup!B6": 595, "Rollup!C6": 625 });
    expect(verdict.modif).toBe(true);
    expect(verdict.acc).toBe(false);
    expect(verdict.unintendedEditCount).toBe(2);
  });

  it("flags overwriting a protected source sheet", () => {
    const verdict = gradeWithEdits(task, { ...CORRECT, "Q4 Data!B2": 999 });
    expect(verdict.acc).toBe(false);
    expect(verdict.checks.find((check) => check.name === "no_mutation").pass).toBe(false);
  });
});
describe("Harness wiring", () => {
  it("discovers task specs on disk", () => {
    expect(listTasks()).toContain("margin-model-doctor");
  });

  it("reports a missing task clearly", () => {
    expect(() => loadTask("no-such-task")).toThrow(/not found/);
  });

  it("keeps the runner self-test aligned with the grader", () => {
    const results = selfTest();
    expect(results.map((result) => result.label)).toEqual(["clean run", "destructive run", "protected-sheet violation", "incomplete run"]);
    expect(results[0].verdict.acc).toBe(true);
    expect(results[1].verdict.modif && !results[1].verdict.acc).toBe(true);
    expect(results[3].verdict.modif).toBe(false);
  });
});