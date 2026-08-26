import { describe, expect, it } from "vitest";
import { validateTask } from "../evals/lib/validate-task.mjs";
import { checkCorpusDiscrimination, listTasks, loadTask, validateAllTasks } from "../evals/run.mjs";

const baseSeed = [{ name: "Inputs", rows: [["Parameter", "Value"], ["Rate", 0.2]] }, { name: "Model", rows: [["Item", "Qty", "Total"], ["A", 5, "=B2*Inputs!$B$2"], ["B", 7, "12"]] }];

function task(overrides = {}) {
  return {
    id: "sample",
    prompt: "Repair the broken total formula on the Model sheet without touching Inputs.",
    protectedSheets: ["Inputs"],
    budget: { maxToolCalls: 10 },
    seed: { sheets: baseSeed },
    targetFixes: { "Model!C3": "=B3*Inputs!$B$2" },
    variants: [{ id: "a", edits: {} }, { id: "b", edits: { "Model!B2": 9 } }],
    ...overrides,
  };
}

describe("eval task spec validation", () => {
  it("accepts a well-formed task", () => {
    const report = validateTask(task());
    expect(report.errors).toEqual([]);
    expect(report.ok).toBe(true);
  });

  it("rejects a task that asserts nothing, which would otherwise score a perfect pass", () => {
    const report = validateTask(task({ targetFixes: {}, expectedValues: {} }));
    expect(report.ok).toBe(false);
    expect(report.errors.join(" ")).toMatch(/asserts nothing/i);
  });

  it("rejects a target fix that the seed already satisfies", () => {
    const report = validateTask(task({ targetFixes: { "Model!C2": "=B2*Inputs!$B$2" } }));
    expect(report.ok).toBe(false);
    expect(report.errors.join(" ")).toMatch(/already matches the seed/i);
  });

  it("rejects a protected sheet that does not exist, because the protection is inert", () => {
    const report = validateTask(task({ protectedSheets: ["Typo"] }));
    expect(report.ok).toBe(false);
    expect(report.errors.join(" ")).toMatch(/not a sheet in the seed/i);
  });

  it("rejects an unsatisfiable value assertion on a formula cell", () => {
    // Generated seeds carry no cached values and there is no recalc oracle.
    const report = validateTask(task({ expectedValues: { "Model!C2": 1 } }));
    expect(report.ok).toBe(false);
    expect(report.errors.join(" ")).toMatch(/recalculation oracle/i);
  });

  it("rejects a corpus entry with fewer than two variants", () => {
    const report = validateTask(task({ variants: [{ id: "a", edits: {} }] }));
    expect(report.ok).toBe(false);
    expect(report.errors.join(" ")).toMatch(/at least two variants/i);
  });

  it("rejects variants that never perturb the seed", () => {
    const report = validateTask(task({ variants: [{ id: "a", edits: {} }, { id: "b", edits: {} }] }));
    expect(report.ok).toBe(false);
    expect(report.errors.join(" ")).toMatch(/must perturb the seed/i);
  });

  it("rejects references to sheets or addresses outside the seed", () => {
    expect(validateTask(task({ targetFixes: { "Ghost!C3": "=1+1" } })).errors.join(" ")).toMatch(/not in the seed/i);
    expect(validateTask(task({ targetFixes: { "Model!C3:C9": "=1+1" } })).errors.join(" ")).toMatch(/single-cell address/i);
  });
});

describe("shipped eval corpus", () => {
  it("has enough tasks to function as a gate", () => {
    expect(listTasks().length).toBeGreaterThanOrEqual(5);
  });

  it("passes spec validation for every shipped task", () => {
    const invalid = validateAllTasks().filter((report) => !report.ok);
    expect(invalid.map((report) => `${report.id}: ${report.errors.join("; ")}`)).toEqual([]);
  });

  it("declares a reference solution and a reference destructive run for every task", () => {
    for (const id of listTasks()) {
      const spec = loadTask(id);
      expect(Object.keys(spec.referenceSolution ?? {}), `${id} referenceSolution`).not.toHaveLength(0);
      expect(Object.keys(spec.referenceDamage ?? {}), `${id} referenceDamage`).not.toHaveLength(0);
    }
  });

  it("separates a clean run from a run that fixes the target but damages healthy cells", () => {
    const reports = checkCorpusDiscrimination();
    expect(reports.length).toBeGreaterThanOrEqual(10);
    const failing = reports.filter((report) => !report.ok);
    expect(failing.map((report) => `${report.id}/${report.variant}: ${report.reason}`)).toEqual([]);
    // Every damaged run must be the dangerous shape: intended fixes landed,
    // overall accuracy failed.
    for (const report of reports) {
      expect(report.damaged.modif, `${report.id}/${report.variant}`).toBe(true);
      expect(report.damaged.acc, `${report.id}/${report.variant}`).toBe(false);
    }
  });
});
