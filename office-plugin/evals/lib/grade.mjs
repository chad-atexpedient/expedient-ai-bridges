// Two-level workbook grader.
//
// Doctrine borrowed from SpreadsheetBench / pi-for-excel:
//   Modif. = did the agent make the intended changes?
//   Acc.   = is the whole workbook still correct?
//
// An agent can score a perfect Modif. while silently rewiring healthy
// formulas elsewhere. That is the failure mode this grader exists to catch,
// so unintended edits are a first-class metric, not a warning.

import { compareAddresses, unionAddresses } from "./workbook.mjs";

const DEFAULT_REL_TOL = 1e-6;
const DEFAULT_ABS_TOL = 1e-9;

function cellOf(sheet, address) {
  return sheet?.get(address) ?? { value: "", formula: "" };
}

function isBlank(cell) {
  const value = cell?.value;
  if (cell?.formula) return false;
  if (value === "" || value === null || value === undefined) return true;
  // Whitespace-only spacer strings count as empty. Documented leniency:
  // layout padding should not register as a destructive edit.
  return typeof value === "string" && value.trim() === "";
}

/**
 * Normalize a formula for comparison, lowercasing and stripping whitespace
 * only OUTSIDE string literals so that ="A b" is not conflated with ="Ab".
 */
export function normalizeFormula(formula = "") {
  const text = String(formula ?? "");
  let output = "";
  let inString = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    if (char === '"') {
      // Doubled quotes inside a literal are an escaped quote, not a terminator.
      if (inString && text[index + 1] === '"') { output += '""'; index += 1; continue; }
      inString = !inString;
      output += char;
      continue;
    }
    if (inString) { output += char; continue; }
    if (/\s/.test(char)) continue;
    output += char.toUpperCase();
  }
  return output;
}

export function formulasEqual(left = "", right = "") {
  return normalizeFormula(left) === normalizeFormula(right);
}

/** Numeric comparison with relative tolerance, falling back to absolute at zero. */
export function valuesEqual(expected, actual, options = {}) {
  const relTol = Number.isFinite(options.relTol) ? options.relTol : DEFAULT_REL_TOL;
  const absTol = Number.isFinite(options.absTol) ? options.absTol : DEFAULT_ABS_TOL;

  // Booleans must stay booleans: TRUE is not 1 for grading purposes.
  if (typeof expected === "boolean" || typeof actual === "boolean") return expected === actual;

  const expectedNumber = typeof expected === "number" ? expected : Number(expected);
  const actualNumber = typeof actual === "number" ? actual : Number(actual);
  const bothNumeric = expected !== "" && actual !== ""
    && Number.isFinite(expectedNumber) && Number.isFinite(actualNumber);

  if (bothNumeric) {
    if (expectedNumber === actualNumber) return true;
    if (expectedNumber === 0) return Math.abs(actualNumber) <= absTol;
    return Math.abs(actualNumber - expectedNumber) / Math.abs(expectedNumber) <= relTol;
  }
  return String(expected ?? "").trim() === String(actual ?? "").trim();
}

function parseRef(ref = "") {
  const bang = String(ref).lastIndexOf("!");
  if (bang < 0) return { sheet: null, address: String(ref).toUpperCase() };
  return {
    sheet: String(ref).slice(0, bang).replace(/^'|'$/g, ""),
    address: String(ref).slice(bang + 1).replace(/\$/g, "").toUpperCase(),
  };
}

/** Check 1: graded output cells match the expected-value oracle. */
export function checkCellsMatch(result, expectedValues = {}, options = {}) {
  const failures = [];
  let checked = 0;

  for (const [ref, expected] of Object.entries(expectedValues)) {
    const { sheet: sheetName, address } = parseRef(ref);
    const sheet = result.sheets.get(sheetName ?? result.sheetNames[0]);
    if (!sheet) {
      failures.push({ ref, reason: "sheet-missing", expected, actual: null });
      continue;
    }
    checked += 1;
    const actual = cellOf(sheet, address).value;
    if (!valuesEqual(expected, actual, options)) failures.push({ ref, reason: "value-mismatch", expected, actual });
  }

  return { name: "cells_match", pass: failures.length === 0, checked, failures };
}

/** Check 2: the intended edits were actually made (formula-level). */
export function checkTargetFixes(result, targetFixes = {}) {
  const failures = [];
  let checked = 0;

  for (const [ref, expectedFormula] of Object.entries(targetFixes)) {
    const { sheet: sheetName, address } = parseRef(ref);
    const sheet = result.sheets.get(sheetName ?? result.sheetNames[0]);
    if (!sheet) {
      failures.push({ ref, reason: "sheet-missing", expected: expectedFormula, actual: null });
      continue;
    }
    checked += 1;
    const actual = cellOf(sheet, address).formula;
    if (!formulasEqual(expectedFormula, actual)) failures.push({ ref, reason: "formula-mismatch", expected: expectedFormula, actual });
  }

  return { name: "target_fixes", pass: failures.length === 0, checked, failures };
}

/** Check 3: protected sheets were not touched, in values OR formulas. */
export function checkNoMutation(seed, result, protectedSheets = []) {
  const failures = [];

  for (const sheetName of protectedSheets) {
    const seedSheet = seed.sheets.get(sheetName);
    const resultSheet = result.sheets.get(sheetName);
    if (!seedSheet) continue;
    if (!resultSheet) {
      failures.push({ ref: sheetName, reason: "sheet-deleted" });
      continue;
    }
    for (const address of unionAddresses(seedSheet, resultSheet)) {
      const before = cellOf(seedSheet, address);
      const after = cellOf(resultSheet, address);
      if (isBlank(before) && isBlank(after)) continue;
      if (!formulasEqual(before.formula, after.formula) || !valuesEqual(before.value, after.value)) {
        failures.push({ ref: `${sheetName}!${address}`, reason: "protected-cell-changed", expected: before, actual: after });
      }
    }
  }

  return { name: "no_mutation", pass: failures.length === 0, checked: protectedSheets.length, failures };
}

/**
 * Check 4: destructive-edit detection.
 *
 * Diffs every cell outside the intended edit set, over the UNION of seed and
 * result addresses so that deletions are caught even when the used range
 * shrank. Formula-to-value replacement is treated as an edit: it is the most
 * common way an agent silently destroys a model.
 */
export function checkUnintendedEdits(seed, result, intendedRefs = [], options = {}) {
  const intended = new Set(intendedRefs.map((ref) => {
    const { sheet, address } = parseRef(ref);
    return `${sheet ?? seed.sheetNames[0]}!${address}`;
  }));
  const ignored = new Set((options.ignoreSheets ?? []).map(String));
  const edits = [];

  const sheetNames = [...new Set([...seed.sheetNames, ...result.sheetNames])];
  for (const sheetName of sheetNames) {
    if (ignored.has(sheetName)) continue;
    const seedSheet = seed.sheets.get(sheetName);
    const resultSheet = result.sheets.get(sheetName);

    if (seedSheet && !resultSheet) {
      edits.push({ ref: sheetName, reason: "sheet-deleted" });
      continue;
    }
    if (!seedSheet) continue; // Newly added sheets are additive, not destructive.

    for (const address of unionAddresses(seedSheet, resultSheet)) {
      const ref = `${sheetName}!${address}`;
      if (intended.has(ref)) continue;
      const before = cellOf(seedSheet, address);
      const after = cellOf(resultSheet, address);
      if (isBlank(before) && isBlank(after)) continue;

      if (before.formula && !after.formula) {
        edits.push({ ref, reason: "formula-replaced-by-value", expected: before.formula, actual: after.value });
        continue;
      }
      if (!formulasEqual(before.formula, after.formula)) {
        edits.push({ ref, reason: "formula-changed", expected: before.formula, actual: after.formula });
        continue;
      }
      if (!before.formula && !valuesEqual(before.value, after.value, options)) {
        edits.push({ ref, reason: "value-changed", expected: before.value, actual: after.value });
      }
    }
  }

  edits.sort((a, b) => String(a.ref).localeCompare(String(b.ref), undefined, { numeric: true }));
  return { name: "unintended_edited_cells", pass: edits.length === 0, checked: edits.length, failures: edits };
}

/**
 * Run all four checks and produce the two-level verdict.
 *
 * modif = intended changes landed (cells_match + target_fixes)
 * acc   = modif AND nothing else broke (no_mutation + unintended edits)
 */
export function gradeRun(seed, result, task = {}, options = {}) {
  const intendedRefs = [
    ...Object.keys(task.targetFixes ?? {}),
    ...Object.keys(task.expectedValues ?? {}),
    ...(task.allowedEdits ?? []),
  ];

  const checks = [
    checkCellsMatch(result, task.expectedValues ?? {}, options),
    checkTargetFixes(result, task.targetFixes ?? {}),
    checkNoMutation(seed, result, task.protectedSheets ?? []),
    checkUnintendedEdits(seed, result, intendedRefs, { ...options, ignoreSheets: task.ignoreSheets ?? [] }),
  ];

  const byName = Object.fromEntries(checks.map((check) => [check.name, check]));
  const modif = byName.cells_match.pass && byName.target_fixes.pass;
  const acc = modif && byName.no_mutation.pass && byName.unintended_edited_cells.pass;

  return {
    task: task.id ?? "unnamed-task",
    variant: task.variant ?? null,
    modif,
    acc,
    checks,
    unintendedEditCount: byName.unintended_edited_cells.failures.length,
    summary: summarize(task, modif, acc, byName),
  };
}

function summarize(task, modif, acc, byName) {
  const unintended = byName.unintended_edited_cells.failures.length;
  const lines = [
    `Task: ${task.id ?? "unnamed-task"}${task.variant ? ` (variant ${task.variant})` : ""}`,
    `Modif.: ${modif ? "PASS" : "FAIL"}   Acc.: ${acc ? "PASS" : "FAIL"}`,
  ];
  for (const check of Object.values(byName)) {
    lines.push(`- ${check.name}: ${check.pass ? "pass" : `FAIL (${check.failures.length})`}`);
  }
  if (modif && unintended > 0) {
    lines.push(`! Made all intended changes but also edited ${unintended} unrelated cell${unintended === 1 ? "" : "s"}. This is a destructive run.`);
  }
  return lines.join("\n");
}

export { compareAddresses };