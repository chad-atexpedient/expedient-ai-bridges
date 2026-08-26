// Task-spec validator.
//
// The corpus is the asset, so a malformed task is worse than a missing one: a
// task that asserts nothing scores a perfect Modif./Acc. pass and quietly
// removes a gate. Every trap below was confirmed against the live grader:
//
//   1. A task with no targetFixes and no expectedValues graded as modif+acc.
//   2. A task whose targetFixes already match the seed required no work to pass.
//   3. protectedSheets naming a nonexistent sheet was silently ignored.
//   4. expectedValues on a formula cell can never match, because generated
//      fixtures carry no cached values and the harness has no recalc oracle.

import { readWorkbook } from "./workbook.mjs";
import { buildWorkbook } from "./fixtures.mjs";
import { formulasEqual } from "./grade.mjs";

const ADDRESS = /^[A-Z]+[1-9]\d*$/;
const QUOTE_EDGES = /^'|'$/g;

function parseRef(ref = "") {
  const bang = String(ref).lastIndexOf("!");
  if (bang < 0) return { sheet: null, address: String(ref).toUpperCase() };
  return {
    sheet: String(ref).slice(0, bang).replace(QUOTE_EDGES, ""),
    address: String(ref).slice(bang + 1).replace(/\$/g, "").toUpperCase(),
  };
}

/**
 * Validate one task spec, optionally against its built seed workbook.
 * Returns { ok, errors, warnings }. Errors must block a release.
 */
export function validateTask(task, options = {}) {
  const errors = [];
  const warnings = [];
  const push = (list, message) => list.push(message);

  if (!task || typeof task !== "object") return { ok: false, errors: ["Task is not an object."], warnings };
  if (!task.id) push(errors, "Task is missing an id.");
  if (!task.prompt || String(task.prompt).trim().length < 20) push(errors, "Task needs a prompt of at least 20 characters, or the agent has no instruction to follow.");
  if (!task.seed?.sheets?.length) push(errors, "Task is missing seed.sheets.");

  const targetFixes = task.targetFixes ?? {};
  const expectedValues = task.expectedValues ?? {};
  const assertionCount = Object.keys(targetFixes).length + Object.keys(expectedValues).length;
  // Trap 1: the vacuous task.
  if (assertionCount === 0) push(errors, "Task asserts nothing: with no targetFixes and no expectedValues every run scores a perfect pass.");

  if (!Array.isArray(task.variants) || task.variants.length < 2) {
    push(errors, "Task needs at least two variants so a memorized answer cannot pass.");
  } else {
    const ids = task.variants.map((variant) => variant.id);
    if (new Set(ids).size !== ids.length) push(errors, "Variant ids must be unique.");
    const substantive = task.variants.filter((variant) => Object.keys(variant.edits ?? {}).length > 0);
    if (!substantive.length) push(errors, "At least one variant must perturb the seed; identical variants add no defense against memorization.");
  }

  if (!Array.isArray(task.protectedSheets) || !task.protectedSheets.length) {
    push(warnings, "Task declares no protectedSheets. Consider marking any sheet the agent must not touch.");
  }
  if (!task.budget?.maxToolCalls) push(warnings, "Task declares no maxToolCalls budget, so a runaway agent is unbounded.");

  if (options.checkSeed === false || !task.seed?.sheets?.length) return { ok: errors.length === 0, errors, warnings };

  let seed;
  try {
    seed = readWorkbook(buildWorkbook(task.seed.sheets));
  } catch (error) {
    push(errors, `Seed workbook could not be built or read: ${error.message}`);
    return { ok: false, errors, warnings };
  }

  const sheetNames = new Set(seed.sheetNames);
  const declared = new Set((task.seed.sheets ?? []).map((sheet) => sheet.name));

  // Trap 3: a protected sheet that does not exist protects nothing.
  for (const sheetName of task.protectedSheets ?? []) {
    if (!declared.has(sheetName)) push(errors, `protectedSheets names "${sheetName}", which is not a sheet in the seed. The protection is silently inert.`);
  }
  for (const sheetName of task.ignoreSheets ?? []) {
    if (!declared.has(sheetName)) push(warnings, `ignoreSheets names "${sheetName}", which is not a seed sheet.`);
  }

  const checkRef = (ref, label) => {
    const { sheet, address } = parseRef(ref);
    const resolved = sheet ?? seed.sheetNames[0];
    if (!ADDRESS.test(address)) { push(errors, `${label} reference "${ref}" is not a single-cell address.`); return null; }
    if (!sheetNames.has(resolved)) { push(errors, `${label} reference "${ref}" points at sheet "${resolved}", which is not in the seed.`); return null; }
    return { sheetName: resolved, address };
  };

  for (const [ref, expectedFormula] of Object.entries(targetFixes)) {
    const resolved = checkRef(ref, "targetFixes");
    if (!resolved) continue;
    if (!String(expectedFormula).startsWith("=")) {
      push(warnings, `targetFixes["${ref}"] is not a formula. Formula-level grading is the point; a literal target is weak.`);
    }
    const cell = seed.sheets.get(resolved.sheetName)?.get(resolved.address) ?? { value: "", formula: "" };
    // Trap 2: the no-op fix.
    if (cell.formula && formulasEqual(cell.formula, expectedFormula)) {
      push(errors, `targetFixes["${ref}"] already matches the seed, so the task passes without the agent doing anything.`);
    }
  }

  for (const ref of Object.keys(expectedValues)) {
    const resolved = checkRef(ref, "expectedValues");
    if (!resolved) continue;
    const cell = seed.sheets.get(resolved.sheetName)?.get(resolved.address) ?? { value: "", formula: "" };
    // Trap 4: generated fixtures omit cached values and there is no recalc
    // oracle, so a value assertion on a formula cell is unsatisfiable.
    if (cell.formula) {
      push(errors, `expectedValues["${ref}"] targets a formula cell. Generated seeds carry no cached values and the harness has no recalculation oracle, so this assertion can never pass. Assert the formula in targetFixes instead.`);
    }
  }

  for (const ref of task.allowedEdits ?? []) checkRef(ref, "allowedEdits");

  for (const variant of task.variants ?? []) {
    for (const ref of Object.keys(variant.edits ?? {})) {
      const resolved = checkRef(ref, `variant "${variant.id}" edits`);
      if (!resolved) continue;
      // A variant that perturbs a graded target changes the answer itself.
      if (Object.keys(targetFixes).some((target) => {
        const parsed = parseRef(target);
        return (parsed.sheet ?? seed.sheetNames[0]) === resolved.sheetName && parsed.address === resolved.address;
      })) {
        push(warnings, `variant "${variant.id}" edits ${ref}, which is also a graded targetFix. Confirm the expected formula still holds.`);
      }
    }
  }

  return { ok: errors.length === 0, errors, warnings };
}
