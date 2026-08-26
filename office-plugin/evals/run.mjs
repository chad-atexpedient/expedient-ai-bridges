// Eval runner: grade a workbook against a task spec.
//
//   node evals/run.mjs --task <id> --result <result.xlsx>
//   node evals/run.mjs --task <id> --seed-only --out seed.xlsx
//   node evals/run.mjs --self-test
//
// Grading is on workbook STATE, never on what the agent claimed it did.

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { buildWorkbook, applyEdits } from "./lib/fixtures.mjs";
import { readWorkbook } from "./lib/workbook.mjs";
import { gradeRun } from "./lib/grade.mjs";
import { validateTask } from "./lib/validate-task.mjs";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const TASK_DIR = path.join(HERE, "tasks");

export function loadTask(id) {
  const file = path.join(TASK_DIR, `${id}.json`);
  if (!fs.existsSync(file)) throw new Error(`Task ${id} not found at ${file}`);
  return JSON.parse(fs.readFileSync(file, "utf8"));
}

export function listTasks() {
  if (!fs.existsSync(TASK_DIR)) return [];
  return fs.readdirSync(TASK_DIR).filter((name) => name.endsWith(".json")).map((name) => name.replace(/\.json$/, ""));
}

/**
 * Apply a variant's literal perturbations to the seed.
 * Hidden variants defeat memorized answers: an agent that hardcodes a value
 * from a known task fails the moment the inputs shift.
 */
export function seedSheetsForVariant(task, variantId) {
  const base = task.seed?.sheets ?? [];
  if (!variantId) return base;
  const variant = (task.variants ?? []).find((item) => item.id === variantId);
  if (!variant) throw new Error(`Variant ${variantId} is not defined for task ${task.id}.`);
  return applyEdits(base, variant.edits ?? {});
}

export function buildSeed(task, variantId) {
  return buildWorkbook(seedSheetsForVariant(task, variantId));
}

/** Grade a result buffer against the task's seed. */
export function gradeResultBuffer(task, resultBuffer, variantId, options = {}) {
  const seed = readWorkbook(buildSeed(task, variantId));
  const result = readWorkbook(resultBuffer);
  return gradeRun(seed, result, { ...task, variant: variantId ?? null }, options);
}

/**
 * Validate every task spec in the corpus.
 *
 * A malformed task is a silently removed gate, so corpus validation is a
 * release check rather than an authoring convenience.
 */
export function validateAllTasks() {
  return listTasks().map((id) => ({ id, ...validateTask(loadTask(id)) }));
}
function parseArgs(argv) {
  const args = {};
  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (!token.startsWith("--")) continue;
    const key = token.slice(2);
    const next = argv[index + 1];
    if (!next || next.startsWith("--")) args[key] = true;
    else { args[key] = next; index += 1; }
  }
  return args;
}

function formatVerdict(verdict, json) {
  if (json) return JSON.stringify(verdict, null, 2);
  const lines = [verdict.summary];
  for (const check of verdict.checks) {
    if (check.pass) continue;
    lines.push("", `${check.name} failures:`);
    for (const failure of check.failures.slice(0, 25)) {
      const expected = failure.expected && typeof failure.expected === "object" ? JSON.stringify(failure.expected) : failure.expected;
      const actual = failure.actual && typeof failure.actual === "object" ? JSON.stringify(failure.actual) : failure.actual;
      lines.push(`  ${failure.ref}  [${failure.reason}]  expected=${expected ?? "(none)"}  actual=${actual ?? "(none)"}`);
    }
    if (check.failures.length > 25) lines.push(`  ...and ${check.failures.length - 25} more.`);
  }
  return lines.join("\n");
}

/**
 * Self-test: prove the grader catches the destructive-but-passing run.
 * This is the regression that justifies the harness existing.
 */
export function selfTest() {
  const task = loadTask("margin-model-doctor");
  const results = [];

  // 1. A clean run: exactly the three intended fixes.
  const clean = buildWorkbook(applyEdits(task.seed.sheets, {
    "Model!F4": "=(E4-B4*D4)/E4",
    "Model!F6": "=(E6-B6*D6)/E6",
    "Model!E7": "=SUM(E2:E6)",
  }));
  results.push({ label: "clean run", verdict: gradeResultBuffer(task, clean) });

  // 2. The dangerous run: all three bugs fixed, but healthy formulas
  //    elsewhere were flattened into literals.
  const destructive = buildWorkbook(applyEdits(task.seed.sheets, {
    "Model!F4": "=(E4-B4*D4)/E4",
    "Model!F6": "=(E6-B6*D6)/E6",
    "Model!E7": "=SUM(E2:E6)",
    "Model!E2": 2500,
    "Model!E3": 3000,
    "Model!E4": 3200,
    "Model!F2": 0.28,
  }));
  results.push({ label: "destructive run", verdict: gradeResultBuffer(task, destructive) });

  // 3. A run that violated the protected sheet.
  const protectedViolation = buildWorkbook(applyEdits(task.seed.sheets, {
    "Model!F4": "=(E4-B4*D4)/E4",
    "Model!F6": "=(E6-B6*D6)/E6",
    "Model!E7": "=SUM(E2:E6)",
    "Inputs!B2": 0.25,
  }));
  results.push({ label: "protected-sheet violation", verdict: gradeResultBuffer(task, protectedViolation) });

  // 4. An incomplete run: only one bug fixed.
  const incomplete = buildWorkbook(applyEdits(task.seed.sheets, { "Model!F4": "=(E4-B4*D4)/E4" }));
  results.push({ label: "incomplete run", verdict: gradeResultBuffer(task, incomplete) });

  return results;
}

/**
 * Corpus-wide discrimination check.
 *
 * For every task, and every variant, assert two things:
 *   - the reference solution scores Modif.=true and Acc.=true
 *   - a run that applies the reference solution AND damages healthy cells
 *     scores Modif.=true but Acc.=false
 *
 * The second half is the whole point: a grader that cannot separate those two
 * runs is not a gate. Reference solutions live in the task spec so the corpus
 * stays self-describing.
 */
export function checkCorpusDiscrimination() {
  const reports = [];
  for (const id of listTasks()) {
    const task = loadTask(id);
    const solution = task.referenceSolution;
    const damage = task.referenceDamage;
    if (!solution || !damage) {
      reports.push({ id, variant: null, ok: false, reason: "task is missing referenceSolution or referenceDamage" });
      continue;
    }
    const variantIds = [null, ...(task.variants ?? []).map((variant) => variant.id).filter((value) => value !== "a")];
    for (const variantId of variantIds) {
      const base = seedSheetsForVariant(task, variantId);
      const clean = gradeResultBuffer(task, buildWorkbook(applyEdits(base, solution)), variantId);
      const damaged = gradeResultBuffer(task, buildWorkbook(applyEdits(base, { ...solution, ...damage })), variantId);
      const ok = clean.modif && clean.acc && damaged.modif && !damaged.acc && damaged.unintendedEditCount > 0;
      reports.push({
        id,
        variant: variantId ?? "a",
        ok,
        clean: { modif: clean.modif, acc: clean.acc },
        damaged: { modif: damaged.modif, acc: damaged.acc, unintended: damaged.unintendedEditCount },
        reason: ok ? null : (!clean.acc ? "reference solution does not pass" : "grader failed to flag the destructive run"),
      });
    }
  }
  return reports;
}
function main() {
  const args = parseArgs(process.argv.slice(2));

  if (args.list) {
    console.log(listTasks().join("\n") || "(no tasks)");
    return 0;
  }

  if (args["check-corpus"]) {
    const reports = checkCorpusDiscrimination();
    if (!reports.length) { console.error("No tasks found to check."); return 1; }
    let failed = 0;
    for (const report of reports) {
      if (!report.ok) failed += 1;
      const detail = report.reason ? ` (${report.reason})` : ` clean=${report.clean.modif}/${report.clean.acc} damaged=${report.damaged.modif}/${report.damaged.acc} unintended=${report.damaged.unintended}`;
      console.log(`${report.ok ? "PASS" : "FAIL"} ${report.id} variant=${report.variant}${detail}`);
    }
    console.log(failed ? `\n${failed} corpus discrimination check(s) failed.` : `\n${reports.length} corpus discrimination check(s) passed.`);
    return failed ? 1 : 0;
  }

  if (args.validate) {
    const reports = validateAllTasks();
    if (!reports.length) { console.error("No tasks found to validate."); return 1; }
    let failed = 0;
    for (const report of reports) {
      console.log(`${report.ok ? "PASS" : "FAIL"} ${report.id}`);
      for (const error of report.errors) console.log(`  ERROR ${error}`);
      if (!args.quiet) for (const warning of report.warnings) console.log(`  warn  ${warning}`);
      if (!report.ok) failed += 1;
    }
    console.log(failed ? `\n${failed} task spec(s) are invalid.` : `\n${reports.length} task spec(s) valid.`);
    return failed ? 1 : 0;
  }

  if (args["self-test"]) {
    let failures = 0;
    for (const { label, verdict } of selfTest()) {
      const expectation = {
        "clean run": verdict.modif && verdict.acc,
        "destructive run": verdict.modif && !verdict.acc && verdict.unintendedEditCount === 4,
        "protected-sheet violation": verdict.modif && !verdict.acc,
        "incomplete run": !verdict.modif,
      }[label];
      if (!expectation) failures += 1;
      console.log(`${expectation ? "PASS" : "FAIL"} ${label}: Modif.=${verdict.modif} Acc.=${verdict.acc} unintended=${verdict.unintendedEditCount}`);
    }
    console.log(failures ? `\n${failures} self-test expectation(s) failed.` : "\nGrader self-test passed.");
    return failures ? 1 : 0;
  }

  const taskId = typeof args.task === "string" ? args.task : null;
  if (!taskId) {
    console.error("Usage: node evals/run.mjs --task <id> --result <file.xlsx>\n       node evals/run.mjs --task <id> --seed-only --out <file.xlsx>\n       node evals/run.mjs --self-test\n       node evals/run.mjs --list\n       node evals/run.mjs --validate");
    return 2;
  }

  const task = loadTask(taskId);
  const variant = typeof args.variant === "string" ? args.variant : null;

  if (args["seed-only"]) {
    const out = typeof args.out === "string" ? args.out : `${taskId}${variant ? `-${variant}` : ""}-seed.xlsx`;
    fs.writeFileSync(out, buildSeed(task, variant));
    console.log(`Wrote seed workbook to ${out}`);
    console.log(`Prompt: ${task.prompt}`);
    return 0;
  }

  if (typeof args.result !== "string") {
    console.error("--result <file.xlsx> is required when grading.");
    return 2;
  }

  const verdict = gradeResultBuffer(task, fs.readFileSync(args.result), variant);
  console.log(formatVerdict(verdict, Boolean(args.json)));
  return verdict.acc ? 0 : 1;
}

if (process.argv[1] && fileURLToPath(import.meta.url) === path.resolve(process.argv[1])) {
  process.exit(main());
}