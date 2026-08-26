import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const read = (relative) => readFile(path.join(root, relative), "utf8");

const packageJson = JSON.parse(await read("package.json"));
const manifest = await read("manifest.xml");
const dockerfile = await read("deploy/production/Dockerfile");
const envExample = await read("deploy/production/.env.example");
const lockfile = await read("pnpm-lock.yaml");

const failures = [];
const requireText = (source, text, label) => {
  if (!source.includes(text)) failures.push(`${label}: missing ${text}`);
};

for (const script of ["build", "test", "validate:manifest", "smoke:ui", "package:prod", "package:dev", "eval:self-test"]) {
  if (!packageJson.scripts?.[script]) failures.push(`package.json: missing release script ${script}`);
}
if (!packageJson.scripts?.["package:prod"]?.includes("scripts/package-release.mjs")) failures.push("package.json: package:prod must build the release directory with scripts/package-release.mjs");
if (!packageJson.scripts?.["package:dev"]?.includes("--dev")) failures.push("package.json: package:dev must create a development/sideload release package");
if (/00000000-0000-0000-0000-000000000000|your-client-id|your-addin-domain/i.test(manifest)) {
  // The checked-in manifest is development-only; this is a guard against accidentally
  // treating it as a production artifact, not a failure of local development.
  requireText(packageJson.scripts?.["manifest:prod"] || "", "make-production-manifest", "package.json");
}
requireText(dockerfile, "pnpm install --frozen-lockfile", "Dockerfile");
requireText(dockerfile, "USER app", "Dockerfile");
requireText(dockerfile, "NODE_ENV=production", "Dockerfile");
for (const variable of ["BYOK_ALLOWED_TARGETS", "IMAGE_ALLOWED_TARGETS", "WEB_FETCH_ALLOWED_TARGETS", "AUTH_VALIDATOR_MODULE", "TENANT_POLICY_MODULE", "SETTINGS_STORE_MODULE", "RATE_LIMITER_MODULE", "ADDIN_ORIGIN", "AUTHORIZED_TENANTS"]) {
  requireText(envExample, variable, "production env example");
}
// The agent eval harness must stay wired in and self-consistent: a grader that
// silently stops detecting destructive runs is worse than no grader at all.
const evalTasks = JSON.parse(await read("evals/tasks/margin-model-doctor.json"));
if (!Array.isArray(evalTasks.variants) || evalTasks.variants.length < 2) {
  failures.push("evals: margin-model-doctor must define hidden variants so memorized answers cannot pass");
}
if (!Array.isArray(evalTasks.protectedSheets) || !evalTasks.protectedSheets.length) {
  failures.push("evals: margin-model-doctor must declare protected sheets");
}
try {
  const { selfTest, validateAllTasks, checkCorpusDiscrimination } = await import(new URL("../evals/run.mjs", import.meta.url));
  // A malformed task spec is a silently removed gate: a task that asserts
  // nothing scores a perfect pass. Validate the corpus, not just the grader.
  const specReports = validateAllTasks();
  if (specReports.length < 3) failures.push(`evals: corpus has only ${specReports.length} task(s); a single-task corpus is not a gate`);
  for (const report of specReports.filter((item) => !item.ok)) {
    failures.push(`evals: task ${report.id} is invalid (${report.errors.join("; ")})`);
  }
  // Every task must separate a clean run from one that fixes the target while
  // damaging healthy formulas. That separation is the reason the harness exists.
  for (const report of checkCorpusDiscrimination().filter((item) => !item.ok)) {
    failures.push(`evals: ${report.id} variant ${report.variant} does not discriminate (${report.reason})`);
  }
  const results = selfTest();
  const destructive = results.find((result) => result.label === "destructive run");
  if (!destructive) failures.push("evals: grader self-test is missing the destructive-run case");
  else if (!(destructive.verdict.modif && !destructive.verdict.acc)) {
    failures.push("evals: grader no longer fails a run that made every intended fix while damaging other formulas");
  }
  const clean = results.find((result) => result.label === "clean run");
  if (clean && !clean.verdict.acc) failures.push("evals: grader rejects a known-correct run");
} catch (error) {
  failures.push(`evals: grader self-test could not run (${error.message})`);
}
if (!lockfile.includes("lockfileVersion: '9.0'")) failures.push("pnpm-lock.yaml: unexpected lockfile version");
if (lockfile.includes("specifier: latest")) failures.push("pnpm-lock.yaml: latest dependency specifier is not reproducible");

if (failures.length) {
  console.error(failures.map((failure) => `FAIL ${failure}`).join("\n"));
  process.exit(1);
}
console.log("PASS release-gates: packaging, container, environment, and lockfile invariants");
