import { cp, mkdir, readFile, rm, stat, writeFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const args = new Set(process.argv.slice(2));
const devMode = args.has("--dev") || args.has("--sideload");
const releaseRoot = path.join(root, "dist", "release");
const releaseDir = path.join(releaseRoot, "ctrl-byok-office-addin");
const rel = (value) => path.relative(root, value).replace(/\\/g, "/");

async function exists(filePath) {
  try { await stat(filePath); return true; } catch { return false; }
}

async function readText(relativePath) {
  return readFile(path.join(root, relativePath), "utf8");
}

async function copyRequired(sourceRelative, targetRelative = sourceRelative) {
  const source = path.join(root, sourceRelative);
  if (!await exists(source)) throw new Error(`Missing required release input: ${sourceRelative}`);
  const target = path.join(releaseDir, targetRelative);
  await mkdir(path.dirname(target), { recursive: true });
  await cp(source, target, { recursive: true });
}

async function sha256(relativePath) {
  const bytes = await readFile(path.join(releaseDir, relativePath));
  return createHash("sha256").update(bytes).digest("hex");
}

const distApp = path.join(root, "dist", "app");
if (!await exists(distApp)) throw new Error("dist/app is missing. Run the production build before packaging a release.");

const manifestSource = devMode ? "manifest.xml" : "dist/manifest.production.xml";
const manifestXml = await readText(manifestSource);
if (!devMode) {
  if (manifestXml.includes("localhost") || manifestXml.includes("00000000-0000-0000-0000-000000000000")) {
    throw new Error("Production release manifest still contains localhost or placeholder Entra IDs.");
  }
  if (!manifestXml.includes("https://")) throw new Error("Production release manifest must contain HTTPS URLs.");
}

const resolvedRelease = path.resolve(releaseDir);
const resolvedReleaseRoot = path.resolve(releaseRoot);
if (!resolvedRelease.startsWith(resolvedReleaseRoot + path.sep)) throw new Error("Refusing to package outside dist/release.");
await rm(releaseDir, { recursive: true, force: true });
await mkdir(releaseDir, { recursive: true });

await copyRequired("dist/app", "app");
await copyRequired(manifestSource, "manifest.xml");
await copyRequired("server", "server");
await copyRequired("package.json", "package.json");
await copyRequired("pnpm-lock.yaml", "pnpm-lock.yaml");
await copyRequired("pnpm-workspace.yaml", "pnpm-workspace.yaml");
await copyRequired("deploy/production/.env.example", "deploy/production/.env.example");
await copyRequired("deploy/production/Dockerfile", "deploy/production/Dockerfile");
await copyRequired("deploy/production/m365-centralized-deployment.md", "deploy/production/m365-centralized-deployment.md");
await copyRequired("deploy/production/scripts/smoke-test.mjs", "deploy/production/scripts/smoke-test.mjs");
await copyRequired("docs/production-launch-checklist.md", "docs/production-launch-checklist.md");
await copyRequired("docs/office-capability-roadmap.md", "docs/office-capability-roadmap.md");

const packageJson = JSON.parse(await readText("package.json"));
const hashes = [];
for (const file of ["manifest.xml", "package.json", "pnpm-lock.yaml", "app/index.html"]) {
  hashes.push(`- ${file}: ${await sha256(file)}`);
}

const report = `# CTRL BYOK Office Add-in release package

Mode: ${devMode ? "development sideload" : "production"}
Package: ${packageJson.name}@${packageJson.version}
Generated: ${new Date().toISOString()}

## Contents

- app/ - built static task-pane assets served by server/server.mjs
- server/ - same-origin production Node server and API handlers
- manifest.xml - ${devMode ? "development sideload manifest" : "production Microsoft 365 deployment manifest"}
- deploy/production/.env.example - required production environment template
- deploy/production/Dockerfile - container build definition
- deploy/production/m365-centralized-deployment.md - Microsoft 365 Admin Center rollout guide
- deploy/production/scripts/smoke-test.mjs - deployed-service smoke test
- docs/production-launch-checklist.md - production readiness checklist
- docs/office-capability-roadmap.md - current capability/roadmap matrix

## Production launch reminder

Before broad rollout, replace the checked-in fail-closed production example modules with real tenant policy, settings store, rate limiter, auth validator, audit sink, and vault-backed secret integrations. Then run the production smoke test against the hosted HTTPS origin.

## Key checksums

${hashes.join("\n")}
`;

await writeFile(path.join(releaseDir, "RELEASE_REPORT.md"), report);
console.log(`Packaged ${devMode ? "development" : "production"} release directory: ${rel(releaseDir)}`);
