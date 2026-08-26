import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";
import path from "node:path";

const app = process.argv[2] || "excel";
const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const nodePath = process.execPath;
const cli = path.join(root, "node_modules", "office-addin-debugging", "cli.js");
const args = [
  cli,
  "start",
  "manifest.xml",
  "desktop",
  "--app",
  app,
  "--no-debug",
  "--no-live-reload",
  "--debug-method",
  "direct",
  "--dev-server",
  "node node_modules/vite/bin/vite.js --host localhost --port 3000",
  "--dev-server-port",
  "3000",
];

const child = spawn(nodePath, args, { cwd: root, stdio: "inherit", shell: false });
child.on("exit", (code) => process.exit(code ?? 0));
