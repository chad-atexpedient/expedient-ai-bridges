import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";
import path from "node:path";

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const nodePath = process.execPath;
const cli = path.join(root, "node_modules", "office-addin-debugging", "cli.js");
const args = [cli, "stop", "manifest.xml"];
const child = spawn(nodePath, args, { cwd: root, stdio: "inherit", shell: false });
child.on("exit", (code) => process.exit(code ?? 0));
