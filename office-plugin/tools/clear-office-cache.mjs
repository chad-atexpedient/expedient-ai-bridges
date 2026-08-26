import fs from "node:fs";
import path from "node:path";
import os from "node:os";

const wef = path.join(os.homedir(), "AppData", "Local", "Microsoft", "Office", "16.0", "Wef");
const targets = [
  path.join(wef, "AddinInfo"),
  path.join(wef, "AppCommands"),
  path.join(wef, "Resources"),
  path.join(wef, "AggregatedCache"),
];
for (const target of targets) {
  if (fs.existsSync(target)) {
    fs.rmSync(target, { recursive: true, force: true });
    console.log(`Removed ${target}`);
  }
}
console.log("Office add-in cache cleared. Close all Office apps before relaunching.");
