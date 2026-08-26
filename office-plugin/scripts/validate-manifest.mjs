import { readFileSync } from "node:fs";
import { XMLParser } from "fast-xml-parser";

const xml = readFileSync("manifest.xml", "utf8");
const parser = new XMLParser({ ignoreAttributes: false });
const parsed = parser.parse(xml);
const app = parsed.OfficeApp;
const errors = [];
if (!app) errors.push("OfficeApp root is missing");
const hosts = app?.Hosts?.Host;
const hostNames = Array.isArray(hosts) ? hosts.map((host) => host["@_Name"]) : [hosts?.["@_Name"]];
for (const required of ["Workbook", "Document", "Presentation"]) {
  if (!hostNames.includes(required)) errors.push(`Missing host ${required}`);
}
const source = app?.DefaultSettings?.SourceLocation?.["@_DefaultValue"];
if (!source?.startsWith("https://localhost:3000")) errors.push("Default SourceLocation must point at https://localhost:3000");
if (!xml.includes("ReadWriteDocument")) errors.push("ReadWriteDocument permission missing");
if (!xml.includes("<WebApplicationInfo>")) errors.push("WebApplicationInfo missing for Office SSO");
if (!xml.includes("<Scope>Files.Read</Scope>")) errors.push("Files.Read SSO scope missing");
if (process.env.PRODUCTION_MANIFEST === "true" && (xml.includes("localhost") || xml.includes("00000000-0000-0000-0000-000000000000"))) errors.push("Production manifest contains placeholders");
if (errors.length) {
  console.error(errors.join("\n"));
  process.exit(1);
}
console.log("Manifest basic validation passed for Excel, Word, and PowerPoint.");
