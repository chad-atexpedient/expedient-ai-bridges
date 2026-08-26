import { readFileSync, writeFileSync } from "node:fs";
import path from "node:path";

const origin = process.argv[2] || process.env.ADDIN_ORIGIN;
if (!origin) {
  console.error("Usage: node tools/make-production-manifest.mjs https://your-addin.example.com");
  process.exit(1);
}
const normalized = origin.replace(/\/+$/, "");
const parsedOrigin = new URL(normalized);
if (parsedOrigin.protocol !== "https:") {
  console.error("Production add-in origin must use HTTPS.");
  process.exit(1);
}
let xml = readFileSync("manifest.xml", "utf8").replaceAll("https://localhost:3000", normalized);
const ssoClientId = process.env.MSAL_CLIENT_ID || process.env.M365_CLIENT_ID;
const ssoResource = process.env.OFFICE_SSO_RESOURCE || (ssoClientId ? `api://${parsedOrigin.host}/${ssoClientId}` : "");
if (!ssoClientId || !ssoResource || ssoClientId === "00000000-0000-0000-0000-000000000000" || ssoResource.includes("00000000-0000-0000-0000-000000000000")) {
  console.error("Production packaging requires a real MSAL_CLIENT_ID and OFFICE_SSO_RESOURCE.");
  process.exit(1);
}
xml = xml.replaceAll("00000000-0000-0000-0000-000000000000", ssoClientId);
xml = xml.replace(/<Resource>[^<]+<\/Resource>/, `<Resource>${ssoResource}</Resource>`);
if (xml.includes("localhost") || xml.includes("00000000-0000-0000-0000-000000000000")) {
  console.error("Production manifest still contains development placeholders.");
  process.exit(1);
}
const out = path.join("dist", "manifest.production.xml");
writeFileSync(out, xml);
console.log(`Wrote ${out} for ${normalized}`);
