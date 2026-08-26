import { chromium } from "playwright";

const origin = (process.argv[2] || "https://localhost:3000").replace(/\/+$/, "");
const browser = await chromium.launch({ headless: true });
const results = [];

for (const viewport of [
  { name: "desktop", width: 1000, height: 760 },
  { name: "narrow", width: 360, height: 640 },
]) {
  const page = await browser.newPage({ viewport: { width: viewport.width, height: viewport.height }, ignoreHTTPSErrors: true });
  const consoleErrors = [];
  page.on("console", (message) => { if (message.type() === "error") consoleErrors.push(message.text()); });
  await page.goto(`${origin}/`, { waitUntil: "networkidle", timeout: 30_000 });
  const setupVisible = await page.locator(".setup-note").count();
  await page.getByRole("button", { name: "Open settings" }).click();
  await page.locator(".settings-scroll select").first().selectOption("managed");
  const draftDisclosure = await page.getByText(/Mode preview only/).count();
  const managedKeyDisabled = await page.locator('input[placeholder="Configured by your organization"]').isDisabled();
  await page.getByRole("button", { name: "Close settings" }).click();
  const filteredErrors = process.env.UI_SMOKE_ALLOW_STATIC_API === "true"
    ? consoleErrors.filter((error) => !error.includes("/api/settings"))
    : consoleErrors;
  results.push({ viewport: viewport.name, setupVisible, draftDisclosure, managedKeyDisabled, consoleErrors: filteredErrors });
  await page.close();
}

await browser.close();
const failures = results.filter((result) => result.draftDisclosure !== 1 || !result.managedKeyDisabled || result.consoleErrors.length);
for (const result of results) console.log(`${failures.includes(result) ? "FAIL" : "PASS"} ${result.viewport} ${JSON.stringify(result)}`);
process.exit(failures.length ? 1 : 0);
