const origin = (process.argv[2] || "").replace(/\/+$/, "");
if (!origin) {
  console.error("Usage: node deploy/production/scripts/smoke-test.mjs https://your-addin-domain.example.com");
  process.exit(1);
}

const checks = [
  { name: "task pane", url: `${origin}/`, expectStatus: 200, expectText: "CTRL" },
  { name: "brain icon 32", url: `${origin}/brain-icon-32.png`, expectStatus: 200 },
  { name: "brain icon 80", url: `${origin}/brain-icon-80.png`, expectStatus: 200 },
  { name: "health", url: `${origin}/healthz`, expectStatus: 200, expectText: '"ok":true' },
  { name: "readiness", url: `${origin}/readyz`, expectStatus: 200, expectText: '"ok":true' },
  { name: "unauthenticated API rejection", url: `${origin}/api/proxy`, expectStatus: 401 },
];

let failed = false;
for (const check of checks) {
  try {
    const response = await fetch(check.url);
    const text = await response.text();
    const statusOk = response.status === check.expectStatus;
    const textOk = check.expectText ? text.includes(check.expectText) : true;
    const ok = statusOk && textOk;
    console.log(`${ok ? "PASS" : "FAIL"} ${check.name}: ${response.status} ${check.url}`);
    if (!ok) failed = true;
  } catch (error) {
    failed = true;
    console.log(`FAIL ${check.name}: ${error instanceof Error ? error.message : String(error)}`);
  }
}

process.exit(failed ? 1 : 0);
