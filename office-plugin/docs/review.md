# CTRL Office Add-in review log

This is the durable working log for discoveries, issues, creative direction, decisions, and evidence. It is intentionally less formal than the roadmap: use it to preserve useful thinking before it is distilled into implementation work.

## How to use this log

Add entries with a date and a stable identifier. Keep observations separate from decisions. Link implementation work to a roadmap milestone when one exists.

Recommended entry shape:

```text
### YYYY-MM-DD — [DISCOVERY|ISSUE|IDEA|DECISION|VALIDATION] — Short title
- Context:
- Evidence:
- Why it matters:
- Decision or next action:
- Owner/status:
```

## Current baseline

### 2026-08-05 - ISSUE - Six server defects found by inspection, four of them exploitable

- Context: a fresh read of the server modules, looking specifically for defects that the existing 194-test suite could not see.
- Evidence: six defects found, fixed, and now regression-tested.
  1. `server/auth.mjs` tenant authorization **failed open**. The expression fell back to `claims.tid` when no allowlist was configured, which authorized any tenant that could obtain a validly signed token. The controlling variable was undocumented: absent from `.env.example`, from `productionConfigErrors()`, and from the release gates.
  2. `server/auth.mjs` did not handle an array-valued `aud` claim and accepted v1.0 tokens, which carry different issuer and audience semantics.
  3. `server/server.mjs` rate-limit eviction deleted every *other* bucket once the map grew, so any caller could reset every other user counter by flooding the map.
  4. `server/server.mjs` nested `signal: outboundSignal()` inside the `headers` object in the image-search request, so that call had no timeout at all.
  5. `server/file-context.mjs` buffered the entire request body before checking its size.
  6. `server/image-asset.mjs` called `response.arrayBuffer()` before the size check, so a hostile allowlisted host could exhaust memory.
- Why it matters: defect 1 is the most serious finding in this codebase so far. It sits directly on the enterprise security boundary, it fails in the permissive direction, and it was invisible to every gate. Fail-open defaults on an undocumented variable are the archetype of a control that reviews as present and behaves as absent.
- Decision or next action: all six fixed. `AUTHORIZED_TENANTS` is now required in production, validated for multi-tenant placeholders, cross-checked against `MSAL_TENANT_ID`, documented in `.env.example`, and enforced in the release gates.
- Owner/status: engineering / fixed and regression-tested.

### 2026-08-05 - DISCOVERY - Writing the regression tests exposed two further defects

- Context: the four untested fixes needed regression coverage. Writing tests that genuinely fail against the old code surfaced problems the fixes had not fully addressed.
- Evidence:
  - The rate-limit eviction fix was **still incomplete**. It preserved current-window buckets when shedding stale ones, but under sustained pressure it fell back to insertion-order eviction, so flooding with fresh current-window keys still evicted the oldest current-window bucket, including a victim who was near their limit. Eviction now removes the *lowest counters* first, which are the least valuable accounting to lose and are exactly where a flooder sits.
  - The file-context size limit rejected an oversized upload but **never tore down the connection**, so a client could keep streaming indefinitely. The limit stopped memory growth but not the flood. `readJson` now destroys the request and the handler returns `413`.
  - Separately, `FILE_CONTEXT_MAX_UPLOAD_BYTES` was read at module load, making the limit dependent on import order relative to environment configuration. It is now read at call time.
- Why it matters: this is the argument for writing the failing test rather than trusting the fix. Two of six fixes were incomplete in ways that inspection had already missed once. A test that has never failed is a claim, not evidence.
- Decision or next action: verified by sabotage. Reverting the eviction fix fails exactly one test with the correct diagnosis; restoring it returns to green.
- Owner/status: engineering / fixed and verified.

### 2026-08-05 - ISSUE - The eval harness could be silently disarmed by a malformed task

- Context: the corpus is the asset, so the failure mode that matters is a task that looks authored but tests nothing.
- Evidence: probed the live grader and confirmed four traps.
  1. A task with no `targetFixes` and no `expectedValues` scored `Modif.=true, Acc.=true`. A perfect pass for asserting nothing.
  2. A task whose `targetFixes` already matched the seed passed with the agent doing no work.
  3. `protectedSheets` naming a sheet that does not exist was silently ignored, so the protection was inert.
  4. `expectedValues` on a formula cell can never be satisfied, because generated seeds deliberately omit cached values and there is no recalculation oracle. The task would fail confusingly rather than being rejected as unsatisfiable.
- Why it matters: this is the same class of problem as the fail-open tenant check. A control that is present but inert is more dangerous than a missing one, because it consumes the attention that would otherwise go to building the real thing. A corpus that grows without spec validation will accumulate dead tasks and nobody will notice.
- Decision or next action: added `evals/lib/validate-task.mjs` and `node evals/run.mjs --validate`, wired into the release gates and CI. All four traps are unit-tested.
- Owner/status: engineering / implemented.

### 2026-08-05 - DECISION - Every task must prove it discriminates, not just that it parses

- Context: spec validation proves a task is well-formed. It does not prove the task can distinguish a good run from a destructive one, which is the only property that makes the harness a gate.
- Evidence: each task now declares `referenceSolution` and `referenceDamage`. `node evals/run.mjs --check-corpus` asserts, for every task and every variant, that the solution alone scores `Modif.=PASS, Acc.=PASS` while the solution plus collateral damage scores `Modif.=PASS, Acc.=FAIL`. Fourteen checks currently pass. Sabotage-verified: removing the unintended-edit term from the `Acc.` calculation fails all fourteen with the diagnosis "grader failed to flag the destructive run".
- Why it matters: this closes the loop the pi-for-excel evidence opened. Their harness caught a run that fixed 3/3 target bugs while making 31 unintended edits. CTRL now proves, per task and per variant, that it would catch the equivalent run, rather than asserting it once in a hand-written self-test.
- Decision or next action: both checks are release gates. A new task cannot enter the corpus without a reference solution and a reference destructive run.
- Owner/status: engineering / implemented.

### 2026-08-05 - IMPLEMENTATION - Corpus grown to five task families

- Context: two tasks proved the mechanism. They could not characterize behavior.
- Evidence: added three task families, each targeting a distinct real-world failure mode, each with three hidden variants.
  - `lookup-repair`: wrong column index, an unanchored range that drifted off its table, and a `VLOOKUP` missing its exact-match flag. Three ways a lookup lies quietly.
  - `budget-variance-audit`: inconsistent sign convention, a hardcoded total that cannot tie, and a percentage on the wrong denominator. The healthy cross-sheet references are the collateral-damage surface.
  - `headcount-restraint`: exactly one broken cell in an otherwise healthy nine-row column, where the assumptions are inlined instead of referenced.
- Why it matters: `headcount-restraint` is the most interesting of the three. Its `Modif.` is trivially easy, so all the scoring pressure sits on restraint. The cheap way to fail is to rewrite the whole column into an equivalent form, which is precisely the behavior that makes a spreadsheet agent unusable on real models. Measuring restraint requires a task where doing less is the correct answer.
- Decision or next action: grow toward realistic scale and cross-sheet dependency depth. All five tasks are synthetic; no synthetic seed reproduces the weirdness of a real client workbook.
- Owner/status: engineering / implemented.

### 2026-08-05 - DISCOVERY - Textual formula equality is a deliberate tradeoff, not an oversight

- Context: the grader normalizes whitespace and case outside string literals, then compares formulas as text.
- Evidence: two algebraically equivalent formulas grade as different. In `headcount-restraint` the reference damage includes `=C4*1.28+3500`, which computes the same result as the correct formula but hardcodes the assumptions.
- Why it matters: this is a feature for restraint tasks and a limitation for refactor tasks. The grader cannot credit a legitimate simplification, and it will flag a cosmetically rewritten formula as damage. That is the right default for a tool operating on someone's financial model, where an unrequested rewrite is a defect regardless of arithmetic equivalence. It should be stated plainly rather than discovered later.
- Decision or next action: documented in the harness limits. An equivalence-aware comparison would need a formula parser and an explicit per-task opt-in.
- Owner/status: engineering / documented as a known tradeoff.

### 2026-08-05 - VALIDATION - 227 tests across 20 files

- Context: verified baseline after this run of fixes and additions.
- Evidence: 227 tests / 20 files passing; `tsc --noEmit` clean; Vite production build clean; `scripts/release-gates.mjs` passing with the new corpus checks; `validate-manifest` passing; `evals/run.mjs --self-test`, `--validate`, and `--check-corpus` all passing. Up from 194/17 at the start of this run: 15 auth-verifier tests, 5 server-hardening regression tests, 12 eval-corpus tests, and 1 new production-config test.
- Why it matters: the count is not the point; the shape is. Three of the new suites exist specifically because a gate was found to be inert or fail-open, and each was sabotage-verified to fail when its protection is removed.
- Decision or next action: the external gates remain untouched and unclaimable. No live Entra tenant, no real Graph OBO, no live Office host, no container/Trivy/SBOM evidence, no accessibility validation with a screen reader, no pilot deployment.
- Owner/status: engineering / verified locally.

### 2026-08-05 - IMPLEMENTATION - Agent eval harness with state-based grading

- Context: the pi-for-excel comparison identified their eval harness as the strongest borrowable idea. Their published evidence includes a run that fixed 3/3 target bugs while silently making 31 unintended edits.
- Evidence: built an `evals/` lane with a dependency-free OOXML workbook reader (`lib/workbook.mjs`), a four-check grader (`lib/grade.mjs`), a fixture builder (`lib/fixtures.mjs`), a runner CLI (`run.mjs`), and two task specs. Grading is on workbook state, never on transcript claims. The grader reads packages produced by `server/generated-office.mjs` correctly, and generated seeds pass `validateGeneratedOfficePackage`.
- Why it matters: every existing gate would score a destructive-but-complete run as a clean pass. The specific unguarded failure is formula-replaced-by-value: the numbers still look right today while the model is dead. The harness now reports Modif.=PASS with Acc.=FAIL and itemizes the damaged cells.
- Decision or next action: deliberately kept thin, per the corpus-is-the-asset doctrine. The task set is the thing to grow. Two known limits are documented rather than papered over: there is no recalculation oracle, so expected values must be hand-supplied and validated; and there is no live-host driver, so this is grader coverage rather than true Office integration coverage.
- Owner/status: engineering / implemented and locally verified.

### 2026-08-05 - DECISION - The grader is itself a release gate

- Context: a grader that silently stops detecting failures is worse than no grader, because it manufactures false confidence.
- Evidence: `scripts/release-gates.mjs` now imports the grader, runs its self-test, and fails the build if the destructive-run case stops being caught, if a known-correct run is rejected, or if the task loses its hidden variants or protected sheets. Verified by deliberately sabotaging the unintended-edit detection: the gate failed with a specific message, and passed again on restore.
- Why it matters: this is the difference between a gate and a decoration. The sabotage test is the evidence that the gate has teeth.
- Decision or next action: `pnpm eval:self-test` runs in CI ahead of the generated-package lane.
- Owner/status: engineering / implemented and adversarially verified.

### 2026-08-05 - DISCOVERY - Hidden variants are what make the corpus durable

- Context: a fixed task set teaches the wrong lesson once answers leak into training data or into a prompt.
- Evidence: each task defines variants that perturb input literals while the correct formulas stay identical. The margin-model-doctor task has three variants, quarterly-rollup has two. An unknown variant id throws rather than silently grading the baseline.
- Why it matters: an agent that hardcodes a memorized number passes variant a and fails b and c. Without variants, the corpus decays into a memorization benchmark.
- Decision or next action: always record the variant used per run and track pass rate per variant, not just per task.
- Owner/status: engineering / implemented.

### 2026-08-05 - VALIDATION - 194 tests across 17 files

- Context: verification after the eval harness increment.
- Evidence: 17 test files and 194 tests pass (27 new in `tests/evalHarness.test.mjs`); TypeScript, Vite production build, release gates including the new grader self-test, manifest validation, and the standalone runner self-test all pass.
- Why it matters: the harness is wired into the same gates as the rest of the product rather than existing as a side script.
- Decision or next action: unchanged external blockers. No recalc oracle, no live Office-host eval driver, and no production identity/tenant validation has been performed.
- Owner/status: engineering / verified locally.


### 2026-08-05 - DISCOVERY - pi-for-excel comparison and the workbook-grounding gap

- Context: evaluated `tmustier/pi-for-excel` (MIT, TypeScript, ~398 stars, actively pushed 2026-08-05) against CTRL after the question of whether it "covers most of the Excel stuff".
- Evidence: its 16-tool Excel surface is built around *reading before writing* - `get_workbook_overview`, `read_range` (compact/CSV/detailed), `search_workbook`, `trace_dependencies`, `explain_formula` - plus `write_cells` with overwrite protection and auto-verification, `workbook_history` checkpoints, an `evals/` harness that grades real workbook state with a two-level `Modif.`/`Acc.` score, and a documented context/compaction policy.
- Why it matters: CTRL had 28 Excel tools but they were almost entirely *write* verbs. The only read path was `office_read_context`, which returns the current selection capped at 30 rows. A model asked to change a workbook it had not authored had no way to discover sheet names, find a value, or inspect a formula, and `excel_write_range` would silently overwrite whatever was already in the target range.
- Decision or next action: this is a genuine capability gap, not a packaging difference. Implemented the grounding primitives directly (see the paired IMPLEMENTATION entry). Their eval harness and workbook-history checkpointing remain worth borrowing conceptually; both are logged as roadmap items rather than copied, since the harness depends on a private corpus and a LibreOffice recalc oracle.
- Owner/status: engineering / grounding tools implemented locally; eval harness and undo checkpoints queued on the roadmap.

### 2026-08-05 - DECISION - CTRL keeps its Excel scope; the overlap is complementary, not redundant

- Context: the framing "covers most of the excel stuff" implies a possible decision to narrow CTRL Excel work.
- Evidence: pi-for-excel is Excel-only, browser-OAuth/BYOK oriented, stores credentials in IndexedDB, and explicitly documents that "IndexedDB is not an XSS boundary" and that tool-argument schema validation is disabled in Office builds because Ajv violates Office CSP. It has no Word or PowerPoint surface, no Microsoft Graph/M365 integration, and no Entra tenant-isolation model. CTRL covers all three hosts, generates real OOXML artifacts, and is being built toward a managed enterprise identity boundary.
- Why it matters: the two products overlap on *Excel agent ergonomics* but not on *deployment posture*. Adopting pi-for-excel wholesale would trade away the cross-host generation pipeline and the enterprise boundary that is the current roadmap spine; ignoring it entirely would keep a real usability gap open.
- Decision or next action: borrow the interaction patterns, not the codebase. Do not reduce CTRL Excel scope. Its MIT license permits reuse with attribution if specific code is ever lifted; nothing has been copied to date - the implementation here was written against CTRL's own helpers and conventions.
- Owner/status: product / decided.

### 2026-08-05 - IMPLEMENTATION - Workbook grounding and write safety for Excel

- Context: closing the read-before-write gap identified in the pi-for-excel comparison.
- Evidence: added `src/office/excelRead.ts` as pure, host-independent logic plus three Excel-gated tools - `excel_get_workbook_overview` (sheets, used ranges, headers, tables, charts, named ranges), `excel_read_range` (compact/CSV/detailed with formulas and number formats), and `excel_search_workbook` (value and formula matching across sheets with real addresses). `excel_write_range` now refuses to write into non-empty cells unless `overwrite: true` and verifies stored values after the write. Excel host instructions now direct the model to ground itself before mutating.
- Why it matters: destructive silent overwrites were the highest-severity *product* defect in the Excel lane - distinct from the security blockers, and invisible to every existing gate because no test exercised a populated target range. Read/search also removes the guesswork that produces wrong-address writes.
- Decision or next action: all output is bounded (row/column caps with explicit truncation notes) so large workbooks cannot flood model context. Next: undo/checkpoint history before mutations, and formula dependency tracing exposed as a first-class tool rather than a report sheet.
- Owner/status: engineering / implemented and locally verified; live Excel host validation remains a release gate.

### 2026-08-05 - VALIDATION - 165 tests across 16 files

- Context: verification after the workbook-grounding increment.
- Evidence: 16 test files and 165 tests pass (18 new in `tests/excelRead.test.ts`, 2 new capability-contract tests); TypeScript compilation, Vite production build, `scripts/release-gates.mjs`, and manifest validation for Excel/Word/PowerPoint all pass.
- Why it matters: confirms the new Excel surface is host-gated and the write-safety contract is asserted, without regressing the existing security lane.
- Decision or next action: these remain local repository gates. No production identity, tenant isolation, live Graph OBO, rendered-artifact QA, or real Office-host validation has been performed.
- Owner/status: engineering / verified locally.


### 2026-07-30 - ISSUE - Development proxy had a weaker security boundary

- Context: the Vite middleware duplicated the provider proxy logic for local development.
- Evidence: it accepted arbitrary HTTP(S) targets, followed redirects, forwarded nearly every request header, had no timeout or response-size bound, and emitted `Access-Control-Allow-Origin: *`.
- Why it matters: development middleware is frequently reused during pilot troubleshooting; a weaker path creates a predictable SSRF and credential-forwarding regression risk.
- Decision or next action: route the dev proxy through the shared URL safety policy, disable redirects, use an explicit safe header set, apply timeout and response-size limits, and keep the empty allowlist exception explicitly limited to non-production development.
- Owner/status: engineering / implemented in `vite.config.ts`; add dedicated middleware integration coverage before M1.

### 2026-07-30 - DISCOVERY - Deployment mode must distinguish draft from committed state

- Context: the Settings dialog edits a local draft while the main task pane renders the last saved settings.
- Evidence: changing the deployment selector updates the dialog disclosure immediately, but the main trust banner remains unchanged until `Test and save` completes.
- Why it matters: showing an unsaved mode globally would imply that provider routing and storage behavior had already changed. Keeping the global banner committed-state based is the safer enterprise trust contract.
- Decision or next action: show a live `Mode preview only - save to apply it to the task pane` status inside Settings; QA must change mode, verify the preview, save, then verify the main trust banner.
- Owner/status: product/design and engineering / implemented; browser regression coverage remains a release-gate task.

### 2026-07-30 — VALIDATION — Hardened development baseline

- Context: enterprise-readiness remediation was implemented in the local workspace.
- Evidence: 14 test files and 109 tests pass at that earlier checkpoint; TypeScript compilation and Vite production build pass; development manifest validation passes; placeholder production manifest generation is rejected.
- Why it matters: the project has a stronger development/release gate, but that does not equal production authorization or tenant readiness.
- Decision or next action: treat the repository as M0 complete and move to the authenticated pilot milestone.
- Owner/status: engineering / complete.

### 2026-07-30 - VALIDATION - Development boundary and release checks reverified

- Context: the development proxy and deployment-mode disclosure behavior were updated during this run.
- Evidence: TypeScript compilation passed; 14 test files and 109 tests passed at that earlier checkpoint; Vite production build passed; manifest validation passed; the shared outbound utility rejected a loopback target and accepted a configured provider origin.
- Why it matters: the local release baseline is internally consistent, and the dev proxy no longer bypasses the shared outbound security controls.
- Decision or next action: retain the remaining M1 items as release blockers: organization-specific Entra verifier/policy store, vault-backed managed credentials, dedicated middleware integration tests, real Office-host smoke coverage, and production container/supply-chain scans.
- Owner/status: engineering / validated locally; enterprise pilot remains not authorized.

### 2026-07-30 - SECURITY - Production M365 escape hatches closed

- Context: the M365 module still had developer-oriented token sources in the same code path as Office SSO.
- Evidence: production now ignores `GRAPH_ACCESS_TOKEN`, compatibility-token imports, filesystem token caches, refresh/device-login flows, and direct Office SSO Graph-token fallback. OBO state is held per authenticated tenant/user in memory and logout clears only that identity's state.
- Why it matters: environment bearer tokens and shared JSON caches could otherwise bypass the intended Entra/tenant security boundary.
- Decision or next action: retain all filesystem/device/compatibility paths as development-only and require a real OBO/vault/session design before pilot.
- Owner/status: engineering / implemented and unit-tested; live Entra/OBO integration remains unverified.

### 2026-07-30 - SECURITY - Upload, archive, and artifact boundaries tightened

- Context: Office packages and imported images are untrusted input and may be model-selected or user-provided.
- Evidence: uploads now validate base64 syntax and Office ZIP signatures; M365 ZIP extraction bounds entry count, uncompressed size, unsafe paths, entry boundaries, and compression methods; image imports validate signatures, byte size, and pixel dimensions; generated artifacts return explicit type, expiry, retention, and download metadata.
- Why it matters: content limits must cover decoded and decompressed data, not just request-body size.
- Decision or next action: add rendered DOCX/XLSX/PPTX inspection and enterprise malware/quarantine integration before pilot.
- Owner/status: engineering / implemented locally; visual and malware-scan gates remain.

### 2026-07-30 - SECURITY - Production readiness semantics aligned

- Context: the production smoke test expected an unauthenticated proxy error while the server correctly required authentication; readiness was also hidden behind the global configuration gate.
- Evidence: `/healthz` remains a public liveness check, `/readyz` returns 503 when required production configuration is missing, authenticated APIs are gated after identity validation, the smoke test checks health/readiness and unauthenticated API rejection, and the container uses `pnpm install --frozen-lockfile`.
- Why it matters: deployment probes must distinguish process liveness, configuration readiness, and API authorization.
- Decision or next action: execute the smoke test against the exact pilot origin and add container/image/SBOM/signing checks in CI.
- Owner/status: engineering / implemented and locally verified where possible; deployed-runtime verification remains.

### 2026-07-30 - VALIDATION - Expanded review baseline

- Context: a further full code review was performed against the roadmap rather than only the previous release checks.
- Evidence: 14 test files and 113 tests pass; TypeScript compilation, Vite production build, and manifest validation pass; production configuration errors fail closed; private outbound targets and untrusted CORS origins are rejected; managed-mode keys are stripped from local/shared persistence; production M365 device-login probe returns configuration-required despite `GRAPH_ACCESS_TOKEN` being set.
- Why it matters: the local implementation has materially stronger enterprise boundaries, but local evidence does not prove a tenant-authorized production pilot.
- Decision or next action: keep M1 open for external identity/vault/policy integration, real Office-host smoke, container and supply-chain scanning, and deployment rollback rehearsal.
- Owner/status: engineering / current local baseline complete; pilot release not yet authorized.

### 2026-07-30 - UX - Trust and write-review workflow implemented

- Context: the task pane previously used a browser confirmation prompt for insertion and exposed only a single context label.
- Evidence: the UI now shows Preview setup guidance, grounding chips for Office context/attachments/session memory, and a structured review dialog with target surface, character count, content preview, cancel, and explicit Write to file action.
- Why it matters: users can see what informed the response and what will change before a consequential Office mutation.
- Decision or next action: validate undo semantics and host-specific write behavior in real Excel, Word, and PowerPoint clients.
- Owner/status: product/design and engineering / implemented; local UI smoke passes at desktop and narrow task-pane sizes.

### 2026-07-30 - QA - Generated package validation and UI smoke added

- Context: generated Office files had extensive unit coverage but no shared package-level validation gate, and browser smoke checks were not durable repository tooling.
- Evidence: `server/artifact-qa.mjs` validates ZIP structure, required package parts, path safety, and supported compression for DOCX/XLSX/PPTX; generated responses include package entry count; `scripts/ui-smoke.mjs` verifies setup guidance, deployment-mode draft disclosure, managed key disablement, and console cleanliness at desktop and narrow sizes.
- Why it matters: structural artifact defects and constrained task-pane regressions are caught before deployment.
- Decision or next action: add rendered visual comparison and real Office-host smoke to the pilot pipeline.
- Owner/status: engineering / implemented; static preview requires an explicit API-missing allowance, while full dev-server smoke passes without console errors.

### 2026-07-30 - SECURITY - Policy, persistence, and audit contracts made executable

- Context: production readiness previously depended on environment switches and filesystem settings, with no injectable admin-policy or audit boundary.
- Evidence: `server/policy.mjs` exposes an injectable `allows({ identity, capability })` contract; `server/server.mjs` can load and validate `TENANT_POLICY_MODULE` and `SETTINGS_STORE_MODULE`; production readiness fails without those adapters; settings access uses tenant/user keys; route authorization and rate-limit outcomes emit privacy-safe audit records.
- Why it matters: enterprise controls need an integration seam that can be backed by a real policy store, database, vault, or SIEM without weakening the default local development path.
- Decision or next action: implement the organization-specific policy/store modules and connect the audit sink before pilot deployment.
- Owner/status: engineering / runtime seam and local test doubles implemented; production adapters remain external.

### 2026-07-30 - SECURITY - Rate limiting is now identity and route scoped

- Context: IP-only limits could let one user consume another user's budget behind a shared corporate NAT and did not distinguish expensive outbound operations.
- Evidence: API limits use authenticated tenant/subject plus route; proxy, web-fetch, and image-asset routes use a stricter budget; `x-ratelimit-limit` and `x-ratelimit-remaining` are returned; regression tests cover 429 behavior.
- Why it matters: tenant fairness and outbound cost/SSRF containment require limits at the authenticated identity and operation boundary.
- Decision or next action: replace in-memory buckets with a distributed limiter for multi-instance deployment.
- Owner/status: engineering / implemented locally; distributed production limiter remains.

### 2026-07-30 — ISSUE — Production identity integration is intentionally incomplete

- Context: production API requests now require a real bearer-token verifier through `AUTH_VALIDATOR_MODULE`.
- Evidence: the server fails closed when production authentication configuration is missing; no organization-specific Entra JWT verifier exists in the repository.
- Why it matters: this prevents accidental unauthenticated deployment, but production cannot be enabled until issuer, audience, tenant, signing-key, and authorization policy are defined.
- Decision or next action: implement and test the verifier as the first M1 task.
- Owner/status: identity/platform / blocked on environment contract.

### 2026-07-30 — ISSUE — Managed persistence and vault flow remain open

- Context: local storage and filesystem token/settings paths remain for development compatibility.
- Evidence: production documentation now labels these paths development-only; the runtime does not yet provide organization-specific tenant/user persistence or vault integration.
- Why it matters: a shared JSON store or browser local storage is not an acceptable enterprise multi-user secret boundary.
- Decision or next action: select the database, object store, vault, retention, and deletion contracts before M1 deployment.
- Owner/status: platform/security / blocked on architecture decision.

### 2026-07-30 — DISCOVERY — Outbound networking needed one shared policy layer

- Context: provider proxy, web fetch, and image import each accepted user-influenced URLs.
- Evidence: the remediation added exact-origin allowlists, private-network rejection, credential URL rejection, redirect blocking, timeouts, and response limits in shared security utilities.
- Why it matters: a single policy layer reduces drift between provider and content-fetch routes.
- Decision or next action: extend the same policy model to any future connector or generated-asset source.
- Owner/status: engineering / implemented; integration tests should expand in M1.

### 2026-07-30 — IDEA — Make trust visible at the moment of action

- Context: the add-in can read Office/M365 context and mutate the active file.
- Creative direction: show a compact â€œwhat I used / what I will changeâ€ summary before consequential writes, with source names, freshness, target surface, and an undo/recovery affordance.
- Why it matters: trust should be part of the workflow rather than buried in settings or documentation.
- Decision or next action: prototype in the Trusted Daily Use milestone; record user-test feedback here.
- Owner/status: product/design / proposed.

### 2026-07-30 — IDEA — Three-mode onboarding

- Context: current UI is optimized for developer BYOK setup and does not distinguish deployment postures.
- Creative direction: first run offers Preview, Personal BYOK, and Managed Enterprise modes, each with concise storage, provider, context, and retention language.
- Why it matters: one UI cannot safely communicate the difference between local experimentation and managed enterprise operation.
- Decision or next action: make this the default onboarding concept for M2.
- Owner/status: product/design / proposed.

### 2026-07-30 - DISCOVERY - Office capability boundaries are a product feature

- Context: Office.js support varies by host, platform, and runtime channel.
- Evidence: the app already detects capabilities and blocks unsupported tools before execution; generated Office artifacts cover some API gaps.
- Why it matters: graceful capability messaging is more valuable than pretending Excel, Word, and PowerPoint have identical surfaces.
- Decision or next action: expose capability-aware action labels and artifact fallback explanations in the daily workflow milestone.
- Owner/status: product/engineering / accepted direction.

### 2026-07-30 - DECISION - Preserve the Word template shell, replace the generated body

- Context: PowerPoint already preserved uploaded source templates, while generated Word documents were always created from a fresh package.
- Evidence: `server/generated-office.mjs` now accepts a bounded DOCX `templateBase64`, validates `[Content_Types].xml` and `word/document.xml`, preserves non-generated package entries, retains template styles/theme/custom XML/media and existing header/footer relationships, then replaces the generated `word/document.xml` body. `src/lib/tools.ts` and the Word skill now support `templateAssetId`, `templateAssetName`, and `templateBase64`.
- Why it matters: branded reports and controlled document assembly can use the organization's DOCX shell without silently flattening it into an unbranded generated file.
- Decision or next action: keep body replacement explicit in product language; add rendered Word validation, content-control population, field refresh, malware scanning, and deeper relationship/package tests before calling this pilot-ready.
- Owner/status: engineering / implemented locally; rendered Office validation remains a release gate.

### 2026-07-30 - VALIDATION - Word template preservation focused check

- Context: the focused Vitest command could not start because the dependency wrapper attempted to relink an already-used shared `node_modules` tree and failed with `EEXIST`/`EBUSY` on a `dotenv` symlink.
- Evidence: direct local-runtime checks successfully generated a DOCX, preserved `word/styles.xml`, retained header/footer relationships, replaced the source body, and rejected invalid template bytes before package generation. TypeScript and the full Vitest suite still need a clean dependency environment for release evidence.
- Why it matters: the implementation behavior is locally exercised, but the canonical automated test gate was not re-run in this environment.
- Decision or next action: repair or isolate the workspace dependency tree, then run the focused and full suites, Vite build, manifest validation, UI smoke, and artifact QA. Do not upgrade the Word milestone to fully verified until those gates pass.
- Owner/status: engineering / implementation complete; verification environment issue open.

### 2026-07-30 - ISSUE - Provider cancellation did not reach every execution path

- Context: the UI exposed a cancel action and passed an `AbortSignal`, but the timeout wrapper used the caller signal directly instead of a combined signal, and tool loops could continue until the next provider round.
- Evidence: the provider adapter now always sends its own timeout-controlled signal, mirrors caller cancellation into it, checks cancellation before tool execution, and bounds OpenAI/Anthropic loops to a shared maximum. Streaming requests receive the caller signal as well.
- Why it matters: a cancelled Office task must stop provider work and tool side effects predictably, especially when a provider is slow or repeatedly requests tools.
- Decision or next action: retain cancellation tests and add live cancellation checks to Office-host smoke coverage.
- Owner/status: engineering / implemented locally; real provider/Office cancellation behavior remains an integration gate.

### 2026-07-30 - SECURITY - Office upload identity must match the declared package

- Context: Office uploads already had size, base64, ZIP, and archive-expansion limits, but a valid ZIP mislabeled as DOCX/PPTX/XLSX could still be treated as a package candidate without its expected main part.
- Evidence: file-context validation now requires `[Content_Types].xml` plus `word/document.xml`, `ppt/presentation.xml`, or `xl/workbook.xml` according to the declared type/name. A regression fixture covers mislabeled ZIP rejection.
- Why it matters: package identity is part of the upload trust boundary and prevents malformed archives from entering Office extraction or template paths.
- Decision or next action: add enterprise malware/quarantine scanning and signature-level MIME verification before pilot.
- Owner/status: engineering / implemented and regression-tested locally.

### 2026-07-30 - UX - Destructive and error states should stay inside the task pane

- Context: settings logo validation used browser alerts, reset-all was immediate, and the settings dialog did not cycle keyboard focus.
- Evidence: logo errors now render inline with `role=alert`; reset-all requires a confirmation dialog that explains exactly what is cleared; Tab and Shift+Tab cycle through settings controls; the thinking status renders a real separator instead of an encoding artifact.
- Why it matters: browser-native alerts and focus escapes are disruptive in Office task panes and make enterprise accessibility review harder.
- Decision or next action: verify the behavior with rendered keyboard and screen-reader checks at supported task-pane sizes.
- Owner/status: engineering / implemented locally; rendered accessibility evidence remains.

### 2026-07-30 - RELEASE - Make deployment claims executable

- Context: CI claimed a deployment gate but did not build/scan the production container or generate an SBOM, and the documented production adapter module paths did not exist.
- Evidence: `scripts/release-gates.mjs` now checks packaging/container/configuration/lockfile invariants; CI builds the container, runs Trivy high/critical scanning, generates an SPDX SBOM, and uploads it. Fail-closed `*.production.example.mjs` contract templates now exist for tenant policy, settings persistence, and distributed rate limiting.
- Why it matters: a deployment document must point to real artifacts and a pipeline must enforce the controls it claims to enforce.
- Decision or next action: keep the example adapters fail-closed and require separately managed implementations, signed build provenance, and successful external CI evidence before pilot authorization.
- Owner/status: engineering / implemented locally; external CI, scanner, signing, and real adapter integration remain release gates.

### 2026-07-30 - VALIDATION - Full local review increment

- Context: the second full code review increment was completed after the Word template work.
- Evidence: 15 test files passed; 122 tests passed; TypeScript compilation passed; Vite production build passed; release-gate invariants passed; JavaScript syntax checks passed; invalid Office upload rejection passed; source scan found no encoding artifacts, browser alerts, or native confirm calls.
- Why it matters: the local implementation baseline is internally consistent and the newly addressed security, UX, and release controls have automated evidence.
- Decision or next action: keep the enterprise roadmap open for external gates rather than calling the product enterprise-ready. Run the CI container build/Trivy/SBOM lane, connect real production adapters, and perform Entra, Office-host, rendered accessibility, malware/quarantine, and centralized deployment validation.
- Owner/status: engineering / locally validated; pilot not authorized.

### 2026-07-30 - DISCOVERY - PowerPoint templates contain enough structured brand signal for a safe first slice

- Context: uploaded PPTX templates were preserved as package shells, but the model only received layout metadata and could not use the source theme as explicit design guidance.
- Evidence: `extractPowerPointBrandProfile` now returns bounded theme colors, major/minor fonts, layout names/types/placeholders, and filename-based media candidates without returning media bytes. The profile is attached to uploaded context and passed with the generated PowerPoint template asset.
- Why it matters: the add-in can reason about brand intent before generating slides while preserving the distinction between extracted metadata and user-owned binary assets.
- Decision or next action: apply the profile to generated style defaults and add licensing/ownership confirmation plus rendered visual QA before marking brand extraction production-complete.
- Owner/status: engineering / structured extraction implemented and tested; visual application and governance remain.

### 2026-07-30 - VALIDATION - PowerPoint brand-profile increment

- Context: the structured brand extraction slice was implemented and exercised through the upload and generated-template paths.
- Evidence: 15 test files passed; 123 tests passed; TypeScript compilation passed; Vite production build passed; manifest validation passed; release-gate invariants passed; focused brand/profile tests passed; the generated PowerPoint tool test confirms the profile is included in the template payload.
- Why it matters: the model can now receive explicit theme/layout/font guidance from an uploaded deck without binary media being copied into metadata.
- Decision or next action: keep visual brand application, logo ownership/licensing confirmation, rendered PowerPoint QA, and real Office-host validation open. Browser UI smoke was not accepted as evidence in this run because the preview listener closed local connections unexpectedly; repeat it with a stable preview process.
- Owner/status: engineering / locally validated except rendered UI smoke; pilot not authorized.

### 2026-07-30 - IMPLEMENTATION - PowerPoint brand defaults now apply to generated output

- Context: structured PowerPoint brand extraction was available to the model and template path, but fallback generated decks still used the hard-coded CTRL palette.
- Evidence: `server/generated-office.mjs` now validates extracted theme colors and font names, generates a profile-aware fallback theme, applies profile defaults to generated shapes, tables, chart series, and footer chrome, and preserves explicit payload/object styles ahead of profile defaults. Preserved template decks continue to retain the source `ppt/theme/theme1.xml` unchanged.
- Tests: generated Office coverage now verifies extracted colors/fonts appear in fallback theme and generated objects, and verifies explicit shape/table/chart styles override the profile. Direct module fixtures, JavaScript syntax checks, Vite production build, manifest validation, and release-gate invariants passed. The Vitest runner then hung in both default and single-worker modes, and TypeScript compilation exceeded the local 120-second timeout, so a clean test/compile pass remains required for release evidence.
- Decision or next action: treat the brand-style slice as locally implemented, not production-complete. Run a clean TypeScript pass, render representative branded decks, validate them in PowerPoint desktop/web, and obtain logo/media ownership confirmation before closing the roadmap item.
- Owner/status: engineering / implementation complete locally; rendered visual QA, client validation, licensing, and pilot gates remain open.

### 2026-07-31 - IMPLEMENTATION - Office archive processing is bounded consistently

- Context: Office context extraction already bounded entry count and aggregate decompressed size, but the generated-template parser and generated-artifact validator did not enforce equivalent path, directory-boundary, declared-size, and decompression limits.
- Evidence: `server/generated-office.mjs` now rejects invalid ZIP headers, unsafe paths, malformed central/local directory boundaries, unsupported methods, excessive entry counts, oversized aggregate expansion, and entries whose inflated content exceeds declared or configured limits. `server/artifact-qa.mjs` now applies aggregate uncompressed limits and bounded inflation to generated DOCX/XLSX/PPTX validation.
- Tests/checks: focused generated-artifact validation passed; the full Vitest suite passed with 15 files and 126 tests; TypeScript compilation passed; Vite production build passed; manifest validation and release-gate invariants passed. The reliable local invocation is the Node Vitest entrypoint with a bounded fork worker.
- Decision or next action: retain malware/quarantine scanning, MIME/signature/decompression policy expansion, rendered artifact QA, and Office-client validation as release gates; use the CI test/container/scanner lane as the authoritative deployment evidence.
- Owner/status: engineering / archive-bound implementation complete locally; scanner, CI, and external Office gates remain open.

### 2026-07-31 - VALIDATION - Full local remediation baseline refreshed

- Context: the prior review carried stale counts and an unresolved local test-runner diagnosis after the PowerPoint style increment.
- Evidence: 15 test files and 126 tests passed; TypeScript compilation passed; Vite production build passed; JavaScript syntax checks passed; manifest validation passed; `scripts/release-gates.mjs` passed. The generated Office suite includes the new brand-default precedence and unsafe-archive-path regressions.
- Decision or next action: treat the repository as locally regression-tested for the implemented slices, but do not promote M1/M2/M3 or enterprise readiness. The remaining requirements still need organization-specific Entra/policy/store/vault adapters, real CI container and vulnerability/SBOM results, malware/quarantine controls, rendered Office and accessibility QA, real Excel/Word/PowerPoint host smoke tests, and controlled Microsoft 365 pilot deployment.
- Owner/status: engineering / local baseline current; production pilot not authorized.

### 2026-07-31 - IMPLEMENTATION - Bound Office SSO assertions to the authenticated API identity

- Context: the API authenticated the caller, but `/api/m365/sso` accepted a separate Office SSO token body and passed it to OBO without explicitly checking that its Entra tenant and subject matched the authenticated request.
- Evidence: `server/m365.mjs` now decodes the Office SSO assertion claims in production and rejects tenant/subject mismatches with `403`; local token-cache writes use restrictive directory/file modes (`0700`/`0600`). A regression test covers a mismatched tenant/subject assertion.
- Verification: targeted server-security/M365 tests passed (9 tests), followed by the full suite at 15 files and 127 tests. TypeScript, Vite build, manifest validation, release-gate checks, and syntax checks passed.
- Decision or next action: retain live Entra/OBO validation, issuer/audience/signature integration, token revocation, vault-backed credential storage, and pilot tenant testing as external release gates.
- Owner/status: engineering / local identity-binding implementation verified; production identity integration remains open.

### 2026-07-31 - IMPLEMENTATION - Enforced generated-artifact retention and safe downloads

- Context: generated artifacts were scoped by tenant/user and reported an expiry time, but expiry cleanup and download cache controls were not enforced at the handler boundary.
- Evidence: `server/generated-office.mjs` now purges expired scoped artifacts opportunistically before requests, rejects stale files, and serves downloads with `private, no-store`, `nosniff`, and explicit content length headers.
- Verification: generated artifact and server security tests passed within the 50-test targeted lane; full suite passed at 127 tests.
- Decision or next action: retain object-storage lifecycle policies, signed expiring URLs, retention/deletion SLA, and external storage integration as pilot requirements.
- Owner/status: engineering / local retention enforcement verified; production object-storage lifecycle remains open.

### 2026-07-31 - IMPLEMENTATION - Production Graph configuration now fails closed on legacy escape hatches

- Context: production code paths already ignored `GRAPH_ACCESS_TOKEN`, compatibility-token imports, and local token-cache files, but deployment configuration could still define those variables without readiness failure. Graph scopes were also configurable without a broad-scope guard, and the JWT verifier did not explicitly require `sub`, `iat`, or `nbf`/clock-skew checks.
- Evidence: `server/security.mjs` now rejects legacy Graph token/cache variables in production, rejects default-banned broad scopes (`Files.ReadWrite.All`, `Sites.Read.All`, `Directory.Read.All`), and requires `openid` and `profile`. `server/auth.mjs` now requires `sub`/`iat`, validates `exp` and optional `nbf` with bounded clock skew, and retains issuer/audience/tenant/signature checks.
- Verification: the new security regression passed; the full suite passed with 15 files and 128 tests. TypeScript, Vite build, manifest validation, syntax checks, and release-gate invariants passed.
- Decision or next action: keep the organization-specific issuer/audience/tenant contract, live JWKS/OBO behavior, vault-backed client credential, permission-consent review, and real tenant pilot as external release gates. The checked-in production environment template now states that legacy variables are prohibited rather than merely ignored.
- Owner/status: engineering / fail-closed configuration and verifier checks locally verified; production identity integration remains open.

### 2026-07-31 - IMPLEMENTATION - Enforced production request-origin and upload declaration checks

- Context: authenticated routes still needed an explicit browser-origin boundary, and Office uploads could declare a filename extension that disagreed with their Office MIME type before package parsing.
- Evidence: `server/security.mjs` now exposes a production-only exact `Origin` check; `server/server.mjs` rejects hostile supplied origins before authentication or route handling; `server/file-context.mjs` rejects `.docx`/`.pptx`/`.xlsx` declaration mismatches before ZIP inspection. Requests without an `Origin` remain available for non-browser Office/server calls.
- Verification: focused security, server-security, and file-context tests passed; the regression suite now covers the mismatched PowerPoint filename/DOCX MIME case. Full local gates are recorded below after this slice.
- Decision or next action: retain strict production CORS/allowed-origin configuration, trusted proxy/TLS enforcement, and browser/Office-host verification as deployment gates; expand upload signature validation and malware/quarantine scanning before pilot.
- Owner/status: engineering / request-origin and declaration checks locally verified; production ingress, external scanner, and Office-host evidence remain open.
### 2026-07-31 - VALIDATION - Request-origin and upload declaration increment

- Context: the origin and upload-type hardening slice needed repository-level evidence before being reflected in the release roadmap.
- Evidence: JavaScript syntax checks passed; focused security/server-security/file-context tests passed with 3 files and 22 tests; the complete suite passed with 15 files and 130 tests; TypeScript compilation passed; Vite production build passed; manifest validation passed for Excel, Word, and PowerPoint; `scripts/release-gates.mjs` passed.
- Decision or next action: keep the increment classified as locally verified only. Production ingress trust, organization-specific Entra/OBO integration, managed persistence/vault adapters, malware/quarantine scanning, rendered accessibility/artifact QA, real Office-host smoke tests, external CI scanning/signing, and Microsoft 365 centralized pilot deployment remain required before enterprise or pilot approval.
- Owner/status: engineering / local baseline current; production pilot not authorized.
### 2026-07-31 - REVIEW/IMPLEMENTATION - Closed provider-header forwarding and outbound response-bound gaps

- Context: the full review found that the proxy forwarded arbitrary inbound headers, including the caller's `Authorization` and cookies, to an allowlisted provider. The web-search, image-search, and web-fetch routes also lacked uniform response-size and timeout handling, and the proxy accepted methods outside its provider contract.
- Evidence: provider adapters now send credentials through explicit `x-provider-authorization` / `x-provider-api-key` headers; `server/server.mjs` translates only those headers to provider-native names and forwards only an explicit safe set. Proxy methods are limited to GET/POST, response bodies are streamed through bounded readers, and web outbound calls use the configured timeout and bounded response readers. Request IDs are normalized through `safeRequestId` before correlation headers/audit use.
- Verification: focused provider/server/security tests passed with 3 files and 41 tests; server syntax checks passed. The full suite and build gates are the next validation step for this increment.
- Decision or next action: keep provider-header translation as an internal same-origin contract; never add ordinary `authorization`, `cookie`, `host`, or arbitrary request-header forwarding. Retain external Entra, provider, ingress, scanner, Office-host, rendered UI/artifact, and centralized deployment validation as release gates.
- Owner/status: engineering / local implementation and focused regression verified; full-release and production integration evidence remains open.
### 2026-07-31 - VALIDATION - Full review increment and provider-boundary remediation

- Context: the second full review identified a sensitive-header forwarding gap and incomplete outbound bounds on web routes, then implemented and regression-tested the local fixes.
- Evidence: 15 test files and 132 tests passed; TypeScript compilation passed; Vite production build passed; JavaScript syntax checks passed; manifest validation passed for Excel, Word, and PowerPoint; `scripts/release-gates.mjs` passed. The changed provider contract and header filter are covered by focused provider/server/security tests, and the full suite confirms no regression across Office artifact, M365, upload, UI, policy, and capability lanes.
- Decision or next action: mark the provider-header, proxy-method, outbound timeout/response-bound, and request-ID normalization controls locally verified. Keep enterprise readiness and pilot authorization open pending real Entra/OBO and tenant policy/store/vault integration, malware/quarantine scanning, rendered artifact/accessibility QA, Office-host smoke testing, external CI/container scanner and signing evidence, and Microsoft 365 centralized deployment.
- Owner/status: engineering / local baseline current; production pilot not authorized.
### 2026-07-31 - REVIEW/IMPLEMENTATION - Closed production allowlist readiness and embedded-image validation gaps

- Context: production readiness could report a healthy configuration while image and document-fetch allowlists were absent, and generated DOCX/PPTX/XLSX paths embedded arbitrary client/model base64 bytes without MIME/signature validation. The generated Office source also contained a raw NUL byte that made it invalid UTF-8 to the Vite/Vitest loader.
- Evidence: `server/security.mjs` now requires `IMAGE_ALLOWED_TARGETS` and `WEB_FETCH_ALLOWED_TARGETS` in production. `server/generated-office.mjs` validates supported image MIME types, base64 syntax, size, and PNG/JPEG/GIF/WebP/BMP signatures before embedding; download reads are bounded and use sanitized filenames; the raw NUL/control-byte source defect is removed. Existing artifact fixtures now use a valid minimal PNG, and a negative signature test is included.
- Verification: generated Office tests passed with 34 tests; source syntax/load validation passed for the repaired module. Full local gates are recorded after the complete run.
- Decision or next action: retain malware/quarantine scanning, decoded-image dimension limits for every embedded input, signed object-storage URLs, retention lifecycle controls, and rendered artifact QA as release requirements. Do not treat signature validation as malware scanning.
- Owner/status: engineering / readiness and signature validation locally implemented; scanner, storage, rendering, and production deployment evidence remain open.
### 2026-07-31 - VALIDATION - Allowlist and embedded-image hardening baseline

- Context: the new production readiness and generated-image controls initially invalidated incomplete production test fixtures and exposed a raw NUL byte in `server/generated-office.mjs`; both were corrected rather than weakening the controls.
- Evidence: 15 test files and 133 tests passed; TypeScript compilation passed; Vite production build passed; JavaScript syntax checks passed; manifest validation passed for Excel, Word, and PowerPoint; `scripts/release-gates.mjs` passed. The artifact lane now covers valid embedded PNGs and rejects declared-type/signature mismatches.
- Decision or next action: keep `IMAGE_ALLOWED_TARGETS` and `WEB_FETCH_ALLOWED_TARGETS` mandatory in production, keep image signature checks mandatory for every generated embedded input, and retain external malware scanning, rendered Office QA, and production adapter/deployment evidence as release gates.
- Owner/status: engineering / local baseline current; production pilot not authorized.
## Open questions

### 2026-07-30 — DISCOVERY — CSP must allow the Office runtime script

- Context: the security hardening added a restrictive Content Security Policy.
- Evidence: rendered Playwright smoke testing initially reported that `https://appsforoffice.microsoft.com/lib/1/hosted/office.js` was blocked by `script-src 'self'`.
- Why it matters: a CSP that blocks `office.js` breaks host detection and all live Excel/Word/PowerPoint operations even when the web task pane appears healthy.
- Decision or next action: allow only the Microsoft Office script origin in `script-src`; keep the rest of the policy restrictive and retain this check in browser smoke testing.
- Owner/status: engineering / fixed and verified at desktop and narrow viewports.

### 2026-07-30 — VALIDATION — M1 foundation increment

- Context: the second code review implemented the missing enterprise foundation pieces.
- Evidence: repository now includes an Entra-compatible JWKS verifier, injectable verifier support, tenant capability policy hooks, tenant/user-scoped settings and generated artifacts, retention expiry, rate limiting, and Preview/Personal/Managed deployment modes.
- Why it matters: these changes move the project from declarative production guidance toward executable security and trust boundaries.
- Decision or next action: continue M1 with organization-specific tenant policy persistence, vault integration, integration tests, and CI security scanning.
- Owner/status: engineering / baseline complete; production integration remains.

- What exact tenant claim and authorization model should govern pilot access?
- Which provider credentials are user-owned, admin-owned, or prohibited from persistence?
- What is the retention period for prompts, uploaded files, M365 context, generated artifacts, and audit events?
- Which Office clients are in the supported matrix for the first pilot?
- Should generated artifacts be downloaded only, or also stored for later project retrieval?

## Review cadence

- Update after each security or architecture decision.
- Add validation evidence after each release milestone.
- Promote confirmed decisions into `docs/roadmap.md` or the relevant implementation documentation.
- Archive resolved entries only when the evidence remains useful for future maintainers.
