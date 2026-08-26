# CTRL Office Add-in roadmap

This is the working roadmap for turning CTRL from a capable Office prototype into a secure, polished, enterprise-deployable product. Items are ordered by release risk and user value, not by implementation convenience.

## Status legend

- **Now** - required for the next safe milestone.
- **Next** - high-value work after the current release gate.
- **Later** - valuable capability that depends on platform, design, or policy decisions.
- **Blocked** - requires an external decision, tenant resource, Office capability, or provider integration.

## Now: production gate and identity integration

**Goal:** make a pilot deployment safe to expose to a controlled tenant.

- **Now — Entra verifier:** integrate and exercise the Entra-compatible verifier with the organization's issuer, audience, tenant authorization, subject, expiry, and signing-key contract. The verifier now has 15 tests using real RSA keys and a stubbed JWKS endpoint, covering wrong audience, array audience, wrong issuer, v1.0 rejection, expiry, `nbf`, `alg=none`, wrong-key signatures, unknown `kid`, and missing claims. A **fail-open tenant check was found and fixed**: with no allowlist configured, authorization fell back to the token's own tenant, admitting any tenant that could obtain a validly signed token.
- **Now — Tenant authorization must be explicit:** `AUTHORIZED_TENANTS` is now required in production, rejected when it names `common`/`organizations`/`consumers`/`*`, and cross-checked against `MSAL_TENANT_ID` so a deployment cannot authorize a tenant it can never sign in. It is documented in `.env.example` and enforced by the release gates. It was previously undocumented and unenforced, which is why the fail-open default went unnoticed.
- **Now — Tenant policy:** define user/tenant policy for provider targets, web access, uploads, generated artifacts, M365 context, and device login; runtime now accepts an injected `TENANT_POLICY_MODULE` contract and fails production readiness without one.
- **Now — Scoped persistence:** keep local development state isolated by authenticated tenant/user; replace JSON/filesystem settings, token, and artifact persistence with managed database/object storage before pilot.
- **Now — Managed credentials:** managed mode never persists personal provider keys; choose and integrate vault-backed organization credentials before pilot.
- **Now — Security regression lane:** API authorization, SSRF, redirect, rate-limit, file-limit, token-redaction, and manifest negative tests are in place. This run added regression coverage for four previously untested fixes: rate-limit bucket eviction as a limit-bypass, an unbounded file-context upload, an unbounded image download, and a missing outbound timeout. Writing those tests exposed two further defects (see below), which is the argument for making every fix carry a test that has actually been observed to fail.
- **Now — Eviction must not be a rate-limit bypass:** bucket eviction previously deleted every other bucket under memory pressure, letting any caller reset all other users' counters. The first fix preserved current-window buckets but still fell back to insertion order, so flooding with fresh current-window keys could still evict a victim near their limit. Eviction now sheds the lowest counters first. A distributed limiter remains required for multi-instance deployment; the in-process limiter is correct but single-node.
- **Now — Reject oversized uploads by tearing down the connection:** the file-context size limit rejected oversized payloads but never destroyed the request, so a client could keep streaming indefinitely. The limit bounded memory but not the flood. The handler now destroys the request and returns `413`. Upload limits are also read at call time rather than module load, so configuration cannot be silently bypassed by import order.
- **Now — Deployment gate:** require security scans, reproducible dependency installation, health/readiness checks, and real production manifest generation before deployment. CI now includes an executable release-gate script, production container build, Trivy high/critical scan, SBOM generation, and artifact upload; signed/reproducible artifact verification remains external CI evidence.
- **Now — Privacy-safe audit seam:** emit correlation-safe auth, policy, rate-limit, and route outcome events through an injectable audit sink without content or secret fields.
- **Now — Dev/prod boundary parity:** keep the Vite development proxy behind the same outbound URL, redirect, header, timeout, and response-size controls as production; provider credentials now use an explicit internal header contract and arbitrary inbound headers are not forwarded; local development remains the only mode where the allowlist may be empty.
- **Now — Identity/routed rate limits:** rate-limit by authenticated tenant/user and route, with a stricter budget for expensive outbound operations.

## Next: trustworthy daily workflow

**Goal:** make the add-in predictable and pleasant in real Office work.

- **Next — First-run setup:** distinguish Preview, Personal BYOK, and Managed Enterprise modes with clear data-flow and storage disclosure; preview setup guidance is implemented locally, with tenant consent/onboarding still required.
- **Next — Settings disclosure semantics:** show unsaved deployment-mode changes as an in-dialog preview and update the main trust banner only after the user saves; implemented and UI-smoke verified.
- **Next — Recovery UX:** add retry, cancellation, expired-session recovery, provider diagnostics, upload failure recovery, and host capability explanations. Cancellation now reaches streaming provider requests and bounded tool loops; upload failures use inline status; remaining work is expired-session recovery and live host diagnostics.
- **Next — Context provenance:** show which workbook/document/deck, attachments, or memory grounded a response; source chips are implemented locally, while freshness and M365 provenance require live context integration.
- **Next — Consequential writes:** add a structured review/confirm step before inserting assistant output; implemented locally, with host-specific undo/recovery still requiring Office-host validation.
- **Next — Accessibility QA:** complete keyboard, focus, contrast, live-region, narrow task-pane, and screen-reader validation. Settings focus trapping, inline validation, reset confirmation, and the thinking-status encoding defect are now addressed locally; rendered keyboard/screen-reader validation remains a release gate.
- **Next — Operational telemetry:** add privacy-safe metrics for auth, provider, tool, latency, rate-limit, and artifact failures without collecting content or secrets.

## Next: professional Office workflows

**Goal:** close the highest-value gaps identified in the Office capability audit.

- **Next — Word templates:** preserve uploaded DOCX templates for branded reports, contracts, and controlled document assembly. The first implementation slice now accepts `templateAssetId`, `templateAssetName`, or `templateBase64`, validates the DOCX package, preserves the package shell and non-generated parts, retains template styles/theme/custom XML/media/header-footer relationships, and replaces the generated document body. Remaining gate: rendered Word validation, content-control/field behavior, and malware/quarantine policy.
- **Next — PowerPoint brand extraction:** infer theme colors, fonts, layouts, and logo treatment from an uploaded deck. Structured extraction now returns theme colors, major/minor fonts, layout/placeholders, and filename-based media candidates through attachment context and the generated-template path; fallback generated decks now apply validated colors/fonts to theme, shapes, tables, charts, and chrome while explicit styles retain precedence. Remaining gates: confirm media ownership/licensing and perform rendered visual QA across Office clients.
- **Next — Excel analyst polish:** improve slicer/pivot-chart support, executive workbook layout, named ranges, print setup, and review annotations.
- **Now — Workbook grounding:** read before write. `excel_get_workbook_overview`, `excel_read_range`, and `excel_search_workbook` are implemented locally, and `excel_write_range` now refuses non-empty targets without `overwrite: true` and verifies stored values. Remaining: live Excel host validation across large and protected workbooks.
- **Next — Mutation history and undo:** capture a pre-mutation checkpoint for every workbook write and expose one-click revert in the task pane. Informed by the pi-for-excel `workbook_history` pattern; needs a storage and retention decision before implementation.
- **Next — First-class formula lineage:** expose precedent/dependent tracing and plain-language formula explanation as direct tool responses rather than only as generated audit sheets.
- **Now — Agent eval harness:** implemented. Grades runs on resulting workbook state with two-level scoring (Modif. = intended changes made; Acc. = whole workbook still correct), first-class destructive-edit detection, protected sheets, and hidden variants. The grader self-test is a release gate, verified by sabotage.
- **Now — Eval corpus integrity gates:** implemented. `--validate` rejects task specs that would silently disarm the grader (no assertions, a fix the seed already satisfies, a nonexistent protected sheet, an unsatisfiable value assertion on a formula cell). `--check-corpus` proves per task and per variant that the reference solution passes while the reference solution plus collateral damage scores Modif.=PASS with Acc.=FAIL. Both are release gates and both are sabotage-verified.
- **Next — Eval task corpus depth:** five task families are shipped (`margin-model-doctor`, `quarterly-rollup`, `lookup-repair`, `budget-variance-audit`, `headcount-restraint`), each with hidden variants and a declared destructive-run reference. The task set is the asset; keep the harness thin. All five are synthetic and none reproduces the scale or dependency depth of a real client workbook. Licensed or customer-derived workbooks belong in a private corpus with the documented leakage scrub applied.
- **Next — Eval recalculation oracle:** verify that fixed formulas compute correct values, not just that they match expected formula text. Needs a headless recalc engine or a real Excel host. The task validator now rejects value assertions on formula cells rather than letting them fail mysteriously, which makes the missing oracle explicit instead of latent.
- **Next — Live-host eval driver:** drive a real Excel session end-to-end so eval tasks become true Office integration coverage rather than grader coverage. This is the single largest remaining gap in the eval lane: everything today grades a file, not a session.
- **Later — Eval regression history:** persist per-task, per-variant pass rates over time so agent or model regressions are detectable, not just individually observable.
- **Later — Formula equivalence grading:** formula comparison is textual, so an algebraically equivalent rewrite grades as damage. Correct for restraint tasks, wrong for refactor tasks. An equivalence-aware comparison needs a formula parser and a per-task opt-in.
- **Later — Context budget policy:** define explicit reinjection and compaction rules for workbook blueprints, tool outputs, and long sessions, optimizing for context headroom rather than billed tokens.
- **Next — Cross-surface workflows:** support Excel-to-PowerPoint board packs, M365 meeting pre-reads, Word-to-Excel transforms, and deck-to-brief workflows.
- **Next — Generated artifact QA:** validate generated DOCX/XLSX/PPTX package structure before presenting them; implemented locally. Rendered visual QA remains a release gate.

## Later: platform and governance maturity

- **Later — Admin controls:** tenant-level provider catalog, feature flags, retention, regional routing, and central branding.
- **Later — Approval workflows:** review queues for generated artifacts and high-impact file mutations.
- **Later — Audit exports:** administrator-visible action history with content-safe event records.
- **Later — Provider abstraction:** policy-aware routing, cost controls, model capability registry, and regional/data-residency constraints.
- **Later — Multi-user collaboration:** shared project memory with explicit retention, access, and deletion semantics.

## Blocked decisions

- **Blocked — Identity contract:** exact Entra token issuer/audience/tenant policy and verifier hosting model.
- **Blocked — Storage contract:** database/object-store/vault choices, data region, retention, and deletion SLA.
- **Blocked — Managed BYOK model:** whether enterprise users enter personal keys, admins configure provider credentials, or both modes are supported.
- **Blocked — Office support matrix:** minimum Office versions, Windows/Mac/web support, and required degraded behavior.
- **Blocked — Product positioning:** broad knowledge-work assistant versus focused Office authoring/workflow product.

## Release milestones

### M0 — Hardened development baseline

Current state: complete.

- Server security primitives and production fail-closed hooks implemented.
- Development proxy now uses the shared outbound safety policy and no longer emits wildcard CORS or forwards arbitrary inbound headers.
- Production manifest rejects placeholders and non-HTTPS origins.
- Graph defaults reduced to least-privilege read-oriented scopes.
- UI modal, focus, and constrained-viewport improvements implemented.
- 227 tests across 20 files, TypeScript, Vite production build, manifest validation, and local release-gate checks pass after the brand-style, archive-bound, SSO identity-binding, artifact-retention, production Graph escape-hatch, request-origin, and upload-type increments. These are local repository gates; production readiness now also fails closed when image or document-fetch allowlists are absent, and all generated embedded images pass local signature validation; external tenant, Office-host, rendered-artifact, malware-scan, CI scanner, and centralized-deployment evidence remains required.
- Structured write review, setup guidance, provenance chips, and generated-package validation are implemented and UI-smoke covered.
- Agent eval harness grades workbook state with two-level scoring; its self-test is enforced by the release gate and verified by deliberate sabotage. Five task families ship with hidden variants, spec validation, and per-variant discrimination checks. Recalculation oracle and live-host eval driver remain outstanding, and all seeds are synthetic.
- Six server defects were found by inspection and fixed this cycle, including a fail-open tenant authorization check in `server/auth.mjs`. Two of the six fixes were incomplete and were corrected while writing their regression tests.
- **Recurring pattern worth naming:** the two most serious findings this cycle were both *inert controls* rather than missing ones. The tenant check appeared present and authorized everyone; a malformed eval task appears authored and asserts nothing. A control that reviews as present and behaves as absent is more dangerous than an obvious gap, because it consumes the attention that would have built the real thing. Prefer gates that fail closed, and verify each one by breaking it on purpose.

### M1 — Authenticated pilot

Exit criteria:

- Real Entra verifier and tenant policy store deployed.
- No global settings/token/artifact state.
- Provider and outbound policies enforced per tenant/user.
- Security tests and container/dependency scans pass.
- Pilot manifest deployed to a controlled tenant group.

### M2 — Trusted daily use

Exit criteria:

- First-run, consent, provenance, recovery, cancellation, and write-confirmation UX complete.
- Accessibility and responsive QA pass across supported task-pane sizes.
- Privacy-safe telemetry and support runbook operational.

### M3 — Professional Office workflows

Exit criteria:

- Word template preservation and PowerPoint brand-profile extraction/style application are implemented locally; licensing review, rendered Office visual QA, Excel analyst polish, and explicit deferral decisions remain.
- Cross-surface workflows have artifact QA and source provenance.

## Roadmap maintenance

Update this document when a milestone changes, a decision becomes unblocked, or evidence changes priority. Record the underlying evidence and rationale in `docs/review.md`; keep this file focused on direction, sequencing, and exit criteria.
