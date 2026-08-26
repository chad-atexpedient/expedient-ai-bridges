# Production launch checklist

Use this checklist when moving from local sideloading to a managed Microsoft 365 rollout.

## 1. Choose production hosting

- [ ] Pick the production add-in origin, for example `https://ctrl-ai.example.com`.
- [ ] Ensure HTTPS uses a trusted production certificate.
- [ ] Confirm the hosted app can serve static task-pane files and same-origin APIs.
- [ ] Set `BYOK_ALLOWED_TARGETS` to approved provider origins only.
- [ ] Decide whether the BYOK model remains user-entered, gateway-managed, or vault-managed.

## 2. Create/configure Entra app registration

Required for Office SSO and Microsoft Graph delegated context.

- [ ] Create or choose an Entra app registration for the Office add-in.
- [ ] Record the Application/client ID.
- [ ] Set supported account/tenant policy.
- [ ] Configure the application ID URI/resource, for example:

```text
api://ctrl-ai.example.com/<client-id>
```

- [ ] Add/verify delegated Microsoft Graph scopes. Start read-only by default:

```text
openid
profile
  User.Read
  Files.Read
```

- [ ] Add `offline_access` or broader scopes only when a documented product requirement and tenant consent decision justify them.
- [ ] If using backend On-Behalf-Of exchange, create/store a client secret or certificate securely.
- [ ] Do not use placeholder app IDs in production.

## 3. Configure server environment

Start from `deploy/production/.env.example`.

Required/recommended:

```text
PORT=3000
HOST=0.0.0.0
BYOK_ALLOWED_TARGETS=https://internal-beta.expedient.cloud,https://api.openai.com,https://api.anthropic.com
IMAGE_ALLOWED_TARGETS=https://images.unsplash.com,https://upload.wikimedia.org
WEB_FETCH_ALLOWED_TARGETS=https://www.microsoft.com,https://learn.microsoft.com
ADDIN_ORIGIN=https://ctrl-ai.example.com
REQUIRE_AUTH=true
AUTH_VALIDATOR_MODULE=./server/auth.mjs
TENANT_POLICY_MODULE=./server/tenant-policy.production.example.mjs
SETTINGS_STORE_MODULE=./server/settings-store.production.example.mjs
AUDIT_LOG_STDOUT=false
REQUEST_MAX_BYTES=2000000
PROXY_MAX_BYTES=2000000
PROXY_MAX_RESPONSE_BYTES=10000000
OUTBOUND_TIMEOUT_MS=15000
RATE_LIMIT_PER_MINUTE=120
RATE_LIMIT_EXPENSIVE_PER_MINUTE=30
MSAL_CLIENT_ID=<real-client-id>
MSAL_TENANT_ID=<tenant-id-or-common>
GRAPH_SCOPES=openid profile User.Read Files.Read
MSAL_CLIENT_SECRET=<secret-or-cert-backed-credential-for-OBO>
OFFICE_SSO_RESOURCE=api://ctrl-ai.example.com/<real-client-id>
```

Production target:

- [ ] Configure the production bearer-token verifier referenced by `AUTH_VALIDATOR_MODULE`; verify issuer, audience, tenant, subject, expiry, and signing keys.
- [ ] Provide `TENANT_POLICY_MODULE` with an admin-backed `allows({ identity, capability })` contract.
- [ ] Provide `SETTINGS_STORE_MODULE` with tenant/user-scoped `get`, `put`, and `delete` methods.
- [ ] Replace the checked-in `*.production.example.mjs` fail-closed contract templates with separately managed implementations; the example modules are not a production data store, policy service, or distributed limiter.
- [ ] Bind settings, artifacts, uploads, and M365 state to authenticated tenant/user identity; do not use the global JSON settings/token fallbacks.
- [ ] Replace local token files with backend session/token storage.
- [ ] Confirm `GRAPH_ACCESS_TOKEN`, compatibility-token paths, and all local M365 token/cache path variables are unset; production readiness rejects them if present.
- [ ] Store secrets in a vault.
- [ ] Connect the privacy-safe audit sink; record provider route, model, source identifiers, and tool outcomes without raw document contents, prompts, API keys, or tokens.
- [ ] Keep `/api/settings` contract but back it with tenant/user settings storage.

## 4. Generate production manifest

From project root:

```powershell
$env:MSAL_CLIENT_ID="<real-client-id>"
$env:OFFICE_SSO_RESOURCE="api://ctrl-ai.example.com/<real-client-id>"
node tools/make-production-manifest.mjs https://ctrl-ai.example.com
node scripts/validate-manifest.mjs
```

Expected output:

```text
dist/manifest.production.xml
```

Before upload:

- [ ] Confirm all `https://localhost:3000` URLs were replaced with the production origin.
- [ ] Confirm `WebApplicationInfo` contains the real app id/resource.
- [ ] Confirm hosts include Workbook, Document, and Presentation.
- [ ] Confirm permissions are intentional. Current permission is `ReadWriteDocument`.
- [ ] Confirm icon URLs point to production-hosted assets.

## 5. Build and deploy app

```powershell
node node_modules/vite/bin/vite.js build
node tools/make-production-manifest.mjs https://ctrl-ai.example.com
node scripts/package-release.mjs
```

Expected handoff output:

```text
dist/release/expedient-ai-workspace/
dist/release/expedient-ai-workspace/RELEASE_REPORT.md
dist/release/expedient-ai-workspace/manifest.xml
```

Before uploading the packaged manifest:

- [ ] Confirm the release package was created in production mode, not with `--dev`.
- [ ] Review `RELEASE_REPORT.md` for package name, version, generation time, contents, and checksums.
- [ ] Confirm `dist/release/expedient-ai-workspace/manifest.xml` contains the final HTTPS origin and real Entra app ID/resource.

Deploy either directly or via container:

```powershell
docker build -f deploy/production/Dockerfile -t expedient-ai-workspace:latest .
docker run --rm -p 3000:3000 --env-file deploy/production/.env.example expedient-ai-workspace:latest
```

Smoke test:

```powershell
node deploy/production/scripts/smoke-test.mjs https://ctrl-ai.example.com
```

## 6. Deploy through Microsoft 365 Admin Center

Follow `deploy/production/m365-centralized-deployment.md`.

- [ ] Upload `dist/manifest.production.xml` as an Office Add-in.
- [ ] Assign to a pilot group first.
- [ ] Ask users to restart Excel, Word, and PowerPoint.
- [ ] Confirm `CTRL AI` appears on the Home ribbon.

## 7. Pilot validation

Run this in each host: Excel, Word, PowerPoint.

- [ ] Task pane loads from production domain.
- [ ] Shared branding/settings load consistently across hosts.
- [ ] Provider connection test succeeds.
- [ ] Temperature/max tokens can stay blank and use gateway defaults.
- [ ] Auto model routing defaults show the intended three model tiers.
- [ ] Office context read works.
- [ ] Native write tools work for the host.
- [ ] Anthropic-compatible tool use works through `tool_use`/`tool_result` with the same Office/web/M365/generated-file executor as OpenAI-compatible routes.
- [ ] Excel range write works.
- [ ] Excel chart creation works where Office chart APIs are available.
- [ ] Excel PivotTable creation works where Office PivotTable APIs are available, and blocked runtimes fall back to summary tables or generated workbooks.
- [ ] Excel worksheet lifecycle, table creation, range clearing/formatting, sorting, filtering, freeze panes, dropdown validation, conditional formatting, comments/notes, named ranges, worksheet protection, and page layout work where runtime APIs are available.
- [ ] Word text, heading, table, style, font/selection formatting, page-break, section-break, header/footer, find/replace, comment, hyperlink, and image tools work where runtime APIs are available.
- [ ] Generated Word documents can include real bulleted/numbered list parts, comments, footnotes/endnotes, and tracked-change style redlines for legal/comms/research review workflows.
- [ ] PowerPoint slide creation, slide duplication, text boxes, hyperlink text boxes, table grids, shapes, slide background, slide cleanup/deletion, speaker notes, image insertion, and generated deck fallback work where runtime APIs are available.
- [ ] PowerPoint uploads expose slide text, notes, layout/theme/media hints for template/content generation.
- [ ] Word uploads expose body text, styles, comments, headers/footers, and media hints for review/summarization/transformation.
- [ ] Excel uploads expose workbook/sheet/table/filter/chart metadata for analysis and transformation.
- [ ] Uploaded images can be inserted as real Office image objects where the installed Office runtime supports image APIs.
- [ ] Topical image requests can call `web_image_search`, show/use source URLs, import the selected URL with `web_image_import`, and pass selected `assetId`, `assetName`, or `imageUrl` values into Office image or generated-file tools.
- [ ] Public web image URLs can be fetched through `/api/image-asset` and inserted, with clear failure messages for unsupported runtimes or invalid image responses.
- [ ] Unknown upload formats return bounded raw/base64 fallback context instead of failing silently.
- [ ] Feature detection hides or blocks unavailable tools with clear messages.
- [ ] Public image use has tenant-approved copyright/license/security policy before broad rollout.
- [ ] `web_image_search` and `web_image_import` have production source allowlists/license filtering before broad rollout, with `web_image_import` as the validation/enforcement point before embedding.
- [ ] Generated PowerPoint `.pptx`, Word `.docx`, and Excel `.xlsx` files can be created, downloaded, opened, and inspected for real Open XML structure; generated files can embed uploaded image assets or public image URLs without requiring the model to manually provide base64.
- [ ] Simple demo/showcase/tour requests render as an in-chat add-in feature card without model/tool calls; explicit demo actions can create a live Office sample with `mode=live` or downloadable showcase files with `mode=artifact`/`surface=all`.
- [ ] Generated Word `.docx` reports can include a TOC field, document-level and section-specific page layout/orientation/margins, real Word columns, hyperlinks, real caption/cross-reference fields, real bulleted/numbered lists, footnotes, headers/footers, comments, tables, and embedded images rather than manual layout instructions.
- [ ] Generated Excel `.xlsx` workbooks can include self-contained bar/line/area/pie/doughnut/scatter/combo chart parts anchored through worksheet drawings rather than chart instructions, with no dangling PowerPoint-only external chart-data relationships.
- [ ] Generated Excel `.xlsx` workbooks include professional document properties in `docProps/core.xml` and `docProps/app.xml` for title, subject, author/creator, keywords, category, company, and manager where supplied.
- [ ] Generated Excel `.xlsx` workbooks can include workbook-scoped and sheet-scoped named ranges for semantic formulas, reusable inputs, and downstream automation.
- [ ] Generated Excel `.xlsx` workbooks can include cell comments/notes for review, assumptions, audit notes, and handoff context; uploaded workbook context extraction surfaces those comments back to the model.
- [ ] Generated Excel `.xlsx` workbooks can include clickable external hyperlinks with display text/screen tips, and uploaded workbook context extraction surfaces hyperlink targets back to the model.
- [ ] Generated Excel `.xlsx` worksheets can include real merged cells, explicit column widths/hidden states, row heights/hidden rows, and plain range AutoFilters for polished report layouts and filterable handoff sheets.
- [ ] Generated Excel `.xlsx` worksheets preserve View/Page Layout settings such as frozen panes, zoom, gridlines/headings visibility, print margins, paper size/orientation, fit-to-page/scale, print area, and repeating title rows/columns.
- [ ] Generated Excel `.xlsx` worksheets can include sheet protection for handoff/accidental-edit prevention, with clear documentation that this is not encryption or strong data security.
- [ ] Generated PowerPoint decks can include real slide backgrounds, footer/date/classification/slide-number text boxes, shapes, hyperlink text boxes, professionally styled table graphic frames, editable bar/line/area/pie/doughnut/scatter/combo chart parts with embedded `.xlsx` chart-data workbooks, workbook-range-backed chart series, speaker notes, and embedded images with fit/fill/crop placement rather than placeholder instructions.
- [ ] When live PowerPoint image insertion is unavailable, the model uses `powerpoint_generate_deck_file` for an image-embedded `.pptx` instead of returning picture placeholders.

## 8. Microsoft 365 context validation

- [ ] `m365_try_office_sso` succeeds for a signed-in Office user.
- [ ] `/api/m365/sso` accepts the Office SSO token.
- [ ] Backend On-Behalf-Of exchange works when `MSAL_CLIENT_SECRET` is configured.
- [ ] `m365_search_files` returns only files the signed-in user can access.
- [ ] `m365_read_file` extracts text from `.docx`, `.pptx`, `.xlsx`, and text-like files.
- [ ] Unsupported files return bounded raw/base64 fallback context, not an opaque failure.
- [ ] Source names are visible to the user when used for grounding.
- [ ] Device login fallback is available only as a controlled fallback.

## 9. Memory/session validation

- [ ] New chat compacts the previous chat into temporary session memory.
- [ ] Relevant prior chat memory is pulled into later prompts.
- [ ] UI indicates when prior chat memory is used.
- [ ] Memory is session-bounded today; define retention policy before adding persistent/vector storage.

## 10. Security review gates

- [ ] No real API keys or tokens in repo files.
- [ ] No placeholder Entra IDs in production manifest.
- [ ] BYOK provider targets are allowlisted.
- [ ] Image and web-fetch target allowlists are configured and tested.
- [ ] Private-network SSRF, unsafe redirects, credential URLs, path traversal, and oversized responses are rejected.
- [ ] Security headers, request IDs, timeouts, response limits, and readiness probes are verified.
- [ ] Production container runs as non-root and dependency/image/secret scans pass.
- [ ] Secrets are vault-backed.
- [ ] Token storage/session strategy is approved.
- [ ] Graph scopes are least-privilege.
- [ ] Logs avoid raw file contents and secrets.
- [ ] Admin/pilot rollout path is approved.
- [ ] Recovery path documented: disable app assignment, rotate secrets, revoke refresh tokens, clear provider allowlist.

## 11. Release commands

Minimum validation before release:

```powershell
node node_modules/typescript/bin/tsc --noEmit
node node_modules/vitest/vitest.mjs run
node node_modules/vite/bin/vite.js build
node scripts/validate-manifest.mjs
node tools/make-production-manifest.mjs https://ctrl-ai.example.com
node scripts/package-release.mjs
node deploy/production/scripts/smoke-test.mjs https://ctrl-ai.example.com
```

The manifest command must receive real `MSAL_CLIENT_ID` and `OFFICE_SSO_RESOURCE` values. It intentionally fails when values are missing, when the origin is not HTTPS, or when placeholders remain.

After these pass, upload `dist/manifest.production.xml` through Microsoft 365 Centralized Deployment.



