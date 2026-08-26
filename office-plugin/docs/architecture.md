# Architecture and file structure

Expedient AI Workspace is one Office task-pane app shared by Excel, Word, and PowerPoint. It has three runtime pieces:

1. React task pane UI served from the add-in domain.
2. Office.js host adapters for reading/writing the active Office file.
3. Same-origin Node server for static hosting, authenticated APIs, BYOK proxying, settings, web search/fetch, and Microsoft 365 delegated context.

## Runtime flow

```text
Excel / Word / PowerPoint ribbon
  -> manifest.xml command
  -> hosted task pane /index.html
  -> React app
  -> Office.js host APIs for active file operations
  -> same-origin server APIs
       /api/settings       authenticated settings seam; development fallback only
       /api/proxy          BYOK provider proxy
       /api/web-search     lightweight web search tool
       /api/web-fetch      public URL fetch tool
       /api/image-asset    public web image fetch/validation/base64 bridge
       /api/generated/*    generated Office file artifacts for API gaps
       /api/m365/*         Office SSO / Graph delegated context tools
       /healthz,/readyz    liveness and production readiness probes
```

## Production security boundary

API requests are development-accessible only when `NODE_ENV` is not `production` and `REQUIRE_AUTH` is not enabled. Production requires a bearer-token verifier exposed through `AUTH_VALIDATOR_MODULE`; requests without valid identity are rejected before reaching API handlers.

Outbound requests use exact-origin allowlists, reject credentials and private/local network destinations, disable redirects, apply timeouts, and enforce request/response limits. Configure `BYOK_ALLOWED_TARGETS`, `IMAGE_ALLOWED_TARGETS`, and `WEB_FETCH_ALLOWED_TARGETS` per environment.

The checked-in `manifest.xml` is for local development and contains SSO placeholders. `tools/make-production-manifest.mjs` requires an HTTPS origin and real `MSAL_CLIENT_ID`/`OFFICE_SSO_RESOURCE`, and rejects output containing `localhost` or placeholder IDs.

The remaining production integration work is organization-specific: provide the Entra JWT verifier contract, an admin-backed tenant policy module, a tenant/user-scoped settings store, vault-backed managed credentials, and object storage/retention for artifacts. The server exposes injectable `TENANT_POLICY_MODULE`, `SETTINGS_STORE_MODULE`, and privacy-safe audit sinks; filesystem token caches and browser-local BYOK settings remain development-only compatibility paths.

## Top-level files

| Path | Purpose |
| --- | --- |
| `manifest.xml` | Local/dev Office add-in manifest for Excel, Word, and PowerPoint. Contains placeholder `WebApplicationInfo` for Office SSO until real Entra app details are configured. |
| `index.html` | Vite entry point for the task pane. |
| `vite.config.ts` | Local dev HTTPS server, proxy routes, shared settings route, M365 route wiring. |
| `package.json` | Local commands for dev, build, test, manifest validation, Office launchers, and production packaging. |
| `open-*-addin.cmd` | Friendly Windows launchers for local sideloading in Excel, Word, and PowerPoint. |
| `stop-addin.cmd` | Stops local Office add-in/dev server processes. |

## Source layout

| Path | Purpose |
| --- | --- |
| `src/App.tsx` | Main task-pane UI: chat, settings drawer, task progress, attachments, new chat/session memory, Office context refresh. |
| `src/components/SettingsPanel.tsx` | BYOK provider, model routing, shared branding, and app settings UI. |
| `src/lib/tools.ts` | Model tool/function schemas and browser-side tool executor routing. |
| `src/lib/capabilities.ts` | Runtime feature detection. Filters tools by current Office host and available APIs. |
| `src/lib/sessionMemory.ts` | Temporary open-session chat memory cache and keyword retrieval. Can later be replaced with vector-backed retrieval. |
| `src/lib/uploadRegistry.ts` | Current-pane uploaded/imported asset registry used by native Office image insertion tools, generated Office files, and temporary PowerPoint template preservation. |
| `src/lib/storage.ts` | Shared settings client. Uses `/api/settings` first, local storage as fallback. |
| `src/lib/artifacts.ts` | Text/table/slide parsing helpers and attachment summaries. |
| `src/office/host.ts` | Office.js host adapter: read context and mutate Excel, Word, PowerPoint with host-specific authoring tools. |
| `src/office/sso.ts` | Office SSO client helper using `OfficeRuntime.auth.getAccessToken()` / `Office.auth.getAccessToken()`. |
| `src/providers/adapters.ts` | OpenAI-compatible and Anthropic-compatible provider adapters, model routing, streaming, native tool loop. |
| `src/styles/app.css` | Task-pane UI styling. |

## Server layout

| Path | Purpose |
| --- | --- |
| `server/server.mjs` | Production/static server. Serves built app and owns `/api/settings`, `/api/proxy`, `/api/web-search`, `/api/web-fetch`, `/api/m365/*`. Production adapter contract templates live beside it as `*.production.example.mjs` and fail closed until replaced. |
| `server/m365.mjs` | Microsoft 365 delegated context implementation: Office SSO token intake, optional On-Behalf-Of exchange, device login fallback, Graph file search/read, Office-file text extraction including PowerPoint layout metadata, raw fallback context. |
| `server/file-context.mjs` | Uploaded file context extraction/fallback for Office files, PDFs, text files, and unknown formats. |
| `server/generated-office.mjs` | Generated Office artifact lane for downloadable `.pptx`, `.docx`, and `.xlsx` files with Open XML structure, embedded media support, template-preserving PowerPoint package cloning, and bounded DOCX template-shell preservation. |
| `server/image-asset.mjs` | Public web image fetch/validation/base64 conversion for `web_image_import`, Office image insertion tools, and generated Office files. |

## Tooling and deployment

| Path | Purpose |
| --- | --- |
| `scripts/validate-manifest.mjs` | Basic manifest sanity check for hosts, source URL, permissions, and SSO block presence. |
| `tools/make-production-manifest.mjs` | Rewrites localhost manifest URLs to a production origin and optionally replaces Office SSO placeholders from env vars. |
| `tools/open-office.mjs` | Local sideload launcher for Excel/Word/PowerPoint. |
| `tools/certs.mjs` | Local Office HTTPS certificate setup/verification. |
| `deploy/production/` | Production deployment lane: Dockerfile, env template, smoke test, centralized deployment guide. |
| `public/` | Static manifest assets, including ribbon icons and source brain logo asset. |
| `tests/` | Unit tests for provider adapters, artifacts, storage, M365 tooling, capabilities, and session memory. |

## Authentication ladder for M365 file context

When the model needs Microsoft 365 files, it should use this order:

1. `m365_try_office_sso` to reuse the Office profile already signed into Excel/Word/PowerPoint.
2. `/api/m365/sso` stores the Office SSO token; if `MSAL_CLIENT_SECRET` is configured, the server attempts Graph On-Behalf-Of exchange.
3. In local development only, controlled fallback token sources may be used: `GRAPH_ACCESS_TOKEN`, an explicit compatibility cache, or device login. Production rejects those escape-hatch variables and requires Office SSO plus backend OBO.
4. If no delegated token exists, device-login fallback can start through `m365_auth_status`/device flow.

Production should prefer Office SSO plus On-Behalf-Of exchange, read-only Graph scopes by default, and visible source attribution in the task pane.

## Production seams

These are intentionally simple local implementations today and should be hardened behind the same contracts for production:

| Seam | Local/dev behavior | Production target |
| --- | --- | --- |
| `/api/settings` | Stores one shared JSON blob for Excel/Word/PowerPoint. | Tenant/user settings from authenticated backend. |
| BYOK key storage | User enters key in task pane/shared settings. | Vault-backed customer secret or gateway-held key. |
| M365 token cache | Local JSON token caches for SSO/device/dev paths. | Backend session/token store with Entra/MSAL and vault controls. |
| Session memory | `sessionStorage` compact chat records. | Optional vector/index store with retention policy and per-file/user scope. |
| Provider allowlist | `BYOK_ALLOWED_TARGETS` env. | Enforced enterprise policy with audit logging. |
