# CTRL BYOK Office Add-in

A shared Microsoft Office task-pane add-in for Excel, Word, and PowerPoint. It supports BYOK OpenAI-compatible/Anthropic-compatible providers, branded settings, native Office tools, Microsoft 365 file context, and temporary session chat memory.

## What is included

- One React/Vite task pane shared by Excel, Word, and PowerPoint.
- Office manifest with Workbook, Document, and Presentation hosts.
- Shared provider/branding settings across Office hosts.
- OpenAI-compatible `/chat/completions` adapter with native tool calling.
- Anthropic-compatible `/messages` adapter with shared context/capability prompts.
- Model routing defaults: `GPT-5-Mini`, `GPT-5.4`, `GPT-5.5`.
- Gateway-default temperature/max tokens when those fields are blank.
- Native Office tools for reading context and writing into Excel, Word, and PowerPoint.
- Microsoft 365 context tools using Office SSO first, then controlled delegated fallbacks.
- Temporary new-chat/session memory while the task pane is open.
- Same-origin server for settings, BYOK proxy, web search/fetch, and M365 Graph context.

## Local quickstart

Easiest path: double-click one launcher from this folder:

```text
open-excel-addin.cmd
open-word-addin.cmd
open-powerpoint-addin.cmd
```

Stop/restart helpers:

```text
stop-addin.cmd
clear-office-addin-cache.cmd
```

Package-script path:

```powershell
pnpm install
pnpm certs
pnpm excel
```

The local add-in runs at `https://localhost:3000` with Microsoft Office development certificates. If Office shows stale ribbon icons or commands, close all Office apps, run `clear-office-addin-cache.cmd`, then launch again.

## Important local notes

- The checked-in manifest is development-only and contains placeholder Office SSO `WebApplicationInfo` values. Production packaging fails unless a real Entra app registration and HTTPS origin are supplied.
- Production requires a concrete Entra tenant, Office SSO resource, and backend On-Behalf-Of exchange. Developer token fallbacks are ignored by the production runtime.
- BYOK keys/settings are local/shared for development only. Managed production mode requires authenticated, tenant-scoped storage and vault-backed provider credentials.
- Production API routes fail closed without authentication configuration, outbound provider allowlists, and an explicit add-in origin. Developer token fallbacks must not be enabled in production.

## Documentation map

- [docs/README.md](</C:/Users/chad.stthomas/Downloads/CTRL Add IN/docs/README.md>) - documentation index and production posture summary.
- [docs/architecture.md](</C:/Users/chad.stthomas/Downloads/CTRL Add IN/docs/architecture.md>) - file structure, runtime flows, server/API seams, auth ladder.
- [docs/production-launch-checklist.md](</C:/Users/chad.stthomas/Downloads/CTRL Add IN/docs/production-launch-checklist.md>) - production launch readiness checklist.
- [docs/office-api-hardening.md](</C:/Users/chad.stthomas/Downloads/CTRL Add IN/docs/office-api-hardening.md>) - long-term hardening guidance for Office APIs, shared settings, M365 context, and memory.
- [deploy/production/README.md](</C:/Users/chad.stthomas/Downloads/CTRL Add IN/deploy/production/README.md>) - production hosting/package lane.
- [deploy/production/m365-centralized-deployment.md](</C:/Users/chad.stthomas/Downloads/CTRL Add IN/deploy/production/m365-centralized-deployment.md>) - Microsoft 365 Admin Center deployment steps.
- [docs/roadmap.md](</C:/Users/chad.stthomas/Downloads/CTRL Add IN/docs/roadmap.md>) - phased product and enterprise-readiness roadmap.
- [docs/review.md](</C:/Users/chad.stthomas/Downloads/CTRL Add IN/docs/review.md>) - ongoing creative, issue, decision, and discovery log.

## Useful commands

```powershell
pnpm test
pnpm build
pnpm validate:manifest
```

Direct commands used by the scripts:

```powershell
node node_modules/typescript/bin/tsc --noEmit
node node_modules/vitest/vitest.mjs run
node node_modules/vite/bin/vite.js build
node scripts/validate-manifest.mjs
node tools/make-production-manifest.mjs https://your-addin-domain.example.com
```

## Production source of truth

Use [docs/production-launch-checklist.md](</C:/Users/chad.stthomas/Downloads/CTRL Add IN/docs/production-launch-checklist.md>) before launch. In short:

1. Host the app on a stable HTTPS production domain.
2. Configure Entra app registration and Office SSO `WebApplicationInfo`.
3. Configure backend Graph On-Behalf-Of exchange and least-privilege scopes.
4. Configure provider allowlists and backend/vault storage for secrets.
5. Build the app and generate `dist/manifest.production.xml`.
6. Deploy through Microsoft 365 Centralized Deployment.
7. Run pilot validation across Excel, Word, and PowerPoint.

## Security notes

- Do not commit real API keys, Graph tokens, client secrets, or production cert material.
- Use `BYOK_ALLOWED_TARGETS` in production.
- Prefer Office SSO plus backend On-Behalf-Of exchange for Microsoft Graph.
- Start with read-only Graph permissions by default.
- Keep source attribution visible when M365 files ground a response.
- The current manifest grants `ReadWriteDocument` because the add-in can write into the active Office file.
