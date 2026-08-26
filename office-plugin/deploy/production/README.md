# Expedient AI Workspace - Production Deployment

This folder is the path from local sideloading to a centrally deployed Microsoft Office add-in.

Before production launch, use `../../docs/production-launch-checklist.md` as the full readiness checklist. The service must run with authentication, tenant/user isolation, explicit outbound allowlists, and real Entra configuration; the local JSON settings/token fallbacks are development-only.

The production deployment is one hosted app:

```text
Excel / Word / PowerPoint ribbon command
        -> https://<addin-domain>/index.html
        -> static React task pane
        -> Office SSO where configured
        -> same-origin /api/proxy
        -> user-selected BYOK provider endpoint
        -> same-origin /api/m365/* for Microsoft Graph delegated context
```

The add-in does not require `C:\Users\chad.stthomas\codex-bridge` in production. The BYOK proxy is built into this project at `server/server.mjs` and is packaged with the web task pane.

The checked-in `server/*.production.example.mjs` files are fail-closed contract templates only. Replace them with separately managed production modules connected to the tenant policy service, database/object store, and distributed rate limiter before allowing pilot traffic.

## Files in this lane

- `Dockerfile` - container image for the hosted add-in service.
- `.env.example` - production environment variables.
- `m365-centralized-deployment.md` - Microsoft 365 Admin Center rollout steps.
- `scripts/smoke-test.mjs` - simple deployed-service smoke test.

## Build production assets

From the project root:

```powershell
$env:MSAL_CLIENT_ID="<real-client-id>"
$env:OFFICE_SSO_RESOURCE="api://your-addin-domain.example.com/<real-client-id>"
node node_modules/vite/bin/vite.js build
node tools/make-production-manifest.mjs https://your-addin-domain.example.com
```

This creates:

```text
dist/app/                  Static task pane assets
dist/manifest.production.xml
```

## Create a handoff release directory

For a production handoff after the build and production manifest are ready:

```powershell
npm run package:prod
```

This creates:

```text
dist/release/expedient-ai-workspace/
```

The release directory contains the built task-pane app, server runtime files, production manifest copied as `manifest.xml`, deployment docs, the production environment template, Dockerfile, smoke-test script, and `RELEASE_REPORT.md` with package metadata and checksums.

For local/sideload QA handoff only, use:

```powershell
npm run package:dev
```

That package intentionally uses the development `manifest.xml` and is not suitable for Microsoft 365 centralized production deployment.

## Run locally in production mode

```powershell
$env:PORT="3000"
$env:BYOK_ALLOWED_TARGETS="https://internal-beta.expedient.cloud,https://api.openai.com,https://api.anthropic.com"
node server/server.mjs
```

Then visit:

```text
http://localhost:3000/
```

In real deployment, terminate TLS at your platform/load balancer and expose the app as HTTPS.

## Container build

From the project root:

```powershell
docker build -f deploy/production/Dockerfile -t expedient-ai-workspace:1.0.0 .
docker run --rm -p 3000:3000 --env-file deploy/production/.env.example expedient-ai-workspace:1.0.0
```

## Smoke test a deployed service

```powershell
node deploy/production/scripts/smoke-test.mjs https://your-addin-domain.example.com
```

The smoke test checks:

- the task pane HTML loads;
- required brain icon assets load;
- `/healthz` reports liveness;
- `/readyz` reports configuration readiness;
- unauthenticated API access is rejected.

## BYOK security model

The user supplies their own API key in the add-in UI. The hosted app forwards that key only to the configured provider endpoint. In production, set `BYOK_ALLOWED_TARGETS` so keys can only be sent to approved provider origins.

Example:

```text
BYOK_ALLOWED_TARGETS=https://internal-beta.expedient.cloud,https://api.openai.com,https://api.anthropic.com
```

If this is empty in production, the service must fail readiness and reject outbound requests.

## Office SSO and Microsoft 365 context

The production add-in should use Office SSO first for Microsoft 365 file context:

1. Configure Entra app registration.
2. Set manifest `WebApplicationInfo` through `MSAL_CLIENT_ID` and `OFFICE_SSO_RESOURCE` before generating `dist/manifest.production.xml`.
3. Configure backend On-Behalf-Of exchange with `MSAL_CLIENT_SECRET`, a concrete `MSAL_TENANT_ID`, and `OFFICE_SSO_RESOURCE`.
4. Keep Graph scopes least-privilege; the default is `openid profile User.Read Files.Read`.
5. Device login, compatibility-token imports, filesystem token caches, and `GRAPH_ACCESS_TOKEN` are prohibited by the production configuration gate; readiness fails closed if any legacy token/cache escape-hatch variable is present.

See `../../docs/office-api-hardening.md` for the full auth ladder and context-reader behavior.

## Microsoft 365 deployment

After hosting the app and generating `dist/manifest.production.xml`, follow `m365-centralized-deployment.md` to deploy it through Microsoft 365 Admin Center so users see the add-in in Excel, Word, and PowerPoint without manual XML sideloading.

If using the packaged handoff directory, upload `dist/release/expedient-ai-workspace/manifest.xml` only after confirming the release was created in production mode and the manifest contains the final HTTPS origin and real Entra app ID/resource.