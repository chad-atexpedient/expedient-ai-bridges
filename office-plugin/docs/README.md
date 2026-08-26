# CTRL BYOK Office Add-in docs

This folder is the project knowledge base. Use it with the production deployment lane under `deploy/production/`.

## Start here

- `architecture.md` - how the add-in is organized, what each major folder owns, and how runtime flows connect.
- `production-launch-checklist.md` - step-by-step launch readiness checklist for hosting, Entra/SSO, Graph access, BYOK policy, manifest packaging, validation, and rollout.
- `office-api-hardening.md` - implementation rules and long-term hardening notes for Office API feature detection, Microsoft 365 delegated context, shared settings, and session memory.
- `office-capability-roadmap.md` - current Office tool coverage matrix and remaining production roadmap toward professional toolbar-grade capability.
- `office-missing-pieces-audit.md` - practical Word/Excel/PowerPoint gap audit against ribbon-level and professional workflow expectations.
- `roadmap.md` - phased product, security, UX, Office capability, and deployment roadmap.
- `review.md` - durable review log for creative direction, issues, decisions, evidence, and discoveries.
- `../evals/README.md` - agent eval harness: grades runs on resulting workbook state with two-level scoring, destructive-edit detection, and hidden variants.

## Deployment-specific docs

- `../deploy/production/README.md` - production hosting/package lane.
- `../deploy/production/m365-centralized-deployment.md` - Microsoft 365 Admin Center deployment steps.
- `../deploy/production/.env.example` - production environment variable template.
- `../scripts/package-release.mjs` - creates `dist/release/ctrl-byok-office-addin/` handoff packages for production or local sideload QA.

## Repo map

```text
src/                  Task pane UI, Office host adapters, provider adapters, tools, storage, memory
server/               Production Node server and Microsoft 365 Graph/SSO implementation
evals/                Agent eval harness, task specs, and corpus policy
public/               Static icon/logo assets referenced by the manifest
tests/                Unit tests for adapters, Office/M365 utilities, storage, and memory
tools/                Local Office launchers, cert helpers, production manifest generator
scripts/              Validation scripts
docs/                 Architecture, hardening, and launch documentation
deploy/production/    Docker, env template, hosted deployment, Microsoft 365 rollout docs
manifest.xml          Local/dev Office add-in manifest with SSO placeholders
```

## Current production posture

The codebase now has a production security boundary and fails closed until the organization-specific identity and tenant policy integrations are supplied. The checked-in manifest remains development-only; production assets require a real Entra app registration and HTTPS origin.

Required before production launch:

- replace the manifest `WebApplicationInfo` placeholder with the real Entra app id/resource;
- provide the production bearer-token verifier through `AUTH_VALIDATOR_MODULE` and bind identity to tenant/user-scoped settings and artifacts;
- configure Office SSO and backend Graph On-Behalf-Of exchange with least-privilege scopes;
- move managed BYOK secrets/settings from local/browser storage to tenant/user backend storage and vault-backed secrets;
- configure `BYOK_ALLOWED_TARGETS` plus image and web-fetch allowlists;
- verify `/healthz` and `/readyz`, security headers, rate limits, request limits, and outbound timeouts;
- deploy over a trusted HTTPS domain;
- deploy the generated production manifest through Microsoft 365 Centralized Deployment.
- create and review the release handoff directory with `npm run package:prod` after production manifest generation.

Current verification baseline from the latest local review run:

- 15 test files and 147 tests pass;
- TypeScript compilation passes;
- Vite production build passes;
- development manifest validation passes;
- UI smoke passes in desktop and narrow task-pane layouts;
- production manifest generation rejects missing or placeholder Entra configuration.
