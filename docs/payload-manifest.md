# v1.1.0 Payload and Publication Manifest

## Release artifacts

The supported Windows artifact is published on the GitHub v1.1.0 release:

- `Expedient-AI-Bridges-Setup.exe`
- Size: `219,711,910 bytes`
- SHA-256: `F8107403263A73CDB30568055EE6F56C8423863F5A508C0F91E2BEA15139B74E`
- Product version: `1.1.0+6635fb02c61a5e5938ef9c8c8a307cc2dac1a9f8`

The setup executable is the distributable package. It contains the program payload needed to install the native companion, current bridge bundles, bundled Node runtime, branding, diagnostics, and optional Office plugin. GitHub also generates source ZIP/TAR archives for the tagged script-edition repository.

## Installed program payload observed

The installed program directory contained 6,687 files (approximately 316 MiB):

| Component | Purpose | Publication route |
|---|---|---|
| `ExpedientBridges.exe` | Native launcher, background companion, chooser, uninstall | Embedded in setup EXE |
| `app/index.js` | Current OpenAI/Codex bridge bundle | Embedded in setup EXE |
| `app/anthropic-bridge.js` | Current Anthropic/Claude bridge bundle | Embedded in setup EXE |
| `runtime/node.exe` | Pinned Windows Node runtime | Embedded in setup EXE |
| `assets/` | Product/provider/Office branding and templates | Embedded in setup EXE; rights review required |
| `office/` | Optional task-pane plugin and local service | Embedded in setup EXE |
| `diagnostics/` | Office E2E diagnostic executable | Embedded in setup EXE; should become optional in a future curated release |
| `config.env.example` | Blank configuration template | Embedded in setup EXE |

The installer and launcher are currently not Authenticode signed. The published checksum is therefore essential until code signing is implemented.

## Why AppData is not uploaded wholesale

`%LOCALAPPDATA%\ExpedientAIBridges` is mutable per-user state, not a release source directory. Uploading it would risk publishing credentials and machine-specific data. The following are intentionally excluded:

- `config.env` — may contain upstream and image API credentials;
- `logs/**` and `*.jsonl` — may contain routing/error metadata;
- `runtime/*pids.json` — stale, machine-specific process state;
- `office/localhost.key.pem` — private TLS key;
- `office/localhost.crt.pem` and certificate thumbprint — generated machine trust state;
- caches, generated Office files, uploads, settings/token stores, dumps, and user documents;
- legacy duplicate `AppData\...\app` payload, which differs from the v1.1.0 installed program.

These files must be generated or preserved locally by the installer. They are not needed to reproduce a user installation and must never be committed or attached to a public release.

## Source provenance boundary

The repository contains the script-based bridge edition. It does not yet contain all first-party source required to reproduce v1.1.0, including:

- unminified current bridge source and source maps;
- native companion/installer project;
- current package build definition and pinned dependency closure;
- complete Office plugin release source as assembled in setup;
- tests and reproducible release pipeline.

The setup EXE is published for installation and inspection with this limitation clearly disclosed. Installed binaries are evidence, not a substitute for missing source.

## Next release gates

1. Recover or recreate all first-party source.
2. Build from an empty allowlist staging directory, never from AppData.
3. Prune Office dependencies to production closure and omit source maps/build tools.
4. Generate SPDX/CycloneDX SBOM and third-party notices, including bundled Node.
5. Confirm redistribution/trademark rights for fonts, logos, and Office templates.
6. Separate local sideload and production Office manifests.
7. Sign and timestamp the launcher, setup, and optional diagnostics.
8. Secret-scan, malware-scan, hash, install-test in a clean profile, and publish provenance.
