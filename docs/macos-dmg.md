# macOS DMG Edition — Architecture and Release Plan

## Status

This document is a build specification, not a claim that a macOS binary currently exists. The Windows v1.1.0 package contains Windows-specific launcher, installer, shortcut, registry, certificate, and Office-host behavior that cannot simply be renamed or wrapped as a DMG.

## Product goal

Ship a signed and notarized universal macOS companion that runs the same two local protocol bridges:

- OpenAI/Codex-compatible endpoint on `127.0.0.1:4001`
- Anthropic/Claude-compatible endpoint on `127.0.0.1:4002`
- Optional Office task-pane plugin service on `https://localhost:3000`

The bridge core should be platform-neutral. The native shell owns configuration, lifecycle, logging, updates, and macOS integration.

## Proposed application bundle

```text
Expedient AI Bridges.app/
├── Contents/
│   ├── Info.plist
│   ├── MacOS/ExpedientAIBridges
│   ├── Resources/
│   │   ├── app/index.js
│   │   ├── app/anthropic-bridge.js
│   │   ├── runtime/node              # universal2 or arch-specific
│   │   ├── office/                   # optional plugin payload
│   │   ├── assets/
│   │   ├── config.env.example
│   │   └── THIRD-PARTY-NOTICES.txt
│   └── _CodeSignature/
```

Do not write mutable files inside the signed bundle. Store them below the user Library:

```text
~/Library/Application Support/Expedient AI Bridges/config.env
~/Library/Application Support/Expedient AI Bridges/runtime/
~/Library/Logs/Expedient AI Bridges/
~/Library/Caches/Expedient AI Bridges/
```

Set configuration permissions to user-only (mode `0600`) and directories to `0700`.

## Native companion responsibilities

1. Resolve bundle resources without assuming the current working directory.
2. Parse and validate configuration without printing credentials.
3. Refuse non-loopback listeners unless an explicit secured mode is configured.
4. Start both bridge children with bounded memory and redirected rotating logs.
5. Track PID plus executable path and process start identity; never kill by stale PID alone.
6. Probe readiness endpoints before reporting healthy.
7. Restart crashed children with bounded exponential backoff.
8. Provide menu-bar actions: status, open configuration, start, stop, restart, diagnostics, Office plugin, quit.
9. Register optional login launch through `SMAppService`; do not install a legacy LaunchAgent silently.
10. Remove login registration and generated trust material during uninstall.

## Runtime strategy

Prefer a pinned Node LTS universal2 binary. Node 25 from the inspected Windows package is not an LTS release. Record the exact runtime version, upstream URL, SHA-256, license, and architecture in the build manifest.

Options:

- **Universal2 app:** one DMG containing arm64 + x86_64 native shell and runtime. Best user experience, largest artifact.
- **Two DMGs:** separate `arm64` and `x64` releases. Smaller artifacts and simpler third-party runtime sourcing.

Never download an executable runtime on first launch without signature and checksum verification.

## Office plugin on macOS

Office remains an optional plugin, not part of the protocol bridge contract. Validate separately on current Microsoft 365 for Mac versions:

- task-pane manifest sideload/deployment behavior;
- trusted localhost HTTPS certificate installation and removal;
- Excel, Word, and PowerPoint Office.js API support gaps;
- file-generation/open flows without Windows COM or registry assumptions;
- Keychain-backed Microsoft identity tokens;
- sandbox and file-picker behavior.

For enterprise deployment, prefer centralized Microsoft 365 add-in deployment rather than modifying Office registration files directly.

## Secrets and Keychain

The compatibility config file may hold non-secret routing settings, but production API credentials should move to macOS Keychain using a service name such as `com.expedient.ai-bridges`. The bridge child should receive credentials through an inherited environment or short-lived IPC channel. Never place keys in command-line arguments, logs, crash reports, plist files, DMG metadata, or GitHub Actions output.

## Signing and notarization

Required Apple assets:

- Apple Developer Program membership;
- Developer ID Application certificate;
- Developer ID Installer certificate only if a PKG is later added;
- App Store Connect API key or notarization keychain profile;
- stable bundle identifier, proposed `com.expedient.ai-bridges`.

Release order:

1. Build universal/native launcher and assemble `.app`.
2. Sign nested executable runtime and helper binaries with hardened runtime.
3. Sign the outer app deeply and verify with `codesign --verify --deep --strict --verbose=2`.
4. Create DMG with Applications symlink and branded background.
5. Sign DMG with Developer ID Application.
6. Submit using `xcrun notarytool submit ... --wait`.
7. Staple using `xcrun stapler staple`.
8. Validate with `spctl --assess --type execute` and `spctl --assess --type open`.
9. Generate SHA-256, SBOM, dependency notices, and provenance attestation.

Never publish an unnotarized artifact as a normal macOS release; Gatekeeper messaging will damage user trust.

## Entitlements

Start with the minimum hardened-runtime entitlements. A normal Developer ID app does not need App Sandbox unless distribution policy requires it. Avoid broad file, automation, microphone, camera, or Apple Events entitlements. If Office automation requires Apple Events, document the exact target and user-facing purpose and isolate it from the core bridge.

## Build tooling

A suitable pipeline can use a native Swift menu-bar app, a .NET 8/9 self-contained universal shell, Tauri, or another auditable shell. The choice must be represented by source in this repository. Do not release a DMG that cannot be reproduced from committed source and pinned dependencies.

Suggested make targets:

```text
make test
make bundle-macos ARCH=arm64
make bundle-macos ARCH=x64
make universal
make sign
make dmg
make notarize
make verify-release
```

## GitHub Actions design

Use a pinned `macos-14` or later runner. Store signing material only in GitHub Actions encrypted secrets or an OIDC-backed secret manager. The workflow should:

1. check out an immutable tag;
2. install pinned toolchains with checksum validation;
3. run bridge translation, lifecycle, and security tests;
4. build both architectures;
5. assemble/sign/notarize/staple;
6. run launch and loopback smoke tests on a clean runner;
7. generate CycloneDX or SPDX SBOM and checksums;
8. upload DMG, checksums, SBOM, and provenance to a draft release;
9. require approval before publishing the release.

Fork-originated pull requests must never receive signing secrets.

## Acceptance gates

A DMG is release-ready only when all are true:

- source for every first-party executable is present;
- clean clone reproduces the unsigned app;
- arm64 and x86_64 launch tests pass;
- ports 4001 and 4002 pass contract tests;
- configuration migration preserves settings and never logs secrets;
- app binds only to loopback by default;
- login-item enable/disable works;
- quit terminates verified child processes;
- Office plugin install/remove is reversible;
- app and DMG pass codesign, Gatekeeper, notarization, and stapling checks;
- license notices, SBOM, SHA-256, privacy/data-flow statement, and release notes ship together.

## Release naming

```text
Expedient-AI-Bridges-1.1.0-universal.dmg
Expedient-AI-Bridges-1.1.0-arm64.dmg
Expedient-AI-Bridges-1.1.0-x64.dmg
SHA256SUMS.txt
Expedient-AI-Bridges-1.1.0.spdx.json
```

## Work still required

The repository now contains the current v1.1.0 compiled bridge bundles and the complete Office plugin source recovered from the local development tree. The remaining implementation step is a cross-platform native shell: define a stable payload/lifecycle contract, implement the macOS menu-bar companion from committed source, and connect it to the same bridge and Office builds. The Windows setup and installed launcher remain behavioral references for matching install, startup, configuration, and uninstall semantics.
