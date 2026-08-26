# Cross-platform companion

This is the source-buildable replacement shell for the native lifecycle behavior distributed in the Windows v1.1.0 setup package.

## Current commands

```text
ExpedientBridges start
ExpedientBridges stop
ExpedientBridges restart
ExpedientBridges status
ExpedientBridges diagnostics
ExpedientBridges --background
ExpedientBridges --choose
```

The companion:

- resolves mutable data to `%LOCALAPPDATA%\ExpedientAIBridges` on Windows;
- resolves mutable data to `~/Library/Application Support/Expedient AI Bridges` on macOS;
- loads the external `config.env` without logging secrets;
- refuses non-loopback binding unless explicitly allowed;
- requires Node.js 22+ and prefers a bundled runtime;
- launches the exact bridge bundles under `app/`;
- records PID and process-start identity;
- checks port readiness and emits redacted diagnostics.

## Build

```powershell
dotnet build src/ExpedientBridges/ExpedientBridges.csproj -c Release
dotnet publish src/ExpedientBridges/ExpedientBridges.csproj -c Release -r win-x64
dotnet publish src/ExpedientBridges/ExpedientBridges.csproj -c Release -r osx-arm64
dotnet publish src/ExpedientBridges/ExpedientBridges.csproj -c Release -r osx-x64
```

## Production hardening backlog

Before replacing the released Windows companion or creating a DMG:

1. make `--background` a persistent single-instance supervisor with a user-only control channel;
2. atomically persist supervisor and child state, including canonical executable/script paths;
3. use application-level HTTP readiness probes rather than TCP-only checks;
4. add graceful shutdown, bounded restart backoff, and size-rotated redacted logs;
5. integrate optional Office lifecycle with explicit loopback TLS and `/healthz`/`/readyz` checks;
6. add Windows tray and macOS menu-bar UX;
7. sign, notarize, staple, generate SBOM/provenance, and run clean-profile installation tests.
