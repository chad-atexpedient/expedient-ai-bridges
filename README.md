# Expedient AI Bridges for Windows

Local compatibility bridges connecting OpenAI/Codex- and Anthropic/Claude-oriented clients to an Expedient or other OpenAI-compatible upstream.

> **Scope:** this repository contains the script-based bridge edition. The separately installed Windows companion may bundle newer bridge builds and optional plugins that are not reproducible from this source tree.

## Architecture report

Explore the interactive bridge-first [project site](https://chad-atexpedient.github.io/expedient-ai-bridges/), then open the [full technical report](https://chad-atexpedient.github.io/expedient-ai-bridges/architecture-report.html).

## Core bridge endpoints

| Listener | Default | Compatible routes |
|---|---|---|
| OpenAI/Codex | `http://127.0.0.1:4001` | Responses, Chat Completions, Models, Images |
| Anthropic/Claude | `http://127.0.0.1:4002` | Messages, Messages token count |

The bridge translates message, tool, model, and streaming formats before forwarding requests to `OPENAI_BASE_URL`. It does not run an AI model locally. Microsoft Office support in the packaged desktop product is an **optional plugin feature**, not the core bridge.

## Requirements

- Windows 10 or 11
- Node.js 22+ or Codex Desktop with its bundled Node runtime
- An OpenAI-compatible upstream URL and API key
- Claude CLI on `PATH` if using the Claude shortcut

## Install

1. Run `install.bat`.
2. Open `%LOCALAPPDATA%\ExpedientAIBridges\config.env`.
3. Set at minimum:

```dotenv
OPENAI_BASE_URL=https://your-openai-compatible-service.example/v1
UPSTREAM_API_KEY=your-key-here
```

4. Start from a generated shortcut or run `run-bridges.ps1`.

Stop the script edition with:

```powershell
powershell -ExecutionPolicy Bypass -File "$env:LOCALAPPDATA\ExpedientAIBridges\app\run-bridges.ps1" -Stop
```

## Configuration

The template is `bridge/config.env.example`. Important settings include `OPENAI_BASE_URL`, `UPSTREAM_API_KEY`, `BRIDGE_HOST`, `PORT`, `ANTHROPIC_PORT`, `LOG_LEVEL`, and optional `IMAGE_GEN_*` settings.

Keep listeners on loopback unless you add appropriate authentication and transport security. Never commit `config.env`, logs, certificates, private keys, or API tokens.

## Installer behavior

The installer copies the payload to `%LOCALAPPDATA%\ExpedientAIBridges\app`, preserves external configuration, protects its ACL, creates shortcuts, and launches both bridge processes.

## Source and release status

The files in `bridge/app` are compiled/minified runtime bundles. This repository currently lacks their unminified source, source maps, dependency manifest, tests, and source for the newer native Windows companion. See the report for the inspected scope and release recommendations.

## Releases and platform plans

- [Windows v1.1.0 setup release](https://github.com/chad-atexpedient/expedient-ai-bridges/releases/tag/v1.1.0)
- [Installed payload publication manifest](docs/payload-manifest.md)
- [macOS DMG architecture and release plan](docs/macos-dmg.md)

Mutable AppData is deliberately excluded from source and release attachments because it contains machine-specific state and can contain credentials, logs, PID records, certificates, and private keys. The setup executable contains the distributable installed program payload.

## License

MIT — see `bridge/LICENSE`.
