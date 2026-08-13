# Expedient AI Bridges for Windows

Run `install.bat`, set `UPSTREAM_API_KEY` in `%LOCALAPPDATA%\ExpedientAIBridges\config.env`, then use either branded desktop shortcut.

The installer:

- installs compiled bridge files under `%LOCALAPPDATA%\ExpedientAIBridges\app`
- preserves an existing configuration during upgrades
- runs both bridges silently at sign-in
- creates branded Codex and Claude Code desktop shortcuts
- attempts to place both shortcuts in the Windows taskbar-pins folder

Windows policy may block silent taskbar pinning. If needed, right-click a desktop shortcut and select **Pin to taskbar**.

Requirements: Windows 10/11 and either Node.js 22+ or Codex Desktop with its bundled Node runtime.

For unattended setup, run `install.bat -Silent` after configuring the key. Full configuration and AI-assisted installation instructions are in the repository README.
