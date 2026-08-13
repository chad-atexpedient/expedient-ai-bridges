$ErrorActionPreference = "Stop"
$installDir = Split-Path -Parent $MyInvocation.MyCommand.Path
& (Join-Path $installDir "run-bridges.ps1")
$env:ANTHROPIC_BASE_URL = "http://127.0.0.1:4002"
if (-not $env:ANTHROPIC_API_KEY) { $env:ANTHROPIC_API_KEY = "sk-local-bridge" }
Start-Process powershell.exe -ArgumentList "-NoExit", "-Command", "claude" -WorkingDirectory $HOME
