$ErrorActionPreference = "Stop"
$installDir = Split-Path -Parent $MyInvocation.MyCommand.Path
& (Join-Path $installDir "run-bridges.ps1")
$codex = Join-Path $env:LOCALAPPDATA "OpenAI\Codex\Codex.exe"
if (-not (Test-Path $codex)) { $codex = Get-ChildItem (Join-Path $env:LOCALAPPDATA "OpenAI\Codex") -Recurse -Filter codex.exe -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName }
if (-not $codex) { throw "Codex Desktop is not installed." }
Start-Process $codex
