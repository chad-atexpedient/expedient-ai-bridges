param([switch]$Stop)

$ErrorActionPreference = "Stop"
$installDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$runtimeDir = Join-Path $env:LOCALAPPDATA "ExpedientAIBridges\runtime"
$configPath = Join-Path $env:LOCALAPPDATA "ExpedientAIBridges\config.env"
$pidPath = Join-Path $runtimeDir "bridge-pids.json"

New-Item -ItemType Directory -Force -Path $runtimeDir | Out-Null

function Stop-Bridges {
    if (Test-Path $pidPath) {
        $pids = Get-Content $pidPath -Raw | ConvertFrom-Json
        foreach ($id in @($pids.codex, $pids.anthropic)) {
            if ($id) { Stop-Process -Id $id -Force -ErrorAction SilentlyContinue }
        }
        Remove-Item $pidPath -Force -ErrorAction SilentlyContinue
    }
}

if ($Stop) { Stop-Bridges; exit 0 }

if (-not (Test-Path $configPath)) { throw "Missing configuration: $configPath" }
foreach ($line in Get-Content $configPath) {
    $line = $line.Trim()
    if (-not $line -or $line.StartsWith("#")) { continue }
    $parts = $line -split "=", 2
    if ($parts.Count -eq 2) { [Environment]::SetEnvironmentVariable($parts[0].Trim(), $parts[1].Trim(), "Process") }
}

$node = Get-Command node.exe -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source
if (-not $node) {
    $root = Join-Path $env:LOCALAPPDATA "OpenAI\Codex\runtimes"
    $node = Get-ChildItem $root -Recurse -Filter node.exe -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
}
if (-not $node) { throw "Node.js 22+ or Codex Desktop is required." }

Stop-Bridges
$env:NODE_OPTIONS = "--max-old-space-size=8192"
$codex = Start-Process $node -ArgumentList (Join-Path $installDir "app\index.js") -WindowStyle Hidden -PassThru
$anthropic = Start-Process $node -ArgumentList (Join-Path $installDir "app\anthropic-bridge.js") -WindowStyle Hidden -PassThru
@{ codex = $codex.Id; anthropic = $anthropic.Id } | ConvertTo-Json | Set-Content $pidPath

Start-Sleep -Seconds 2
if ($codex.HasExited -or $anthropic.HasExited) { throw "A bridge failed to start. Check Windows Event Viewer or run the script manually for diagnostics." }
