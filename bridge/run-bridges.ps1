param([switch]$Stop)

$ErrorActionPreference = "Stop"
$installDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$dataDir = Join-Path $env:LOCALAPPDATA "ExpedientAIBridges"
$runtimeDir = Join-Path $dataDir "runtime"
$logsDir = Join-Path $dataDir "logs"
$configPath = Join-Path $dataDir "config.env"
$pidPath = Join-Path $runtimeDir "bridge-pids.json"

New-Item -ItemType Directory -Force -Path $runtimeDir, $logsDir | Out-Null

function Get-TrackedProcess($entry) {
    if (-not $entry) { return $null }
    $id = if ($entry -is [int] -or $entry -is [long]) { [int]$entry } else { [int]$entry.pid }
    if (-not $id) { return $null }
    $process = Get-Process -Id $id -ErrorAction SilentlyContinue
    if (-not $process) { return $null }
    if ($entry -isnot [int] -and $entry -isnot [long] -and $entry.started) {
        try {
            $expected = [datetime]::Parse($entry.started).ToUniversalTime()
            $actual = $process.StartTime.ToUniversalTime()
            if ([math]::Abs(($actual - $expected).TotalSeconds) -gt 2) { return $null }
        } catch { return $null }
    }
    return $process
}

function Stop-Bridges {
    if (-not (Test-Path $pidPath)) { return }
    try {
        $tracked = Get-Content $pidPath -Raw | ConvertFrom-Json
        foreach ($entry in @($tracked.codex, $tracked.anthropic)) {
            $process = Get-TrackedProcess $entry
            if ($process) { Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue }
        }
    } finally {
        Remove-Item $pidPath -Force -ErrorAction SilentlyContinue
    }
}

if ($Stop) { Stop-Bridges; exit 0 }
if (-not (Test-Path $configPath)) { throw "Missing configuration: $configPath" }

foreach ($line in Get-Content $configPath) {
    $line = $line.Trim()
    if (-not $line -or $line.StartsWith("#")) { continue }
    $parts = $line -split "=", 2
    if ($parts.Count -eq 2 -and $parts[0].Trim() -match '^[A-Za-z_][A-Za-z0-9_]*$') {
        [Environment]::SetEnvironmentVariable($parts[0].Trim(), $parts[1].Trim(), "Process")
    }
}
if (-not $env:OPENAI_BASE_URL) { throw "OPENAI_BASE_URL must be configured in $configPath" }

$node = Get-Command node.exe -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source
if (-not $node) {
    $bundled = Join-Path (Split-Path $installDir -Parent) "runtime\node.exe"
    if (Test-Path $bundled) { $node = $bundled }
}
if (-not $node) {
    $root = Join-Path $env:LOCALAPPDATA "OpenAI\Codex\runtimes"
    $node = Get-ChildItem $root -Recurse -Filter node.exe -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
}
if (-not $node) { throw "Node.js 22+ or Codex Desktop is required." }
$major = & $node -p "Number(process.versions.node.split('.')[0])"
if ([int]$major -lt 22) { throw "Node.js 22+ is required; found $(& $node --version)." }

$hostName = if ($env:BRIDGE_HOST) { $env:BRIDGE_HOST } else { "127.0.0.1" }
if ($hostName -notin @("127.0.0.1", "::1", "localhost") -and $env:ALLOW_NON_LOOPBACK -ne "1") {
    throw "Refusing non-loopback BRIDGE_HOST=$hostName without ALLOW_NON_LOOPBACK=1."
}
$ports = @(
    if ($env:PORT) { [int]$env:PORT } else { 4001 },
    if ($env:ANTHROPIC_PORT) { [int]$env:ANTHROPIC_PORT } else { 4002 }
)
foreach ($port in $ports) {
    $owner = Get-NetTCPConnection -State Listen -LocalPort $port -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($owner) { throw "Port $port is already in use by PID $($owner.OwningProcess). Stop that process before starting the bridges." }
}

Stop-Bridges
$env:NODE_OPTIONS = "--max-old-space-size=8192"
$env:BRIDGE_METRICS_DIR = $logsDir
$stdout = Join-Path $logsDir "bridges-output.log"
$stderr = Join-Path $logsDir "bridges-error.log"
$codexScript = Join-Path $installDir "app\index.js"
$anthropicScript = Join-Path $installDir "app\anthropic-bridge.js"
$codex = Start-Process $node -ArgumentList @($codexScript) -WindowStyle Hidden -RedirectStandardOutput $stdout -RedirectStandardError $stderr -PassThru
$anthropic = Start-Process $node -ArgumentList @($anthropicScript) -WindowStyle Hidden -RedirectStandardOutput $stdout -RedirectStandardError $stderr -PassThru
@{
    codex = @{ pid = $codex.Id; started = $codex.StartTime.ToUniversalTime().ToString("o"); script = $codexScript }
    anthropic = @{ pid = $anthropic.Id; started = $anthropic.StartTime.ToUniversalTime().ToString("o"); script = $anthropicScript }
} | ConvertTo-Json -Depth 4 | Set-Content $pidPath

$deadline = (Get-Date).AddSeconds(10)
do {
    Start-Sleep -Milliseconds 250
    if ($codex.HasExited -or $anthropic.HasExited) { throw "A bridge exited during startup. See $stderr" }
    $ready = $true
    foreach ($port in $ports) {
        if (-not (Get-NetTCPConnection -State Listen -LocalPort $port -ErrorAction SilentlyContinue)) { $ready = $false }
    }
} until ($ready -or (Get-Date) -ge $deadline)
if (-not $ready) { Stop-Bridges; throw "Bridge readiness timed out. See $stderr" }
