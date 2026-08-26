param()
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$expected = @{
    "bridge/app/index.js" = "09C1A9BB9C1156916BF4D03AA61C879342E8329F196C70C4D7761844A8D107C2"
    "bridge/app/anthropic-bridge.js" = "936D6F959DC486582C2A8F965F0CAB07D94ADED3C04BDDC10AE5ABCF3F780615"
    "bridge/config.env.example" = "1975B4BE1F8730FF52A5E344816205F565A73168279C8EB0DF4EAE2208EDD3F6"
}
foreach ($relative in $expected.Keys) {
    $path = Join-Path $root ($relative -replace '/', '\')
    if (-not (Test-Path $path)) { throw "Missing $relative" }
    $actual = (Get-FileHash $path -Algorithm SHA256).Hash
    if ($actual -ne $expected[$relative]) { throw "Hash mismatch for ${relative}: $actual" }
}
$requiredConfig = @("OPENAI_BASE_URL", "BRIDGE_HOST", "PORT", "ANTHROPIC_PORT", "MAX_REQUEST_BODY_BYTES", "MAX_RESPONSES_IN_FLIGHT", "MAX_ANTHROPIC_IN_FLIGHT", "UPSTREAM_API_KEY")
$config = Get-Content (Join-Path $root "bridge\config.env.example")
foreach ($name in $requiredConfig) {
    if (-not ($config -match "^$name=")) { throw "Config example is missing $name" }
}
$forbidden = @(("github" + "_pat_"), ("gh" + "p_"), ("BEGIN " + "PRIVATE KEY"), ("BEGIN RSA " + "PRIVATE KEY"), ("BEGIN OPENSSH " + "PRIVATE KEY"))
$files = Get-ChildItem $root -Recurse -File | Where-Object { $_.FullName -notmatch '[\\/]\.git[\\/]' -and $_.Length -lt 5MB }
foreach ($needle in $forbidden) {
    $match = Select-String -Path $files.FullName -SimpleMatch $needle -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($match) { throw "Forbidden secret marker '$needle' in $($match.Path):$($match.LineNumber)" }
}
Write-Host "Bridge release verification passed."
