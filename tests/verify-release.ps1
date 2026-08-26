param()
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$expectedBlobs = @{
    "bridge/app/index.js" = "89675cdaa0b5f977cd96dfa6efe40518866b8048"
    "bridge/app/anthropic-bridge.js" = "6356e24b3a69dbb69bbc3f980dc1a74c3f1a94e2"
    "bridge/config.env.example" = "a31685d8d547a0d4c9e3f567099663c6baf3d891"
}
foreach ($relative in $expectedBlobs.Keys) {
    $path = Join-Path $root ($relative -replace '/', '\')
    if (-not (Test-Path $path)) { throw "Missing $relative" }
    $actual = (& git -C $root rev-parse "HEAD:$relative").Trim()
    if ($actual -ne $expectedBlobs[$relative]) { throw "Git blob mismatch for ${relative}: $actual" }
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
