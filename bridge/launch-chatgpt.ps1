$ErrorActionPreference = "Stop"
$installDir = Split-Path -Parent $MyInvocation.MyCommand.Path
& (Join-Path $installDir "run-bridges.ps1")

$env:OPENAI_BASE_URL = "http://127.0.0.1:4001/v1"
if (-not $env:OPENAI_API_KEY) { $env:OPENAI_API_KEY = "sk-local-bridge" }

$candidates = @(
    (Join-Path $env:LOCALAPPDATA "OpenAI\ChatGPT\ChatGPT.exe"),
    (Join-Path $env:LOCALAPPDATA "Programs\ChatGPT\ChatGPT.exe")
)
foreach ($candidate in $candidates) {
    if (Test-Path $candidate) {
        Start-Process $candidate
        exit 0
    }
}

$package = Get-AppxPackage | Where-Object { $_.Name -match "ChatGPT" -or $_.PackageFamilyName -match "ChatGPT" } | Select-Object -First 1
if ($package) {
    [Environment]::SetEnvironmentVariable("OPENAI_BASE_URL", $env:OPENAI_BASE_URL, "User")
    [Environment]::SetEnvironmentVariable("OPENAI_API_KEY", $env:OPENAI_API_KEY, "User")
    Start-Process explorer.exe "shell:AppsFolder\$($package.PackageFamilyName)!App"
    exit 0
}

throw "ChatGPT Desktop is not installed. Install it, then run this shortcut again."
