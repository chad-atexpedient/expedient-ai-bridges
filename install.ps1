param([switch]$Silent)

$ErrorActionPreference = "Stop"
$source = Join-Path $PSScriptRoot "bridge"
$installDir = Join-Path $env:LOCALAPPDATA "ExpedientAIBridges\app"
$dataDir = Join-Path $env:LOCALAPPDATA "ExpedientAIBridges"
$configPath = Join-Path $dataDir "config.env"

New-Item -ItemType Directory -Force -Path $installDir, $dataDir | Out-Null
Copy-Item "$source\*" $installDir -Recurse -Force
if (-not (Test-Path $configPath)) { Copy-Item (Join-Path $source "config.env.example") $configPath }
$acl = Get-Acl $configPath
$acl.SetAccessRuleProtection($true, $false)
$currentUserSid = [Security.Principal.WindowsIdentity]::GetCurrent().User
$acl.SetAccessRule((New-Object Security.AccessControl.FileSystemAccessRule($currentUserSid, "FullControl", "Allow")))
$acl.SetAccessRule((New-Object Security.AccessControl.FileSystemAccessRule("SYSTEM", "FullControl", "Allow")))
Set-Acl $configPath $acl

$logo = Join-Path $installDir "assets\expbrain.png"
$icon = Join-Path $installDir "assets\expbrain.ico"
Add-Type -AssemblyName System.Drawing
$bitmap = [System.Drawing.Bitmap]::FromFile($logo)
$handle = $bitmap.GetHicon()
$ico = [System.Drawing.Icon]::FromHandle($handle)
$stream = [IO.File]::Create($icon)
$ico.Save($stream)
$stream.Close(); $ico.Dispose(); $bitmap.Dispose()

$shell = New-Object -ComObject WScript.Shell
function New-Link($path, $target, $arguments, $description) {
    $link = $shell.CreateShortcut($path)
    $link.TargetPath = "powershell.exe"
    $link.Arguments = "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$target`" $arguments"
    $link.WorkingDirectory = $installDir
    $link.IconLocation = "$icon,0"
    $link.Description = $description
    $link.Save()
}

$desktop = [Environment]::GetFolderPath("Desktop")
New-Link (Join-Path $desktop "Codex - Expedient AI.lnk") (Join-Path $installDir "launch-codex.ps1") "" "Start the bridges and open Codex"
New-Link (Join-Path $desktop "Claude Code - Expedient AI.lnk") (Join-Path $installDir "launch-claude.ps1") "" "Start the bridges and open Claude Code"
New-Link (Join-Path $desktop "ChatGPT - Expedient AI.lnk") (Join-Path $installDir "launch-chatgpt.ps1") "" "Start the bridges and open ChatGPT"
New-Link (Join-Path ([Environment]::GetFolderPath("Startup")) "Expedient AI Bridges.lnk") (Join-Path $installDir "run-bridges.ps1") "" "Start Expedient AI bridges silently"

$taskbar = Join-Path $env:APPDATA "Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
New-Item -ItemType Directory -Force -Path $taskbar | Out-Null
Copy-Item (Join-Path $desktop "Codex - Expedient AI.lnk") $taskbar -Force
Copy-Item (Join-Path $desktop "Claude Code - Expedient AI.lnk") $taskbar -Force
Copy-Item (Join-Path $desktop "ChatGPT - Expedient AI.lnk") $taskbar -Force

& (Join-Path $installDir "run-bridges.ps1")
if (-not $Silent) { Write-Host "Installed. Configure your key in $configPath" }
