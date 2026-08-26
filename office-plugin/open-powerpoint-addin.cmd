@echo off
setlocal
cd /d "%~dp0"
"C:\Users\chad.stthomas\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe" tools\certs.mjs install
"C:\Users\chad.stthomas\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe" tools\open-office.mjs powerpoint
if errorlevel 1 pause
endlocal
