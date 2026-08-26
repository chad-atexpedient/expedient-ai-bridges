@echo off
setlocal
cd /d "%~dp0"
"C:\Users\chad.stthomas\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe" tools\stop-office.mjs
endlocal
