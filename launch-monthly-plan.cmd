@echo off
setlocal
set "ROOT=%~dp0"
start "ITI Monthly Plan Export Server" "C:\Users\akash\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe" "%ROOT%server.mjs"
timeout /t 2 /nobreak >nul
start "" "%ROOT%index.html"
