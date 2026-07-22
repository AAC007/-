@echo off
setlocal
set PY=C:\Users\dpd\Documents\Codex\1C_Diagnostics\.venv\Scripts\python.exe
if not exist "%PY%" (
  echo Python venv not found: %PY%
  exit /b 2
)
start "1C stock lookup server" /min "%PY%" "%~dp0tools\stock_lookup_server.py"
endlocal
