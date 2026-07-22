@echo off
setlocal
set ROOT=%~dp0..
set PY=C:\Users\dpd\Documents\Codex\1C_Diagnostics\.venv\Scripts\python.exe
if not exist "%PY%" (
  echo Python venv not found: %PY%
  exit /b 2
)
"%PY%" "%~dp0tools\quick_stock_lookup.py" %*
endlocal

