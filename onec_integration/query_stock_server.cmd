@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0query_stock_server.ps1" %*
