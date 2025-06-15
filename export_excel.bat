@echo off
echo Executing...

powershell -ExecutionPolicy Bypass -NoProfile -File "%~dp0export_excel.ps1"

timeout /t 5 >nul
