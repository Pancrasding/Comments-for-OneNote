@echo off
setlocal EnableExtensions EnableDelayedExpansion
title Comments for OneNote - Install
cd /d "%~dp0"

set "SILENT=0"
if /I "%~1"=="--silent" set "SILENT=1"

echo Installing Comments for OneNote...
if "!SILENT!"=="1" (
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1" -NoLaunch
) else (
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1"
)
set "EXIT_CODE=!ERRORLEVEL!"

echo.
if not "!EXIT_CODE!"=="0" (
  echo Installation failed. See the message above.
) else (
  echo Comments for OneNote installed successfully.
)

if "!SILENT!"=="0" pause
exit /b !EXIT_CODE!