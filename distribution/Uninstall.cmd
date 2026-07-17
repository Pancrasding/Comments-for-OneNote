@echo off
setlocal EnableExtensions EnableDelayedExpansion
title Comments for OneNote - Uninstall
cd /d "%~dp0"

set "SILENT=0"
if /I "%~1"=="--silent" set "SILENT=1"

echo Removing Comments for OneNote...
if "!SILENT!"=="1" (
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0uninstall.ps1" -NoLaunch
) else (
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0uninstall.ps1"
)
set "EXIT_CODE=!ERRORLEVEL!"

echo.
if not "!EXIT_CODE!"=="0" (
  echo Uninstall failed. See the message above.
) else (
  echo Comments for OneNote was removed successfully.
  echo The official OneMore add-in is active again.
)

if "!SILENT!"=="0" pause
exit /b !EXIT_CODE!