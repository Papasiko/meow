@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title Org Evaluation - Offline Package Builder

REM Delegate all logic to PowerShell to avoid CMD parser/parenthesis issues.
if not exist "%~dp0prepare_offline_packages.ps1" (
  echo ERROR: prepare_offline_packages.ps1 is missing.
  pause
  exit /b 1
)

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0prepare_offline_packages.ps1"
set "RESULT=%ERRORLEVEL%"
if not "%RESULT%"=="0" (
  echo.
  echo Preparation failed with exit code %RESULT%.
  echo Please take a screenshot of the error above.
)
pause
exit /b %RESULT%
