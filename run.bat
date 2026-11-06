@echo off
setlocal
cd /d "%~dp0"
title Organizational Evaluation System
if not exist "python\python.exe" (
  echo ERROR: Bundled Python runtime is missing. Extract the entire ZIP again.
  pause
  exit /b 1
)
"python\python.exe" -c "import flask, openpyxl, waitress" >nul 2>&1
if errorlevel 1 (
  echo ERROR: Bundled Python libraries could not be loaded.
  pause
  exit /b 1
)
if not defined HOST set "HOST=0.0.0.0"
if not defined PORT set "PORT=5000"
if not defined WAITRESS_THREADS set "WAITRESS_THREADS=8"
"python\python.exe" run_server.py
if errorlevel 1 pause
