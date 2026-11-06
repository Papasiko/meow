@echo off
cd /d "%~dp0"
if exist "python\python.exe" (
  echo Portable Python is already included. No installation or Internet is required.
  exit /b 0
)
echo ERROR: Portable Python folder is missing. Extract the complete ZIP again.
exit /b 1
