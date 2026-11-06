@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
 echo ابتدا install_offline.bat را اجرا کنيد.
 pause
 exit /b 1
)
".venv\Scripts\python.exe" backup_db.py
if errorlevel 1 echo خطا در بکاپ. مطمئن شو سامانه یک بار اجرا شده و پایگاه داده ساخته شده است.
pause
