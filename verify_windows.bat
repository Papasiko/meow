@echo off
setlocal EnableExtensions
chcp 65001 >nul
cd /d "%~dp0"
echo ====================================================
echo بررسی سامانه ارزیابی - Windows / Python 3.12
echo ====================================================
if not exist ".venv\Scripts\python.exe" (
  echo محیط مجازی پیدا نشد. ابتدا install_offline.bat را اجرا کنید.
  pause
  exit /b 1
)
".venv\Scripts\python.exe" -c "import sys,flask,openpyxl,waitress; print('Python:',sys.version); print('Flask:',flask.__version__); print('openpyxl:',openpyxl.__version__); print('Waitress installed')"
if errorlevel 1 goto failed
".venv\Scripts\python.exe" -m pip check
if errorlevel 1 goto failed
".venv\Scripts\python.exe" -m py_compile app.py run_server.py wsgi.py backup_db.py
if errorlevel 1 goto failed
if not exist "wheelhouse\*.whl" echo NOTE: wheelhouse is empty; the current installation may be working, but cannot be re-installed offline from this folder.
echo.
echo Static and dependency checks passed.
echo For live check, run run_windows.bat and visit http://127.0.0.1:5000/api/health.
pause
exit /b 0
:failed
echo.
echo Check failed. Please capture the full error message.
pause
exit /b 1
