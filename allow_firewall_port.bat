@echo off
setlocal
chcp 65001 >nul
net session >nul 2>&1
if errorlevel 1 (
  echo ERROR: اين فايل بايد با Run as administrator اجرا شود.
  echo اگر دسترسي مدير نداريد، از مسئول IT بخواهيد پورت TCP 5000 را فقط براي شبکه داخلي باز کند.
  pause
  exit /b 1
)
echo Creating an inbound firewall rule for TCP port 5000 on Domain and Private profiles...
netsh advfirewall firewall add rule name="Org Evaluation System TCP 5000" dir=in action=allow protocol=TCP localport=5000 profile=domain,private
if errorlevel 1 (
  echo ERROR: ساخت قانون Firewall انجام نشد.
  pause
  exit /b 1
)
echo Rule added. Access is still limited by network routing and organizational policy.
pause
