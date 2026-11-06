@echo off
setlocal
chcp 65001 >nul
net session >nul 2>&1
if errorlevel 1 (
  echo اين فايل را با Run as administrator اجرا کنيد.
  pause
  exit /b 1
)
netsh advfirewall firewall delete rule name="Org Evaluation System TCP 5000"
pause
