@echo off
setlocal

REM Requires administrator privileges to add firewall rule.
net session >nul 2>&1
if %errorlevel% neq 0 (
  echo Please run this file as Administrator.
  pause
  exit /b 1
)

set /p PORT=Enter port to open (1-65535): 
if "%PORT%"=="" (
  echo No port entered.
  pause
  exit /b 1
)

for /f "delims=0123456789" %%a in ("%PORT%") do (
  echo Invalid port.
  pause
  exit /b 1
)

if %PORT% lss 1 (
  echo Invalid port.
  pause
  exit /b 1
)
if %PORT% gtr 65535 (
  echo Invalid port.
  pause
  exit /b 1
)

set RULE_NAME=MSFS Flight Data Export %PORT%

netsh advfirewall firewall show rule name="%RULE_NAME%" >nul 2>&1
if %errorlevel% equ 0 (
  echo Firewall rule already exists: %RULE_NAME%
  pause
  exit /b 0
)

netsh advfirewall firewall add rule name="%RULE_NAME%" dir=in action=allow protocol=TCP localport=%PORT%
if %errorlevel% neq 0 (
  echo Failed to add firewall rule.
  pause
  exit /b 1
)

echo Firewall rule added: %RULE_NAME%
pause
