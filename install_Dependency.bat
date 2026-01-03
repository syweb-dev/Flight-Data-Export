@echo off
setlocal

python -m pip --version >nul 2>&1
if %errorlevel% neq 0 (
  echo pip not found. Please install Python and check "Add Python to PATH".
  pause
  exit /b 1
)

python -m pip install --upgrade pip
python -m pip install SimConnect
python -m pip install pyuipc
if %errorlevel% neq 0 (
  echo Failed to install SimConnect.
  pause
  exit /b 1
)

echo SimConnect installed successfully.
pause
