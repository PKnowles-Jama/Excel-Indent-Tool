@echo off

where pip > nul 2>&1
if %ERRORLEVEL% == 1 (
  echo pip is not installed. Please install pip first.
  exit /b 1
)

if not exist requirements.txt (
  echo requirements.txt not found in the current directory.
  exit /b 1
)

pip install -r requirements.txt

echo Requirements installed successfully!