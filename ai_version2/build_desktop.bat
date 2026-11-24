@echo off
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo Creating local virtual environment...
  python -m venv .venv
  if errorlevel 1 (
    echo Could not create .venv. Make sure Python venv support is installed.
    exit /b 1
  )
)

.venv\Scripts\python -m pip install -r requirements-desktop.txt
if errorlevel 1 exit /b 1

.venv\Scripts\python ..\setup.py build_exe %*
