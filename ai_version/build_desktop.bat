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

set "ICON="
if exist "assets\app.ico" set "ICON=--icon assets\app.ico"

.venv\Scripts\python -m PyInstaller --noconfirm --clean --onedir --windowed --name "ExcelSQL Studio" --add-data "templates;templates" --add-data "static;static" %ICON% desktop.py
if errorlevel 1 exit /b 1

echo.
echo Desktop build ready: dist\ExcelSQL Studio\
echo Remove --windowed from this script if you need to debug startup errors.
