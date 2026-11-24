@echo off
cd /d "%~dp0"

python -c "import flask" 2>NUL
if errorlevel 1 (
  echo Installing Flask...
  python -m pip install -r requirements.txt
  if errorlevel 1 exit /b 1
)

python app.py %*
