#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

if [[ ! -x .venv/bin/python ]]; then
  echo "Creating local virtual environment..."
  if ! python3 -m venv .venv; then
    echo "Could not create .venv. On Debian/Ubuntu install python3-venv first."
    exit 1
  fi
fi

.venv/bin/python -m pip install -r requirements-desktop.txt

ICON_ARGS=()
if [[ -f assets/app.icns ]]; then
  ICON_ARGS+=(--icon assets/app.icns)
elif [[ -f assets/app.ico ]]; then
  ICON_ARGS+=(--icon assets/app.ico)
fi

.venv/bin/python -m PyInstaller \
  --noconfirm \
  --clean \
  --onedir \
  --windowed \
  --name "ExcelSQL Studio" \
  --add-data "templates:templates" \
  --add-data "static:static" \
  "${ICON_ARGS[@]}" \
  desktop.py

echo
echo "Desktop build ready: dist/ExcelSQL Studio/"
echo "Remove --windowed from this script if you need to debug startup errors."
