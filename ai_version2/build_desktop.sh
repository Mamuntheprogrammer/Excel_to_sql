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
.venv/bin/python ../setup.py build_exe "$@"
