#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")"

PYTHON_BIN=""
for candidate in python3.12 python3.13; do
  if command -v "$candidate" >/dev/null 2>&1; then
    PYTHON_BIN="$candidate"
    break
  fi
done
if [ -z "$PYTHON_BIN" ]; then
  echo "ERROR: instala Python 3.12 o 3.13 antes de continuar."
  exit 1
fi

VENV_DIR=".venv-${PYTHON_BIN}"
if [ ! -d "$VENV_DIR" ]; then "$PYTHON_BIN" -m venv "$VENV_DIR"; fi
source "$VENV_DIR/bin/activate"
python -m pip install --upgrade pip setuptools wheel
python -m pip install --only-binary=:all: -r requirements.txt
python app.py
