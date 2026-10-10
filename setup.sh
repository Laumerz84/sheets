#!/bin/bash
# Ekxel setup for macOS (Apple Silicon or Intel) and Linux: makes .venv next to this file,
# installs the libraries into it, then checks everything. Safe to run again.
# Nothing is installed system-wide.
set -e
cd "$(dirname "$0")"

PY=""
for c in python3.14 python3.13 python3.12 python3.11 python3; do
  if command -v "$c" >/dev/null 2>&1 && "$c" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 11) else 1)' 2>/dev/null; then
    PY="$c"
    break
  fi
done
if [ -z "$PY" ]; then
  echo "Python 3.11 or newer not found (the one that comes with macOS is too old)."
  echo "Install it with Homebrew:  brew install python@3.12   (or from https://www.python.org/downloads/macos/)"
  echo "then run this again."
  exit 1
fi

if [ ! -x .venv/bin/python ]; then
  echo "Creating .venv with $PY ..."
  "$PY" -m venv .venv
fi
echo "Installing libraries ..."
PIP_CACHE_DIR="$PWD/.pip-cache" .venv/bin/python -m pip install --disable-pip-version-check -q -r requirements.txt
chmod +x Ekxel.command 2>/dev/null || true
.venv/bin/python tools/doctor.py
