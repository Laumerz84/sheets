#!/bin/bash
# macOS: double-click in Finder to start Ekxel (run ./setup.sh once first).
# Files: drag them onto this, or use File > Open in Ekxel.
cd "$(dirname "$0")"
if [ ! -x .venv/bin/python ]; then
  echo "Ekxel isn't set up yet. In Terminal, run:  cd \"$PWD\" && ./setup.sh"
  read -r -n 1 -p "Press any key to close."
  exit 1
fi
# Start detached so closing this Terminal window doesn't close Ekxel.
nohup .venv/bin/python launch.pyw "$@" >/dev/null 2>&1 &
disown
exit 0
