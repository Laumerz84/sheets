"""Start Sheets. Works when double-clicked with any Python: re-launches itself
with the project's virtual environment if PySide6 isn't available."""
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
VENV_PYW = os.path.join(HERE, ".venv", "Scripts", "pythonw.exe")

try:
    import PySide6  # noqa: F401
except ImportError:
    if os.path.exists(VENV_PYW) and os.path.normcase(sys.executable) != os.path.normcase(VENV_PYW):
        subprocess.Popen([VENV_PYW, os.path.abspath(__file__)] + sys.argv[1:], cwd=HERE)
        sys.exit(0)
    raise

sys.path.insert(0, HERE)
from sheets.app import main  # noqa: E402

sys.exit(main(sys.argv[1:]))
