r"""Start Ekxel (package name: Sheets). Works when started with any Python: re-launches itself
with the project's virtual environment if PySide6 isn't available.
Windows: double-click, or .venv\Scripts\pythonw.exe launch.pyw.  macOS: double-click Ekxel.command."""
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if sys.platform == "win32":
    VENV_PY = os.path.join(HERE, ".venv", "Scripts", "pythonw.exe")
else:
    VENV_PY = os.path.join(HERE, ".venv", "bin", "python")

try:
    import PySide6  # noqa: F401
except ImportError:
    if os.path.exists(VENV_PY) and os.path.normcase(sys.executable) != os.path.normcase(VENV_PY):
        subprocess.Popen([VENV_PY, os.path.abspath(__file__)] + sys.argv[1:], cwd=HERE)
        sys.exit(0)
    raise

sys.path.insert(0, HERE)
from sheets.app import main  # noqa: E402

sys.exit(main(sys.argv[1:]))
