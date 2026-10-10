r"""Check that Ekxel and its Claude panel are set up. Read-only: installs and changes nothing.

usage: Windows  .venv\Scripts\python.exe tools\doctor.py
       macOS    .venv/bin/python tools/doctor.py
Prints one PASS / FAIL / NOTE line per check, with the fix for each FAIL. Exit code 0 = ready."""
import importlib
import os
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
IS_WIN = sys.platform == "win32"
IS_MAC = sys.platform == "darwin"
VENV_PY = os.path.join(ROOT, ".venv", "Scripts", "python.exe") if IS_WIN else os.path.join(ROOT, ".venv", "bin", "python")
SETUP = "setup.bat" if IS_WIN else "./setup.sh"
if IS_WIN:
    INSTALL_CLAUDE = "in PowerShell run  irm https://claude.ai/install.ps1 | iex"
else:
    INSTALL_CLAUDE = "in Terminal run  curl -fsSL https://claude.ai/install.sh | bash"
results = []


def check(ok, label, fix=""):
    results.append(ok)
    print(f"{'PASS' if ok else 'FAIL'}  {label}")
    if not ok and fix:
        print(f"      FIX: {fix}")


def note(label):
    print(f"NOTE  {label}")


def main():
    print(f"Ekxel setup check ({ROOT}, {sys.platform})\n")
    py_fix = ("Install Python 3.11 or newer (winget install Python.Python.3.12)" if IS_WIN else
              "Install Python 3.11 or newer (brew install python@3.12)") + f", then run {SETUP} again."
    check(sys.version_info >= (3, 11), f"Python {sys.version.split()[0]} (need 3.11+)", py_fix)
    in_venv = os.path.normcase(os.path.abspath(sys.executable)).startswith(
        os.path.normcase(os.path.join(ROOT, ".venv")))
    check(in_venv, "running from the repo's .venv", f"Run {SETUP}, then run this with {VENV_PY}")

    packages = [("PySide6", "PySide6"), ("openpyxl", "openpyxl"), ("xlrd", "xlrd"), ("mcp", "mcp")]
    if IS_WIN:
        packages.append(("win32com", "pywin32"))  # taskbar name/icon; Windows only
    for mod, pkg in packages:
        try:
            importlib.import_module(mod)
            check(True, f"package {pkg}")
        except Exception as e:  # noqa: BLE001
            check(False, f"package {pkg} ({type(e).__name__}: {e})",
                  f'"{VENV_PY}" -m pip install -r "{os.path.join(ROOT, "requirements.txt")}"')

    try:
        import platform

        from PySide6.QtCore import qVersion
        note(f"Qt {qVersion()} on {platform.machine()}")
    except Exception:  # noqa: BLE001
        pass

    try:
        importlib.import_module("sheets.ai.mcp_server")
        check(True, "Claude bridge (sheets/ai/mcp_server.py) imports")
    except Exception as e:  # noqa: BLE001
        check(False, f"Claude bridge doesn't import ({type(e).__name__}: {e})",
              "Usually the mcp package is missing: re-run the pip install above.")

    from sheets.ai.panel import find_claude
    cmd = find_claude()
    if not cmd:
        check(False, "Claude Code (claude) found",
              f"Install Claude Code: {INSTALL_CLAUDE}  (it goes to ~/.local/bin, which Ekxel finds even "
              "off PATH). Or set the env var SHEETS_CLAUDE_CMD to a JSON list like [\"/path/to/claude\"].")
    else:
        try:
            out = subprocess.run(cmd + ["--version"], capture_output=True, text=True, timeout=60,
                                 creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
            ver = (out.stdout or out.stderr).strip().splitlines()[0] if (out.stdout or out.stderr) else "?"
            check(out.returncode == 0, f"Claude Code runs: {cmd[0]} ({ver})",
                  "Reinstall Claude Code (see above) and try `claude --version` in a terminal.")
        except Exception as e:  # noqa: BLE001
            check(False, f"Claude Code at {cmd[0]} doesn't run ({e})", f"Reinstall Claude Code: {INSTALL_CLAUDE}")
    e2e = os.path.join("tools", "e2e_claude.py")
    note("Sign-in can't be checked without using Claude: the user must have run `claude` once in a "
         "terminal and logged in. To test the whole chain (uses a little Claude usage, ask first): "
         f'"{VENV_PY}" {e2e} "put =1+1 in A1"')

    if IS_WIN:
        lnk = os.path.join(os.environ.get("APPDATA", ""), r"Microsoft\Windows\Start Menu\Programs")
        has_lnk = os.path.isdir(lnk) and any("ekxel" in f.lower() for f in os.listdir(lnk))
        note("Start menu shortcut: " + ("present" if has_lnk else
             f'none (optional, ask the user): "{VENV_PY}" -c "from sheets import winshell; '
             'winshell.write_shortcut(winshell.START_MENU_LNK)"'))
        start = "start Ekxel"
    else:
        cmd_file = os.path.join(ROOT, "Ekxel.command")
        check(os.access(cmd_file, os.X_OK), "Ekxel.command is executable", f'chmod +x "{cmd_file}"')
        start = "double-click Ekxel.command (or run .venv/bin/python launch.pyw)"

    keys = "Cmd+Shift+A" if IS_MAC else "Ctrl+Shift+A"
    ready = all(results)
    print(f"\nREADY: {start}, then press {keys} for the Claude panel." if ready
          else "\nNOT READY: do each FIX above, then run this again.")
    return 0 if ready else 1


if __name__ == "__main__":
    sys.exit(main())
