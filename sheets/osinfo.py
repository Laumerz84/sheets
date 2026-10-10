"""Per-OS differences in one place (Windows is the main target; macOS and Linux are supported)."""
import os
import sys

IS_WIN = sys.platform == "win32"
IS_MAC = sys.platform == "darwin"

HOME = os.path.expanduser("~")


def settings_dir():
    """Where Ekxel keeps settings, the cursor library and the wishlist counts."""
    if IS_WIN:
        return os.path.join(os.environ.get("APPDATA", HOME), "Sheets")
    if IS_MAC:
        return os.path.join(HOME, "Library", "Application Support", "Sheets")
    return os.path.join(os.environ.get("XDG_CONFIG_HOME", os.path.join(HOME, ".config")), "Sheets")


def default_spreadsheet_folder():
    """Start folder for Open / Save As until the user picks one (File > Set Default Folder)."""
    if IS_WIN:
        return r"G:\Spreadsheets" if os.path.isdir("G:\\") else os.path.join(HOME, "Documents", "Spreadsheets")
    return os.path.join(HOME, "Documents", "Spreadsheets")


def user_name():
    return os.environ.get("USERNAME") or os.environ.get("USER") or "user"


# Interface fonts: the platform's own UI font. Point sizes differ: macOS draws 1 pt = 1 px,
# Windows 1 pt = 1.33 px, so the same look needs a bigger number on a Mac.
if IS_WIN:
    UI_FONT, UI_PT, MONO_FONT = "Segoe UI", 9, "Consolas"
elif IS_MAC:
    UI_FONT, UI_PT, MONO_FONT = "Helvetica Neue", 13, "Menlo"
else:
    UI_FONT, UI_PT, MONO_FONT = "DejaVu Sans", 10, "DejaVu Sans Mono"


def ui_pt(windows_pt):
    """Scale a size chosen for Windows to this platform."""
    return round(windows_pt * UI_PT / 9) if not IS_WIN else windows_pt


def claude_candidates():
    """Places Claude Code's installers put the `claude` command (used when it isn't on PATH,
    e.g. a Mac app started from Finder gets a minimal PATH)."""
    if IS_WIN:
        return [os.path.join(HOME, ".local", "bin", "claude.exe")]
    return [os.path.join(HOME, ".local", "bin", "claude"), "/opt/homebrew/bin/claude", "/usr/local/bin/claude",
            os.path.join(HOME, ".claude", "local", "claude"), os.path.join(HOME, ".npm-global", "bin", "claude")]


def venv_python(root, windowless=False):
    """The repo's .venv interpreter (pythonw on Windows when no console is wanted)."""
    if IS_WIN:
        return os.path.join(root, ".venv", "Scripts", "pythonw.exe" if windowless else "python.exe")
    return os.path.join(root, ".venv", "bin", "python")
