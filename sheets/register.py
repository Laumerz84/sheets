"""Per-user Windows file association ("Open with" entries) for Sheets.

Only writes under HKEY_CURRENT_USER\\Software\\Classes; no admin rights needed.
Windows requires the user to confirm the *default* app themselves."""
import os
import sys

PROGID = "Sheets.Spreadsheet"
EXTS = [".csv", ".tsv", ".xlsx", ".xlsm", ".xls"]
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PYW = os.path.join(ROOT, ".venv", "Scripts", "pythonw.exe")
LAUNCHER = os.path.join(ROOT, "launch.pyw")
ICON = os.path.join(ROOT, "sheets.ico")


def command():
    return f'"{PYW}" "{LAUNCHER}" "%1"'


def describe():
    return ("This adds Sheets to the 'Open with' list for " + ", ".join(EXTS) + " files "
            "(for your Windows user only).")


def make_icon():
    from PySide6.QtGui import QImage
    from .ui.mainwindow import app_icon
    img = app_icon().pixmap(64, 64).toImage()
    img.save(ICON, "ICO")
    return ICON


def register():
    import winreg
    make_icon()
    base = r"Software\Classes"
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, rf"{base}\{PROGID}") as k:
        winreg.SetValueEx(k, "", 0, winreg.REG_SZ, "Sheets Spreadsheet")
        winreg.SetValueEx(k, "FriendlyTypeName", 0, winreg.REG_SZ, "Sheets Spreadsheet")
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, rf"{base}\{PROGID}\DefaultIcon") as k:
        winreg.SetValueEx(k, "", 0, winreg.REG_SZ, ICON)
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, rf"{base}\{PROGID}\shell\open") as k:
        winreg.SetValueEx(k, "FriendlyAppName", 0, winreg.REG_SZ, "Sheets")
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, rf"{base}\{PROGID}\shell\open\command") as k:
        winreg.SetValueEx(k, "", 0, winreg.REG_SZ, command())
    for ext in EXTS:
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, rf"{base}\{ext}\OpenWithProgids") as k:
            winreg.SetValueEx(k, PROGID, 0, winreg.REG_NONE, b"")
    try:
        import ctypes
        ctypes.windll.shell32.SHChangeNotify(0x08000000, 0, None, None)
    except Exception:
        pass


def unregister():
    import winreg
    base = r"Software\Classes"
    for ext in EXTS:
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, rf"{base}\{ext}\OpenWithProgids", 0,
                                winreg.KEY_SET_VALUE) as k:
                winreg.DeleteValue(k, PROGID)
        except OSError:
            pass
    for sub in (r"shell\open\command", r"shell\open", "shell", "DefaultIcon", ""):
        try:
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER, rf"{base}\{PROGID}" + (("\\" + sub) if sub else ""))
        except OSError:
            pass


if __name__ == "__main__":
    if "--remove" in sys.argv:
        unregister()
        print("Removed Sheets file associations.")
    else:
        register()
        print("Registered. Right-click a file > Open with > Sheets.")
