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
    return ("This adds Exkel to the 'Open with' list for " + ", ".join(EXTS) + " files "
            "(for your Windows user only).")


def make_icon(path=ICON):
    """Write the icon as a multi-size .ico (PNG entries), so Windows picks a sharp one at every size."""
    import struct
    from PySide6.QtCore import QBuffer, QIODevice
    from .ui.mainwindow import ICON_SIZES, paint_app_icon
    pngs = []
    for px in ICON_SIZES:
        buf = QBuffer()
        buf.open(QIODevice.WriteOnly)
        paint_app_icon(px).save(buf, "PNG")
        pngs.append((px, bytes(buf.data())))
    offset = 6 + 16 * len(pngs)
    out = [struct.pack("<HHH", 0, 1, len(pngs))]
    for px, data in pngs:
        dim = 0 if px >= 256 else px
        out.append(struct.pack("<BBBBHHII", dim, dim, 0, 0, 1, 32, len(data), offset))
        offset += len(data)
    out += [data for _, data in pngs]
    with open(path, "wb") as fh:
        fh.write(b"".join(out))
    return path


def register():
    import winreg
    make_icon()
    base = r"Software\Classes"
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, rf"{base}\{PROGID}") as k:
        winreg.SetValueEx(k, "", 0, winreg.REG_SZ, "Exkel Spreadsheet")
        winreg.SetValueEx(k, "FriendlyTypeName", 0, winreg.REG_SZ, "Exkel Spreadsheet")
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, rf"{base}\{PROGID}\DefaultIcon") as k:
        winreg.SetValueEx(k, "", 0, winreg.REG_SZ, ICON)
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, rf"{base}\{PROGID}\shell\open") as k:
        winreg.SetValueEx(k, "FriendlyAppName", 0, winreg.REG_SZ, "Exkel")
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
        print("Removed Exkel file associations.")
    else:
        register()
        print("Registered. Right-click a file > Open with > Exkel.")
