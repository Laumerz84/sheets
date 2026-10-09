"""Windows taskbar identity: makes Sheets' windows show and pin as "Sheets"
instead of the Python interpreter that runs them. Needs pywin32; silently
does nothing without it (or off Windows)."""
import os
import sys

APP_ID = "Sheets.Spreadsheet"
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PYW = os.path.join(ROOT, ".venv", "Scripts", "pythonw.exe")
LAUNCHER = os.path.join(ROOT, "launch.pyw")
ICON = os.path.join(ROOT, "sheets.ico")
START_MENU_LNK = os.path.join(os.environ.get("APPDATA", ""), "Microsoft", "Windows", "Start Menu",
                              "Programs", "Sheets.lnk")


def relaunch_command():
    return f'"{PYW}" "{LAUNCHER}"'


def set_window_identity(hwnd):
    """Tag a top-level window so 'Pin to taskbar' pins Sheets (command, name, icon)."""
    if sys.platform != "win32":
        return False
    try:
        from win32com.propsys import propsys, pscon
        store = propsys.SHGetPropertyStoreForWindow(int(hwnd), propsys.IID_IPropertyStore)
        for key, value in ((pscon.PKEY_AppUserModel_ID, APP_ID),
                           (pscon.PKEY_AppUserModel_RelaunchCommand, relaunch_command()),
                           (pscon.PKEY_AppUserModel_RelaunchDisplayNameResource, "Sheets"),
                           (pscon.PKEY_AppUserModel_RelaunchIconResource, ICON + ",0")):
            store.SetValue(key, propsys.PROPVARIANTType(value))
        store.Commit()
        return True
    except Exception:
        return False


def write_shortcut(path, load_existing=False):
    """Create (or rewrite in place) a .lnk that starts Sheets with its icon and app id."""
    import pythoncom
    from win32com.propsys import propsys, pscon
    from win32com.shell import shell
    link = pythoncom.CoCreateInstance(shell.CLSID_ShellLink, None, pythoncom.CLSCTX_INPROC_SERVER,
                                      shell.IID_IShellLink)
    pf = link.QueryInterface(pythoncom.IID_IPersistFile)
    if load_existing and os.path.exists(path):
        pf.Load(path, 2)  # STGM_READWRITE, so the property store can be changed
    link.SetPath(PYW)
    link.SetArguments(f'"{LAUNCHER}"')
    link.SetWorkingDirectory(ROOT)
    link.SetIconLocation(ICON, 0)
    link.SetDescription("Sheets spreadsheet")
    store = link.QueryInterface(propsys.IID_IPropertyStore)
    store.SetValue(pscon.PKEY_AppUserModel_ID, propsys.PROPVARIANTType(APP_ID))
    store.Commit()
    pf.Save(path, 0)
    return path


def refresh_shell_icons():
    try:
        import ctypes
        ctypes.windll.shell32.SHChangeNotify(0x08000000, 0, None, None)  # SHCNE_ASSOCCHANGED
    except Exception:
        pass
