"""Application entry point (single instance: later launches hand their files over)."""
import json
import os
import sys

from PySide6.QtCore import QEvent, QObject, Qt, QTimer
from PySide6.QtGui import QGuiApplication
from PySide6.QtNetwork import QLocalServer, QLocalSocket
from PySide6.QtWidgets import QApplication

SERVER_NAME = os.environ.get("SHEETS_SERVER_NAME") or ("SheetsSpreadsheetApp-" + (os.environ.get("USERNAME") or os.environ.get("USER") or "user"))


def _forward(paths, quit_app=False):
    """If Sheets is already running, send it our files (or a quit request) and return True."""
    sock = QLocalSocket()
    sock.connectToServer(SERVER_NAME)
    if not sock.waitForConnected(300):
        return False
    msg = {"quit": True} if quit_app else {"open": paths}
    sock.write(json.dumps(msg).encode("utf-8"))
    sock.flush()
    sock.waitForBytesWritten(1000)
    sock.disconnectFromServer()
    return True


def handle_rpc(msg):
    """A tool call from the Claude bridge: {"rpc": tool, "args": {...}, "target": window token}."""
    from .ui.mainwindow import WINDOWS
    win = next((w for w in WINDOWS if w.ai_token == msg.get("target")), None)
    if win is None:
        return {"error": "That Ekxel window was closed."}
    return win.ai_tools.call(msg.get("rpc"), msg.get("args") or {})


def _start_server(on_paths, name=None, rpc=handle_rpc):
    """Local socket server.  Two kinds of client:
    - launcher hand-offs: one JSON object, then disconnect ({"open": [...]} / {"quit": true})
    - the Claude bridge: newline-terminated {"rpc": ...} requests, each answered with one JSON line."""
    server = QLocalServer()
    name = name or SERVER_NAME
    QLocalServer.removeServer(name)
    server.listen(name)

    def on_conn():
        conn = server.nextPendingConnection()
        if conn is None:
            return
        state = {"buf": b"", "rpc": False}

        def read():
            state["buf"] += bytes(conn.readAll())
            while b"\n" in state["buf"]:
                line, state["buf"] = state["buf"].split(b"\n", 1)
                try:
                    msg = json.loads(line.decode("utf-8"))
                except ValueError:
                    continue
                if "rpc" not in msg:
                    continue
                state["rpc"] = True
                try:
                    reply = rpc(msg)
                except Exception as e:  # keep the bridge alive whatever happens
                    reply = {"error": f"{type(e).__name__}: {e}"}
                conn.write((json.dumps(reply, default=str) + "\n").encode("utf-8"))
                conn.flush()

        def done():
            read()
            if not state["rpc"]:
                try:
                    msg = json.loads(state["buf"].decode("utf-8") or "{}")
                except ValueError:
                    msg = {}
                if msg.get("quit"):
                    QTimer.singleShot(0, quit_all)
                elif "open" in msg:
                    on_paths(msg.get("open", []))
            conn.deleteLater()
        conn.readyRead.connect(read)
        conn.disconnected.connect(done)
    server.newConnection.connect(on_conn)
    return server


class _FileOpenEvents(QObject):
    def eventFilter(self, obj, ev):
        if ev.type() == QEvent.FileOpen and ev.file():
            QTimer.singleShot(0, lambda p=ev.file(): open_paths([p]))
            return True
        return False


def _install_error_hook():
    """Big-file mode refuses some cell-by-cell work by raising bigdata.TooBig: show its message."""
    prev = sys.excepthook

    def hook(etype, value, tb):
        if etype.__name__ == "TooBig":
            from PySide6.QtWidgets import QMessageBox
            while QApplication.overrideCursor() is not None:
                QApplication.restoreOverrideCursor()
            QMessageBox.information(QApplication.activeWindow(), "Ekxel", str(value))
            return
        prev(etype, value, tb)
    sys.excepthook = hook


def quit_all():
    """Close every window the normal way, so unsaved work still gets the Save? prompt."""
    from .ui.mainwindow import WINDOWS
    for w in list(WINDOWS):
        w.raise_()
        w.activateWindow()
        if not w.close():
            return


def open_paths(paths):
    from .ui.mainwindow import WINDOWS, MainWindow
    if not paths:
        w = MainWindow()
        w.show()
        w.raise_()
        w.activateWindow()
        return
    for p in paths:
        host = next((w for w in WINDOWS if w.is_pristine()), None)
        if host is None:
            host = MainWindow()
        w = host.open_path(p) or host
        if not host.isVisible() and host.wb.path is None and w is not host:
            host.close()
        w.show()
        w.raise_()
        w.activateWindow()


def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    paths = [os.path.abspath(a) for a in argv if not a.startswith("-")]
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("Sheets.Spreadsheet")
        except Exception:
            pass
    QGuiApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    app = QApplication(sys.argv[:1])
    app.setApplicationName("Sheets")
    app.setApplicationDisplayName("Macrosoft Ekxel® 2003 Private Reserve Special Cuvée")
    app.setOrganizationName("Sheets")
    if "--quit" in argv:  # ask a running copy to close (used by App Launcher's Stop)
        _forward([], quit_app=True)
        return 0
    if "--new-instance" not in argv and _forward(paths):
        return 0
    from .ui.mainwindow import app_icon
    from .ui.style import apply_palette
    apply_palette(app)
    app.setWindowIcon(app_icon())
    _install_error_hook()
    global SERVER_NAME
    server = _start_server(open_paths)
    if not server.isListening():
        # e.g. --new-instance while another copy owns the name: use our own, so the Claude
        # panel in this process can still reach its windows
        SERVER_NAME = f"{SERVER_NAME}-{os.getpid()}"
        server = _start_server(open_paths)
    # macOS: Finder (double-click, Open With, dropping on the Dock icon) sends files as events, not argv
    file_events = _FileOpenEvents(app)
    app.installEventFilter(file_events)
    open_paths(paths)
    app.setQuitOnLastWindowClosed(True)
    rc = app.exec()
    server.close()
    return rc


if __name__ == "__main__":
    sys.exit(main())
