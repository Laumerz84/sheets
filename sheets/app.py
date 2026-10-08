"""Application entry point (single instance: later launches hand their files over)."""
import json
import os
import sys

from PySide6.QtCore import QByteArray, Qt, QTimer
from PySide6.QtGui import QGuiApplication
from PySide6.QtNetwork import QLocalServer, QLocalSocket
from PySide6.QtWidgets import QApplication

SERVER_NAME = "SheetsSpreadsheetApp-" + os.environ.get("USERNAME", "user")


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


def _start_server(on_paths):
    server = QLocalServer()
    QLocalServer.removeServer(SERVER_NAME)
    server.listen(SERVER_NAME)

    def on_conn():
        conn = server.nextPendingConnection()
        if conn is None:
            return
        buf = QByteArray()

        def read():
            buf.append(conn.readAll())

        def done():
            read()
            try:
                msg = json.loads(bytes(buf).decode("utf-8") or "{}")
            except ValueError:
                msg = {}
            if msg.get("quit"):
                QTimer.singleShot(0, quit_all)
            else:
                on_paths(msg.get("open", []))
            conn.deleteLater()
        conn.readyRead.connect(read)
        conn.disconnected.connect(done)
    server.newConnection.connect(on_conn)
    return server


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
    server = _start_server(open_paths)
    open_paths(paths)
    app.setQuitOnLastWindowClosed(True)
    rc = app.exec()
    server.close()
    return rc


if __name__ == "__main__":
    sys.exit(main())
