"""Claude integration: workbook tools, the MCP bridge, and the panel (with a fake claude)."""
import json
import os
import subprocess
import sys
import threading
import time
import uuid

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QT_QPA_FONTDIR", r"C:\Windows\Fonts")

import pytest
from PySide6.QtWidgets import QApplication, QMessageBox

from sheets import app as sheets_app
from sheets import ops
from sheets.refs import key, parse_addr
from sheets.ui import mainwindow as mw
from sheets.ui.style import apply_palette

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


@pytest.fixture(scope="session")
def app():
    a = QApplication.instance() or QApplication([])
    apply_palette(a)
    return a


@pytest.fixture
def win(app, monkeypatch, tmp_path):
    monkeypatch.setattr(mw, "SETTINGS_DIR", str(tmp_path))
    monkeypatch.setattr(mw, "SETTINGS_FILE", str(tmp_path / "settings.json"))
    monkeypatch.setattr(QMessageBox, "question", staticmethod(lambda *a, **k: QMessageBox.Yes))
    w = mw.MainWindow()
    w.resize(1200, 700)
    w.show()
    app.processEvents()
    yield w
    w.undo.setClean()
    w.close()
    app.processEvents()


def put(w, a, text):
    sh = w.grid.sheet
    r, c = parse_addr(a)
    w._push_states({key(r, c): ops.input_state(sh, r, c, text)}, "test")


def val(w, a):
    return w.grid.sheet.value(*parse_addr(a))


# ---------------------------------------------------------------- tools

def test_tools_read_write_and_single_undo_step(win, app):
    t = win.ai_tools
    put(win, "A1", "10")
    put(win, "A2", "20")
    before = win.undo.index()
    t.begin_turn("make a total")
    res = t.call("write_range", {"start_cell": "A3", "rows": [["=SUM(A1:A2)", "label"], ["=1/0", None]]})
    assert res["written"].endswith("A3:B4")
    assert res["formula_errors"] == {"A4": "#DIV/0!"}
    t.call("format_range", {"range": "A3:B3", "bold": True, "fill_color": "lightyellow",
                            "number_format": '"$"#,##0'})
    t.call("add_sheet", {"name": "Summary"})
    t.call("write_range", {"start_cell": "B2", "rows": [["hi"]], "sheet": "Summary"})
    t.end_turn()
    assert val(win, "A3") == 30
    assert win.grid.sheet.style(2, 0).bold and win.grid.sheet.style(2, 0).fill == "#FFF2CC"
    assert win.wb.get_sheet("Summary").value(1, 1) == "hi"
    assert win.grid.sheet.name == "Sheet1"            # user stays where they were
    assert win.undo.index() == before + 1             # one undo step for the whole request
    win.undo.undo()
    assert val(win, "A3") is None and win.wb.get_sheet("Summary") is None
    r = t.read_range("A1:A3")
    assert r["values"] == [[10], [20], [None]]


def test_tools_read_formulas_info_find_sort(win, app):
    t = win.ai_tools
    for i, (n, q) in enumerate([("Name", "Qty"), ("b", "2"), ("a", "3"), ("c", "1")], start=1):
        put(win, f"A{i}", n)
        put(win, f"B{i}", q)
    put(win, "C2", "=B2*2")
    info = t.workbook_info()
    assert info["sheets"][0]["used_range"] == "A1:C4"
    r = t.read_range("A:C")
    assert r["formulas"] == {"C2": "=B2*2"} and len(r["values"]) == 4
    assert [m["cell"] for m in t.find("a", whole_cell=True)["matches"]] == ["Sheet1!A3"]
    t.begin_turn("sort")
    assert t.call("sort_range", {"range": "A1:C4", "column": "B"})["changed"]
    t.end_turn()
    assert [val(win, f"A{i}") for i in (2, 3, 4)] == ["c", "b", "a"]
    assert win.grid.sheet.formula_text(2, 2) == "=B3*2"  # formula moved with its row


def test_tools_errors_are_reported_not_raised(win, app):
    t = win.ai_tools
    assert "error" in t.call("read_range", {"range": "nonsense!!"})
    assert "error" in t.call("read_range", {"range": "A1", "sheet": "Nope"})
    assert "error" in t.call("format_range", {"range": "A1", "fill_color": "not-a-color"})
    assert "error" in t.call("write_range", {"start_cell": "A1", "rows": []})
    assert "error" in t.call("bogus_tool", {})
    assert "error" in t.call("read_range", {"rng": "A1"})


def test_tools_structure_and_select(win, app):
    t = win.ai_tools
    put(win, "A1", "1")
    put(win, "A2", "2")
    put(win, "B1", "=SUM(A1:A2)")
    t.begin_turn("insert")
    t.call("insert_or_delete", {"action": "insert_rows", "at": "2", "count": 2})
    t.end_turn()
    assert win.grid.sheet.formula_text(0, 1) == "=SUM(A1:A4)"
    t.select_range("A4")
    assert win.grid.sel.active == (3, 0)


# ---------------------------------------------------------------- bridge (real MCP over stdio)

def test_bridge_round_trip(win, app):
    name = "SheetsTest-" + uuid.uuid4().hex[:8]
    server = sheets_app._start_server(lambda paths: None, name=name)
    win.ai_tools.begin_turn("bridge test")
    env = dict(os.environ, SHEETS_SERVER_NAME=name, SHEETS_TARGET=win.ai_token)
    proc = subprocess.Popen([sys.executable, os.path.join(ROOT, "sheets", "ai", "mcp_server.py")],
                            stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE, env=env)
    replies = []

    def client():
        def send(obj):
            proc.stdin.write((json.dumps(obj) + "\n").encode())
            proc.stdin.flush()

        def recv():
            return json.loads(proc.stdout.readline())
        send({"jsonrpc": "2.0", "id": 1, "method": "initialize",
              "params": {"protocolVersion": "2025-06-18", "capabilities": {},
                         "clientInfo": {"name": "test", "version": "1"}}})
        replies.append(recv())
        send({"jsonrpc": "2.0", "method": "notifications/initialized"})
        send({"jsonrpc": "2.0", "id": 2, "method": "tools/list"})
        replies.append(recv())
        send({"jsonrpc": "2.0", "id": 3, "method": "tools/call",
              "params": {"name": "write_range", "arguments": {"start_cell": "D5", "rows": [["=6*7"]]}}})
        replies.append(recv())
    th = threading.Thread(target=client, daemon=True)
    th.start()
    deadline = time.time() + 30
    while th.is_alive() and time.time() < deadline:
        app.processEvents()
        time.sleep(0.01)
    proc.kill()
    server.close()
    assert len(replies) == 3, proc.stderr.read().decode()[-2000:]
    names = {t["name"] for t in replies[1]["result"]["tools"]}
    assert {"workbook_info", "read_range", "write_range", "format_range"} <= names
    assert "D5" in json.dumps(replies[2]["result"])
    assert val(win, "D5") == 42


# ---------------------------------------------------------------- panel with a fake claude

def test_panel_runs_turn_and_resumes(win, app, monkeypatch, tmp_path):
    name = "SheetsTest-" + uuid.uuid4().hex[:8]
    monkeypatch.setattr(sheets_app, "SERVER_NAME", name)
    server = sheets_app._start_server(lambda paths: None, name=name)
    log = tmp_path / "fake.log"
    monkeypatch.setenv("FAKE_CLAUDE_LOG", str(log))
    monkeypatch.setenv("SHEETS_CLAUDE_CMD", json.dumps([sys.executable, os.path.join(ROOT, "tests", "fake_claude.py")]))
    put(win, "A1", "10")
    put(win, "A2", "20")
    before = win.undo.index()
    panel = win.claude_panel
    panel.show()

    def run(text):
        panel.input.setPlainText(text)
        panel.send()
        deadline = time.time() + 30
        while panel.busy() and time.time() < deadline:
            app.processEvents()
            time.sleep(0.01)
        app.processEvents()

    run("add a total")
    assert val(win, "A3") == 30
    assert win.undo.index() == before + 1
    assert "Done: A3 = 30." in panel.md and "Wrote A3" in panel.md
    assert panel.session_id == "sess-123"
    run("thanks")
    calls = [json.loads(x) for x in log.read_text(encoding="utf-8").splitlines()]
    assert "--resume" not in calls[0]["argv"]
    assert calls[1]["argv"][calls[1]["argv"].index("--resume") + 1] == "sess-123"
    assert calls[0]["prompt"].startswith("[Sheets: ") and calls[0]["prompt"].endswith("add a total")
    assert calls[0]["argv"][calls[0]["argv"].index("--tools") + 1] == ""
    server.close()


# ---------------------------------------------------------------- review fixes

def test_user_edits_blocked_while_claude_works(win, app):
    t = win.ai_tools
    put(win, "A1", "1")
    before = win.undo.index()
    t.begin_turn("work")
    win.grid.read_only = True
    t.call("write_range", {"start_cell": "B1", "rows": [["claude"]]})
    put(win, "C1", "user")                 # user's own change is refused, not merged
    win.do_undo()                          # undo is refused too
    assert val(win, "C1") is None and val(win, "B1") == "claude"
    win.grid.begin_edit(text="x")
    assert not win.grid.editing
    t.end_turn()
    win.grid.read_only = False
    assert win.undo.index() == before + 1
    put(win, "C1", "user")
    assert val(win, "C1") == "user"


def test_noop_and_stop(win, app):
    t = win.ai_tools
    put(win, "A1", "same")
    before = win.undo.index()
    t.begin_turn("noop")
    t.call("write_range", {"start_cell": "A1", "rows": [["same"]]})
    t.end_turn()
    assert win.undo.index() == before          # nothing changed -> no undo step
    t.begin_turn("stop")
    t.stopped = True
    assert "error" in t.call("write_range", {"start_cell": "A2", "rows": [["late"]]})
    t.end_turn()
    assert val(win, "A2") is None


def test_numeric_limits(win, app):
    t = win.ai_tools
    t.begin_turn("limits")
    for args in ({"columns": "A", "width": 1e9}, {"columns": "A", "width": -5}):
        assert "error" in t.call("set_column_width", args)
    for size in (float("nan"), 0, 1e9):
        assert "error" in t.call("format_range", {"range": "A1", "font_size": size})
    t.end_turn()
    win.grid.grab()
