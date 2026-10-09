import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QT_QPA_FONTDIR", r"C:\Windows\Fonts")

import pytest
from PySide6.QtCore import Qt
from PySide6.QtGui import QGuiApplication
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication, QMessageBox

from sheets import ops
from sheets.refs import key, parse_addr
from sheets.ui import keytips as kt
from sheets.ui import mainwindow as mw
from sheets.ui.style import apply_palette


@pytest.fixture(scope="session")
def app():
    a = QApplication.instance() or QApplication([])
    apply_palette(a)
    return a


@pytest.fixture
def win(app, monkeypatch, tmp_path):
    monkeypatch.setattr(mw, "SETTINGS_DIR", str(tmp_path))
    monkeypatch.setattr(mw, "SETTINGS_FILE", str(tmp_path / "settings.json"))
    monkeypatch.setenv("APPDATA", str(tmp_path))
    monkeypatch.setattr(QMessageBox, "question", staticmethod(lambda *a, **k: QMessageBox.Yes))
    w = mw.MainWindow()
    w.resize(1200, 700)
    w.show()
    w.activateWindow()
    w.grid.setFocus()
    app.processEvents()
    yield w
    w.undo.setClean()
    w.close()
    app.processEvents()


def put(w, a, text):
    sh = w.grid.sheet
    r, c = parse_addr(a)
    w._push_states({key(r, c): ops.input_state(sh, r, c, text)}, "test")


def tap_alt(w):
    QTest.keyPress(w.grid, Qt.Key_Alt, Qt.AltModifier)
    QTest.keyRelease(w.grid, Qt.Key_Alt, Qt.NoModifier)


def keys(w, app, text):
    for ch in text:
        QTest.keyClick(w.grid, getattr(Qt, f"Key_{ch}"))
    app.processEvents()
    QTest.qWait(20)


def test_alt_h_o_i_autofits_columns(win, app):
    put(win, "A1", "a fairly long piece of text")
    tap_alt(win)
    assert win.keytips.active and win.keytips.hints.isVisible()
    keys(win, app, "HOI")
    assert not win.keytips.active
    assert win.grid.sheet.col_widths.get(0, 64) > 100


def test_hold_alt_style_and_2003_sequence(win, app):
    put(win, "A1", "5")
    put(win, "B1", "=A1*2")
    win.grid.set_selection([(0, 1, 0, 1)])
    win.copy()
    win.grid.set_selection([(2, 1, 2, 1)])
    QTest.keyPress(win.grid, Qt.Key_Alt, Qt.AltModifier)
    QTest.keyClick(win.grid, Qt.Key_E, Qt.AltModifier)
    QTest.keyRelease(win.grid, Qt.Key_Alt, Qt.NoModifier)
    keys(win, app, "SV")
    sh = win.grid.sheet
    assert sh.value(2, 1) == 10 and sh.formula_text(2, 1) is None   # pasted as a value
    QGuiApplication.clipboard().setText("")


def test_unbuilt_tip_is_recorded(win, app):
    tap_alt(win)
    keys(win, app, "NV")
    tried = kt.load_tried()
    assert tried["NV"]["count"] == 1 and "wishlist" in win.statusBar().currentMessage()


def test_escape_backs_out_and_bad_sequence_cancels(win, app):
    tap_alt(win)
    keys(win, app, "HO")
    QTest.keyClick(win.grid, Qt.Key_Escape)
    assert win.keytips.active and win.keytips.buf == "H"
    keys(win, app, "Q")
    assert not win.keytips.active


def test_typing_in_a_cell_is_not_captured(win, app):
    win.grid.begin_edit(text="", mode="enter")
    tap_alt(win)
    assert not win.keytips.active


def test_table_is_consistent(win):
    seqs = [s for s, *_ in kt.KEYTIPS]
    assert len(seqs) == len(set(seqs)), "duplicate sequence"
    for s in seqs:
        for other in seqs:
            assert other == s or not other.startswith(s), f"{s} hides {other}"
    known = {"showMenu", "trigger", "set_zoom", "_focus", "_font_step", "_indent", "font_box", "size_box",
             "numfmt_box", "grid", "a_painter", "font_color_btn", "fill_color_btn"}
    for seq, label, action, excel in kt.KEYTIPS:
        if action is None:
            assert len(excel) > 20, f"{seq} needs an Excel description"
            continue
        if action.__name__ != "<lambda>":
            continue  # a helper function defined in keytips.py
        for name in action.__code__.co_names:
            assert hasattr(win, name) or name in known or hasattr(kt, name), f"{seq}: unknown {name}"
