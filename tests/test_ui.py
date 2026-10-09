import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QT_QPA_FONTDIR", r"C:\Windows\Fonts")

import pytest
from PySide6.QtCore import QPoint, Qt
from PySide6.QtGui import QGuiApplication
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication, QMessageBox

from sheets import ops
from sheets.refs import key, parse_addr
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
    monkeypatch.setattr(QMessageBox, "question", staticmethod(lambda *a, **k: QMessageBox.Yes))
    w = mw.MainWindow()
    w.resize(1200, 700)
    w.show()
    app.processEvents()
    yield w
    w.undo.setClean()
    w.close()
    app.processEvents()


def cell_center(w, a):
    r, c = parse_addr(a)
    return w.grid.cell_rect(r, c).center()


def click(w, a, mods=Qt.NoModifier):
    QTest.mouseClick(w.grid, Qt.LeftButton, mods, cell_center(w, a))


def type_text(w, text, enter=True):
    target = w.grid.editor if w.grid.editing else w.grid
    for ch in text:
        target = w.grid.editor if w.grid.editing else w.grid
        QTest.keyClicks(target, ch)
    if enter:
        QTest.keyClick(w.grid.editor, Qt.Key_Return)


def val(w, a):
    return w.grid.sheet.value(*parse_addr(a))


def put(w, a, text):
    sh = w.grid.sheet
    r, c = parse_addr(a)
    w._push_states({key(r, c): ops.input_state(sh, r, c, text)}, "test")


def test_typing_and_undo(win, app):
    click(win, "B2")
    type_text(win, "hello")
    assert val(win, "B2") == "hello"
    assert win.grid.sel.active == (2, 1)
    click(win, "C2")
    type_text(win, "=2*21")
    assert val(win, "C2") == 42
    win.undo.undo()
    assert val(win, "C2") is None
    win.undo.redo()
    assert val(win, "C2") == 42
    assert "•" in win.windowTitle()


def test_point_mode_click_inserts_reference(win, app):
    put(win, "A1", "5")
    click(win, "B1")
    type_text(win, "=", enter=False)
    QTest.mouseClick(win.grid, Qt.LeftButton, Qt.NoModifier, cell_center(win, "A1"))
    assert win.grid.editing
    assert win.grid.editor.text() == "=A1"
    QTest.keyClicks(win.grid.editor, "*2")
    QTest.keyClick(win.grid.editor, Qt.Key_Return)
    assert val(win, "B1") == 10


def test_arrow_point_mode(win, app):
    put(win, "A1", "7")
    click(win, "B1")
    type_text(win, "=", enter=False)
    QTest.keyClick(win.grid.editor, Qt.Key_Left)
    assert win.grid.editor.text() == "=A1"
    QTest.keyClick(win.grid.editor, Qt.Key_Return)
    assert val(win, "B1") == 7


def test_enter_mode_arrow_commits(win, app):
    click(win, "A1")
    type_text(win, "abc", enter=False)
    QTest.keyClick(win.grid.editor, Qt.Key_Right)
    assert val(win, "A1") == "abc"
    assert win.grid.sel.active == (0, 1)


def test_copy_paste_shifts_formula(win, app):
    put(win, "A1", "1")
    put(win, "A2", "2")
    put(win, "B1", "=A1*10")
    click(win, "B1")
    win.copy()
    click(win, "B2")
    win.paste()
    assert win.grid.sheet.formula_text(1, 1) == "=A2*10"
    assert val(win, "B2") == 20


def test_external_text_paste(win, app):
    QGuiApplication.clipboard().setText("a\tb\r\n1\t2\r\n")
    click(win, "C3")
    win.paste()
    assert val(win, "C3") == "a" and val(win, "D4") == 2


def test_fill_handle_drag(win, app):
    put(win, "A1", "1")
    put(win, "A2", "2")
    g = win.grid
    g.set_selection([(0, 0, 1, 0)])
    app.processEvents()
    R = g.cell_rect(1, 0)
    handle = QPoint(R.right(), R.bottom())
    assert g._on_fill_handle(handle)
    QTest.mousePress(g, Qt.LeftButton, Qt.NoModifier, handle)
    target = cell_center(win, "A5")
    QTest.mouseMove(g, target)
    from PySide6.QtGui import QMouseEvent
    from PySide6.QtCore import QEvent, QPointF
    ev = QMouseEvent(QEvent.MouseMove, QPointF(target), QPointF(g.mapToGlobal(target)), Qt.NoButton,
                     Qt.LeftButton, Qt.NoModifier)
    QApplication.sendEvent(g, ev)
    QTest.mouseRelease(g, Qt.LeftButton, Qt.NoModifier, target)
    assert [val(win, f"A{i}") for i in range(1, 6)] == [1, 2, 3, 4, 5]


def test_insert_delete_rows_undo(win, app):
    put(win, "A1", "1")
    put(win, "A2", "2")
    put(win, "B1", "=SUM(A1:A2)")
    win.grid.set_selection([(1, 0, 1, 0)])
    win.insert_rows_cols("row")
    assert val(win, "A3") == 2
    assert win.grid.sheet.formula_text(0, 1) == "=SUM(A1:A3)"
    win.undo.undo()
    assert val(win, "A2") == 2
    assert win.grid.sheet.formula_text(0, 1) == "=SUM(A1:A2)"


def test_filter_and_sort(win, app):
    for i, (n, q) in enumerate([("Name", "Qty"), ("b", "2"), ("a", "3"), ("c", "1")]):
        put(win, f"A{i + 1}", n)
        put(win, f"B{i + 1}", q)
    click(win, "A2")
    win.toggle_filter()
    sh = win.grid.sheet
    assert sh.autofilter == (0, 0, 3, 1)
    win.apply_filter(0, {"values": {"a", "c"}, "blanks": False})
    assert sh.filter_hidden == {1}
    assert win.grid.rows.size(1) == 0
    assert "2 of 3" in win.filter_lbl.text()
    win.grid.update()
    app.processEvents()
    win.clear_filters()
    assert sh.filter_hidden == set()
    click(win, "B2")
    win.quick_sort(True)
    assert [val(win, f"A{i}") for i in (2, 3, 4)] == ["c", "b", "a"]
    win.undo.undo()
    assert [val(win, f"A{i}") for i in (2, 3, 4)] == ["b", "a", "c"]


def test_freeze_and_scroll_paint(win, app):
    for r in range(1, 200):
        put(win, f"A{r}", str(r))
    click(win, "B2")
    win.freeze_panes()
    g = win.grid
    assert g.sheet.freeze == (1, 1)
    g.scroll_rows(50)
    assert g.top >= 51
    g.grab()
    assert g.row_at(g.hh + 2) == 0  # frozen row stays at the top
    win.undo.undo()
    assert g.sheet.freeze == (0, 0)


def test_formatting_and_merge_render(win, app):
    put(win, "A1", "A long piece of text that overflows into the neighbours")
    put(win, "A3", "1234.5")
    win.grid.set_selection([(2, 0, 2, 0)])
    win.apply_numfmt("$#,##0.00")
    assert win.grid.sheet.style(2, 0).numfmt == "$#,##0.00"
    win.toggle_font("bold")
    assert win.grid.sheet.style(2, 0).bold
    win.grid.set_selection([(4, 0, 5, 2)])
    win.merge_center()
    assert (4, 0, 5, 2) in win.grid.sheet.merges
    win.grid.set_selection([(7, 0, 9, 2)])
    win.apply_borders("all")
    win.grid.grab()
    win.change_decimals(1)


def test_save_and_reopen(win, app, tmp_path):
    put(win, "A1", "x")
    put(win, "B1", "=1+1")
    p = str(tmp_path / "t.xlsx")
    assert win.do_save(p)
    assert win.undo.isClean()
    from sheets.fileio import open_file
    wb = open_file(p)
    assert wb.sheets[0].value(0, 1) == 2
    p2 = str(tmp_path / "t.csv")
    assert win.do_save(p2)
    assert open(p2, encoding="utf-8-sig").read().startswith("x,2")


def test_find_next(win, app):
    put(win, "A5", "needle")
    put(win, "C9", "another needle")
    win.show_find(False)
    p = win.find_dlg.params()
    p["needle"] = "needle"
    click(win, "A1")
    win.find_next(p)
    assert win.grid.sel.active == (4, 0)
    win.find_next(p)
    assert win.grid.sel.active == (8, 2)
    win.find_dlg.close()


def test_ctrl_arrow_and_ctrl_a(win, app):
    for r in range(1, 11):
        put(win, f"A{r}", str(r))
    click(win, "A1")
    QTest.keyClick(win.grid, Qt.Key_Down, Qt.ControlModifier)
    assert win.grid.sel.active == (9, 0)
    QTest.keyClick(win.grid, Qt.Key_A, Qt.ControlModifier)
    assert win.grid.sel.rects == [(0, 0, 9, 0)]


def test_sheets_add_rename_delete(win, app):
    win.add_sheet()
    assert len(win.wb.sheets) == 2
    sh = win.grid.sheet
    win.undo.push(mw.SnapshotCommand(win, sh, "Rename", lambda: win.wb.rename_sheet(sh, "Data"), sheets=[]))
    assert win.tabs.tabText(1) == "Data"
    win.delete_sheet()
    assert len(win.wb.sheets) == 1
    win.undo.undo()
    assert len(win.wb.sheets) == 2


def test_column_resize_drag(win, app):
    g = win.grid
    x = g.col_x(0) + g.cols.size(0) - 1
    y = g.hh // 2
    QTest.mousePress(g, Qt.LeftButton, Qt.NoModifier, QPoint(x, y))
    from PySide6.QtGui import QMouseEvent
    from PySide6.QtCore import QEvent, QPointF
    ev = QMouseEvent(QEvent.MouseMove, QPointF(x + 50, y), QPointF(g.mapToGlobal(QPoint(x + 50, y))),
                     Qt.NoButton, Qt.LeftButton, Qt.NoModifier)
    QApplication.sendEvent(g, ev)
    QTest.mouseRelease(g, Qt.LeftButton, Qt.NoModifier, QPoint(x + 50, y))
    assert g.sheet.col_widths[0] > 100
    win.undo.undo()
    assert 0 not in g.sheet.col_widths


def test_autowiden_date(win, app):
    click(win, "F2")
    type_text(win, "2024-01-15")
    assert win.grid.sheet.col_widths.get(5, 64) > 64


def test_format_painter(win, app):
    put(win, "A1", "x")
    win.grid.set_selection([(0, 0, 0, 0)])
    win.toggle_font("bold")
    win.set_style_field("fill", "#FFFF00")
    win.a_painter.setChecked(True)
    win.start_format_painter()
    QTest.mouseClick(win.grid, Qt.LeftButton, Qt.NoModifier, cell_center(win, "C3"))
    st = win.grid.sheet.style(2, 2)
    assert st.bold and st.fill == "#FFFF00"
    assert not win.a_painter.isChecked()


def test_remove_duplicates(win, app):
    from sheets.ui.extras import dedupe_states
    rows = [("k", "v"), ("a", "1"), ("A", "1"), ("b", "2"), ("a", "1"), ("c", "3")]
    for i, (a, b) in enumerate(rows, start=1):
        put(win, f"A{i}", a)
        put(win, f"B{i}", b)
    sh = win.grid.sheet
    states, removed, kept = dedupe_states(sh, (0, 0, 5, 1), [0, 1], True)
    assert (removed, kept) == (2, 3)
    win._push_states(states, "dedupe")
    assert [val(win, f"A{i}") for i in range(1, 7)] == ["k", "a", "b", "c", None, None]


def test_double_click_edit_then_click_away_commits(win, app):
    put(win, "A1", "abc")
    QTest.mouseDClick(win.grid, Qt.LeftButton, Qt.NoModifier, cell_center(win, "A1"))
    assert win.grid.editing and win.grid.editor.text() == "abc"
    QTest.keyClicks(win.grid.editor, "d")
    QTest.mouseClick(win.grid, Qt.LeftButton, Qt.NoModifier, cell_center(win, "C5"))
    assert not win.grid.editing
    assert val(win, "A1") == "abcd"
    assert win.grid.sel.active == (4, 2)


def test_formula_bar_editing(win, app):
    click(win, "B2")
    win.fbar.setFocus()
    app.processEvents()
    assert win.grid.editing and win.grid.edit_widget is win.fbar
    QTest.keyClicks(win.fbar, "=3*3")
    assert win.grid.editor.text() == "=3*3"
    QTest.keyClick(win.fbar, Qt.Key_Return)
    assert val(win, "B2") == 9
    assert win.fbar.text() == ""  # now showing B3


def test_tab_enter_and_ctrl_enter(win, app):
    click(win, "A1")
    type_text(win, "1", enter=False)
    QTest.keyClick(win.grid.editor, Qt.Key_Tab)
    assert win.grid.sel.active == (0, 1)
    win.grid.set_selection([(0, 3, 2, 3)])
    win.grid.begin_edit(text="=ROW()", mode="enter")
    QTest.keyClick(win.grid.editor, Qt.Key_Return, Qt.ControlModifier)
    assert [val(win, f"D{i}") for i in (1, 2, 3)] == [1, 2, 3]


def test_header_and_multi_selection(win, app):
    g = win.grid
    x = g.col_x(2) + 5
    QTest.mouseClick(g, Qt.LeftButton, Qt.NoModifier, QPoint(x, g.hh // 2))
    assert g.sel.rects[-1][1] == 2 and g.sel.rects[-1][2] >= 1_000_000
    click(win, "A1")
    click(win, "C3", Qt.ShiftModifier)
    assert g.sel.rects == [(0, 0, 2, 2)]
    click(win, "E5", Qt.ControlModifier)
    assert len(g.sel.rects) == 2
    put(win, "A1", "5")
    put(win, "E5", "7")
    g.set_selection([(0, 0, 0, 0), (4, 4, 4, 4)])
    win._update_stats()
    assert "Sum: 12" in win.stats_lbl.text()


def test_delete_and_backspace(win, app):
    put(win, "A1", "x")
    click(win, "A1")
    QTest.keyClick(win.grid, Qt.Key_Delete)
    assert val(win, "A1") is None
    put(win, "A1", "y")
    QTest.keyClick(win.grid, Qt.Key_Backspace)
    assert win.grid.editing and win.grid.editor.text() == ""
    QTest.keyClick(win.grid.editor, Qt.Key_Escape)
    assert val(win, "A1") == "y"


def test_sheet_switch_keeps_selection(win, app):
    click(win, "C3")
    win.add_sheet()
    assert win.grid.sheet is win.wb.sheets[1]
    click(win, "B2")
    type_text(win, "=Sheet1!A1+1")
    assert val(win, "B2") == 1
    win.cycle_sheet(-1)
    assert win.grid.sel.active == (2, 2)
    put(win, "A1", "41")
    win.cycle_sheet(1)
    assert val(win, "B2") == 42


def test_undo_restores_view(win, app):
    click(win, "D7")
    type_text(win, "hi")
    click(win, "A1")
    win.undo.undo()
    assert win.grid.sel.active == (6, 3)


def test_freeze_while_scrolled(win, app, tmp_path):
    for r in range(1, 300):
        put(win, f"A{r}", str(r))
    g = win.grid
    g.set_top(50)
    g.set_selection([(60, 0, 60, 0)])
    win.freeze_panes()
    sh = g.sheet
    assert sh.freeze == (60, 0) and sh.freeze_origin == (50, 0)
    assert g.frozen_h() < g.height() // 2
    assert g.row_at(g.hh + 2) == 50
    g.scroll_rows(10)
    assert g.top >= 70
    p = str(tmp_path / "f.xlsx")
    assert win.do_save(p)
    from sheets.fileio import open_file
    s2 = open_file(p).sheets[0]
    assert s2.freeze == (60, 0) and s2.freeze_origin == (50, 0)


def test_cut_with_hidden_rows_keeps_hidden_data(win, app):
    for i in range(1, 6):
        put(win, f"A{i}", str(i))
    sh = win.grid.sheet
    win._push_meta({"hidden_rows": (set(), {2})}, "hide")
    win.grid.set_selection([(0, 0, 4, 0)])
    win.cut()
    win.grid.set_selection([(0, 2, 0, 2)])
    win.paste()
    assert val(win, "A3") == 3
    assert [val(win, f"C{i}") for i in range(1, 5)] == [1, 2, 4, 5]


def test_stale_cut_is_cancelled(win, app):
    put(win, "A1", "old")
    click(win, "A1")
    win.cut()
    put(win, "A1", "new")
    assert win.clip is None
    QGuiApplication.clipboard().setText("")


def test_delete_overlapping_row_selections(win, app):
    for i in range(1, 21):
        put(win, f"A{i}", str(i))
    win.grid.set_selection([(4, 0, 9, 0), (7, 0, 11, 0)])
    win.delete_rows_cols("row")
    assert val(win, "A5") == 13


def test_header_then_shift_arrow(win, app):
    g = win.grid
    click(win, "A1")
    y = g.row_y(4) + 3
    QTest.mouseClick(g, Qt.LeftButton, Qt.NoModifier, QPoint(g.rw // 2, y))
    QTest.keyClick(g, Qt.Key_Down, Qt.ShiftModifier)
    r1, c1, r2, c2 = g.sel.rects[-1]
    assert (r1, r2, c1) == (4, 5, 0) and c2 >= 16000


def test_undo_while_editing_cancels_edit(win, app):
    put(win, "A4", "r4")
    click(win, "A4")
    type_text(win, "changed", enter=False)
    win.do_undo()
    assert not win.grid.editing
    assert val(win, "A4") == "r4"


def test_default_folder(win, app, tmp_path, monkeypatch):
    target = tmp_path / "Spreadsheets"
    monkeypatch.setattr(mw, "DEFAULT_FOLDER", str(target))
    assert win.default_dir() == str(target) and target.is_dir()  # created on first use
    chosen = tmp_path / "mine"
    chosen.mkdir()
    from PySide6.QtWidgets import QFileDialog
    monkeypatch.setattr(QFileDialog, "getExistingDirectory", staticmethod(lambda *a, **k: str(chosen)))
    win.set_default_folder()
    put(win, "A1", "x")
    assert win.do_save(str(tmp_path / "elsewhere.csv"))
    assert win.default_dir() == str(chosen)  # saving elsewhere doesn't change it
