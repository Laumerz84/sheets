"""Excel-style KeyTips: press and release Alt, then type letters (Alt H O I = AutoFit Column
Width). Also accepts the classic Excel 2003 sequences (Alt E S V, Alt O C A, ...).

KEYTIPS is the single source of truth: an entry with an action is built; an entry with
action None is a feature Sheets doesn't have yet - pressing it records the attempt in
%APPDATA%\\Sheets\\wishlist-tried.json, and tools/make_wishlist.py lists it in
docs/excel-wishlist.md."""
import datetime as dt
import json
import os

from ..osinfo import UI_FONT, UI_PT, settings_dir
from PySide6.QtCore import QEvent, QObject, QPoint, Qt, QTimer
from PySide6.QtWidgets import (QAbstractSpinBox, QApplication, QComboBox, QFrame,
                               QInputDialog, QLabel, QLineEdit, QPlainTextEdit,
                               QTextEdit, QVBoxLayout, QWidget)


# ---------------------------------------------------------------- small actions
def _font_step(w, d):
    sizes = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72]
    cur = w.sheet.style(*w.grid.sel.active).size or 11
    if d > 0:
        nxt = next((s for s in sizes if s > cur), cur + 8)
    else:
        nxt = next((s for s in reversed(sizes) if s < cur), max(1, cur - 1))
    w.set_font_size(str(nxt))


def _indent(w, d):
    w.restyle(lambda s: s.with_(indent=max(0, min(15, s.indent + d)),
                                halign=s.halign if s.halign in ("left", "right") else "left"), "Indent")


def _focus(widget, popup=False):
    widget.setFocus()
    if isinstance(widget, QComboBox):
        if widget.isEditable():
            widget.lineEdit().selectAll()
        if popup:
            widget.showPopup()


def _tab_color(w):
    from .dialogs import ColorMenu
    m = ColorMenu(w, "No Color")
    m.color_chosen.connect(w.set_tab_color)
    m.exec(w.tabs.mapToGlobal(QPoint(0, 0)) - QPoint(0, m.sizeHint().height()))


def _zoom_dialog(w):
    v, ok = QInputDialog.getInt(w, "Zoom", "Zoom (%):", int(round(w.grid.zoom * 100)), 25, 400, 10)
    if ok:
        w.grid.set_zoom(v / 100)


ACCOUNTING = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'

# (sequence, label, action or None, what it does in Excel - used for the wishlist)
KEYTIPS = [
    # ---------------- Home
    ("H1", "Bold", lambda w: w.toggle_font("bold"), ""),
    ("H2", "Italic", lambda w: w.toggle_font("italic"), ""),
    ("H3", "Underline", lambda w: w.toggle_font("underline"), ""),
    ("H4", "Strikethrough", lambda w: w.toggle_font("strike"), ""),
    ("HFF", "Font name", lambda w: _focus(w.font_box), ""),
    ("HFS", "Font size", lambda w: _focus(w.size_box), ""),
    ("HFG", "Increase font size", lambda w: _font_step(w, 1), ""),
    ("HFK", "Decrease font size", lambda w: _font_step(w, -1), ""),
    ("HFC", "Font color", lambda w: w.font_color_btn.showMenu(), ""),
    ("HH", "Fill color", lambda w: w.fill_color_btn.showMenu(), ""),
    ("HBO", "Bottom border", lambda w: w.apply_borders("bottom"), ""),
    ("HBP", "Top border", lambda w: w.apply_borders("top"), ""),
    ("HBL", "Left border", lambda w: w.apply_borders("left"), ""),
    ("HBR", "Right border", lambda w: w.apply_borders("right"), ""),
    ("HBN", "No border", lambda w: w.apply_borders("none"), ""),
    ("HBA", "All borders", lambda w: w.apply_borders("all"), ""),
    ("HBS", "Outside borders", lambda w: w.apply_borders("outline"), ""),
    ("HBT", "Thick outside border", lambda w: w.apply_borders("thick_outline"), ""),
    ("HAL", "Align left", lambda w: w.set_align("halign", "left"), ""),
    ("HAC", "Center", lambda w: w.set_align("halign", "center"), ""),
    ("HAR", "Align right", lambda w: w.set_align("halign", "right"), ""),
    ("HAT", "Top align", lambda w: w.set_align("valign", "top"), ""),
    ("HAM", "Middle align", lambda w: w.set_align("valign", "center"), ""),
    ("HAB", "Bottom align", lambda w: w.set_align("valign", None), ""),
    ("HAN", "Accounting format", lambda w: w.apply_numfmt(ACCOUNTING), ""),
    ("H5", "Decrease indent", lambda w: _indent(w, -1), ""),
    ("H6", "Increase indent", lambda w: _indent(w, 1), ""),
    ("HW", "Wrap text", lambda w: w.toggle_wrap(), ""),
    ("HMC", "Merge & Center", lambda w: w.merge_center(), ""),
    ("HMA", "Merge across", lambda w: w.merge_across(), ""),
    ("HMM", "Merge cells", lambda w: w.merge_center(center=False), ""),
    ("HMU", "Unmerge cells", lambda w: w.unmerge(), ""),
    ("HN", "Number format", lambda w: _focus(w.numfmt_box, popup=True), ""),
    ("HP", "Percent style", lambda w: w.apply_numfmt("0%"), ""),
    ("HK", "Comma style", lambda w: w.apply_numfmt("#,##0.00"), ""),
    ("H0", "Increase decimal", lambda w: w.change_decimals(1), ""),
    ("H9", "Decrease decimal", lambda w: w.change_decimals(-1), ""),
    ("HL", "Conditional Formatting", None,
     "Opens Conditional Formatting: highlight rules (greater than, text contains, duplicates), "
     "top/bottom rules, data bars, color scales and icon sets, plus Manage Rules. Ekxel shows "
     "conditional formatting from Excel files but can't create or edit rules yet."),
    ("HT", "Format as Table", None,
     "Turns the range into an Excel Table with a style gallery: banded rows, header filter "
     "buttons, a total row, and structured references like Table1[Sales]."),
    ("HJ", "Cell Styles", None,
     "Gallery of named cell styles (Good, Bad, Neutral, Heading 1-4, Title, Total, Currency...) "
     "applied in one click."),
    ("HIR", "Insert sheet rows", lambda w: w.insert_rows_cols("row"), ""),
    ("HIC", "Insert sheet columns", lambda w: w.insert_rows_cols("col"), ""),
    ("HIS", "Insert sheet", lambda w: w.add_sheet(), ""),
    ("HII", "Insert cells", None,
     "Insert Cells dialog: insert blank cells and shift the existing ones right or down, "
     "instead of inserting whole rows or columns."),
    ("HDR", "Delete sheet rows", lambda w: w.delete_rows_cols("row"), ""),
    ("HDC", "Delete sheet columns", lambda w: w.delete_rows_cols("col"), ""),
    ("HDS", "Delete sheet", lambda w: w.delete_sheet(), ""),
    ("HDD", "Delete cells", None,
     "Delete Cells dialog: delete the selected cells and shift the cells around them left or up."),
    ("HOI", "AutoFit column width", lambda w: w.autofit_selected_cols(), ""),
    ("HOA", "AutoFit row height", lambda w: w.autofit_selected_rows(), ""),
    ("HOW", "Column width...", lambda w: w.column_width_dialog(), ""),
    ("HOH", "Row height...", lambda w: w.row_height_dialog(), ""),
    ("HOUR", "Hide rows", lambda w: w.hide_rows_cols("row", True), ""),
    ("HOUC", "Hide columns", lambda w: w.hide_rows_cols("col", True), ""),
    ("HOUO", "Unhide rows", lambda w: w.hide_rows_cols("row", False), ""),
    ("HOUL", "Unhide columns", lambda w: w.hide_rows_cols("col", False), ""),
    ("HOR", "Rename sheet", lambda w: w.rename_sheet(), ""),
    ("HOM", "Move or copy sheet (copies it)", lambda w: w.duplicate_sheet(), ""),
    ("HOT", "Tab color", _tab_color, ""),
    ("HOE", "Format Cells...", lambda w: w.format_cells(), ""),
    ("HOP", "Protect sheet", None,
     "Protect Sheet: lock the sheet with an optional password so locked cells can't be edited; "
     "choose what users may still do (select, format, sort, filter...)."),
    ("HOL", "Lock cell", None,
     "Toggles the Locked property of the selected cells (takes effect when the sheet is protected)."),
    ("HUS", "AutoSum", lambda w: w.autosum("SUM"), ""),
    ("HUA", "Average", lambda w: w.autosum("AVERAGE"), ""),
    ("HUC", "Count numbers", lambda w: w.autosum("COUNT"), ""),
    ("HUM", "Max", lambda w: w.autosum("MAX"), ""),
    ("HUI", "Min", lambda w: w.autosum("MIN"), ""),
    ("HUF", "More functions...", lambda w: w.insert_function_dialog(), ""),
    ("HFID", "Fill down", lambda w: w.fill_dir("down"), ""),
    ("HFIR", "Fill right", lambda w: w.fill_dir("right"), ""),
    ("HFIS", "Fill series...", None,
     "Series dialog: fill a linear, growth or date series with a chosen step and stop value "
     "(Ekxel has drag-to-fill series, but not this dialog)."),
    ("HEA", "Clear all", lambda w: w.clear_all(), ""),
    ("HEF", "Clear formats", lambda w: w.clear_formats(), ""),
    ("HEC", "Clear contents", lambda w: w.clear_contents(), ""),
    ("HSS", "Sort A to Z", lambda w: w.quick_sort(True), ""),
    ("HSO", "Sort Z to A", lambda w: w.quick_sort(False), ""),
    ("HSU", "Custom sort...", lambda w: w.sort_dialog(), ""),
    ("HSF", "Filter", lambda w: w.toggle_filter(), ""),
    ("HSC", "Clear filter", lambda w: w.clear_filters(), ""),
    ("HSY", "Reapply filter", lambda w: w.reapply_filters(), ""),
    ("HFDF", "Find...", lambda w: w.show_find(False), ""),
    ("HFDR", "Replace...", lambda w: w.show_find(True), ""),
    ("HFDG", "Go To...", lambda w: w.goto_dialog(), ""),
    ("HFDS", "Go To Special...", None,
     "Go To Special: select every cell of one kind - blanks, constants, formulas, errors, "
     "visible cells only, current region, differences between rows/columns."),
    ("HVP", "Paste", lambda w: w.paste(), ""),
    ("HVF", "Paste formulas", lambda w: w.paste(), ""),
    ("HVV", "Paste values", lambda w: w.paste(values_only=True), ""),
    ("HVR", "Paste formatting", lambda w: w.paste(formats_only=True), ""),
    ("HVT", "Paste transposed", lambda w: w.paste(transpose=True), ""),
    ("HVS", "Paste Special...", None,
     "Paste Special dialog: paste only values/formulas/formats/comments/validation, combine with "
     "an operation (add, subtract, multiply, divide), skip blanks, transpose, paste link."),
    ("HC", "Copy", lambda w: w.copy(), ""),
    ("HX", "Cut", lambda w: w.cut(), ""),
    ("HFP", "Format Painter", lambda w: w.a_painter.trigger(), ""),
    # ---------------- Data
    ("ASA", "Sort A to Z", lambda w: w.quick_sort(True), ""),
    ("ASD", "Sort Z to A", lambda w: w.quick_sort(False), ""),
    ("ASS", "Sort...", lambda w: w.sort_dialog(), ""),
    ("AT", "Filter", lambda w: w.toggle_filter(), ""),
    ("AC", "Clear filter", lambda w: w.clear_filters(), ""),
    ("AY", "Reapply filter", lambda w: w.reapply_filters(), ""),
    ("AM", "Remove duplicates...", lambda w: w.remove_duplicates(), ""),
    ("AE", "Text to Columns", None,
     "Text to Columns wizard: split one column into several by a delimiter (comma, tab, space...) "
     "or fixed widths, choosing each new column's data type."),
    ("AVV", "Data Validation", None,
     "Data Validation: restrict what can be typed in cells (whole numbers, decimals, a dropdown "
     "list, dates, text length, custom formula) with input messages and error alerts. Ekxel keeps "
     "validation from Excel files but can't create or enforce it."),
    ("AGG", "Group rows/columns", None,
     "Group: outline rows or columns so they can be collapsed and expanded with +/- buttons."),
    ("AUU", "Ungroup", None, "Ungroup: remove an outline group."),
    ("AWG", "Goal Seek", None,
     "What-If Analysis > Goal Seek: change one input cell until a formula cell reaches a target value."),
    ("AFF", "Flash Fill", None,
     "Flash Fill (Ctrl+E): fill a column by recognizing the pattern in examples you typed "
     "(e.g. split or combine names)."),
    # ---------------- View
    ("WFF", "Freeze / unfreeze panes", lambda w: w.freeze_panes(), ""),
    ("WFR", "Freeze top row", lambda w: w.freeze_top_row(), ""),
    ("WFC", "Freeze first column", lambda w: w.freeze_first_col(), ""),
    ("WVG", "Gridlines", lambda w: w.toggle_gridlines(), ""),
    ("WJ", "Zoom to 100%", lambda w: w.grid.set_zoom(1.0), ""),
    ("WQ", "Zoom...", _zoom_dialog, ""),
    ("WS", "Split panes", None,
     "Split: divide the window into up to four independently scrolling panes at the active cell."),
    ("WN", "New window", None, "New Window: open a second window onto the same workbook."),
    ("WI", "Page Break Preview", None, "Page Break Preview: show and drag the page breaks used for printing."),
    # ---------------- Formulas
    ("MUS", "AutoSum", lambda w: w.autosum("SUM"), ""),
    ("MF", "Insert function...", lambda w: w.insert_function_dialog(), ""),
    ("MH", "Show formulas", lambda w: w.toggle_show_formulas(), ""),
    ("MB", "Calculate now", lambda w: w.recalc_all(), ""),
    ("MMD", "Define name", None,
     "Define Name: give a cell, range or formula a name (e.g. TaxRate) to use in formulas. Ekxel "
     "evaluates names from Excel files but can't create them."),
    ("MN", "Name Manager", None, "Name Manager: list, edit and delete the workbook's defined names."),
    ("MP", "Trace precedents", None,
     "Trace Precedents: draw arrows from the cells a formula uses to the formula cell."),
    ("MD", "Trace dependents", None,
     "Trace Dependents: draw arrows from a cell to the formulas that use it."),
    ("MV", "Evaluate formula", None,
     "Evaluate Formula: step through a formula's calculation one part at a time."),
    # ---------------- Insert
    ("NV", "PivotTable", None,
     "Insert PivotTable: summarize a table by dragging fields into Rows, Columns, Values and "
     "Filters (sum/count/average...), with grouping and refresh."),
    ("NT", "Table", None, "Insert Table: same as Format as Table (banded rows, filters, total row, structured references)."),
    ("NC", "Charts", None,
     "Insert a chart (column, bar, line, pie, scatter, area...) from the selected data, with chart "
     "titles, axes, legend and styling. Ekxel also can't keep charts when re-saving Excel files."),
    ("NI", "Hyperlink", None, "Insert Link (Ctrl+K): make a cell a clickable link to a web page, file or another cell."),
    ("NP", "Pictures", None, "Insert Pictures: place an image on the sheet."),
    ("NSH", "Shapes", None, "Insert Shapes: draw rectangles, arrows, callouts and other shapes."),
    ("NX", "Text box", None, "Insert Text Box: a floating box of text over the grid."),
    ("NU", "Symbol", None, "Insert Symbol: pick a special character to insert."),
    # ---------------- Review
    ("RC", "New comment / note", None,
     "New Comment/Note: attach a note to a cell (shown by a red corner marker). Ekxel keeps notes "
     "from Excel files but can't show or add them."),
    ("RS", "Spelling", None, "Spelling (F7): check spelling in the sheet."),
    ("RPS", "Protect sheet", None, "Protect Sheet: see Home > Format > Protect Sheet."),
    # ---------------- Page Layout / printing
    ("PO", "Orientation", None,
     "Page setup and printing: orientation, paper size, margins, print area, print titles, "
     "scaling, headers/footers, and Print / Print Preview (Ctrl+P)."),
    ("FP", "Print", None, "Print (Ctrl+P): print preview and printing with page setup options."),
    # ---------------- File
    ("FS", "Save", lambda w: w.file_save(), ""),
    ("FA", "Save as...", lambda w: w.file_save_as(), ""),
    ("FO", "Open...", lambda w: w.file_open(), ""),
    ("FN", "New workbook", lambda w: w.file_new(), ""),
    ("FC", "Close", lambda w: w.close(), ""),
    # ---------------- Excel 2003 menu sequences (still work in Excel)
    ("ESV", "Paste values (2003)", lambda w: w.paste(values_only=True), ""),
    ("EST", "Paste formats (2003)", lambda w: w.paste(formats_only=True), ""),
    ("ESE", "Paste transposed (2003)", lambda w: w.paste(transpose=True), ""),
    ("ESF", "Paste formulas (2003)", lambda w: w.paste(), ""),
    ("ED", "Delete rows/columns (2003)", lambda w: w.delete_smart(), ""),
    ("EL", "Delete sheet (2003)", lambda w: w.delete_sheet(), ""),
    ("EM", "Move or copy sheet (2003)", lambda w: w.duplicate_sheet(), ""),
    ("EID", "Fill down (2003)", lambda w: w.fill_dir("down"), ""),
    ("EIR", "Fill right (2003)", lambda w: w.fill_dir("right"), ""),
    ("EAA", "Clear all (2003)", lambda w: w.clear_all(), ""),
    ("EAF", "Clear formats (2003)", lambda w: w.clear_formats(), ""),
    ("EAC", "Clear contents (2003)", lambda w: w.clear_contents(), ""),
    ("OCA", "AutoFit column (2003)", lambda w: w.autofit_selected_cols(), ""),
    ("OCW", "Column width (2003)", lambda w: w.column_width_dialog(), ""),
    ("OCH", "Hide columns (2003)", lambda w: w.hide_rows_cols("col", True), ""),
    ("OCU", "Unhide columns (2003)", lambda w: w.hide_rows_cols("col", False), ""),
    ("ORA", "AutoFit row (2003)", lambda w: w.autofit_selected_rows(), ""),
    ("ORE", "Row height (2003)", lambda w: w.row_height_dialog(), ""),
    ("ORH", "Hide rows (2003)", lambda w: w.hide_rows_cols("row", True), ""),
    ("ORU", "Unhide rows (2003)", lambda w: w.hide_rows_cols("row", False), ""),
    ("OHR", "Rename sheet (2003)", lambda w: w.rename_sheet(), ""),
    ("OE", "Format Cells (2003)", lambda w: w.format_cells(), ""),
    ("OD", "Conditional Formatting (2003)", None, "Same as Home > Conditional Formatting."),
    ("IR", "Insert rows (2003)", lambda w: w.insert_rows_cols("row"), ""),
    ("IC", "Insert columns (2003)", lambda w: w.insert_rows_cols("col"), ""),
    ("IW", "Insert worksheet (2003)", lambda w: w.add_sheet(), ""),
    ("IF", "Insert function (2003)", lambda w: w.insert_function_dialog(), ""),
    ("IE", "Insert cells (2003)", None, "Same as Home > Insert > Insert Cells (shift cells right/down)."),
    ("DS", "Sort (2003)", lambda w: w.sort_dialog(), ""),
    ("DFF", "AutoFilter (2003)", lambda w: w.toggle_filter(), ""),
    ("DFS", "Show all / clear filter (2003)", lambda w: w.clear_filters(), ""),
    ("DE", "Text to Columns (2003)", None, "Same as Data > Text to Columns."),
]

GROUPS = {
    "H": "Home", "HF": "Font, Find, Fill, Format Painter", "HFD": "Find & Select", "HFI": "Fill",
    "HA": "Alignment", "HB": "Borders", "HM": "Merge", "HI": "Insert", "HD": "Delete", "HO": "Format",
    "HOU": "Hide & Unhide", "HU": "AutoSum", "HE": "Clear", "HS": "Sort & Filter", "HV": "Paste",
    "A": "Data", "AS": "Sort", "AV": "Data Validation", "AG": "Group", "AU": "Ungroup", "AW": "What-If",
    "AF": "Flash Fill", "W": "View", "WF": "Freeze Panes", "WV": "Show", "M": "Formulas", "MU": "AutoSum",
    "MM": "Define Name", "N": "Insert", "NS": "Shapes", "R": "Review", "RP": "Protect", "P": "Page Layout",
    "F": "File", "E": "Edit (Excel 2003)", "ES": "Paste Special", "EI": "Fill", "EA": "Clear",
    "O": "Format (Excel 2003)", "OC": "Column", "OR": "Row", "OH": "Sheet", "I": "Insert (Excel 2003)",
    "D": "Data (Excel 2003)", "DF": "Filter",
}

BY_SEQ = {seq: (label, action, excel) for seq, label, action, excel in KEYTIPS}


def tried_log_path():
    return os.path.join(settings_dir(), "wishlist-tried.json")


def load_tried():
    try:
        with open(tried_log_path(), encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, ValueError):
        return {}


def record_tried(seq):
    data = load_tried()
    label = BY_SEQ[seq][0]
    e = data.setdefault(seq, {"label": label, "count": 0})
    e["label"] = label
    e["count"] = e.get("count", 0) + 1
    e["last"] = dt.datetime.now().isoformat(timespec="seconds")
    try:
        os.makedirs(os.path.dirname(tried_log_path()), exist_ok=True)
        with open(tried_log_path(), "w", encoding="utf-8") as fh:
            json.dump(data, fh, indent=1)
    except OSError:
        pass


def next_options(prefix):
    """{next_char: label} for sequences continuing prefix."""
    out = {}
    for seq, label, action, _ in KEYTIPS:
        if seq.startswith(prefix) and len(seq) > len(prefix):
            ch = seq[len(prefix)]
            nxt = prefix + ch
            if ch in out:
                continue
            if nxt in BY_SEQ:
                lab, act, _ = BY_SEQ[nxt]
                out[ch] = lab + ("" if act else "  (not yet)")
            else:
                out[ch] = GROUPS.get(nxt, "...") + " ›"
    return out


_TEXT_INPUTS = (QLineEdit, QPlainTextEdit, QTextEdit, QAbstractSpinBox)


class KeyTipHints(QFrame):
    def __init__(self, win):
        super().__init__(win, Qt.ToolTip | Qt.FramelessWindowHint)
        self.setStyleSheet("KeyTipHints { background: #2B2B2B; border: 1px solid #1A1A1A; border-radius: 4px; }"
                           f"QLabel {{ color: #F0F0F0; font: {UI_PT}pt '{UI_FONT}'; }}")
        lay = QVBoxLayout(self)
        lay.setContentsMargins(10, 6, 10, 8)
        self.label = QLabel()
        self.label.setTextFormat(Qt.RichText)
        lay.addWidget(self.label)

    def show_for(self, prefix, opts, anchor):
        path = " › ".join(["Alt"] + list(prefix))
        cells = []
        for ch, lab in sorted(opts.items(), key=lambda kv: (not kv[0].isdigit(), kv[0])):
            cells.append(f"<td style='padding:1px 10px 1px 0'><span style='background:#F0F0F0;color:#111;"
                         f"padding:0 4px;font-weight:bold'>&nbsp;{ch}&nbsp;</span>&nbsp;{lab}</td>")
        rows = ["<tr>" + "".join(cells[i:i + 4]) + "</tr>" for i in range(0, len(cells), 4)]
        title = GROUPS.get(prefix, "KeyTips") if prefix else "KeyTips: type a letter (Esc to cancel)"
        self.label.setText(f"<b>{path}</b> &nbsp; <span style='color:#BBBBBB'>{title}</span>"
                           f"<table style='margin-top:4px'>{''.join(rows)}</table>")
        self.adjustSize()
        self.move(anchor - QPoint(0, self.height()))
        self.show()
        self.raise_()


class KeyTipController(QObject):
    """Application event filter for one main window."""

    def __init__(self, win):
        super().__init__(win)
        self.win = win
        self.active = False
        self.buf = ""
        self.alt_down = False
        self.alt_used = False
        self.hints = KeyTipHints(win)
        self._timeout = QTimer(self)
        self._timeout.setSingleShot(True)
        self._timeout.setInterval(8000)
        self._timeout.timeout.connect(self.cancel)
        QApplication.instance().installEventFilter(self)

    def _mine(self, obj):
        return isinstance(obj, QWidget) and obj.window() is self.win

    def _typing(self):
        fw = QApplication.focusWidget()
        if isinstance(fw, _TEXT_INPUTS):
            return True
        return isinstance(fw, QComboBox) and fw.isEditable()

    def eventFilter(self, obj, ev):
        t = ev.type()
        if t not in (QEvent.KeyPress, QEvent.KeyRelease, QEvent.ShortcutOverride, QEvent.MouseButtonPress,
                     QEvent.WindowDeactivate):
            return False
        if t in (QEvent.MouseButtonPress, QEvent.WindowDeactivate):
            if self.active:
                self.cancel()
            return False
        if not self._mine(obj):
            return False
        key = ev.key()
        mods = ev.modifiers()
        if t == QEvent.ShortcutOverride:
            # stop Alt+letter menu mnemonics so the KeyTip sequence gets the key
            if (self.active or (mods & Qt.AltModifier and not mods & Qt.ControlModifier)) and \
                    self._is_tip_key(key) and not self._typing():
                ev.accept()
                return True
            return False
        if self._typing() and not self.active:
            return False
        if t == QEvent.KeyPress:
            if ev.isAutoRepeat():
                return self.active
            if key == Qt.Key_Alt:
                if self.active:
                    self.cancel()
                else:
                    self.alt_down = True
                    self.alt_used = False
                return True
            if self.active:
                if key == Qt.Key_Escape:
                    self.back()
                    return True
                if self._is_tip_key(key):
                    self.feed(chr(key))
                    return True
                self.cancel()
                return False
            if self.alt_down and mods & Qt.AltModifier and not mods & Qt.ControlModifier and self._is_tip_key(key):
                self.alt_used = True
                self.active = True
                self.buf = ""
                self.feed(chr(key))
                return True
            return False
        if t == QEvent.KeyRelease and key == Qt.Key_Alt:
            down, used = self.alt_down, self.alt_used
            self.alt_down = False
            if down and not used and not self.active:
                self.start()
            return True
        return self.active

    @staticmethod
    def _is_tip_key(key):
        return (Qt.Key_A <= key <= Qt.Key_Z) or (Qt.Key_0 <= key <= Qt.Key_9)

    # ------------------------------------------------------------ state machine
    def start(self):
        if self.win.grid.editing:
            return
        self.active = True
        self.buf = ""
        self._show()

    def cancel(self):
        self.active = False
        self.buf = ""
        self.hints.hide()
        self._timeout.stop()

    def back(self):
        if self.buf:
            self.buf = self.buf[:-1]
            self._show()
        else:
            self.cancel()

    def _show(self):
        g = self.win.grid
        anchor = g.mapToGlobal(QPoint(g.rw + 4, g.height() - 6))
        self.hints.show_for(self.buf, next_options(self.buf), anchor)
        self._timeout.start()

    def feed(self, ch):
        self.buf += ch.upper()
        seq = self.buf
        entry = BY_SEQ.get(seq)
        if entry is not None:
            label, action, _ = entry
            self.cancel()
            spoken = "Alt " + " ".join(seq)
            if action is None:
                record_tried(seq)
                self.win.statusBar().showMessage(
                    f"{spoken} ({label}) isn't in Ekxel yet - added to your feature wishlist "
                    f"(Help > Feature Wishlist).", 6000)
                return
            self.win.statusBar().showMessage(f"{spoken}: {label}", 2500)
            QTimer.singleShot(0, lambda: action(self.win))
            return
        if next_options(seq):
            self._show()
            return
        self.cancel()
        self.win.statusBar().showMessage(f"No KeyTip for Alt {' '.join(seq)}", 3000)
