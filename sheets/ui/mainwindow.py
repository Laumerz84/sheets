"""Main application window."""
import datetime as dt
import json
import math
import os

from PySide6.QtCore import QPoint, QSize, Qt, QTimer, QUrl
from PySide6.QtGui import (QAction, QActionGroup, QColor, QFont, QFontMetrics,
                           QIcon, QKeySequence, QPainter, QPixmap, QUndoStack)
from PySide6.QtWidgets import (QApplication, QComboBox, QDialog,
                               QDialogButtonBox, QFileDialog, QFontComboBox,
                               QFrame, QHBoxLayout, QInputDialog, QLabel,
                               QLineEdit, QListWidget, QMainWindow, QMenu,
                               QMessageBox, QProgressDialog, QPushButton,
                               QSizePolicy, QSlider, QTabBar, QToolBar,
                               QToolButton, QVBoxLayout, QWidget,
                               QPlainTextEdit, QStyle)

from .. import ops
from ..errors import XLError
from ..fileio import (OPEN_FILTER, READABLE, SAVE_FILTERS, WRITABLE,
                      open_file, save_file)
from ..formula import shift_formula
from ..functions import ALL_NAMES, SIGNATURES
from ..numfmt import (compile_format, datetime_to_serial, format_value,
                      is_date_format, parse_input)
from ..refs import MAX_COLS, MAX_ROWS, addr, col_name, key, parse_range, range_addr
from ..values import is_num
from ..workbook import DEFAULT_STYLE, Sheet, intern_style, new_workbook
from . import style as S
from .commands import MetaCommand, SnapshotCommand, StatesCommand
from .dialogs import (ColorMenu, FilterPopup, FindDialog, FormatCellsDialog,
                      SortDialog, category_of)
from .editor import CellEditor
from .grid import Grid

APP_NAME = "Sheets"
WINDOWS = []
SETTINGS_DIR = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "Sheets")
SETTINGS_FILE = os.path.join(SETTINGS_DIR, "settings.json")
DEFAULT_FOLDER = r"G:\Spreadsheets"  # used until the user picks one (File > Set Default Folder)

NUMBER_PRESETS = [
    ("General", "General"), ("Number", "0.00"), ("Currency", "$#,##0.00"),
    ("Accounting", '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'),
    ("Short Date", "m/d/yyyy"), ("Long Date", "dddd, mmmm d, yyyy"), ("Time", "h:mm:ss AM/PM"),
    ("Percentage", "0.00%"), ("Fraction", "# ?/?"), ("Scientific", "0.00E+00"), ("Text", "@"),
]


def load_settings():
    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, ValueError):
        return {}


def save_settings(data):
    try:
        os.makedirs(SETTINGS_DIR, exist_ok=True)
        with open(SETTINGS_FILE, "w", encoding="utf-8") as fh:
            json.dump(data, fh, indent=1)
    except OSError:
        pass


def app_icon():
    pm = QPixmap(64, 64)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    p.setBrush(S.ACCENT)
    p.setPen(Qt.NoPen)
    p.drawRoundedRect(4, 4, 56, 56, 10, 10)
    p.setPen(QColor("#FFFFFF"))
    for i in range(3):
        y = 22 + i * 12
        p.drawLine(14, y, 50, y)
    for i in range(2):
        x = 26 + i * 12
        p.drawLine(x, 14, x, 50)
    p.setPen(QColor(255, 255, 255, 230))
    p.drawRect(14, 14, 36, 36)
    p.end()
    return QIcon(pm)


class MainWindow(QMainWindow):
    def __init__(self, wb=None):
        super().__init__()
        self.wb = wb or new_workbook()
        self.settings = load_settings()
        from .commands import GuardedUndoStack
        self.undo = GuardedUndoStack(self, lambda: self.ai_tools.user_locked(), self._claude_busy_message)
        self.undo.setUndoLimit(300)
        self.undo.cleanChanged.connect(lambda _: self.update_title())
        self.undo.indexChanged.connect(lambda _: self.update_title())
        self.undo.indexChanged.connect(self._cancel_stale_cut)
        self._pasting = False
        self.clip = None
        self.clip_text = None
        self.find_dlg = None
        self._csv_warned = False
        self._lost_ack = False
        self._syncing = False
        self.setWindowIcon(app_icon())
        self.setAcceptDrops(True)
        self.setAttribute(Qt.WA_DeleteOnClose)

        self.grid = Grid()
        import uuid
        from ..ai.panel import ClaudePanel
        from ..ai.tools import WorkbookTools
        self.ai_token = uuid.uuid4().hex   # identifies this window to the Claude bridge
        self.ai_tools = WorkbookTools(self)
        self.claude_panel = ClaudePanel(self)
        self.addDockWidget(Qt.RightDockWidgetArea, self.claude_panel)
        self.claude_panel.hide()
        self._build_actions()
        self._build_menus()
        self._build_toolbar()
        self._build_central()
        self._build_statusbar()
        self._connect_grid()

        self.stats_timer = QTimer(self)
        self.stats_timer.setSingleShot(True)
        self.stats_timer.setInterval(60)
        self.stats_timer.timeout.connect(self._update_stats)

        self.rebuild_tabs()
        self.show_sheet(self.wb.sheets[self.wb.active])
        self.update_title()
        geo = self.settings.get("geometry")
        if geo:
            try:
                self.restoreGeometry(bytes.fromhex(geo))
            except ValueError:
                self.resize(1400, 860)
        else:
            self.resize(1400, 860)
        WINDOWS.append(self)
        QTimer.singleShot(0, self.grid.setFocus)

    def showEvent(self, e):
        super().showEvent(e)
        if not getattr(self, "_identity_set", False):
            from ..winshell import set_window_identity
            self._identity_set = set_window_identity(self.winId())

    # ================================================================ building
    def _act(self, text, slot, shortcut=None, icon=None, checkable=False, tip=None):
        a = QAction(text, self)
        if shortcut:
            if isinstance(shortcut, (list, tuple)):
                a.setShortcuts([QKeySequence(s) for s in shortcut])
            else:
                a.setShortcut(QKeySequence(shortcut))
        if icon is not None:
            a.setIcon(icon)
        if checkable:
            a.setCheckable(True)
        if tip:
            a.setToolTip(tip)
        elif shortcut:
            sc = shortcut[0] if isinstance(shortcut, (list, tuple)) else shortcut
            a.setToolTip(f"{text.replace('&', '')} ({QKeySequence(sc).toString(QKeySequence.NativeText)})")
        a.triggered.connect(slot)
        self.addAction(a)
        return a

    def _build_actions(self):
        A = self._act
        G = S.glyph_icon
        self.a_new = A("&New", self.file_new, "Ctrl+N", G(S.G_NEW))
        self.a_open = A("&Open...", self.file_open, "Ctrl+O", G(S.G_OPEN))
        self.a_save = A("&Save", self.file_save, "Ctrl+S", G(S.G_SAVE))
        self.a_saveas = A("Save &As...", self.file_save_as, ["F12", "Ctrl+Shift+S"])
        self.a_default_dir = A("Set &Default Folder...", self.set_default_folder,
                               tip="Choose the folder Open and Save As start in")
        self.a_close = A("&Close", self.close, "Ctrl+W")
        self.a_exit = A("E&xit", self.quit_all)
        self.a_undo = A("&Undo", self.do_undo, "Ctrl+Z", G(S.G_UNDO))
        self.a_redo = A("&Redo", self.do_redo, ["Ctrl+Y", "Ctrl+Shift+Z"], G(S.G_REDO))
        self.undo.canUndoChanged.connect(self.a_undo.setEnabled)
        self.undo.canRedoChanged.connect(self.a_redo.setEnabled)
        self.a_undo.setEnabled(False)
        self.a_redo.setEnabled(False)
        self.a_cut = A("Cu&t", self.cut, "Ctrl+X", G(S.G_CUT))
        self.a_copy = A("&Copy", self.copy, "Ctrl+C", G(S.G_COPY))
        self.a_paste = A("&Paste", self.paste, "Ctrl+V", G(S.G_PASTE))
        self.a_paste_values = A("Paste &Values", lambda: self.paste(values_only=True), "Ctrl+Shift+V")
        self.a_paste_formats = A("Paste &Formatting", lambda: self.paste(formats_only=True))
        self.a_paste_transpose = A("Paste &Transposed", lambda: self.paste(transpose=True))
        self.a_fill_down = A("Fill &Down", lambda: self.fill_dir("down"), "Ctrl+D")
        self.a_fill_right = A("Fill &Right", lambda: self.fill_dir("right"), "Ctrl+R")
        self.a_clear_contents = A("Clear &Contents", self.clear_contents)
        self.a_clear_formats = A("Clear &Formats", self.clear_formats, icon=G(S.G_CLEAR))
        self.a_clear_all = A("Clear &All", self.clear_all)
        self.a_find = A("&Find...", lambda: self.show_find(False), "Ctrl+F", G(S.G_FIND))
        self.a_replace = A("&Replace...", lambda: self.show_find(True), "Ctrl+H")
        self.a_goto = A("&Go To...", self.goto_dialog, ["Ctrl+G", "F5"])
        self.a_recalc = A("Recalculate &Now", self.recalc_all, "F9")

        self.a_bold = A("&Bold", lambda: self.toggle_font("bold"), ["Ctrl+B", "Ctrl+2"],
                        S.text_icon("B", bold=True, family="Georgia", px=15), checkable=True)
        self.a_italic = A("&Italic", lambda: self.toggle_font("italic"), ["Ctrl+I", "Ctrl+3"],
                          S.text_icon("I", italic=True, family="Georgia", px=15), checkable=True)
        self.a_underline = A("&Underline", lambda: self.toggle_font("underline"), ["Ctrl+U", "Ctrl+4"],
                             S.text_icon("U", underline=True, family="Georgia", px=15), checkable=True)
        self.a_strike = A("S&trikethrough", lambda: self.toggle_font("strike"), "Ctrl+5",
                          S.text_icon("ab", strike=True, px=12), checkable=True)
        self.a_left = A("Align &Left", lambda: self.set_align("halign", "left"), icon=S.lines_icon("left"), checkable=True)
        self.a_center = A("&Center", lambda: self.set_align("halign", "center"), icon=S.lines_icon("center"), checkable=True)
        self.a_right = A("Align &Right", lambda: self.set_align("halign", "right"), icon=S.lines_icon("right"), checkable=True)
        self.a_top = A("Top Align", lambda: self.set_align("valign", "top"), icon=S.lines_icon("top"), checkable=True)
        self.a_middle = A("Middle Align", lambda: self.set_align("valign", "center"), icon=S.lines_icon("middle"), checkable=True)
        self.a_bottom = A("Bottom Align", lambda: self.set_align("valign", None), icon=S.lines_icon("bottom"), checkable=True)
        self.a_wrap = A("&Wrap Text", self.toggle_wrap, icon=S.lines_icon("wrap"), checkable=True)
        self.a_merge = A("&Merge && Center", self.merge_center, icon=S.lines_icon("merge"), checkable=True)
        self.a_merge_across = A("Merge &Across", self.merge_across)
        self.a_merge_cells = A("Merge Ce&lls", lambda: self.merge_center(center=False))
        self.a_unmerge = A("&Unmerge Cells", self.unmerge)
        self.a_format_cells = A("Format &Cells...", lambda: self.format_cells(), "Ctrl+1")
        self.a_currency = A("Currency Format", lambda: self.apply_numfmt("$#,##0.00"), "Ctrl+Shift+$",
                            S.text_icon("$", px=14, bold=True))
        self.a_percent = A("Percent Style", lambda: self.apply_numfmt("0%"), "Ctrl+Shift+%",
                           S.text_icon("%", px=14, bold=True))
        self.a_comma = A("Comma Style", lambda: self.apply_numfmt("#,##0.00"), "Ctrl+Shift+!",
                         S.text_icon(",", px=18, bold=True))
        self.a_datefmt = A("Date Format", lambda: self.apply_numfmt("d-mmm-yy"), "Ctrl+Shift+#")
        self.a_timefmt = A("Time Format", lambda: self.apply_numfmt("h:mm AM/PM"), "Ctrl+Shift+@")
        self.a_generalfmt = A("General Format", lambda: self.apply_numfmt("General"), "Ctrl+Shift+~")
        self.a_inc_dec = A("Increase Decimal", lambda: self.change_decimals(1), icon=S.lines_icon("dec_inc"))
        self.a_dec_dec = A("Decrease Decimal", lambda: self.change_decimals(-1), icon=S.lines_icon("dec_dec"))

        self.a_ins_rows = A("Insert Sheet &Rows", lambda: self.insert_rows_cols("row"), icon=S.lines_icon("insert_row"))
        self.a_ins_cols = A("Insert Sheet &Columns", lambda: self.insert_rows_cols("col"))
        self.a_del_rows = A("Delete Sheet Ro&ws", lambda: self.delete_rows_cols("row"), icon=S.lines_icon("delete_row"))
        self.a_del_cols = A("Delete Sheet Colu&mns", lambda: self.delete_rows_cols("col"))
        self.a_insert_smart = A("Insert", self.insert_smart, ["Ctrl++", "Ctrl+Shift+="])
        self.a_delete_smart = A("Delete", self.delete_smart, "Ctrl+-")
        self.a_hide_rows = A("Hide Rows", lambda: self.hide_rows_cols("row", True), "Ctrl+9")
        self.a_unhide_rows = A("Unhide Rows", lambda: self.hide_rows_cols("row", False), "Ctrl+Shift+9")
        self.a_hide_cols = A("Hide Columns", lambda: self.hide_rows_cols("col", True), "Ctrl+0")
        self.a_unhide_cols = A("Unhide Columns", lambda: self.hide_rows_cols("col", False), "Ctrl+Shift+0")
        self.a_col_width = A("Column &Width...", self.column_width_dialog)
        self.a_row_height = A("Row &Height...", self.row_height_dialog)
        self.a_autofit_cols = A("AutoFit Column Width", self.autofit_selected_cols)
        self.a_autofit_rows = A("AutoFit Row Height", self.autofit_selected_rows)

        self.a_new_sheet = A("Insert &Sheet", self.add_sheet, "Shift+F11", G(S.G_ADD))
        self.a_del_sheet = A("Delete Sheet", self.delete_sheet)
        self.a_rename_sheet = A("Rename Sheet...", self.rename_sheet)
        self.a_dup_sheet = A("Duplicate Sheet", self.duplicate_sheet)
        self.a_next_sheet = A("Next Sheet", lambda: self.cycle_sheet(1), "Ctrl+PgDown")
        self.a_prev_sheet = A("Previous Sheet", lambda: self.cycle_sheet(-1), "Ctrl+PgUp")

        self.a_autosum = A("&AutoSum", lambda: self.autosum("SUM"), "Alt+=", S.text_icon("Σ", px=17, family="Cambria Math"))
        self.a_insert_func = A("Insert &Function...", self.insert_function_dialog, "Shift+F3",
                               S.text_icon("fx", italic=True, family="Cambria", px=14))
        self.a_today = A("Insert Today's Date", lambda: self.insert_now(False), "Ctrl+;")
        self.a_now = A("Insert Current Time", lambda: self.insert_now(True), "Ctrl+Shift+;")

        self.a_sort_az = A("Sort A to Z", lambda: self.quick_sort(True), icon=S.lines_icon("sort_az"))
        self.a_sort_za = A("Sort Z to A", lambda: self.quick_sort(False), icon=S.lines_icon("sort_za"))
        self.a_sort = A("Custom &Sort...", self.sort_dialog)
        self.a_filter = A("&Filter", self.toggle_filter, "Ctrl+Shift+L", S.lines_icon("filter"), checkable=True)
        self.a_clear_filter = A("&Clear Filters", self.clear_filters)
        self.a_reapply = A("Re&apply Filters", self.reapply_filters, "Ctrl+Alt+L")
        self.a_filter_value = A("Filter by Selected Cell's Value", self.filter_by_value)
        self.a_dedupe = A("Remove &Duplicates...", self.remove_duplicates)
        self.a_painter = A("Format Painter", self.start_format_painter, icon=G(S.G_BRUSH, fallback="P"),
                           checkable=True, tip="Format Painter: copy formatting from the selection to the next cells you select")

        self.a_freeze = A("&Freeze Panes", self.freeze_panes, icon=S.lines_icon("freeze"))
        self.a_freeze_row = A("Freeze Top &Row", self.freeze_top_row)
        self.a_freeze_col = A("Freeze First &Column", self.freeze_first_col)
        self.a_unfreeze = A("&Unfreeze Panes", lambda: self.set_freeze((0, 0)))
        self.a_gridlines = A("&Gridlines", self.toggle_gridlines, checkable=True)
        self.a_show_formulas = A("Show F&ormulas", self.toggle_show_formulas, "Ctrl+`", checkable=True)
        self.a_zoom_in = A("Zoom &In", lambda: self.grid.set_zoom(self.grid.zoom + 0.1), "Ctrl+Alt+=", G(S.G_ZOOMIN))
        self.a_zoom_out = A("Zoom &Out", lambda: self.grid.set_zoom(self.grid.zoom - 0.1), "Ctrl+Alt+-", G(S.G_ZOOMOUT))
        self.a_zoom_100 = A("Zoom &100%", lambda: self.grid.set_zoom(1.0), "Ctrl+Alt+0")
        self.a_claude = self.claude_panel.toggleViewAction()
        self.a_claude.setText("&Claude")
        self.a_claude.setShortcut(QKeySequence("Ctrl+Shift+A"))
        self.a_claude.setIcon(S.text_icon("✳", px=17, color="#D97757", family="Segoe UI Symbol"))
        self.a_claude.setToolTip("Claude: ask it to work on this workbook (Ctrl+Shift+A)")
        self.a_claude.toggled.connect(lambda on: on and QTimer.singleShot(0, self.claude_panel.input.setFocus))
        self.addAction(self.a_claude)
        self.a_shortcuts = A("&Keyboard Shortcuts", self.show_shortcuts, "F1")
        self.a_about = A("&About Sheets", self.about)
        self.a_register = A("Make Sheets the default for CSV/Excel files...", self.register_file_types)

    def _build_menus(self):
        mb = self.menuBar()
        m = mb.addMenu("&File")
        for a in (self.a_new, self.a_open):
            m.addAction(a)
        self.recent_menu = m.addMenu("Open &Recent")
        self.recent_menu.aboutToShow.connect(self._fill_recent)
        m.addSeparator()
        m.addAction(self.a_save)
        m.addAction(self.a_saveas)
        m.addSeparator()
        m.addAction(self.a_default_dir)
        m.addAction(self.a_register)
        m.addSeparator()
        m.addAction(self.a_close)
        m.addAction(self.a_exit)

        m = mb.addMenu("&Edit")
        for a in (self.a_undo, self.a_redo, None, self.a_cut, self.a_copy, self.a_paste):
            m.addSeparator() if a is None else m.addAction(a)
        ps = m.addMenu("Paste &Special")
        for a in (self.a_paste_values, self.a_paste_formats, self.a_paste_transpose):
            ps.addAction(a)
        m.addSeparator()
        fm = m.addMenu("F&ill")
        fm.addAction(self.a_fill_down)
        fm.addAction(self.a_fill_right)
        cm = m.addMenu("C&lear")
        for a in (self.a_clear_all, self.a_clear_formats, self.a_clear_contents):
            cm.addAction(a)
        m.addSeparator()
        for a in (self.a_find, self.a_replace, self.a_goto):
            m.addAction(a)

        m = mb.addMenu("&View")
        fz = m.addMenu("&Freeze Panes")
        for a in (self.a_freeze, self.a_freeze_row, self.a_freeze_col, self.a_unfreeze):
            fz.addAction(a)
        m.addAction(self.a_gridlines)
        m.addAction(self.a_show_formulas)
        m.addAction(self.a_claude)
        m.addSeparator()
        for a in (self.a_zoom_in, self.a_zoom_out, self.a_zoom_100):
            m.addAction(a)

        m = mb.addMenu("&Insert")
        for a in (self.a_ins_rows, self.a_ins_cols, self.a_new_sheet, None, self.a_autosum,
                  self.a_insert_func, None, self.a_today, self.a_now):
            m.addSeparator() if a is None else m.addAction(a)

        m = mb.addMenu("F&ormat")
        m.addAction(self.a_format_cells)
        m.addSeparator()
        for a in (self.a_bold, self.a_italic, self.a_underline, self.a_strike):
            m.addAction(a)
        m.addSeparator()
        nm = m.addMenu("&Number Format")
        for label, fmt in NUMBER_PRESETS:
            nm.addAction(label, lambda f=fmt: self.apply_numfmt(f))
        am = m.addMenu("&Alignment")
        for a in (self.a_left, self.a_center, self.a_right, None, self.a_top, self.a_middle, self.a_bottom,
                  None, self.a_wrap):
            am.addSeparator() if a is None else am.addAction(a)
        mm = m.addMenu("&Merge")
        for a in (self.a_merge, self.a_merge_across, self.a_merge_cells, self.a_unmerge):
            mm.addAction(a)
        m.addSeparator()
        rm = m.addMenu("&Rows")
        for a in (self.a_row_height, self.a_autofit_rows, self.a_hide_rows, self.a_unhide_rows):
            rm.addAction(a)
        cm = m.addMenu("&Columns")
        for a in (self.a_col_width, self.a_autofit_cols, self.a_hide_cols, self.a_unhide_cols):
            cm.addAction(a)
        sm = m.addMenu("&Sheet")
        for a in (self.a_rename_sheet, self.a_dup_sheet, self.a_del_sheet):
            sm.addAction(a)
        m.addSeparator()
        m.addAction(self.a_clear_formats)

        m = mb.addMenu("&Data")
        for a in (self.a_sort_az, self.a_sort_za, self.a_sort, None, self.a_filter, self.a_clear_filter,
                  self.a_reapply, None, self.a_dedupe, None, self.a_recalc):
            m.addSeparator() if a is None else m.addAction(a)

        m = mb.addMenu("&Help")
        m.addAction(self.a_shortcuts)
        m.addAction(self.a_about)

    def _tool_button(self, icon, tip, menu=None, slot=None, text=None):
        b = QToolButton()
        b.setIcon(icon)
        b.setToolTip(tip)
        b.setIconSize(QSize(S.ICON_SIZE, S.ICON_SIZE))
        if text:
            b.setText(text)
            b.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        if menu is not None:
            b.setMenu(menu)
            b.setPopupMode(QToolButton.MenuButtonPopup if slot else QToolButton.InstantPopup)
        if slot:
            b.clicked.connect(slot)
        return b

    def _build_toolbar(self):
        tb = QToolBar("Home")
        tb.setMovable(False)
        tb.setIconSize(QSize(S.ICON_SIZE, S.ICON_SIZE))
        self.addToolBar(tb)
        self.toolbar = tb
        for a in (self.a_new, self.a_open, self.a_save):
            tb.addAction(a)
        tb.addSeparator()
        tb.addAction(self.a_undo)
        tb.addAction(self.a_redo)
        tb.addSeparator()
        tb.addAction(self.a_cut)
        tb.addAction(self.a_copy)
        pm = QMenu(self)
        pm.addAction(self.a_paste)
        pm.addAction(self.a_paste_values)
        pm.addAction(self.a_paste_formats)
        pm.addAction(self.a_paste_transpose)
        tb.addWidget(self._tool_button(S.glyph_icon(S.G_PASTE), "Paste (Ctrl+V)", pm, self.paste))
        tb.addAction(self.a_painter)
        tb.addSeparator()

        self.font_box = QFontComboBox()
        self.font_box.setFixedWidth(150)
        self.font_box.setCurrentFont(QFont("Calibri"))
        self.font_box.activated.connect(lambda _: self.set_font_family(self.font_box.currentFont().family()))
        tb.addWidget(self.font_box)
        self.size_box = QComboBox()
        self.size_box.setEditable(True)
        self.size_box.setFixedWidth(56)
        self.size_box.addItems([str(s) for s in (8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72)])
        self.size_box.setCurrentText("11")
        self.size_box.lineEdit().returnPressed.connect(lambda: self.set_font_size(self.size_box.currentText()))
        self.size_box.activated.connect(lambda _: self.set_font_size(self.size_box.currentText()))
        tb.addWidget(self.size_box)
        for a in (self.a_bold, self.a_italic, self.a_underline, self.a_strike):
            tb.addAction(a)
        self.font_color = "#C00000"
        fcm = ColorMenu(self, "Automatic")
        fcm.color_chosen.connect(self._font_color_chosen)
        self.font_color_btn = self._tool_button(S.glyph_icon(S.G_FONTCOLOR, bar=self.font_color, fallback="A"),
                                                "Font Color", fcm, lambda: self.set_style_field("color", self.font_color))
        tb.addWidget(self.font_color_btn)
        self.fill_color = "#FFFF00"
        fim = ColorMenu(self, "No Fill")
        fim.color_chosen.connect(self._fill_color_chosen)
        self.fill_color_btn = self._tool_button(S.glyph_icon(S.G_FILL, bar=self.fill_color, fallback="▣"),
                                                "Fill Color", fim, lambda: self.set_style_field("fill", self.fill_color))
        tb.addWidget(self.fill_color_btn)
        bm = QMenu(self)
        for label, mode in (("Bottom Border", "bottom"), ("Top Border", "top"), ("Left Border", "left"),
                            ("Right Border", "right"), (None, None), ("No Border", "none"),
                            ("All Borders", "all"), ("Outside Borders", "outline"),
                            ("Thick Outside Borders", "thick_outline"), ("Inside Borders", "inside")):
            if label is None:
                bm.addSeparator()
            else:
                bm.addAction(label, lambda m=mode: self.apply_borders(m))
        self._last_border = "all"
        tb.addWidget(self._tool_button(S.lines_icon("borders"), "Borders", bm,
                                       lambda: self.apply_borders(self._last_border)))
        tb.addSeparator()
        for a in (self.a_left, self.a_center, self.a_right, self.a_top, self.a_middle, self.a_bottom, self.a_wrap):
            tb.addAction(a)
        mm = QMenu(self)
        for a in (self.a_merge, self.a_merge_across, self.a_merge_cells, self.a_unmerge):
            mm.addAction(a)
        self.merge_btn = self._tool_button(S.lines_icon("merge"), "Merge & Center", mm, self.merge_center)
        self.merge_btn.setCheckable(True)
        tb.addWidget(self.merge_btn)
        tb.addSeparator()
        self.numfmt_box = QComboBox()
        self.numfmt_box.setFixedWidth(120)
        for label, fmt in NUMBER_PRESETS:
            self.numfmt_box.addItem(label, fmt)
        self.numfmt_box.addItem("More Formats...", None)
        self.numfmt_box.activated.connect(self._numfmt_chosen)
        tb.addWidget(self.numfmt_box)
        for a in (self.a_currency, self.a_percent, self.a_comma, self.a_inc_dec, self.a_dec_dec):
            tb.addAction(a)
        tb.addSeparator()
        im = QMenu(self)
        im.addAction(self.a_ins_rows)
        im.addAction(self.a_ins_cols)
        im.addAction(self.a_new_sheet)
        tb.addWidget(self._tool_button(S.lines_icon("insert_row"), "Insert", im, self.insert_smart))
        dm = QMenu(self)
        dm.addAction(self.a_del_rows)
        dm.addAction(self.a_del_cols)
        dm.addAction(self.a_del_sheet)
        tb.addWidget(self._tool_button(S.lines_icon("delete_row"), "Delete", dm, self.delete_smart))
        fmenu = QMenu(self)
        for a in (self.a_row_height, self.a_autofit_rows, self.a_col_width, self.a_autofit_cols, None,
                  self.a_hide_rows, self.a_unhide_rows, self.a_hide_cols, self.a_unhide_cols, None,
                  self.a_format_cells):
            fmenu.addSeparator() if a is None else fmenu.addAction(a)
        tb.addWidget(self._tool_button(S.glyph_icon("", fallback="F"), "Format rows, columns and cells", fmenu))
        tb.addSeparator()
        sm = QMenu(self)
        for name in ("SUM", "AVERAGE", "COUNT", "MAX", "MIN"):
            sm.addAction(name.title() if name != "COUNT" else "Count Numbers", lambda n=name: self.autosum(n))
        sm.addSeparator()
        sm.addAction(self.a_insert_func)
        tb.addWidget(self._tool_button(S.text_icon("Σ", px=17, family="Cambria Math"), "AutoSum (Alt+=)", sm,
                                       lambda: self.autosum("SUM")))
        srt = QMenu(self)
        for a in (self.a_sort_az, self.a_sort_za, self.a_sort, None, self.a_filter, self.a_clear_filter, self.a_reapply):
            srt.addSeparator() if a is None else srt.addAction(a)
        tb.addWidget(self._tool_button(S.lines_icon("sort_az"), "Sort & Filter", srt, None))
        tb.addAction(self.a_filter)
        tb.addAction(self.a_find)
        fzm = QMenu(self)
        for a in (self.a_freeze, self.a_freeze_row, self.a_freeze_col, self.a_unfreeze):
            fzm.addAction(a)
        tb.addWidget(self._tool_button(S.lines_icon("freeze"), "Freeze Panes", fzm, None))
        tb.addSeparator()
        tb.addAction(self.a_claude)

    def _build_central(self):
        central = QWidget()
        v = QVBoxLayout(central)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(0)
        # formula bar row
        fb = QWidget()
        fb.setStyleSheet("background: #FFFFFF;")
        h = QHBoxLayout(fb)
        h.setContentsMargins(6, 4, 6, 4)
        h.setSpacing(4)
        self.name_box = QLineEdit()
        self.name_box.setFixedWidth(110)
        self.name_box.setToolTip("Name Box: type a cell or range (e.g. B7 or A1:D20) and press Enter")
        self.name_box.returnPressed.connect(self._name_box_go)
        h.addWidget(self.name_box)
        self.fb_cancel = QToolButton()
        self.fb_cancel.setText("✕")
        self.fb_cancel.setToolTip("Cancel (Esc)")
        self.fb_enter = QToolButton()
        self.fb_enter.setText("✓")
        self.fb_enter.setToolTip("Enter")
        self.fb_fx = QToolButton()
        self.fb_fx.setIcon(S.text_icon("fx", italic=True, family="Cambria", px=14))
        self.fb_fx.setToolTip("Insert Function (Shift+F3)")
        for b in (self.fb_cancel, self.fb_enter, self.fb_fx):
            b.setAutoRaise(True)
            b.setFixedSize(26, 24)
            h.addWidget(b)
        self.fbar = CellEditor(in_bar=True)
        self.fbar.setFixedHeight(26)
        self.fbar.setStyleSheet("QPlainTextEdit { border: 1px solid #C8C8C8; background: white; }")
        self.fbar.setFont(QFont("Segoe UI", 10))
        h.addWidget(self.fbar, 1)
        self.fb_expand = QToolButton()
        self.fb_expand.setText("⌄")
        self.fb_expand.setCheckable(True)
        self.fb_expand.setAutoRaise(True)
        self.fb_expand.setToolTip("Expand formula bar")
        self.fb_expand.toggled.connect(lambda on: self.fbar.setFixedHeight(90 if on else 26))
        h.addWidget(self.fb_expand)
        v.addWidget(fb)
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setStyleSheet("color: #D4D4D4;")
        v.addWidget(line)
        # grid + vertical scrollbar
        gh = QHBoxLayout()
        gh.setContentsMargins(0, 0, 0, 0)
        gh.setSpacing(0)
        gh.addWidget(self.grid, 1)
        gh.addWidget(self.grid.vbar)
        v.addLayout(gh, 1)
        # sheet tabs + horizontal scrollbar
        bottom = QWidget()
        bottom.setStyleSheet("background: #F3F3F3;")
        bh = QHBoxLayout(bottom)
        bh.setContentsMargins(4, 0, 0, 0)
        bh.setSpacing(2)
        self.tabs = QTabBar()
        self.tabs.setMovable(True)
        self.tabs.setExpanding(False)
        self.tabs.setDrawBase(False)
        self.tabs.setUsesScrollButtons(True)
        self.tabs.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tabs.currentChanged.connect(self._tab_changed)
        self.tabs.tabMoved.connect(self._tab_moved)
        self.tabs.tabBarDoubleClicked.connect(lambda i: self.rename_sheet())
        self.tabs.customContextMenuRequested.connect(self._tab_menu)
        bh.addWidget(self.tabs, 0)
        add = QToolButton()
        add.setIcon(S.glyph_icon(S.G_ADD, color=S.ACCENT.name(), size=16, fallback="+"))
        add.setToolTip("New sheet (Shift+F11)")
        add.setAutoRaise(True)
        add.clicked.connect(self.add_sheet)
        bh.addWidget(add)
        bh.addStretch(1)
        self.grid.hbar.setFixedWidth(420)
        bh.addWidget(self.grid.hbar)
        v.addWidget(bottom)
        self.setCentralWidget(central)

        self.grid.attach_formula_bar(self.fbar)
        self.fb_cancel.clicked.connect(self.grid.cancel_edit)
        self.fb_enter.clicked.connect(lambda: self.grid.finish_edit(0, 0, False))
        self.fb_fx.clicked.connect(self.insert_function_dialog)

    def _build_statusbar(self):
        sb = self.statusBar()
        self.mode_lbl = QLabel("Ready")
        self.mode_lbl.setMinimumWidth(60)
        sb.addWidget(self.mode_lbl)
        self.filter_lbl = QLabel("")
        sb.addWidget(self.filter_lbl)
        self.stats_lbl = QLabel("")
        sb.addPermanentWidget(self.stats_lbl)
        zo = QToolButton()
        zo.setText("−")
        zo.setAutoRaise(True)
        zo.clicked.connect(lambda: self.grid.set_zoom(self.grid.zoom - 0.1))
        sb.addPermanentWidget(zo)
        self.zoom_slider = QSlider(Qt.Horizontal)
        self.zoom_slider.setRange(25, 400)
        self.zoom_slider.setFixedWidth(130)
        self.zoom_slider.setValue(100)
        self.zoom_slider.valueChanged.connect(lambda v: self.grid.set_zoom(v / 100))
        sb.addPermanentWidget(self.zoom_slider)
        zi = QToolButton()
        zi.setText("+")
        zi.setAutoRaise(True)
        zi.clicked.connect(lambda: self.grid.set_zoom(self.grid.zoom + 0.1))
        sb.addPermanentWidget(zi)
        self.zoom_lbl = QLabel("100%")
        self.zoom_lbl.setMinimumWidth(42)
        sb.addPermanentWidget(self.zoom_lbl)

    def _connect_grid(self):
        g = self.grid
        g.selection_changed.connect(self._selection_changed)
        g.commit_requested.connect(self.commit_cell)
        g.edit_state_changed.connect(self.mode_lbl.setText)
        g.context_menu_requested.connect(self._context_menu)
        g.fill_requested.connect(self.fill_drag)
        g.filter_popup_requested.connect(self.show_filter_popup)
        g.clear_requested.connect(self.clear_contents)
        g.col_widths_changed.connect(lambda o, n: self._push_meta({"col_widths": self._dict_pair(self.sheet.col_widths, o, n)}, "Column Width"))
        g.row_heights_changed.connect(lambda o, n: self._push_meta({"row_heights": self._dict_pair(self.sheet.row_heights, o, n)}, "Row Height"))
        g.autofit_cols_requested.connect(self.autofit_cols)
        g.autofit_rows_requested.connect(self.autofit_rows)
        g.zoom_changed.connect(self._zoom_changed)
        g.edit_text_changed.connect(self._edit_text_changed)
        g.selection_done.connect(self._selection_done)
        g.escape_pressed.connect(self._cancel_painter)
        g.read_only_hit.connect(self._claude_busy_message)

    # ================================================================ state helpers
    @property
    def sheet(self):
        return self.grid.sheet

    def update_title(self):
        name = os.path.basename(self.wb.path) if self.wb.path else "Book1"
        try:
            mod = "" if self.undo.isClean() else " •"
        except RuntimeError:  # window being destroyed
            return
        self.setWindowTitle(f"{name}{mod} - {APP_NAME}")

    def _claude_busy_message(self):
        self.statusBar().showMessage("Claude is working on this workbook - wait for it to finish, or press Stop "
                                     "in the Claude panel.", 4000)

    def is_pristine(self):
        return self.wb.path is None and self.undo.count() == 0 and not self.claude_panel.busy() and \
            not self.wb.sheets[0].values \
            and not self.wb.sheets[0].formulas

    def show_sheet(self, sheet):
        if sheet not in self.wb.sheets:
            return
        idx = self.wb.sheets.index(sheet)
        self.wb.active = idx
        if self.tabs.currentIndex() != idx:
            self.tabs.blockSignals(True)
            self.tabs.setCurrentIndex(idx)
            self.tabs.blockSignals(False)
        if self.grid.sheet is not sheet:
            self.grid.set_sheet(sheet)
        self._sync_view_actions()

    def rebuild_tabs(self):
        self.tabs.blockSignals(True)
        while self.tabs.count():
            self.tabs.removeTab(0)
        for sh in self.wb.sheets:
            i = self.tabs.addTab(sh.name)
            if sh.tab_color:
                self.tabs.setTabIcon(i, S.color_square_icon(sh.tab_color, 10))
        self.tabs.setCurrentIndex(self.wb.active)
        self.tabs.blockSignals(False)

    def after_change(self, sheet, relayout=False):
        if sheet is self.grid.sheet:
            self.grid.invalidate(relayout=relayout)
        self._selection_changed()
        self._update_filter_label()
        self.update_title()

    def after_structure(self):
        self.rebuild_tabs()
        active = self.wb.sheets[min(self.wb.active, len(self.wb.sheets) - 1)]
        for sh in list(self.grid._view_states):
            if sh not in self.wb.sheets:
                self.grid.forget_sheet(sh)
        if self.grid.sheet is not active:
            if self.grid.sheet not in self.wb.sheets:
                self.grid.sheet = None  # deleted: nothing to remember
            self.grid.set_sheet(active)
        else:
            self.grid.relayout()
        self.grid.invalidate()
        self._selection_changed()
        self._update_filter_label()
        self._sync_view_actions()
        self.update_title()

    def _dict_pair(self, current, old_vals, new_vals):
        old = dict(current)
        for k, v in old_vals.items():
            if v is None:
                old.pop(k, None)
            else:
                old[k] = v
        return (old, dict(current))

    def _push_meta(self, changes, text, relayout=True):
        self.undo.push(MetaCommand(self, self.sheet, changes, text, relayout))

    def _push_states(self, states, text, select=None, sheet=None):
        sheet = sheet or self.sheet
        if not states:
            return False
        # drop no-op entries (skipped for huge batches where it costs more than it saves)
        if len(states) > 100_000:
            clean = states
        else:
            clean = {k: st for k, st in states.items() if sheet.get_state(k) != st}
        if not clean:
            return False
        self.undo.push(StatesCommand(self, sheet, clean, text, select))
        return True

    def do_undo(self):
        if self.grid.editing:
            self.grid.cancel_edit()
            return
        self.undo.undo()

    def do_redo(self):
        if self.grid.editing:
            self.grid.cancel_edit()
            return
        self.undo.redo()

    def _prep(self):
        """Finish any in-progress cell edit before running a command."""
        if self.grid.editing:
            self.grid.finish_edit(0, 0, False)

    def hidden_rows(self):
        sh = self.sheet
        return sh.hidden_rows | sh.filter_hidden

    def visible_rects(self, rect):
        """Split a rect into pieces that skip hidden rows (filtered data)."""
        hidden = self.hidden_rows()
        r1, c1, r2, c2 = self.grid.clamp_rect(rect)
        if not hidden or not any(r1 <= r <= r2 for r in hidden):
            return [(r1, c1, r2, c2)]
        out = []
        start = None
        for r in range(r1, r2 + 1):
            if r in hidden:
                if start is not None:
                    out.append((start, c1, r - 1, c2))
                    start = None
            elif start is None:
                start = r
        if start is not None:
            out.append((start, c1, r2, c2))
        return out

    def sel_rects(self, visible=True):
        out = []
        for rect in self.grid.selected_rects():
            out.extend(self.visible_rects(rect) if visible else [self.grid.clamp_rect(rect)])
        return out

    # ================================================================ selection feedback
    def _selection_changed(self):
        if self.sheet is None:
            return
        r, c = self.grid.sel.active
        rects = self.grid.sel.rects
        last = rects[-1]
        if last[0] != last[2] or last[1] != last[3]:
            if self.grid.drag_mode == "select":
                self.name_box.setText(f"{last[2] - last[0] + 1}R x {last[3] - last[1] + 1}C")
            else:
                self.name_box.setText(addr(r, c))
        else:
            self.name_box.setText(addr(r, c))
        if not self.grid.editing:
            self.fbar.set_text(ops.edit_text_for(self.sheet, r, c))
            self.fbar.hide_popups()
        self._sync_format_controls()
        self.stats_timer.start()

    def _sync_format_controls(self):
        st = self.sheet.style(*self.grid.sel.active)
        self._syncing = True
        self.a_bold.setChecked(st.bold)
        self.a_italic.setChecked(st.italic)
        self.a_underline.setChecked(st.underline)
        self.a_strike.setChecked(st.strike)
        self.a_left.setChecked(st.halign == "left")
        self.a_center.setChecked(st.halign == "center")
        self.a_right.setChecked(st.halign == "right")
        self.a_top.setChecked(st.valign == "top")
        self.a_middle.setChecked(st.valign == "center")
        self.a_bottom.setChecked(st.valign in (None, "bottom"))
        self.a_wrap.setChecked(st.wrap)
        merged = self.sheet.merge_at(*self.grid.sel.active) is not None
        self.a_merge.setChecked(merged)
        self.merge_btn.setChecked(merged)
        self.font_box.blockSignals(True)
        self.font_box.setCurrentFont(QFont(st.font or S.DEFAULT_FONT_FAMILY))
        self.font_box.blockSignals(False)
        self.size_box.blockSignals(True)
        size = st.size or S.DEFAULT_FONT_SIZE
        self.size_box.setCurrentText(str(int(size)) if float(size).is_integer() else str(size))
        self.size_box.blockSignals(False)
        idx = next((i for i, (_, f) in enumerate(NUMBER_PRESETS) if f == st.numfmt), -1)
        self.numfmt_box.blockSignals(True)
        if idx >= 0:
            self.numfmt_box.setCurrentIndex(idx)
        else:
            cat = category_of(st.numfmt)
            labels = [l for l, _ in NUMBER_PRESETS]
            self.numfmt_box.setCurrentIndex(labels.index("Short Date") if cat == "Date" else
                                            labels.index("Time") if cat == "Time" else
                                            labels.index("Number") if cat in ("Custom",) else 0)
        self.numfmt_box.blockSignals(False)
        self._syncing = False

    def _sync_view_actions(self):
        if self.sheet is None:
            return
        self.a_gridlines.setChecked(self.sheet.show_grid)
        self.a_filter.setChecked(self.sheet.autofilter is not None)
        self.a_show_formulas.setChecked(self.grid.show_formulas)
        self.zoom_slider.blockSignals(True)
        self.zoom_slider.setValue(int(round(self.grid.zoom * 100)))
        self.zoom_slider.blockSignals(False)
        self.zoom_lbl.setText(f"{int(round(self.grid.zoom * 100))}%")
        fr, fc = self.sheet.freeze
        self.a_freeze.setText("&Unfreeze Panes" if (fr or fc) else "&Freeze Panes")
        self._update_filter_label()

    def _update_stats(self):
        if self.sheet is None:
            return
        sh = self.sheet
        rects = self.sel_rects()
        total = 0
        for r1, c1, r2, c2 in rects:
            total += (r2 - r1 + 1) * (c2 - c1 + 1)
        if total <= 1:
            self.stats_lbl.setText("")
            return
        count = 0
        nums = 0
        s = 0.0
        seen = set()
        for rect in rects:
            for k in ops.rect_keys(sh, rect, include_styles=False):
                if k in seen:
                    continue
                seen.add(k)
                v = sh.value_at_key(k)
                if v is None or v == "":
                    continue
                count += 1
                if is_num(v):
                    nums += 1
                    s += v
        if count < 2 and nums < 2:
            self.stats_lbl.setText("")
            return
        parts = []
        st = sh.style(*self.grid.sel.active)
        fmt = st.numfmt if st.numfmt != "General" and not st.numfmt == "@" else None

        def show(x):
            if fmt:
                return format_value(x, fmt)[0].strip()
            return format_value(round(x, 10), "General")[0]
        if nums:
            parts.append(f"Average: {show(s / nums)}")
        parts.append(f"Count: {count}")
        if nums:
            parts.append(f"Sum: {show(s)}")
        self.stats_lbl.setText("      ".join(parts))

    def _update_filter_label(self):
        sh = self.sheet
        if sh is not None and sh.autofilter and sh.filters:
            r1, _, r2, _ = sh.autofilter
            total = r2 - r1
            shown = total - len([r for r in sh.filter_hidden if r1 < r <= r2])
            self.filter_lbl.setText(f"{shown} of {total} records found")
        else:
            self.filter_lbl.setText("")

    def _zoom_changed(self, z):
        self._sync_view_actions()

    def _edit_text_changed(self, text):
        pass

    def _name_box_go(self):
        text = self.name_box.text().strip()
        self.goto_ref(text)
        self.grid.setFocus()

    def goto_ref(self, text):
        sheet = self.sheet
        ref = text
        if "!" in text:
            sname, ref = text.rsplit("!", 1)
            sname = sname.strip("'")
            target = self.wb.get_sheet(sname)
            if target is None:
                QMessageBox.warning(self, APP_NAME, f"There's no sheet named '{sname}'.")
                return
            sheet = target
        b = parse_range(ref)
        if b is None and text.upper() in self.wb.names:
            return self.goto_ref(self.wb.names[text.upper()].replace("$", ""))
        if b is None:
            QMessageBox.warning(self, APP_NAME, f"'{text}' isn't a valid cell reference.")
            return
        self.show_sheet(sheet)
        self.grid.set_selection([b], active=(b[0], b[1]))
        self.grid.ensure_visible(b[0], b[1])

    # ================================================================ commits
    def commit_cell(self, r, c, text, fill_selection):
        sh = self.sheet
        states = {}
        if fill_selection:
            for rect in self.sel_rects():
                r1, c1, r2, c2 = rect
                if (r2 - r1 + 1) * (c2 - c1 + 1) > 2_000_000:
                    continue
                for rr in range(r1, r2 + 1):
                    for cc in range(c1, c2 + 1):
                        t = shift_formula(text, rr - r, cc - c) if text.startswith("=") else text
                        states[key(rr, cc)] = ops.input_state(sh, rr, cc, t)
        else:
            st = ops.input_state(sh, r, c, text)
            if "\n" in text and not st[1].wrap:
                st = (st[0], st[1].with_(wrap=True))
            states[key(r, c)] = st
        widen = None if fill_selection else self._auto_widen(r, c, states.get(key(r, c)))
        if widen is None:
            self._push_states(states, "Typing")
            return
        self.undo.beginMacro("Typing")
        self._push_states(states, "Typing")
        self._push_meta({"col_widths": (dict(sh.col_widths), {**sh.col_widths, c: widen})}, "Typing")
        self.undo.endMacro()

    def _auto_widen(self, r, c, state):
        """Excel widens a default-width column when a typed number/date doesn't fit."""
        sh = self.sheet
        if state is None or state[0] is None or state[0][0] != "v" or c in sh.col_widths:
            return None
        v = state[0][1]
        st = state[1]
        if not is_num(v) or st.numfmt == "General":
            return None
        text = format_value(v, st.numfmt)[0]
        font, fm = self.grid.font_for(st)
        need = fm.horizontalAdvance(text) / self.grid.zoom + 10
        if need <= sh.col_width(c):
            return None
        return int(math.ceil(need))

    # ================================================================ clipboard
    def copy(self, cut=False):
        self._prep()
        sh = self.sheet
        if len(self.grid.sel.rects) > 1:
            QMessageBox.information(self, APP_NAME, "This command can't be used on multiple selections.")
            return
        rect = self.grid.clamp_rect(self.grid.sel.rects[0])
        hidden = self.hidden_rows()
        rows = [r for r in range(rect[0], rect[2] + 1) if r not in hidden] if hidden else None
        if rows == []:
            self.statusBar().showMessage("Nothing to copy: every selected row is hidden", 4000)
            return
        self.clip = ops.copy_rect(sh, rect, rows)
        self.clip.cut = cut
        self.clip_text = ops.clip_to_text(sh, self.clip, rows)
        QApplication.clipboard().setText(self.clip_text)
        self.grid.set_marquee(rect)
        self.statusBar().showMessage("Select destination and press Enter or choose Paste" if not cut else
                                     "Select destination and press Ctrl+V", 4000)

    def cut(self):
        self.copy(cut=True)

    def _cancel_stale_cut(self, _=None):
        """Like Excel: any edit after Cut cancels it (its addresses may no longer be right)."""
        if self.clip is not None and self.clip.cut and not self._pasting:
            self.clip = None
            self.clip_text = None
            try:
                self.grid.set_marquee(None)
            except RuntimeError:
                pass

    def paste(self, values_only=False, formats_only=False, transpose=False):
        self._prep()
        sh = self.sheet
        text = QApplication.clipboard().text()
        rect = self.grid.sel.rects[-1]
        r0, c0 = rect[0], rect[1]
        internal = self.clip is not None and text == self.clip_text
        if internal:
            clip = self.clip
            h, w = (clip.nrows, clip.ncols) if not transpose else (clip.ncols, clip.nrows)
            sel_h, sel_w = rect[2] - rect[0] + 1, rect[3] - rect[1] + 1
            if sel_h % h == 0 and sel_w % w == 0 and (sel_h > h or sel_w > w) and sel_h * sel_w < 2_000_000:
                dest = rect
            else:
                dest = (r0, c0, r0 + h - 1, c0 + w - 1)
            states = ops.paste_states(sh, clip, dest, values_only, formats_only, transpose)
            if clip.cut and clip.sheet is not None:
                src = clip.rect
                for rr in clip.rows:  # only the rows that were actually cut (not filtered-out ones)
                    for cc in range(src[1], src[3] + 1):
                        k = key(rr, cc)
                        if dest[0] <= rr <= dest[2] and dest[1] <= cc <= dest[3] and clip.sheet is sh:
                            continue
                        if clip.sheet is sh:
                            states.setdefault(k, (None, DEFAULT_STYLE))
                if clip.sheet is not sh:
                    self.undo.beginMacro("Cut and Paste")
                    clear = {key(rr, cc): (None, DEFAULT_STYLE) for rr in clip.rows
                             for cc in range(src[1], src[3] + 1)}
                    self._pasting = True
                    try:
                        self._push_states(clear, "Cut", sheet=clip.sheet)
                        self._push_states(states, "Paste", select=([dest], (dest[0], dest[1])))
                    finally:
                        self._pasting = False
                        self.undo.endMacro()
                else:
                    self._pasting = True
                    try:
                        self._push_states(states, "Cut and Paste", select=([dest], (dest[0], dest[1])))
                    finally:
                        self._pasting = False
                self.clip = None
                self.clip_text = None
                self.grid.set_marquee(None)
            else:
                self._push_states(states, "Paste", select=([dest], (dest[0], dest[1])))
            return
        if not text:
            return
        rows = ops.parse_clipboard_text(text)
        if transpose:
            width = max(len(r) for r in rows)
            rows = [[r[i] if i < len(r) else "" for r in rows] for i in range(width)]
        if len(rows) == 1 and len(rows[0]) == 1 and (rect[2] > rect[0] or rect[3] > rect[1]):
            # single value into a range: fill it (visible cells only)
            states = {}
            for rr1, cc1, rr2, cc2 in self.visible_rects(rect):
                for rr in range(rr1, rr2 + 1):
                    for cc in range(cc1, cc2 + 1):
                        states[key(rr, cc)] = ops.input_state(sh, rr, cc, rows[0][0])
            self._push_states(states, "Paste")
            return
        states = ops.text_paste_states(sh, rows, r0, c0)
        h = len(rows)
        w = max(len(r) for r in rows)
        self._push_states(states, "Paste", select=([(r0, c0, r0 + h - 1, c0 + w - 1)], (r0, c0)))

    # ================================================================ editing commands
    def clear_contents(self):
        self._prep()
        self._push_states(ops.clear_contents(self.sheet, self.sel_rects()), "Clear Contents")

    def clear_formats(self):
        self._prep()
        self._push_states(ops.clear_formats(self.sheet, self.sel_rects()), "Clear Formats")

    def clear_all(self):
        self._prep()
        self._push_states(ops.clear_all(self.sheet, self.sel_rects()), "Clear All")

    def fill_dir(self, direction):
        self._prep()
        states = {}
        for rect in self.sel_rects(visible=False):
            states.update(ops.fill_down_states(self.sheet, rect, direction))
        self._push_states(states, "Fill Down" if direction == "down" else "Fill Right")

    def fill_drag(self, src, target):
        self._push_states(ops.fill_states(self.sheet, src, target), "AutoFill",
                          select=([target], self.grid.sel.active))

    def insert_now(self, time):
        now = dt.datetime.now()
        text = now.strftime("%I:%M %p").lstrip("0") if time else f"{now.month}/{now.day}/{now.year}"
        w = self.grid.edit_widget if self.grid.editing else None
        if w is not None:
            w.insertPlainText(text)
            return
        self.grid.begin_edit(text=text, mode="enter")

    def recalc_all(self):
        self.wb.recalc(full=True)
        self.grid.invalidate()
        self.statusBar().showMessage("Recalculated", 1500)

    # ================================================================ formatting
    def restyle(self, fn, text):
        self._prep()
        rects = self.sel_rects()
        self._push_states(ops.restyle(self.sheet, rects, fn), text)

    def toggle_font(self, field):
        st = self.sheet.style(*self.grid.sel.active)
        new = not getattr(st, field)
        self.restyle(lambda s: s.with_(**{field: new}), field.title())

    def set_style_field(self, field, value):
        self.restyle(lambda s: s.with_(**{field: value}), "Format")

    def set_font_family(self, fam):
        if self._syncing:
            return
        self.set_style_field("font", None if fam == S.DEFAULT_FONT_FAMILY else fam)
        self.grid.setFocus()

    def set_font_size(self, text):
        if self._syncing:
            return
        try:
            v = float(text)
        except ValueError:
            return
        if not 1 <= v <= 409:
            return
        self.set_style_field("size", None if v == S.DEFAULT_FONT_SIZE else v)
        self.grid.setFocus()

    def _font_color_chosen(self, c):
        self.font_color = c
        self.font_color_btn.setIcon(S.glyph_icon(S.G_FONTCOLOR, bar=c or "#000000", fallback="A"))
        self.set_style_field("color", c)

    def _fill_color_chosen(self, c):
        self.fill_color = c
        self.fill_color_btn.setIcon(S.glyph_icon(S.G_FILL, bar=c or "#FFFFFF", fallback="▣"))
        self.set_style_field("fill", c)

    def set_align(self, field, value):
        st = self.sheet.style(*self.grid.sel.active)
        if field == "halign" and st.halign == value:
            value = None
        self.restyle(lambda s: s.with_(**{field: value}), "Align")

    def toggle_wrap(self):
        new = not self.sheet.style(*self.grid.sel.active).wrap
        self.restyle(lambda s: s.with_(wrap=new), "Wrap Text")

    def apply_numfmt(self, fmt):
        self.restyle(lambda s: s.with_(numfmt=fmt), "Number Format")

    def _numfmt_chosen(self, idx):
        fmt = self.numfmt_box.itemData(idx)
        if fmt is None:
            self.format_cells(tab=0)
        else:
            self.apply_numfmt(fmt)
        self.grid.setFocus()

    def change_decimals(self, delta):
        sh = self.sheet
        r, c = self.grid.sel.active
        st = sh.style(r, c)
        fmt = st.numfmt
        if fmt == "General":
            v = sh.value(r, c)
            dec = 0
            if is_num(v):
                s = format_value(v, "General")[0]
                if "." in s and "E" not in s:
                    dec = len(s.split(".")[1])
            base = "0" + ("." + "0" * dec if dec else "")
        else:
            base = fmt

        def adjust(f):
            if f == "General":
                f = base
            secs = f.split(";")
            out = []
            for sec in secs:
                # find the last run of number placeholders
                idx = max(sec.rfind("0"), sec.rfind("#"))
                if idx < 0:
                    out.append(sec)
                    continue
                # locate decimal part immediately before idx+1
                j = idx
                while j >= 0 and sec[j] in "0#":
                    j -= 1
                has_dot = j >= 0 and sec[j] == "."
                if delta > 0:
                    if has_dot:
                        sec = sec[:idx + 1] + "0" + sec[idx + 1:]
                    else:
                        sec = sec[:idx + 1] + ".0" + sec[idx + 1:]
                else:
                    if has_dot:
                        ndec = idx - j
                        if ndec <= 1:
                            sec = sec[:j] + sec[idx + 1:]
                        else:
                            sec = sec[:idx] + sec[idx + 1:]
                out.append(sec)
            return ";".join(out)
        new = adjust(fmt)
        self.restyle(lambda s: s.with_(numfmt=adjust(s.numfmt) if s.numfmt != "General" else new), "Decimals")

    def apply_borders(self, mode, side=("thin", "#000000")):
        self._prep()
        self._last_border = mode
        states = {}
        for rect in self.sel_rects():
            states.update(ops.borders_states(self.sheet, rect, mode, side))
        self._push_states(states, "Borders")

    def format_cells(self, tab=0):
        self._prep()
        sh = self.sheet
        r, c = self.grid.sel.active
        st = sh.style(r, c)
        v = sh.value(r, c)
        dlg = FormatCellsDialog(self, st, v if is_num(v) else None, tab)
        if dlg.exec() != QDialog.Accepted:
            return
        changes = dict(dlg.changes)
        border = dlg.border_choice()
        self.undo.beginMacro("Format Cells")
        if changes:
            self.restyle(lambda s: s.with_(**changes), "Format Cells")
        if border:
            self.apply_borders(*border)
        self.undo.endMacro()

    # ================================================================ format painter
    def start_format_painter(self):
        from .extras import copy_styles
        if not self.a_painter.isChecked():
            self.painter_styles = None
            self.grid.unsetCursor()
            return
        rect = self.grid.clamp_rect(self.grid.sel.rects[-1])
        if (rect[2] - rect[0] + 1) * (rect[3] - rect[1] + 1) > 200000:
            self.a_painter.setChecked(False)
            return
        self.painter_styles = copy_styles(self.sheet, rect)
        self.grid.set_marquee(rect)
        self.statusBar().showMessage("Select the cells to paint the formatting onto (Esc cancels)", 5000)

    def _cancel_painter(self):
        if getattr(self, "painter_styles", None):
            self.painter_styles = None
            self.a_painter.setChecked(False)

    def _selection_done(self):
        styles = getattr(self, "painter_styles", None)
        if not styles:
            return
        from .extras import paint_states
        self.painter_styles = None
        self.a_painter.setChecked(False)
        self.grid.set_marquee(None)
        dest = self.grid.clamp_rect(self.grid.sel.rects[-1])
        self._push_states(paint_states(self.sheet, styles, dest), "Format Painter")

    def remove_duplicates(self):
        from .extras import RemoveDuplicatesDialog, dedupe_states
        self._prep()
        rect, header = self._sort_range()
        dlg = RemoveDuplicatesDialog(self, self.sheet, rect, header)
        if dlg.exec() != QDialog.Accepted:
            return
        cols = dlg.columns()
        if not cols:
            return
        states, removed, kept = dedupe_states(self.sheet, rect, cols, dlg.header_chk.isChecked())
        if removed:
            self._push_states(states, "Remove Duplicates")
        QMessageBox.information(self, APP_NAME, f"{removed} duplicate row(s) found and removed; "
                                f"{kept} unique row(s) remain." if removed else "No duplicate values found.")

    # ================================================================ merge
    def merge_center(self, center=True):
        self._prep()
        sh = self.sheet
        rect = self.grid.clamp_rect(self.grid.sel.rects[-1])
        if any(m[0] <= rect[2] and m[2] >= rect[0] and m[1] <= rect[3] and m[3] >= rect[1] for m in sh.merges):
            return self.unmerge()
        if rect[0] == rect[2] and rect[1] == rect[3]:
            return
        if (rect[2] - rect[0] + 1) * (rect[3] - rect[1] + 1) > 100000:
            QMessageBox.information(self, APP_NAME, "That selection is too large to merge.")
            return
        self._merge_rects([rect], center)

    def merge_across(self):
        self._prep()
        r1, c1, r2, c2 = self.grid.clamp_rect(self.grid.sel.rects[-1])
        if c1 == c2:
            return
        self._merge_rects([(r, c1, r, c2) for r in range(r1, r2 + 1)], False)

    def _merge_rects(self, rects, center):
        sh = self.sheet
        lost = 0
        states = {}
        for rect in rects:
            r1, c1, r2, c2 = rect
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    if (r, c) != (r1, c1) and sh.has_content(r, c):
                        lost += 1
                        states[key(r, c)] = (None, sh.style(r, c))
            if center:
                k = key(r1, c1)
                content, st = sh.get_state(k)
                states[k] = (content, st.with_(halign="center"))
        if lost:
            if QMessageBox.question(self, APP_NAME, "Merging cells only keeps the upper-left value and discards "
                                    "the other values. Continue?") != QMessageBox.Yes:
                return
        new_merges = [m for m in sh.merges if not any(
            m[0] <= R[2] and m[2] >= R[0] and m[1] <= R[3] and m[3] >= R[1] for R in rects)] + list(rects)
        self.undo.beginMacro("Merge Cells")
        self._push_states(states, "Merge")
        self._push_meta({"merges": (list(sh.merges), new_merges)}, "Merge")
        self.undo.endMacro()
        self.grid.set_selection([rects[0] if len(rects) == 1 else (rects[0][0], rects[0][1], rects[-1][2], rects[-1][3])])

    def unmerge(self):
        self._prep()
        sh = self.sheet
        rect = self.grid.clamp_rect(self.grid.sel.rects[-1])
        keep = [m for m in sh.merges if not (m[0] <= rect[2] and m[2] >= rect[0] and m[1] <= rect[3] and m[3] >= rect[1])]
        if len(keep) != len(sh.merges):
            self._push_meta({"merges": (list(sh.merges), keep)}, "Unmerge")

    # ================================================================ rows / cols
    def _sel_rows(self):
        rows = set()
        for r1, c1, r2, c2 in self.grid.sel.rects:
            rows.update(range(r1, min(r2, max(r1, self.sheet.used_extent()[0] + 1000)) + 1))
        return sorted(rows)

    def _sel_cols(self):
        cols = set()
        for r1, c1, r2, c2 in self.grid.sel.rects:
            cols.update(range(c1, min(c2, max(c1, self.sheet.used_extent()[1] + 100)) + 1))
        return sorted(cols)

    def _full_cols_selected(self):
        return all(r1 == 0 and r2 >= MAX_ROWS - 1 for r1, c1, r2, c2 in self.grid.sel.rects)

    def _full_rows_selected(self):
        return all(c1 == 0 and c2 >= MAX_COLS - 1 for r1, c1, r2, c2 in self.grid.sel.rects)

    def insert_smart(self):
        self._prep()
        self.insert_rows_cols("col" if self._full_cols_selected() and not self._full_rows_selected() else "row")

    def delete_smart(self):
        self._prep()
        self.delete_rows_cols("col" if self._full_cols_selected() and not self._full_rows_selected() else "row")

    def insert_rows_cols(self, axis):
        self._prep()
        sh = self.sheet
        r1, c1, r2, c2 = self.grid.sel.rects[-1]
        if axis == "row":
            at, n = r1, min(r2 - r1 + 1, 10000)
            if sh.max_row + n >= MAX_ROWS:
                QMessageBox.warning(self, APP_NAME, "Can't insert rows: data would be pushed off the sheet.")
                return
            action = lambda: sh.insert_rows(at, n)
            text = "Insert Rows"
        else:
            at, n = c1, min(c2 - c1 + 1, 1000)
            if sh.max_col + n >= MAX_COLS:
                QMessageBox.warning(self, APP_NAME, "Can't insert columns: data would be pushed off the sheet.")
                return
            action = lambda: sh.insert_cols(at, n)
            text = "Insert Columns"
        self.undo.push(SnapshotCommand(self, sh, text, action, sheets=[sh]))

    def delete_rows_cols(self, axis):
        self._prep()
        sh = self.sheet
        raw = sorted((r1, r2) if axis == "row" else (c1, c2) for r1, c1, r2, c2 in self.grid.sel.rects)
        # merge overlapping/adjacent spans so no row is deleted twice
        spans = []
        for a, b in raw:
            if spans and a <= spans[-1][1] + 1:
                spans[-1] = (spans[-1][0], max(spans[-1][1], b))
            else:
                spans.append((a, b))
        spans.reverse()  # delete from the bottom/right up so earlier indexes stay valid

        def action():
            for a, b in spans:
                if axis == "row":
                    b = min(b, max(a, sh.used_extent()[0]))
                    sh.delete_rows(a, b - a + 1)
                else:
                    b = min(b, max(a, sh.used_extent()[1]))
                    sh.delete_cols(a, b - a + 1)
        self.undo.push(SnapshotCommand(self, sh, "Delete Rows" if axis == "row" else "Delete Columns", action, sheets=[sh]))

    def hide_rows_cols(self, axis, hide):
        self._prep()
        sh = self.sheet
        if axis == "row":
            rows = set(self._sel_rows())
            if not hide:
                lo, hi = min(rows), max(rows)
                rows = set(range(max(0, lo - 1), hi + 2))
            new = (sh.hidden_rows | rows) if hide else (sh.hidden_rows - rows)
            heights = dict(sh.row_heights)
            if not hide:
                for r in rows:
                    if heights.get(r) == 0:
                        heights.pop(r)
            self._push_meta({"hidden_rows": (set(sh.hidden_rows), new),
                             "row_heights": (dict(sh.row_heights), heights)}, "Hide Rows" if hide else "Unhide Rows")
        else:
            cols = set(self._sel_cols())
            if not hide:
                lo, hi = min(cols), max(cols)
                cols = set(range(max(0, lo - 1), hi + 2))
            new = (sh.hidden_cols | cols) if hide else (sh.hidden_cols - cols)
            widths = dict(sh.col_widths)
            if not hide:
                for c in cols:
                    if widths.get(c) == 0:
                        widths.pop(c)
            self._push_meta({"hidden_cols": (set(sh.hidden_cols), new),
                             "col_widths": (dict(sh.col_widths), widths)}, "Hide Columns" if hide else "Unhide Columns")

    def column_width_dialog(self):
        sh = self.sheet
        c = self.grid.sel.active[1]
        cur = (sh.col_width(c) - 5) / 7
        v, ok = QInputDialog.getDouble(self, "Column Width", "Column width (characters):", round(cur, 2), 0, 255, 2)
        if ok:
            px = int(round(v * 7 + 5)) if v > 0 else 0
            new = dict(sh.col_widths)
            for col in self._sel_cols():
                new[col] = px
            self._push_meta({"col_widths": (dict(sh.col_widths), new)}, "Column Width")

    def row_height_dialog(self):
        sh = self.sheet
        r = self.grid.sel.active[0]
        cur = sh.row_height(r) * 72 / 96
        v, ok = QInputDialog.getDouble(self, "Row Height", "Row height (points):", round(cur, 2), 0, 409, 2)
        if ok:
            px = int(round(v * 96 / 72))
            new = dict(sh.row_heights)
            for row in self._sel_rows():
                new[row] = px
            self._push_meta({"row_heights": (dict(sh.row_heights), new)}, "Row Height")

    def _text_width(self, k, st):
        text, color, v = self.grid.display(k, st)
        if not text:
            return 0
        font, fm = self.grid.font_for(st)
        lines = text.split("\n")
        return max(fm.horizontalAdvance(ln) for ln in lines) / self.grid.zoom

    def autofit_cols(self, cols):
        sh = self.sheet
        new = dict(sh.col_widths)
        by_col = {}
        keys = list(sh.values) + list(sh.formulas)
        wanted = set(cols)
        n = 0
        for k in keys:
            c = k & 0x3FFF
            if c not in wanted:
                continue
            r = k >> 14
            if r in sh.filter_hidden or r in sh.hidden_rows:
                continue
            if sh.merges and sh.merge_at(r, c):
                continue
            st = sh.styles.get(k, DEFAULT_STYLE)
            if st.wrap:
                continue
            w = self._text_width(k, st) + 2 * 3 + 4
            if w > by_col.get(c, 0):
                by_col[c] = w
            n += 1
            if n > 300000:
                break
        for c in cols:
            if c in by_col:
                new[c] = int(min(1200, max(12, math.ceil(by_col[c]))))
            else:
                new.pop(c, None)
        self._push_meta({"col_widths": (dict(sh.col_widths), new)}, "AutoFit Column Width")

    def autofit_rows(self, rows):
        sh = self.sheet
        new = dict(sh.row_heights)
        wanted = set(rows)
        heights = {}
        for k in list(sh.values) + list(sh.formulas):
            r = k >> 14
            if r not in wanted:
                continue
            c = k & 0x3FFF
            st = sh.styles.get(k, DEFAULT_STYLE)
            font, fm = self.grid.font_for(st)
            text, _, _ = self.grid.display(k, st)
            if st.wrap:
                w = int(sh.col_width(c) * self.grid.zoom) - 6
                br = fm.boundingRect(0, 0, max(10, w), 100000, Qt.TextWordWrap, text)
                h = br.height() / self.grid.zoom + 6
            else:
                h = fm.height() / self.grid.zoom + 5
            heights[r] = max(heights.get(r, 0), h)
        for r in rows:
            h = heights.get(r)
            if h is None or abs(h - 20) <= 1 or h < 20:
                new.pop(r, None)
            else:
                new[r] = int(math.ceil(h))
        self._push_meta({"row_heights": (dict(sh.row_heights), new)}, "AutoFit Row Height")

    def autofit_selected_cols(self):
        self.autofit_cols(self._sel_cols())

    def autofit_selected_rows(self):
        self.autofit_rows(self._sel_rows())

    # ================================================================ sheets
    def _tab_changed(self, idx):
        if 0 <= idx < len(self.wb.sheets):
            self.show_sheet(self.wb.sheets[idx])
            self.grid.setFocus()

    def _tab_moved(self, frm, to):
        sh = self.wb.sheets[frm]

        def action():
            self.wb.move_sheet(sh, to)
            self.wb.active = to
        self.undo.push(SnapshotCommand(self, sh, "Move Sheet", action, sheets=[]))

    def cycle_sheet(self, d):
        self._prep()
        i = (self.wb.active + d) % len(self.wb.sheets)
        self.show_sheet(self.wb.sheets[i])

    def add_sheet(self):
        self._prep()
        idx = self.wb.active + 1
        holder = {}

        def action():
            sh = self.wb.add_sheet(None, idx)
            holder["sheet"] = sh
            self.wb.active = idx
        cmd = SnapshotCommand(self, self.sheet, "Insert Sheet", action, sheets=[])
        self.undo.push(cmd)
        self.show_sheet(self.wb.sheets[idx])

    def delete_sheet(self):
        self._prep()
        if len(self.wb.sheets) == 1:
            QMessageBox.information(self, APP_NAME, "A workbook must contain at least one sheet.")
            return
        sh = self.sheet
        if (sh.values or sh.formulas) and QMessageBox.question(
                self, APP_NAME, f"Delete sheet '{sh.name}'? (You can undo this with Ctrl+Z.)") != QMessageBox.Yes:
            return

        def action():
            self.wb.remove_sheet(sh)
        self.undo.push(SnapshotCommand(self, sh, "Delete Sheet", action, sheets=[sh]))

    def rename_sheet(self):
        self._prep()
        sh = self.sheet
        name, ok = QInputDialog.getText(self, "Rename Sheet", "Sheet name:", text=sh.name)
        name = name.strip()
        if not ok or not name or name == sh.name:
            return
        bad = set('[]:*?/\\')
        if any(ch in bad for ch in name) or len(name) > 31 or name.startswith("'"):
            QMessageBox.warning(self, APP_NAME, "Sheet names can't exceed 31 characters or contain [ ] : * ? / \\")
            return
        other = self.wb.get_sheet(name)
        if other is not None and other is not sh:
            QMessageBox.warning(self, APP_NAME, f"There's already a sheet named '{name}'.")
            return
        self.undo.push(SnapshotCommand(self, sh, "Rename Sheet", lambda: self.wb.rename_sheet(sh, name), sheets=[]))

    def duplicate_sheet(self):
        self._prep()
        src = self.sheet
        base = src.name[:27]
        n = 2
        while self.wb.get_sheet(f"{base} ({n})"):
            n += 1
        name = f"{base} ({n})"
        idx = self.wb.active + 1

        def action():
            sh = Sheet(self.wb, name)
            self.wb.sheets.insert(idx, sh)
            sh.values = dict(src.values)
            sh.styles = dict(src.styles)
            for k, f in src.formulas.items():
                sh.set_content_key(k, ("f", f.text), f.fallback)
            sh.col_widths = dict(src.col_widths)
            sh.row_heights = dict(src.row_heights)
            sh.hidden_rows = set(src.hidden_rows)
            sh.hidden_cols = set(src.hidden_cols)
            sh.merges = list(src.merges)
            sh.freeze = src.freeze
            sh.freeze_origin = src.freeze_origin
            sh.show_grid = src.show_grid
            sh.zoom = src.zoom
            sh.recompute_extent()
            self.wb.rebuild_dependencies()
            self.wb.recalc(full=True)
            self.wb.active = idx
        self.undo.push(SnapshotCommand(self, src, "Duplicate Sheet", action, sheets=[]))
        self.show_sheet(self.wb.sheets[idx])

    def set_tab_color(self, color):
        sh = self.sheet
        old = sh.tab_color

        def action():
            sh.tab_color = color
        self.undo.push(SnapshotCommand(self, sh, "Tab Color", action, sheets=[]))
        _ = old

    def _tab_menu(self, pos):
        i = self.tabs.tabAt(pos)
        if i >= 0:
            self.show_sheet(self.wb.sheets[i])
        m = QMenu(self)
        m.addAction(self.a_new_sheet)
        m.addAction(self.a_del_sheet)
        m.addAction(self.a_rename_sheet)
        m.addAction(self.a_dup_sheet)
        m.addSeparator()
        m.addAction("Move Left", lambda: self._move_sheet_by(-1))
        m.addAction("Move Right", lambda: self._move_sheet_by(1))
        cm = ColorMenu(m, "No Color")
        cm.setTitle("Tab Color")
        cm.color_chosen.connect(self.set_tab_color)
        m.addMenu(cm)
        m.exec(self.tabs.mapToGlobal(pos))

    def _move_sheet_by(self, d):
        i = self.wb.active
        j = i + d
        if 0 <= j < len(self.wb.sheets):
            sh = self.sheet

            def action():
                self.wb.move_sheet(sh, j)
                self.wb.active = j
            self.undo.push(SnapshotCommand(self, sh, "Move Sheet", action, sheets=[]))

    # ================================================================ view
    def freeze_panes(self):
        self._prep()
        sh = self.sheet
        if sh.freeze != (0, 0):
            return self.set_freeze((0, 0))
        r, c = self.grid.sel.active
        # like Excel: freeze what's on screen above/left of the active cell
        self.set_freeze((r, c), (min(self.grid.top, r), min(self.grid.left, c)))

    def freeze_top_row(self):
        t = self.grid.top
        self.set_freeze((t + 1, 0), (t, 0))

    def freeze_first_col(self):
        c = self.grid.left
        self.set_freeze((0, c + 1), (0, c))

    def set_freeze(self, fz, origin=(0, 0)):
        self._prep()
        sh = self.sheet
        fz = tuple(fz)
        origin = (origin[0] if fz[0] else 0, origin[1] if fz[1] else 0)
        self._push_meta({"freeze": (sh.freeze, fz), "freeze_origin": (sh.freeze_origin, origin)}, "Freeze Panes")
        self.grid.top = max(self.grid.top, fz[0])
        self.grid.left = max(self.grid.left, fz[1])
        self.grid.relayout()
        self._sync_view_actions()

    def toggle_gridlines(self):
        sh = self.sheet
        self._push_meta({"show_grid": (sh.show_grid, not sh.show_grid)}, "Gridlines", relayout=False)
        self._sync_view_actions()

    def toggle_show_formulas(self):
        self.grid.show_formulas = not self.grid.show_formulas
        self.grid.invalidate()
        self._sync_view_actions()

    # ================================================================ formulas helpers
    def autosum(self, func="SUM"):
        self._prep()
        sh = self.sheet
        rect = self.grid.sel.rects[-1]
        if rect[0] != rect[2] or rect[1] != rect[3]:
            # multi-cell: put totals below each column (or right of a single row)
            r1, c1, r2, c2 = self.grid.clamp_rect(rect)
            states = {}
            if r1 == r2:
                tgt = (r1, c2 + 1)
                states[key(*tgt)] = ops.input_state(sh, *tgt, f"={func}({addr(r1, c1)}:{addr(r1, c2)})")
            else:
                for c in range(c1, c2 + 1):
                    states[key(r2 + 1, c)] = ops.input_state(sh, r2 + 1, c, f"={func}({addr(r1, c)}:{addr(r2, c)})")
            self._push_states(states, "AutoSum")
            return
        r, c = self.grid.sel.active
        rng = ops.autosum_range(sh, r, c)
        text = f"={func}({range_addr(*rng)})" if rng else f"={func}()"
        self.grid.begin_edit(text=text, mode="edit")
        if not rng:
            ed = self.grid.editor
            cur = ed.textCursor()
            cur.setPosition(len(text) - 1)
            ed.setTextCursor(cur)

    def insert_function_dialog(self):
        dlg = QDialog(self)
        dlg.setWindowTitle("Insert Function")
        dlg.resize(460, 420)
        lay = QVBoxLayout(dlg)
        search = QLineEdit()
        search.setPlaceholderText("Search for a function")
        lay.addWidget(search)
        lst = QListWidget()
        lst.addItems(ALL_NAMES)
        lay.addWidget(lst, 1)
        sig = QLabel("")
        sig.setWordWrap(True)
        sig.setStyleSheet("padding: 6px; background: #FAFAFA; border: 1px solid #E0E0E0;")
        lay.addWidget(sig)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        lay.addWidget(bb)
        bb.accepted.connect(dlg.accept)
        bb.rejected.connect(dlg.reject)
        lst.itemDoubleClicked.connect(lambda _: dlg.accept())

        def refilter(t):
            t = t.upper()
            for i in range(lst.count()):
                it = lst.item(i)
                it.setHidden(t not in it.text())
        search.textChanged.connect(refilter)
        lst.currentTextChanged.connect(lambda n: sig.setText(SIGNATURES.get(n, n + "()")))
        lst.setCurrentRow(0)
        if dlg.exec() != QDialog.Accepted or lst.currentItem() is None:
            return
        name = lst.currentItem().text()
        if self.grid.editing:
            w = self.grid.edit_widget or self.grid.editor
            t = w.text()
            ins = name + "("
            if not t.startswith("="):
                ins = "=" + ins
            w.insertPlainText(ins)
            w.setFocus()
        else:
            self.grid.begin_edit(text=f"={name}(", mode="edit")

    # ================================================================ sort & filter
    def _sort_range(self):
        rect = self.grid.sel.rects[-1]
        r, c = self.grid.sel.active
        sh = self.sheet
        if rect[0] == rect[2] and rect[1] == rect[3]:
            if sh.autofilter and sh.autofilter[0] <= r <= ops.filter_extent(sh)[2] and \
                    sh.autofilter[1] <= c <= sh.autofilter[3]:
                return ops.filter_extent(sh), True
            region = ops.current_region(sh, r, c)
            return region, ops.guess_header(sh, region)
        rect = self.grid.clamp_rect(rect)
        return rect, ops.guess_header(self.sheet, rect)

    def quick_sort(self, asc, col=None):
        self._prep()
        rect, header = self._sort_range()
        c = col if col is not None else self.grid.sel.active[1]
        if not rect[1] <= c <= rect[3]:
            c = rect[1]
        self._do_sort(rect, [(c, asc)], header)

    def _do_sort(self, rect, keys, header, case=False):
        sh = self.sheet
        states = ops.sort_states(sh, rect, keys, header, case)
        if not states:
            self.statusBar().showMessage("Already sorted", 2000)
            return
        self.undo.beginMacro("Sort")
        self._push_states(states, "Sort")
        if sh.filters:
            hidden = ops.compute_filter_hidden(sh)
            self._push_meta({"filter_hidden": (set(sh.filter_hidden), hidden)}, "Sort")
        self.undo.endMacro()

    def sort_dialog(self):
        self._prep()
        rect, header = self._sort_range()
        dlg = SortDialog(self, self.sheet, rect, header)
        if dlg.exec() == QDialog.Accepted:
            self._do_sort(rect, dlg.keys(), dlg.header_chk.isChecked(), dlg.case_chk.isChecked())

    def toggle_filter(self):
        self._prep()
        sh = self.sheet
        if sh.autofilter:
            self._push_meta({"autofilter": (sh.autofilter, None), "filters": (dict(sh.filters), {}),
                             "filter_hidden": (set(sh.filter_hidden), set())}, "Remove Filter")
        else:
            rect = self.grid.sel.rects[-1]
            r, c = self.grid.sel.active
            if rect[0] == rect[2] and rect[1] == rect[3]:
                rect = ops.current_region(sh, r, c)
            rect = self.grid.clamp_rect(rect)
            if rect[0] == rect[2]:
                QMessageBox.information(self, APP_NAME, "Select a cell inside your data (with a header row) first.")
                self.a_filter.setChecked(False)
                return
            self._push_meta({"autofilter": (None, rect)}, "Filter")
        self._sync_view_actions()

    def show_filter_popup(self, col, pos):
        sh = self.sheet
        if not sh.autofilter:
            return
        r1, c1, r2, c2 = ops.filter_extent(sh)
        # values offered: rows not hidden by *other* columns' filters
        other = {k: v for k, v in sh.filters.items() if k != col}
        rows = [r for r in range(r1 + 1, r2 + 1)
                if not any(not ops.row_passes(sh, r, spec, oc) for oc, spec in other.items())]
        seen = {}
        blanks = False
        for r in rows:
            t = ops.display_text(sh, r, col)
            if t == "":
                blanks = True
            elif t not in seen:
                seen[t] = sh.value(r, col)
        from ..values import sort_key
        values = [t for t, _ in sorted(seen.items(), key=lambda kv: sort_key(kv[1]) if kv[1] is not None else (9,))]
        title = str(sh.value(r1, col) or col_name(col))
        pop = FilterPopup(self, col, values, blanks, sh.filters.get(col), title)
        pop.applied.connect(self.apply_filter)
        pop.sort_requested.connect(lambda c, asc: self._do_sort(ops.filter_extent(sh), [(c, asc)], True))
        pop.adjustSize()
        screen = QApplication.screenAt(pos)
        if screen:
            g = screen.availableGeometry()
            x = min(pos.x(), g.right() - pop.width())
            y = pos.y() if pos.y() + pop.height() < g.bottom() else max(g.top(), g.bottom() - pop.height())
            pos = QPoint(x, y)
        pop.move(pos)
        pop.show()

    def apply_filter(self, col, spec):
        sh = self.sheet
        filters = dict(sh.filters)
        if spec is None:
            filters.pop(col, None)
        else:
            filters[col] = spec
        old = (sh.autofilter, dict(sh.filters), set(sh.filter_hidden))
        sh.filters = filters
        hidden = ops.compute_filter_hidden(sh)
        new_af = sh.autofilter
        sh.autofilter, sh.filters, sh.filter_hidden = old[0], old[1], old[2]
        self._push_meta({"autofilter": (old[0], new_af), "filters": (old[1], filters),
                         "filter_hidden": (old[2], hidden)}, "Filter")
        af = sh.autofilter
        r, c = self.grid.sel.active
        if r in hidden:
            self.grid.set_active(af[0], c)

    def clear_filters(self):
        sh = self.sheet
        if sh.filters:
            self._push_meta({"filters": (dict(sh.filters), {}), "filter_hidden": (set(sh.filter_hidden), set())},
                            "Clear Filters")

    def reapply_filters(self):
        sh = self.sheet
        if sh.autofilter and sh.filters:
            hidden = ops.compute_filter_hidden(sh)
            self._push_meta({"filter_hidden": (set(sh.filter_hidden), hidden)}, "Reapply Filter")

    def filter_by_value(self):
        sh = self.sheet
        r, c = self.grid.sel.active
        if not sh.autofilter:
            region = ops.current_region(sh, r, c)
            if region[0] == region[2]:
                return
            self._push_meta({"autofilter": (None, region)}, "Filter")
        t = ops.display_text(sh, r, c)
        spec = {"values": {t}, "blanks": False} if t else {"values": set(), "blanks": True}
        self.apply_filter(c, spec)

    # ================================================================ find & replace
    def show_find(self, replace):
        self._prep()
        if self.find_dlg is None:
            self.find_dlg = FindDialog(self, replace)
            self.find_dlg.find_next.connect(self.find_next)
            self.find_dlg.find_all.connect(self.find_all)
            self.find_dlg.replace_one.connect(self.replace_one)
            self.find_dlg.replace_all.connect(self.replace_all)
            self.find_dlg.goto.connect(lambda s, r, c: self.goto_ref(f"'{s}'!{addr(r, c)}"))
        self.find_dlg.set_replace_mode(replace)
        r, c = self.grid.sel.active
        if not self.find_dlg.find_edit.text():
            v = self.sheet.value(r, c)
            if isinstance(v, str) and len(v) < 60:
                self.find_dlg.find_edit.setText(v)
        self.find_dlg.show()
        self.find_dlg.raise_()
        self.find_dlg.activateWindow()
        self.find_dlg.find_edit.setFocus()
        self.find_dlg.find_edit.selectAll()

    def _matches(self, p, sheet):
        return ops.find_matches(sheet, p["needle"], p["case"], p["whole"], p["formulas"])

    def find_next(self, p):
        if not p["needle"]:
            return
        sheets = self.wb.sheets if p["workbook"] else [self.sheet]
        start_idx = sheets.index(self.sheet) if self.sheet in sheets else 0
        order = sheets[start_idx:] + sheets[:start_idx] + [sheets[start_idx]]
        cur = self.grid.sel.active
        for n, sh in enumerate(order):
            ms = self._matches(p, sh)
            if n == 0:
                ms = [m for m in ms if m > cur]
            elif n == len(order) - 1:
                ms = [m for m in ms if m <= cur]
            if ms:
                self.show_sheet(sh)
                r, c = ms[0]
                self.grid.set_active(r, c)
                self.find_dlg.status.setText(f"Found at {sh.name}!{addr(r, c)}")
                return True
        self.find_dlg.status.setText("Sheets couldn't find what you were looking for.")
        return False

    def find_all(self, p):
        if not p["needle"]:
            return
        items = []
        for sh in (self.wb.sheets if p["workbook"] else [self.sheet]):
            for r, c in self._matches(p, sh):
                items.append((sh.name, r, c, ops.edit_text_for(sh, r, c)[:80]))
        self.find_dlg.show_results(items)

    def replace_one(self, p):
        sh = self.sheet
        r, c = self.grid.sel.active
        if (r, c) in self._matches(dict(p, formulas=True), sh):
            states = ops.replace_states(sh, [(r, c)], p["needle"], p["repl"], p["case"], p["whole"])
            self._push_states(states, "Replace")
        self.find_next(p)

    def replace_all(self, p):
        if not p["needle"]:
            return
        total = 0
        self.undo.beginMacro("Replace All")
        for sh in (self.wb.sheets if p["workbook"] else [self.sheet]):
            ms = ops.find_matches(sh, p["needle"], p["case"], p["whole"], True)
            states = ops.replace_states(sh, ms, p["needle"], p["repl"], p["case"], p["whole"])
            if states:
                total += len(states)
                self._push_states(states, "Replace", sheet=sh)
        self.undo.endMacro()
        self.find_dlg.status.setText(f"All done. Made {total} replacement(s).")

    def goto_dialog(self):
        self._prep()
        text, ok = QInputDialog.getText(self, "Go To", "Reference (e.g. B12, A1:D50, Sheet2!C3):")
        if ok and text.strip():
            self.goto_ref(text.strip())

    # ================================================================ context menu
    def _context_menu(self, area, pos):
        m = QMenu(self)
        for a in (self.a_cut, self.a_copy, self.a_paste):
            m.addAction(a)
        ps = m.addMenu("Paste Special")
        for a in (self.a_paste_values, self.a_paste_formats, self.a_paste_transpose):
            ps.addAction(a)
        m.addSeparator()
        if area == "row":
            m.addAction("Insert", lambda: self.insert_rows_cols("row"))
            m.addAction("Delete", lambda: self.delete_rows_cols("row"))
            m.addAction(self.a_clear_contents)
            m.addSeparator()
            m.addAction(self.a_format_cells)
            m.addAction(self.a_row_height)
            m.addAction(self.a_autofit_rows)
            m.addAction(self.a_hide_rows)
            m.addAction(self.a_unhide_rows)
        elif area == "col":
            m.addAction("Insert", lambda: self.insert_rows_cols("col"))
            m.addAction("Delete", lambda: self.delete_rows_cols("col"))
            m.addAction(self.a_clear_contents)
            m.addSeparator()
            m.addAction(self.a_format_cells)
            m.addAction(self.a_col_width)
            m.addAction(self.a_autofit_cols)
            m.addAction(self.a_hide_cols)
            m.addAction(self.a_unhide_cols)
        else:
            im = m.addMenu("Insert")
            im.addAction("Entire Row", lambda: self.insert_rows_cols("row"))
            im.addAction("Entire Column", lambda: self.insert_rows_cols("col"))
            dm = m.addMenu("Delete")
            dm.addAction("Entire Row", lambda: self.delete_rows_cols("row"))
            dm.addAction("Entire Column", lambda: self.delete_rows_cols("col"))
            m.addAction(self.a_clear_contents)
            m.addSeparator()
            fm = m.addMenu("Filter")
            fm.addAction(self.a_filter_value)
            fm.addAction(self.a_clear_filter)
            sm = m.addMenu("Sort")
            sm.addAction(self.a_sort_az)
            sm.addAction(self.a_sort_za)
            sm.addAction(self.a_sort)
            m.addSeparator()
            m.addAction(self.a_format_cells)
            m.addAction(self.a_merge)
        m.exec(pos)

    # ================================================================ files
    def _fill_recent(self):
        self.recent_menu.clear()
        recent = [p for p in self.settings.get("recent", []) if os.path.exists(p)]
        if not recent:
            a = self.recent_menu.addAction("(none)")
            a.setEnabled(False)
            return
        for i, p in enumerate(recent[:12]):
            self.recent_menu.addAction(f"&{i + 1}  {p}", lambda path=p: self.open_path(path))

    def _remember(self, path):
        s = load_settings()
        rec = [p for p in s.get("recent", []) if os.path.normcase(p) != os.path.normcase(path)]
        rec.insert(0, path)
        s["recent"] = rec[:15]
        s["last_dir"] = os.path.dirname(path)
        save_settings(s)
        self.settings = s

    def default_dir(self):
        """Where Open / Save As start: the folder set with File > Set Default Folder,
        else DEFAULT_FOLDER (created on first use), else Documents."""
        d = load_settings().get("default_dir") or DEFAULT_FOLDER  # fresh: may be changed in another window
        try:
            os.makedirs(d, exist_ok=True)
            return d
        except OSError:
            from PySide6.QtCore import QStandardPaths
            return QStandardPaths.writableLocation(QStandardPaths.DocumentsLocation) or os.path.expanduser("~")

    def set_default_folder(self):
        d = QFileDialog.getExistingDirectory(self, "Default folder for Open and Save As", self.default_dir())
        if not d:
            return
        s = load_settings()
        s["default_dir"] = os.path.normpath(d)
        save_settings(s)
        self.settings = s
        self.statusBar().showMessage(f"Open and Save As will now start in {os.path.normpath(d)}", 5000)

    def file_new(self):
        w = MainWindow(new_workbook())
        w.show()

    def file_open(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "Open", self.default_dir(), OPEN_FILTER)
        for p in paths:
            self.open_path(p)

    def open_path(self, path):
        path = os.path.abspath(path)
        for w in WINDOWS:
            if w.wb.path and os.path.normcase(w.wb.path) == os.path.normcase(path):
                w.raise_()
                w.activateWindow()
                return w
        target = self if self.is_pristine() else None
        wb = load_with_progress(self, path)
        if wb is None:
            return None
        self._remember(path)
        if target is not None:
            target.load_workbook(wb)
            return target
        w = MainWindow(wb)
        w.show()
        return w

    def load_workbook(self, wb):
        self.wb = wb
        self.undo.clear()
        self.grid._view_states = {}
        self.grid.sheet = None
        self.clip = None
        self._csv_warned = False
        self._lost_ack = False
        self.rebuild_tabs()
        self.show_sheet(wb.sheets[wb.active])
        self.update_title()

    def file_save(self):
        self._prep()
        p = self.wb.path
        if not p or os.path.splitext(p)[1].lower() not in WRITABLE:
            return self.file_save_as()
        return self.do_save(p)

    def file_save_as(self):
        self._prep()
        p = self.wb.path
        start_dir = os.path.dirname(p) if p else self.default_dir()
        base = os.path.splitext(os.path.basename(p))[0] if p else "Book1"
        ext = os.path.splitext(p)[1].lower() if p else ".xlsx"
        if ext not in WRITABLE:
            ext = ".xlsx"
        filters = ";;".join(SAVE_FILTERS.values())
        chosen, flt = QFileDialog.getSaveFileName(self, "Save As", os.path.join(start_dir, base + ext),
                                                  filters, SAVE_FILTERS[ext])
        if not chosen:
            return False
        cext = os.path.splitext(chosen)[1].lower()
        if cext not in WRITABLE:
            for e, f in SAVE_FILTERS.items():
                if f == flt:
                    chosen += e
                    break
        return self.do_save(chosen, save_as=True)

    def do_save(self, path, save_as=False):
        ext = os.path.splitext(path)[1].lower()
        wb = self.wb
        if ext == ".xlsm" and getattr(wb.xl_book, "vba_archive", None) is None:
            # a macro-enabled file needs macros; without them Excel refuses to open it
            path = os.path.splitext(path)[0] + ".xlsx"
            ext = ".xlsx"
            self.statusBar().showMessage("No macros to keep, so the file is saved as .xlsx", 5000)
        if ext in (".csv", ".tsv", ".txt"):
            if len(wb.sheets) > 1 and not self._csv_warned:
                r = QMessageBox.question(
                    self, APP_NAME,
                    f"CSV files hold just one sheet, so only '{self.sheet.name}' will be saved.\n\n"
                    "Formatting, formulas and the other sheets aren't stored in CSV. "
                    "Use Save As → Excel workbook (.xlsx) to keep everything.\n\nSave the current sheet as CSV?")
                if r != QMessageBox.Yes:
                    return False
                self._csv_warned = True
        elif wb.xl_lost_features and not self._lost_ack and wb.path and \
                os.path.normcase(path) == os.path.normcase(wb.path):
            box = QMessageBox(self)
            box.setWindowTitle(APP_NAME)
            box.setIcon(QMessageBox.Warning)
            box.setText(f"This workbook contains {', '.join(wb.xl_lost_features)} that Sheets can't keep.\n\n"
                        "Saving over the original file will remove them. You can Save As a new file instead "
                        "to keep the original intact.")
            save_btn = box.addButton("Save Anyway", QMessageBox.AcceptRole)
            as_btn = box.addButton("Save As...", QMessageBox.ActionRole)
            box.addButton(QMessageBox.Cancel)
            box.exec()
            if box.clickedButton() is as_btn:
                return self.file_save_as()
            if box.clickedButton() is not save_btn:
                return False
            self._lost_ack = True
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            save_file(wb, path, self.sheet)
        except PermissionError:
            QApplication.restoreOverrideCursor()
            QMessageBox.critical(self, APP_NAME, f"Couldn't save '{os.path.basename(path)}'.\n\n"
                                 "It may be open in another program (like Excel), or read-only.")
            return False
        except Exception as e:
            QApplication.restoreOverrideCursor()
            QMessageBox.critical(self, APP_NAME, f"Couldn't save the file:\n\n{e}")
            return False
        QApplication.restoreOverrideCursor()
        self.undo.setClean()
        self._remember(path)
        self.update_title()
        self.statusBar().showMessage(f"Saved {path}", 4000)
        return True

    def closeEvent(self, e):
        self._prep()
        if self.claude_panel.busy():
            if QMessageBox.question(self, APP_NAME, "Claude is still working on this workbook. Stop it and close?") \
                    != QMessageBox.Yes:
                e.ignore()
                return
            self.claude_panel.stop(finish_now=True)
        if not self.undo.isClean():
            name = os.path.basename(self.wb.path) if self.wb.path else "Book1"
            r = QMessageBox.question(self, APP_NAME, f"Save changes to '{name}'?",
                                     QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel)
            if r == QMessageBox.Cancel:
                e.ignore()
                return
            if r == QMessageBox.Save and not self.file_save():
                e.ignore()
                return
        s = load_settings()
        s["geometry"] = bytes(self.saveGeometry()).hex()
        save_settings(s)
        if self.find_dlg:
            self.find_dlg.close()
        if self in WINDOWS:
            WINDOWS.remove(self)
        e.accept()

    def quit_all(self):
        for w in list(WINDOWS):
            if not w.close():
                return

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()

    def dropEvent(self, e):
        for url in e.mimeData().urls():
            p = url.toLocalFile()
            if p and os.path.splitext(p)[1].lower() in READABLE:
                self.open_path(p)

    # ================================================================ misc
    def register_file_types(self):
        from ..register import describe, register
        r = QMessageBox.question(
            self, APP_NAME,
            describe() + "\n\nWindows will then list Sheets under 'Open with' for these files, and you can pick it "
            "as the default app (Windows asks you to confirm that part yourself).\n\nContinue?")
        if r != QMessageBox.Yes:
            return
        try:
            register()
        except Exception as e:
            QMessageBox.critical(self, APP_NAME, f"Couldn't register file types:\n\n{e}")
            return
        QMessageBox.information(
            self, APP_NAME,
            "Done. To make Sheets the default: right-click a .csv or .xlsx file → Open with → Choose another app → "
            "Sheets → tick 'Always'.")

    def show_shortcuts(self):
        text = """
<table cellpadding=3>
<tr><td><b>Ctrl+N / O / S</b></td><td>New, Open, Save (F12 Save As)</td></tr>
<tr><td><b>Ctrl+Z / Ctrl+Y</b></td><td>Undo / Redo</td></tr>
<tr><td><b>Ctrl+X / C / V</b></td><td>Cut, Copy, Paste &nbsp; (Ctrl+Shift+V paste values)</td></tr>
<tr><td><b>F2</b></td><td>Edit cell &nbsp; (Alt+Enter new line, Ctrl+Enter fill selection)</td></tr>
<tr><td><b>F4</b></td><td>Toggle $ absolute reference while editing</td></tr>
<tr><td><b>Ctrl+Arrows</b></td><td>Jump to edge of data (add Shift to select)</td></tr>
<tr><td><b>Ctrl+Home / End</b></td><td>First / last used cell</td></tr>
<tr><td><b>Ctrl+A</b></td><td>Select the current table, again for whole sheet</td></tr>
<tr><td><b>Ctrl+Space / Shift+Space</b></td><td>Select column / row</td></tr>
<tr><td><b>Ctrl+D / Ctrl+R</b></td><td>Fill down / right</td></tr>
<tr><td><b>Ctrl++ / Ctrl+-</b></td><td>Insert / delete rows (or columns when columns are selected)</td></tr>
<tr><td><b>Ctrl+9 / Ctrl+0</b></td><td>Hide rows / columns (add Shift to unhide)</td></tr>
<tr><td><b>Ctrl+B / I / U / 5</b></td><td>Bold, Italic, Underline, Strikethrough</td></tr>
<tr><td><b>Ctrl+1</b></td><td>Format Cells</td></tr>
<tr><td><b>Ctrl+Shift+$ % ! # @ ~</b></td><td>Currency, Percent, Number, Date, Time, General</td></tr>
<tr><td><b>Alt+=</b></td><td>AutoSum</td></tr>
<tr><td><b>Ctrl+; / Ctrl+Shift+;</b></td><td>Insert date / time</td></tr>
<tr><td><b>Ctrl+F / Ctrl+H / Ctrl+G</b></td><td>Find, Replace, Go To</td></tr>
<tr><td><b>Ctrl+Shift+L</b></td><td>Turn filter on/off</td></tr>
<tr><td><b>Ctrl+PgUp / PgDn</b></td><td>Previous / next sheet</td></tr>
<tr><td><b>Ctrl+`</b></td><td>Show formulas</td></tr>
<tr><td><b>Ctrl+Mouse wheel</b></td><td>Zoom</td></tr>
<tr><td><b>F9</b></td><td>Recalculate</td></tr>
</table>"""
        box = QMessageBox(self)
        box.setWindowTitle("Keyboard Shortcuts")
        box.setTextFormat(Qt.RichText)
        box.setText(text)
        box.exec()

    def about(self):
        QMessageBox.about(self, "About Sheets",
                          "<b>Sheets</b><br>A lightweight spreadsheet for CSV and Excel files.<br><br>"
                          "Opens .xlsx, .xlsm, .xls, .csv and .tsv; saves .xlsx and .csv.")


def load_with_progress(parent, path):
    dlg = QProgressDialog(f"Opening {os.path.basename(path)}...", None, 0, 100, parent)
    dlg.setWindowTitle(APP_NAME)
    dlg.setMinimumDuration(400)
    dlg.setWindowModality(Qt.WindowModal)
    dlg.setValue(0)
    QApplication.setOverrideCursor(Qt.WaitCursor)

    def progress(f):
        dlg.setValue(int(f * 100))
        QApplication.processEvents()
    try:
        return open_file(path, progress)
    except PermissionError:
        QApplication.restoreOverrideCursor()
        QMessageBox.critical(parent, APP_NAME, f"Couldn't open '{os.path.basename(path)}'.\n\n"
                             "It may be open in another program or you don't have permission.")
        return None
    except Exception as e:
        QApplication.restoreOverrideCursor()
        QMessageBox.critical(parent, APP_NAME, f"Couldn't open '{os.path.basename(path)}':\n\n{e}")
        return None
    finally:
        if QApplication.overrideCursor() is not None:
            QApplication.restoreOverrideCursor()
        dlg.close()
