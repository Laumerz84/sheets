"""Dialogs: Format Cells, Sort, Find & Replace, filter dropdown, color pickers."""
import colorsys

from PySide6.QtCore import QSize, Qt, Signal
from PySide6.QtGui import QColor, QFont
from PySide6.QtWidgets import (QAbstractItemView, QButtonGroup, QCheckBox,
                               QColorDialog, QComboBox, QDialog,
                               QDialogButtonBox, QFontComboBox, QFrame,
                               QGridLayout, QGroupBox, QHBoxLayout, QLabel,
                               QLineEdit, QListWidget, QListWidgetItem, QMenu,
                               QPushButton, QRadioButton, QSpinBox,
                               QTabWidget, QToolButton, QVBoxLayout, QWidget,
                               QWidgetAction, QTableWidget, QTableWidgetItem,
                               QHeaderView)

from ..numfmt import format_value
from ..refs import col_name
from .style import ACCENT, color_square_icon

# ================================================================ colors

THEME_BASE = ["#FFFFFF", "#000000", "#E7E6E6", "#44546A", "#4472C4", "#ED7D31",
              "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47"]
STANDARD = ["#C00000", "#FF0000", "#FFC000", "#FFFF00", "#92D050", "#00B050",
            "#00B0F0", "#0070C0", "#002060", "#7030A0"]


def _tint(hexc, t):
    r, g, b = (int(hexc[i:i + 2], 16) / 255 for i in (1, 3, 5))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    l = l * (1 + t) if t < 0 else l * (1 - t) + t
    r, g, b = colorsys.hls_to_rgb(h, max(0, min(1, l)), s)
    return "#%02X%02X%02X" % (round(r * 255), round(g * 255), round(b * 255))


def theme_grid():
    rows = [THEME_BASE]
    for t in (0.8, 0.6, 0.4, -0.25, -0.5):
        row = []
        for i, c in enumerate(THEME_BASE):
            if i == 0:
                row.append(_tint(c, {0.8: -0.05, 0.6: -0.15, 0.4: -0.25, -0.25: -0.35, -0.5: -0.5}[t]))
            elif i == 1:
                row.append(_tint(c, {0.8: 0.5, 0.6: 0.35, 0.4: 0.25, -0.25: 0.15, -0.5: 0.05}[t]))
            else:
                row.append(_tint(c, t))
        rows.append(row)
    return rows


class ColorMenu(QMenu):
    color_chosen = Signal(object)  # '#RRGGBB' or None

    def __init__(self, parent=None, none_label="No Color"):
        super().__init__(parent)
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(8, 6, 8, 6)
        lay.setSpacing(4)
        none_btn = QPushButton(none_label)
        none_btn.setFlat(True)
        none_btn.setStyleSheet("text-align: left; padding: 3px 6px; border: none;")
        none_btn.clicked.connect(lambda: self._pick(None))
        lay.addWidget(none_btn)
        lay.addWidget(QLabel("Theme Colors"))
        lay.addLayout(self._grid(theme_grid()))
        lay.addWidget(QLabel("Standard Colors"))
        lay.addLayout(self._grid([STANDARD]))
        more = QPushButton("More Colors...")
        more.setFlat(True)
        more.setStyleSheet("text-align: left; padding: 3px 6px; border: none;")
        more.clicked.connect(self._more)
        lay.addWidget(more)
        act = QWidgetAction(self)
        act.setDefaultWidget(w)
        self.addAction(act)

    def _grid(self, rows):
        g = QGridLayout()
        g.setSpacing(2)
        for i, row in enumerate(rows):
            for j, c in enumerate(row):
                b = QToolButton()
                b.setFixedSize(QSize(16, 16))
                b.setToolTip(c)
                b.setStyleSheet(f"QToolButton {{ background: {c}; border: 1px solid #C8C8C8; border-radius: 0; padding: 0; }}"
                                f"QToolButton:hover {{ border: 1px solid #E07000; }}")
                b.clicked.connect(lambda _=False, col=c: self._pick(col))
                g.addWidget(b, i, j)
        return g

    def _pick(self, c):
        self.close()
        self.color_chosen.emit(c)

    def _more(self):
        self.close()
        col = QColorDialog.getColor(QColor("#FFFFFF"), self.parentWidget(), "Choose color")
        if col.isValid():
            self.color_chosen.emit(col.name().upper())


class ColorButton(QToolButton):
    """Button showing a color; click opens the palette."""
    changed = Signal(object)

    def __init__(self, color=None, none_label="Automatic", parent=None):
        super().__init__(parent)
        self.color = color
        self.menu_ = ColorMenu(self, none_label)
        self.menu_.color_chosen.connect(self._set)
        self.setMenu(self.menu_)
        self.setPopupMode(QToolButton.InstantPopup)
        self.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        self._refresh()

    def _set(self, c):
        self.color = c
        self._refresh()
        self.changed.emit(c)

    def set_color(self, c):
        self.color = c
        self._refresh()

    def _refresh(self):
        self.setIcon(color_square_icon(self.color, 16))
        self.setText(self.color or "Automatic")


# ================================================================ Format Cells

NUMBER_CATEGORIES = ["General", "Number", "Currency", "Accounting", "Date", "Time",
                     "Percentage", "Fraction", "Scientific", "Text", "Custom"]
DATE_FORMATS = ["m/d/yyyy", "dddd, mmmm d, yyyy", "m/d", "m/d/yy", "mm/dd/yy", "d-mmm",
                "d-mmm-yy", "dd-mmm-yy", "mmm-yy", "mmmm-yy", "mmmm d, yyyy", "yyyy-mm-dd",
                "m/d/yyyy h:mm", "yyyy-mm-dd hh:mm:ss", "d/m/yyyy", "dd/mm/yyyy"]
TIME_FORMATS = ["h:mm:ss AM/PM", "h:mm AM/PM", "h:mm", "h:mm:ss", "mm:ss", "[h]:mm:ss",
                "m/d/yyyy h:mm"]
CUSTOM_FORMATS = ["General", "0", "0.00", "#,##0", "#,##0.00", "#,##0_);(#,##0)",
                  "#,##0.00_);[Red](#,##0.00)", "$#,##0_);($#,##0)", "$#,##0.00",
                  "0%", "0.00%", "0.00E+00", "# ?/?", "m/d/yyyy", "d-mmm-yy", "mmm-yy",
                  "h:mm AM/PM", "h:mm:ss", "[h]:mm:ss", "mm:ss", "yyyy-mm-dd", "@",
                  '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)', "000-00-0000",
                  "(000) 000-0000", '0.0,,"M"', '0.0,"K"']
SYMBOLS = ["$", "€", "£", "¥", "None"]
BORDER_STYLES = [("thin", "Thin"), ("medium", "Medium"), ("thick", "Thick"),
                 ("dashed", "Dashed"), ("dotted", "Dotted"), ("double", "Double"),
                 ("hair", "Hair")]


def build_number_format(cat, decimals=2, thousands=True, symbol="$", negative=0, date=None, time=None, custom=None):
    d = ("." + "0" * decimals) if decimals > 0 else ""
    if cat == "General":
        return "General"
    if cat == "Number":
        base = ("#,##0" if thousands else "0") + d
        return [base, f"{base};[Red]{base}", f"{base}_);({base})", f"{base}_);[Red]({base})"][negative]
    if cat == "Currency":
        sym = "" if symbol == "None" else (symbol if symbol == "$" else f'"{symbol}"')
        base = f"{sym}#,##0{d}"
        return [base, f"{base};[Red]{base}", f"{base}_);({base})", f"{base}_);[Red]({base})"][negative]
    if cat == "Accounting":
        sym = "" if symbol == "None" else (symbol if symbol == "$" else f'"{symbol}"')
        dd = "?" * decimals
        z = f'"-"{dd}' if decimals else '"-"'
        return f'_({sym}* #,##0{d}_);_({sym}* (#,##0{d});_({sym}* {z}_);_(@_)'
    if cat == "Date":
        return date or "m/d/yyyy"
    if cat == "Time":
        return time or "h:mm:ss AM/PM"
    if cat == "Percentage":
        return "0" + d + "%"
    if cat == "Fraction":
        return "# ?/?"
    if cat == "Scientific":
        return "0" + d + "E+00"
    if cat == "Text":
        return "@"
    return custom or "General"


def category_of(fmt):
    if not fmt or fmt == "General":
        return "General"
    if fmt == "@":
        return "Text"
    if fmt in DATE_FORMATS:
        return "Date"
    if fmt in TIME_FORMATS:
        return "Time"
    if fmt.endswith("%") and set(fmt[:-1]) <= set("0."):
        return "Percentage"
    if "E+" in fmt:
        return "Scientific"
    if fmt.startswith("_(") and "*" in fmt:
        return "Accounting"
    return "Custom"


class FormatCellsDialog(QDialog):
    def __init__(self, parent, style, sample, tab=0):
        super().__init__(parent)
        self.setWindowTitle("Format Cells")
        self.setMinimumSize(560, 460)
        self.st = style
        self.sample = sample if sample is not None else 1234.5678
        self.changes = {}
        self.border_op = None
        lay = QVBoxLayout(self)
        self.tabs = QTabWidget()
        lay.addWidget(self.tabs)
        self.tabs.addTab(self._number_tab(), "Number")
        self.tabs.addTab(self._align_tab(), "Alignment")
        self.tabs.addTab(self._font_tab(), "Font")
        self.tabs.addTab(self._border_tab(), "Border")
        self.tabs.addTab(self._fill_tab(), "Fill")
        self.tabs.setCurrentIndex(tab)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)

    # ---------------------------------------------------------------- number
    def _number_tab(self):
        w = QWidget()
        h = QHBoxLayout(w)
        self.cat_list = QListWidget()
        self.cat_list.addItems(NUMBER_CATEGORIES)
        self.cat_list.setFixedWidth(130)
        h.addWidget(self.cat_list)
        right = QVBoxLayout()
        h.addLayout(right, 1)
        sample_box = QGroupBox("Sample")
        sl = QVBoxLayout(sample_box)
        self.sample_lbl = QLabel()
        self.sample_lbl.setMinimumHeight(22)
        sl.addWidget(self.sample_lbl)
        right.addWidget(sample_box)
        opts = QGridLayout()
        right.addLayout(opts)
        self.dec_spin = QSpinBox()
        self.dec_spin.setRange(0, 15)
        self.dec_spin.setValue(2)
        self.thousand_chk = QCheckBox("Use 1000 Separator (,)")
        self.sym_combo = QComboBox()
        self.sym_combo.addItems(SYMBOLS)
        self.neg_list = QListWidget()
        self.neg_list.setMaximumHeight(90)
        self.type_list = QListWidget()
        self.custom_edit = QLineEdit()
        self.custom_list = QListWidget()
        self.custom_list.addItems(CUSTOM_FORMATS)
        self.lbl_dec = QLabel("Decimal places:")
        self.lbl_sym = QLabel("Symbol:")
        self.lbl_neg = QLabel("Negative numbers:")
        self.lbl_type = QLabel("Type:")
        opts.addWidget(self.lbl_dec, 0, 0)
        opts.addWidget(self.dec_spin, 0, 1)
        opts.addWidget(self.lbl_sym, 1, 0)
        opts.addWidget(self.sym_combo, 1, 1)
        opts.addWidget(self.thousand_chk, 2, 0, 1, 2)
        opts.addWidget(self.lbl_neg, 3, 0, 1, 2)
        opts.addWidget(self.neg_list, 4, 0, 1, 2)
        opts.addWidget(self.lbl_type, 5, 0, 1, 2)
        opts.addWidget(self.custom_edit, 6, 0, 1, 2)
        opts.addWidget(self.type_list, 7, 0, 1, 2)
        opts.addWidget(self.custom_list, 8, 0, 1, 2)
        right.addStretch(1)

        fmt = self.st.numfmt
        cat = category_of(fmt)
        self.custom_edit.setText(fmt)
        if cat == "Percentage":
            self.dec_spin.setValue(len(fmt.split(".")[1]) - 1 if "." in fmt else 0)
        self.cat_list.setCurrentRow(NUMBER_CATEGORIES.index(cat))
        self.cat_list.currentRowChanged.connect(self._cat_changed)
        for wdg in (self.dec_spin,):
            wdg.valueChanged.connect(self._num_changed)
        self.thousand_chk.toggled.connect(self._num_changed)
        self.sym_combo.currentIndexChanged.connect(self._num_changed)
        self.neg_list.currentRowChanged.connect(self._num_changed)
        self.type_list.currentRowChanged.connect(self._num_changed)
        self.custom_edit.textChanged.connect(self._num_changed)
        self.custom_list.currentTextChanged.connect(self.custom_edit.setText)
        self._initial_fmt = fmt
        self._cat_changed(self.cat_list.currentRow(), initial=True)
        return w

    def _cat_changed(self, row, initial=False):
        cat = NUMBER_CATEGORIES[row]
        show = {
            "dec": cat in ("Number", "Currency", "Accounting", "Percentage", "Scientific"),
            "sym": cat in ("Currency", "Accounting"),
            "thousand": cat == "Number",
            "neg": cat in ("Number", "Currency"),
            "type": cat in ("Date", "Time"),
            "custom": cat == "Custom",
        }
        self.lbl_dec.setVisible(show["dec"])
        self.dec_spin.setVisible(show["dec"])
        self.lbl_sym.setVisible(show["sym"])
        self.sym_combo.setVisible(show["sym"])
        self.thousand_chk.setVisible(show["thousand"])
        self.lbl_neg.setVisible(show["neg"])
        self.neg_list.setVisible(show["neg"])
        self.lbl_type.setVisible(show["type"] or show["custom"])
        self.type_list.setVisible(show["type"])
        self.custom_edit.setVisible(show["custom"])
        self.custom_list.setVisible(show["custom"])
        self.type_list.blockSignals(True)
        self.type_list.clear()
        if cat == "Date":
            for f in DATE_FORMATS:
                self.type_list.addItem(format_value(45356.5833, f)[0] + "    (" + f + ")")
            idx = DATE_FORMATS.index(self._initial_fmt) if self._initial_fmt in DATE_FORMATS else 0
            self.type_list.setCurrentRow(idx)
        elif cat == "Time":
            for f in TIME_FORMATS:
                self.type_list.addItem(format_value(45356.5833, f)[0] + "    (" + f + ")")
            idx = TIME_FORMATS.index(self._initial_fmt) if self._initial_fmt in TIME_FORMATS else 0
            self.type_list.setCurrentRow(idx)
        self.type_list.blockSignals(False)
        self.neg_list.blockSignals(True)
        self.neg_list.clear()
        if show["neg"]:
            for i in range(4):
                fmt = build_number_format(cat, 2, True, "$", i)
                text, color = format_value(-1234.1, fmt)
                it = QListWidgetItem(text)
                if color:
                    it.setForeground(QColor(color))
                self.neg_list.addItem(it)
            self.neg_list.setCurrentRow(0)
        self.neg_list.blockSignals(False)
        if not initial:
            self._num_changed()
        else:
            self._update_sample(self._initial_fmt)

    def current_number_format(self):
        cat = NUMBER_CATEGORIES[self.cat_list.currentRow()]
        return build_number_format(
            cat, self.dec_spin.value(), self.thousand_chk.isChecked(), self.sym_combo.currentText(),
            max(0, self.neg_list.currentRow()),
            DATE_FORMATS[self.type_list.currentRow()] if cat == "Date" and self.type_list.currentRow() >= 0 else None,
            TIME_FORMATS[self.type_list.currentRow()] if cat == "Time" and self.type_list.currentRow() >= 0 else None,
            self.custom_edit.text())

    def _num_changed(self, *_):
        fmt = self.current_number_format()
        self._update_sample(fmt)
        if fmt != self._initial_fmt:
            self.changes["numfmt"] = fmt
        else:
            self.changes.pop("numfmt", None)

    def _update_sample(self, fmt):
        try:
            text, color = format_value(self.sample, fmt)
        except Exception:
            text, color = "(invalid format)", None
        self.sample_lbl.setText(text)
        self.sample_lbl.setStyleSheet(f"color: {color or '#000'};")

    # ---------------------------------------------------------------- alignment
    def _align_tab(self):
        w = QWidget()
        g = QGridLayout(w)
        self.h_combo = QComboBox()
        self.h_combo.addItems(["General", "Left", "Center", "Right"])
        self.h_combo.setCurrentIndex({None: 0, "left": 1, "center": 2, "right": 3}.get(self.st.halign, 0))
        self.v_combo = QComboBox()
        self.v_combo.addItems(["Top", "Center", "Bottom"])
        self.v_combo.setCurrentIndex({"top": 0, "center": 1}.get(self.st.valign, 2))
        self.indent_spin = QSpinBox()
        self.indent_spin.setRange(0, 15)
        self.indent_spin.setValue(self.st.indent)
        self.wrap_chk = QCheckBox("Wrap text")
        self.wrap_chk.setChecked(self.st.wrap)
        g.addWidget(QLabel("Horizontal:"), 0, 0)
        g.addWidget(self.h_combo, 0, 1)
        g.addWidget(QLabel("Indent:"), 0, 2)
        g.addWidget(self.indent_spin, 0, 3)
        g.addWidget(QLabel("Vertical:"), 1, 0)
        g.addWidget(self.v_combo, 1, 1)
        g.addWidget(self.wrap_chk, 2, 0, 1, 2)
        g.setRowStretch(3, 1)
        g.setColumnStretch(4, 1)
        self.h_combo.currentIndexChanged.connect(
            lambda i: self._set("halign", [None, "left", "center", "right"][i], self.st.halign))
        self.v_combo.currentIndexChanged.connect(
            lambda i: self._set("valign", ["top", "center", None][i], self.st.valign))
        self.indent_spin.valueChanged.connect(lambda v: self._set("indent", v, self.st.indent))
        self.wrap_chk.toggled.connect(lambda v: self._set("wrap", v, self.st.wrap))
        return w

    def _set(self, field, value, initial):
        if value != initial:
            self.changes[field] = value
        else:
            self.changes.pop(field, None)
        if field in ("font", "size", "bold", "italic", "underline", "strike", "color"):
            self._update_font_preview()

    # ---------------------------------------------------------------- font
    def _font_tab(self):
        w = QWidget()
        g = QGridLayout(w)
        self.font_combo = QFontComboBox()
        self.font_combo.setCurrentFont(QFont(self.st.font or "Calibri"))
        self.size_combo = QComboBox()
        self.size_combo.setEditable(True)
        self.size_combo.addItems([str(s) for s in (8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72)])
        self.size_combo.setCurrentText(_fmt_size(self.st.size or 11))
        self.b_chk = QCheckBox("Bold")
        self.b_chk.setChecked(self.st.bold)
        self.i_chk = QCheckBox("Italic")
        self.i_chk.setChecked(self.st.italic)
        self.u_chk = QCheckBox("Underline")
        self.u_chk.setChecked(self.st.underline)
        self.s_chk = QCheckBox("Strikethrough")
        self.s_chk.setChecked(self.st.strike)
        self.fcolor = ColorButton(self.st.color, "Automatic")
        self.preview = QLabel("AaBbCcYyZz")
        self.preview.setAlignment(Qt.AlignCenter)
        self.preview.setMinimumHeight(70)
        self.preview.setStyleSheet("background: white; border: 1px solid #D0D0D0;")
        g.addWidget(QLabel("Font:"), 0, 0)
        g.addWidget(self.font_combo, 0, 1)
        g.addWidget(QLabel("Size:"), 0, 2)
        g.addWidget(self.size_combo, 0, 3)
        g.addWidget(self.b_chk, 1, 0)
        g.addWidget(self.i_chk, 1, 1)
        g.addWidget(self.u_chk, 2, 0)
        g.addWidget(self.s_chk, 2, 1)
        g.addWidget(QLabel("Color:"), 3, 0)
        g.addWidget(self.fcolor, 3, 1)
        g.addWidget(QLabel("Preview:"), 4, 0)
        g.addWidget(self.preview, 5, 0, 1, 4)
        g.setRowStretch(6, 1)
        initial_font = self.st.font or "Calibri"
        self.font_combo.currentFontChanged.connect(
            lambda f: self._set("font", None if f.family() == "Calibri" else f.family(), self.st.font))
        self.size_combo.currentTextChanged.connect(self._size_changed)
        self.b_chk.toggled.connect(lambda v: self._set("bold", v, self.st.bold))
        self.i_chk.toggled.connect(lambda v: self._set("italic", v, self.st.italic))
        self.u_chk.toggled.connect(lambda v: self._set("underline", v, self.st.underline))
        self.s_chk.toggled.connect(lambda v: self._set("strike", v, self.st.strike))
        self.fcolor.changed.connect(lambda c: self._set("color", c, self.st.color))
        _ = initial_font
        self._update_font_preview()
        return w

    def _size_changed(self, t):
        try:
            v = float(t)
        except ValueError:
            return
        if 1 <= v <= 409:
            self._set("size", None if v == 11 else v, self.st.size)

    def _update_font_preview(self):
        if not hasattr(self, "preview"):
            return
        f = QFont(self.font_combo.currentFont().family())
        try:
            f.setPointSizeF(float(self.size_combo.currentText()))
        except ValueError:
            pass
        f.setBold(self.b_chk.isChecked())
        f.setItalic(self.i_chk.isChecked())
        f.setUnderline(self.u_chk.isChecked())
        f.setStrikeOut(self.s_chk.isChecked())
        self.preview.setFont(f)
        col = self.fcolor.color or "#000000"
        self.preview.setStyleSheet(f"background: white; border: 1px solid #D0D0D0; color: {col};")

    # ---------------------------------------------------------------- border
    def _border_tab(self):
        w = QWidget()
        g = QGridLayout(w)
        g.addWidget(QLabel("Presets:"), 0, 0)
        row = QHBoxLayout()
        self.border_mode = None
        for label, mode in (("None", "none"), ("Outline", "outline"), ("Inside", "inside"), ("All", "all")):
            b = QPushButton(label)
            b.setCheckable(True)
            b.clicked.connect(lambda _=False, m=mode, btn=b: self._border_pick(m, btn))
            row.addWidget(b)
        self._border_buttons = [row.itemAt(i).widget() for i in range(row.count())]
        g.addLayout(row, 0, 1)
        g.addWidget(QLabel("Edges:"), 1, 0)
        row2 = QHBoxLayout()
        for label, mode in (("Top", "top"), ("Bottom", "bottom"), ("Left", "left"), ("Right", "right")):
            b = QPushButton(label)
            b.setCheckable(True)
            b.clicked.connect(lambda _=False, m=mode, btn=b: self._border_pick(m, btn))
            row2.addWidget(b)
            self._border_buttons.append(b)
        g.addLayout(row2, 1, 1)
        self.bstyle = QComboBox()
        for key, label in BORDER_STYLES:
            self.bstyle.addItem(label, key)
        self.bcolor = ColorButton("#000000", "Automatic")
        g.addWidget(QLabel("Line style:"), 2, 0)
        g.addWidget(self.bstyle, 2, 1)
        g.addWidget(QLabel("Color:"), 3, 0)
        g.addWidget(self.bcolor, 3, 1)
        g.setRowStretch(4, 1)
        return w

    def _border_pick(self, mode, btn):
        for b in self._border_buttons:
            if b is not btn:
                b.setChecked(False)
        self.border_mode = mode if btn.isChecked() else None

    def border_choice(self):
        if not self.border_mode:
            return None
        return self.border_mode, (self.bstyle.currentData(), self.bcolor.color or "#000000")

    # ---------------------------------------------------------------- fill
    def _fill_tab(self):
        w = QWidget()
        v = QVBoxLayout(w)
        v.addWidget(QLabel("Background color:"))
        self.fill_btn = ColorButton(self.st.fill, "No Color")
        self.fill_btn.changed.connect(lambda c: self._set("fill", c, self.st.fill))
        v.addWidget(self.fill_btn)
        v.addStretch(1)
        return w


def _fmt_size(s):
    return str(int(s)) if float(s).is_integer() else str(s)


# ================================================================ Sort

class SortDialog(QDialog):
    def __init__(self, parent, sheet, rect, header):
        super().__init__(parent)
        self.setWindowTitle("Sort")
        self.sheet = sheet
        self.rect = rect
        self.setMinimumWidth(520)
        lay = QVBoxLayout(self)
        top = QHBoxLayout()
        add = QPushButton("+ Add Level")
        rm = QPushButton("Delete Level")
        top.addWidget(add)
        top.addWidget(rm)
        top.addStretch(1)
        self.header_chk = QCheckBox("My data has headers")
        self.header_chk.setChecked(header)
        top.addWidget(self.header_chk)
        lay.addLayout(top)
        self.table = QTableWidget(0, 2)
        self.table.setHorizontalHeaderLabels(["Column", "Order"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        lay.addWidget(self.table)
        self.case_chk = QCheckBox("Case sensitive")
        lay.addWidget(self.case_chk)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)
        add.clicked.connect(self.add_level)
        rm.clicked.connect(self.remove_level)
        self.header_chk.toggled.connect(self._relabel)
        self.add_level()

    def _labels(self):
        r1, c1, r2, c2 = self.rect
        out = []
        for c in range(c1, c2 + 1):
            if self.header_chk.isChecked():
                v = self.sheet.value(r1, c)
                out.append(str(v) if v not in (None, "") else f"Column {col_name(c)}")
            else:
                out.append(f"Column {col_name(c)}")
        return out

    def add_level(self):
        row = self.table.rowCount()
        self.table.insertRow(row)
        cb = QComboBox()
        cb.addItems(self._labels())
        if row < cb.count():
            cb.setCurrentIndex(row)
        oc = QComboBox()
        oc.addItems(["A to Z / Smallest to Largest", "Z to A / Largest to Smallest"])
        self.table.setCellWidget(row, 0, cb)
        self.table.setCellWidget(row, 1, oc)

    def remove_level(self):
        if self.table.rowCount() > 1:
            self.table.removeRow(self.table.currentRow() if self.table.currentRow() >= 0 else self.table.rowCount() - 1)

    def _relabel(self):
        labels = self._labels()
        for row in range(self.table.rowCount()):
            cb = self.table.cellWidget(row, 0)
            i = cb.currentIndex()
            cb.clear()
            cb.addItems(labels)
            cb.setCurrentIndex(i)

    def keys(self):
        c1 = self.rect[1]
        out = []
        for row in range(self.table.rowCount()):
            col = c1 + self.table.cellWidget(row, 0).currentIndex()
            asc = self.table.cellWidget(row, 1).currentIndex() == 0
            out.append((col, asc))
        return out


# ================================================================ Find & Replace

class FindDialog(QDialog):
    find_next = Signal(dict)
    find_all = Signal(dict)
    replace_one = Signal(dict)
    replace_all = Signal(dict)
    goto = Signal(str, int, int)

    def __init__(self, parent, replace=False):
        super().__init__(parent)
        self.setWindowTitle("Find and Replace")
        self.setModal(False)
        self.setMinimumWidth(520)
        lay = QVBoxLayout(self)
        g = QGridLayout()
        self.find_edit = QLineEdit()
        self.repl_edit = QLineEdit()
        g.addWidget(QLabel("Find what:"), 0, 0)
        g.addWidget(self.find_edit, 0, 1)
        self.repl_lbl = QLabel("Replace with:")
        g.addWidget(self.repl_lbl, 1, 0)
        g.addWidget(self.repl_edit, 1, 1)
        lay.addLayout(g)
        opts = QHBoxLayout()
        self.case_chk = QCheckBox("Match case")
        self.whole_chk = QCheckBox("Match entire cell contents")
        self.scope = QComboBox()
        self.scope.addItems(["Sheet", "Workbook"])
        self.lookin = QComboBox()
        self.lookin.addItems(["Formulas", "Values"])
        opts.addWidget(self.case_chk)
        opts.addWidget(self.whole_chk)
        opts.addStretch(1)
        opts.addWidget(QLabel("Within:"))
        opts.addWidget(self.scope)
        opts.addWidget(QLabel("Look in:"))
        opts.addWidget(self.lookin)
        lay.addLayout(opts)
        btns = QHBoxLayout()
        self.b_all = QPushButton("Find All")
        self.b_next = QPushButton("Find Next")
        self.b_repl = QPushButton("Replace")
        self.b_repl_all = QPushButton("Replace All")
        close = QPushButton("Close")
        for b in (self.b_repl_all, self.b_repl, self.b_all, self.b_next, close):
            btns.addWidget(b)
        btns.insertStretch(0, 1)
        lay.addLayout(btns)
        self.results = QListWidget()
        self.results.setVisible(False)
        self.results.setMinimumHeight(140)
        lay.addWidget(self.results)
        self.status = QLabel("")
        self.status.setStyleSheet("color: #555;")
        lay.addWidget(self.status)
        self.b_next.setDefault(True)
        self.b_next.clicked.connect(lambda: self.find_next.emit(self.params()))
        self.b_all.clicked.connect(lambda: self.find_all.emit(self.params()))
        self.b_repl.clicked.connect(lambda: self.replace_one.emit(self.params()))
        self.b_repl_all.clicked.connect(lambda: self.replace_all.emit(self.params()))
        close.clicked.connect(self.close)
        self.results.itemActivated.connect(self._goto_item)
        self.results.itemClicked.connect(self._goto_item)
        self.set_replace_mode(replace)

    def set_replace_mode(self, on):
        for w in (self.repl_lbl, self.repl_edit, self.b_repl, self.b_repl_all):
            w.setVisible(on)
        self.setWindowTitle("Find and Replace" if on else "Find")

    def params(self):
        return {"needle": self.find_edit.text(), "repl": self.repl_edit.text(),
                "case": self.case_chk.isChecked(), "whole": self.whole_chk.isChecked(),
                "workbook": self.scope.currentIndex() == 1,
                "formulas": self.lookin.currentIndex() == 0}

    def show_results(self, items):
        self.results.clear()
        for sheet_name, r, c, text in items:
            it = QListWidgetItem(f"{sheet_name}!{col_name(c)}{r + 1}    {text}")
            it.setData(Qt.UserRole, (sheet_name, r, c))
            self.results.addItem(it)
        self.results.setVisible(True)
        self.status.setText(f"{len(items)} cell(s) found")

    def _goto_item(self, it):
        name, r, c = it.data(Qt.UserRole)
        self.goto.emit(name, r, c)


# ================================================================ filter dropdown

class FilterPopup(QFrame):
    applied = Signal(int, object)   # col, spec or None (clear)
    sort_requested = Signal(int, bool)

    CONDITIONS = [("", "(no condition)"), ("contains", "Contains"), ("not_contains", "Does not contain"),
                  ("begins", "Begins with"), ("ends", "Ends with"), ("equals", "Equals"),
                  ("not_equals", "Does not equal"), (">", "Greater than"), (">=", "Greater or equal"),
                  ("<", "Less than"), ("<=", "Less or equal")]

    def __init__(self, parent, col, values, has_blanks, spec, title):
        super().__init__(parent, Qt.Popup)
        self.col = col
        self.setFrameShape(QFrame.StyledPanel)
        self.setStyleSheet("FilterPopup { background: white; border: 1px solid #A0A0A0; }")
        self.setMinimumWidth(270)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(6, 6, 6, 6)
        for text, asc in (("Sort A to Z", True), ("Sort Z to A", False)):
            b = QPushButton(text)
            b.setFlat(True)
            b.setStyleSheet("text-align: left; padding: 4px 6px; border: none;")
            b.clicked.connect(lambda _=False, a=asc: (self.close(), self.sort_requested.emit(self.col, a)))
            lay.addWidget(b)
        clear = QPushButton(f'Clear Filter From "{title}"')
        clear.setFlat(True)
        clear.setStyleSheet("text-align: left; padding: 4px 6px; border: none;")
        clear.setEnabled(spec is not None)
        clear.clicked.connect(lambda: (self.close(), self.applied.emit(self.col, None)))
        lay.addWidget(clear)
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setStyleSheet("color: #E0E0E0;")
        lay.addWidget(line)
        cond_row = QHBoxLayout()
        self.cond = QComboBox()
        for k, label in self.CONDITIONS:
            self.cond.addItem(label, k)
        self.cond_val = QLineEdit()
        self.cond_val.setPlaceholderText("value")
        cond_row.addWidget(self.cond)
        cond_row.addWidget(self.cond_val, 1)
        lay.addLayout(cond_row)
        self.search = QLineEdit()
        self.search.setPlaceholderText("Search")
        lay.addWidget(self.search)
        self.list = QListWidget()
        self.list.setMinimumHeight(220)
        lay.addWidget(self.list)
        allowed = spec.get("values") if spec else None
        blanks_ok = spec.get("blanks", True) if spec else True
        self.all_item = QListWidgetItem("(Select All)")
        self.all_item.setFlags(self.all_item.flags() | Qt.ItemIsUserCheckable)
        self.list.addItem(self.all_item)
        self.items = []
        shown = values[:10000]
        for t in shown:
            it = QListWidgetItem(t)
            it.setFlags(it.flags() | Qt.ItemIsUserCheckable)
            it.setCheckState(Qt.Checked if allowed is None or t in allowed else Qt.Unchecked)
            self.list.addItem(it)
            self.items.append(it)
        self.blank_item = None
        if has_blanks:
            self.blank_item = QListWidgetItem("(Blanks)")
            self.blank_item.setFlags(self.blank_item.flags() | Qt.ItemIsUserCheckable)
            self.blank_item.setCheckState(Qt.Checked if blanks_ok else Qt.Unchecked)
            self.list.addItem(self.blank_item)
            self.items.append(self.blank_item)
        if spec and spec.get("cond"):
            op, arg = spec["cond"]
            for i in range(self.cond.count()):
                if self.cond.itemData(i) == op:
                    self.cond.setCurrentIndex(i)
            self.cond_val.setText(str(arg))
        self._sync_all()
        self.list.itemChanged.connect(self._item_changed)
        self.search.textChanged.connect(self._filter_list)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self._ok)
        bb.rejected.connect(self.close)
        lay.addWidget(bb)
        self.full_values = values
        self._busy = False

    def _sync_all(self):
        self._busy = True
        states = {it.checkState() for it in self.items if not it.isHidden()}
        if states == {Qt.Checked}:
            self.all_item.setCheckState(Qt.Checked)
        elif states == {Qt.Unchecked}:
            self.all_item.setCheckState(Qt.Unchecked)
        else:
            self.all_item.setCheckState(Qt.PartiallyChecked)
        self._busy = False

    def _item_changed(self, it):
        if getattr(self, "_busy", False):
            return
        if it is self.all_item:
            st = Qt.Checked if it.checkState() != Qt.Unchecked else Qt.Unchecked
            self._busy = True
            for x in self.items:
                if not x.isHidden():
                    x.setCheckState(st)
            self.all_item.setCheckState(st)
            self._busy = False
        else:
            self._sync_all()

    def _filter_list(self, text):
        t = text.lower()
        for it in self.items:
            it.setHidden(bool(t) and t not in it.text().lower())
        if t:
            self._busy = True
            for it in self.items:
                it.setCheckState(Qt.Checked if not it.isHidden() else Qt.Unchecked)
            self._busy = False
        self._sync_all()

    def _ok(self):
        values = {it.text() for it in self.items if it is not self.blank_item and it.checkState() == Qt.Checked}
        blanks = self.blank_item is None or self.blank_item.checkState() == Qt.Checked
        all_values = len(values) == len([it for it in self.items if it is not self.blank_item])
        spec = {}
        if not (all_values and blanks):
            spec["values"] = values
            spec["blanks"] = blanks
        op = self.cond.currentData()
        if op:
            spec["cond"] = (op, self.cond_val.text())
        self.close()
        self.applied.emit(self.col, spec or None)
