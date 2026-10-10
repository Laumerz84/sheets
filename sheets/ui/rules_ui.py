"""Home > Conditional Formatting (menu, quick rules, New/Edit Rule, Manage Rules) and
Data > Data Validation. Every change is one undoable step (a snapshot of the sheet)."""
from PySide6.QtCore import Qt
from PySide6.QtGui import QBrush, QColor, QFont
from PySide6.QtWidgets import (QCheckBox, QComboBox, QDialog, QDialogButtonBox, QFormLayout, QGridLayout,
                               QHBoxLayout, QLabel, QLineEdit, QListWidget, QMenu, QMessageBox,
                               QPlainTextEdit, QPushButton, QSpinBox, QStackedWidget, QTabWidget,
                               QTreeWidget, QTreeWidgetItem, QVBoxLayout, QWidget)

from .. import condfmt as CF
from .. import validation as V
from ..refs import parse_range, range_addr
from ..values import is_num
from .dialogs import ColorButton

CELL_OPS = [("between", "between"), ("notBetween", "not between"), ("equal", "equal to"),
            ("notEqual", "not equal to"), ("greaterThan", "greater than"), ("lessThan", "less than"),
            ("greaterThanOrEqual", "greater than or equal to"), ("lessThanOrEqual", "less than or equal to")]
TEXT_OPS = [("containsText", "containing"), ("notContainsText", "not containing"),
            ("beginsWith", "beginning with"), ("endsWith", "ending with")]


def _rects_text(rects):
    return ", ".join(range_addr(*r) for r in rects)


def _parse_rects(text):
    out = []
    for part in text.replace(";", ",").split(","):
        part = part.strip().lstrip("=").replace("$", "")
        if not part:
            continue
        b = parse_range(part)
        if not b:
            return None
        out.append(tuple(b))
    return out or None


def _selection_average(win):
    nums = []
    for r1, c1, r2, c2 in win.sel_rects(visible=False):
        r2, c2 = min(r2, win.sheet.max_row), min(c2, win.sheet.max_col)
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                v = win.sheet.value(r, c)
                if is_num(v) and not isinstance(v, bool):
                    nums.append(v)
                if len(nums) > 20000:
                    break
    if not nums:
        return ""
    avg = sum(nums) / len(nums)
    return str(int(avg)) if avg == int(avg) else f"{avg:.2f}"


# ================================================================ format choice
class FormatChoice(QWidget):
    """Excel's preset looks plus 'Custom Format...' (fill, font colour, bold, italic)."""

    def __init__(self, parent=None, current=None):
        super().__init__(parent)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        self.combo = QComboBox()
        for label, fill, color in CF.PRESETS:
            self.combo.addItem(label, {"fill": fill, "color": color, "bold": None, "italic": None})
        self.combo.addItem("Custom Format...", None)
        lay.addWidget(self.combo, 1)
        self.custom = None
        self.preview = QLabel("AaBbCcYyZz")
        self.preview.setAlignment(Qt.AlignCenter)
        self.preview.setMinimumWidth(110)
        lay.addWidget(self.preview)
        self._last = 0
        if current is not None:
            for i in range(len(CF.PRESETS)):
                d = self.combo.itemData(i)
                if d["fill"] == current.get("fill") and d["color"] == current.get("color") \
                        and not current.get("bold") and not current.get("italic"):
                    self.combo.setCurrentIndex(i)
                    break
            else:
                self.custom = dict(current)
                self.combo.setItemText(len(CF.PRESETS), "Custom Format")
                self.combo.setCurrentIndex(len(CF.PRESETS))
        self._last = self.combo.currentIndex()
        self.combo.activated.connect(self._chosen)
        self._show()

    def _chosen(self, i):
        if self.combo.itemData(i) is None:  # Custom Format...
            got = CustomFormatDialog(self, self.custom).run()
            if got is None:
                self.combo.setCurrentIndex(self._last)
                return
            self.custom = got
            self.combo.setItemText(i, "Custom Format")
        self._last = i
        self._show()

    def value(self):
        d = self.combo.currentData()
        return dict(self.custom) if d is None else dict(d)

    def _show(self):
        fmt = self.value() if (self.combo.currentData() is not None or self.custom) else {}
        css = "border: 1px solid #A0A0A0; padding: 2px 6px;"
        css += f"background: {fmt.get('fill') or '#FFFFFF'}; color: {fmt.get('color') or '#000000'};"
        if fmt.get("bold"):
            css += "font-weight: bold;"
        if fmt.get("italic"):
            css += "font-style: italic;"
        self.preview.setStyleSheet(css)


class CustomFormatDialog(QDialog):
    def __init__(self, parent, current=None):
        super().__init__(parent)
        self.setWindowTitle("Format Cells")
        cur = current or {}
        form = QFormLayout(self)
        self.fill = ColorButton(cur.get("fill"), "No Fill")
        self.color = ColorButton(cur.get("color"), "Automatic")
        self.bold = QCheckBox("Bold")
        self.bold.setChecked(bool(cur.get("bold")))
        self.italic = QCheckBox("Italic")
        self.italic.setChecked(bool(cur.get("italic")))
        form.addRow("Fill color:", self.fill)
        form.addRow("Font color:", self.color)
        form.addRow("", self.bold)
        form.addRow("", self.italic)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        form.addRow(bb)

    def run(self):
        if self.exec() != QDialog.Accepted:
            return None
        return {"fill": self.fill.color, "color": self.color.color,
                "bold": True if self.bold.isChecked() else None,
                "italic": True if self.italic.isChecked() else None}


# ================================================================ quick rules (the menu items)
class QuickRuleDialog(QDialog):
    """'Greater Than...', 'Text that Contains...', 'Top 10 Items...' and friends: one or two
    values (or a count) plus a format, like Excel's small Highlight/Top-Bottom dialogs."""

    def __init__(self, win, title, prompt, values=1, initial="", spin=None, choice=None):
        super().__init__(win)
        self.setWindowTitle(title)
        v = QVBoxLayout(self)
        v.addWidget(QLabel(prompt))
        row = QHBoxLayout()
        self.edits = []
        self.spin = None
        self.choice = None
        if choice:
            self.choice = QComboBox()
            for label, data in choice:
                self.choice.addItem(label, data)
            row.addWidget(self.choice)
            row.addWidget(QLabel("values with"))
        if spin is not None:
            self.spin = QSpinBox()
            self.spin.setRange(1, 1000)
            self.spin.setValue(spin)
            row.addWidget(self.spin)
        for i in range(values):
            if i:
                row.addWidget(QLabel("and"))
            e = QLineEdit(initial if i == 0 else "")
            e.setMinimumWidth(120)
            self.edits.append(e)
            row.addWidget(e, 1)
        if values or spin is not None:
            row.addWidget(QLabel("with"))
        self.fmt = FormatChoice(self)
        row.addWidget(self.fmt, 2)
        v.addLayout(row)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self._ok)
        bb.rejected.connect(self.reject)
        v.addWidget(bb)
        if self.edits:
            self.edits[0].selectAll()
            self.edits[0].setFocus()

    def _ok(self):
        if any(not e.text().strip() for e in self.edits):
            QMessageBox.warning(self, self.windowTitle(), "Type a value (or a cell reference like =B1).")
            return
        self.accept()


# ================================================================ New / Edit Rule
RULE_TYPES = ["Format all cells based on their values (color scales, data bars)",
              "Format only cells that contain",
              "Format only top or bottom ranked values",
              "Format only values that are above or below average",
              "Format only unique or duplicate values",
              "Use a formula to determine which cells to format"]


class RuleEditor(QDialog):
    """Excel's New Formatting Rule / Edit Formatting Rule dialog."""

    def __init__(self, win, rects, rule=None):
        super().__init__(win)
        self.rects, self.rule = rects, rule
        self.setWindowTitle("Edit Formatting Rule" if rule else "New Formatting Rule")
        self.resize(560, 420)
        v = QVBoxLayout(self)
        v.addWidget(QLabel("Select a Rule Type:"))
        self.types = QListWidget()
        self.types.addItems(RULE_TYPES)
        self.types.setMaximumHeight(130)
        v.addWidget(self.types)
        v.addWidget(QLabel("Edit the Rule Description:"))
        self.pages = QStackedWidget()
        v.addWidget(self.pages, 1)
        fmt_now = None
        if rule is not None and rule.kind not in ("colorScale", "dataBar"):
            fmt_now = {"fill": rule.fill, "color": rule.color, "bold": rule.bold, "italic": rule.italic}
        self._build_scale_page(rule)
        self._build_contains_page(rule)
        self._build_top_page(rule)
        self._build_avg_page(rule)
        self._build_dup_page(rule)
        self._build_formula_page(rule)
        self.fmt = FormatChoice(self, fmt_now)
        self.fmt_row = QWidget()
        h = QHBoxLayout(self.fmt_row)
        h.setContentsMargins(0, 0, 0, 0)
        h.addWidget(QLabel("Format:"))
        h.addWidget(self.fmt, 1)
        v.addWidget(self.fmt_row)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self._ok)
        bb.rejected.connect(self.reject)
        v.addWidget(bb)
        self.types.currentRowChanged.connect(self._type_changed)
        self.types.setCurrentRow(self._type_of(rule))
        self.result = None

    @staticmethod
    def _type_of(rule):
        if rule is None:
            return 1
        k = rule.kind
        if k in ("colorScale", "dataBar"):
            return 0
        if k == "top10":
            return 2
        if k == "aboveAverage":
            return 3
        if k in ("duplicateValues", "uniqueValues"):
            return 4
        if k == "expression":
            return 5
        return 1

    def _type_changed(self, i):
        self.pages.setCurrentIndex(i)
        self.fmt_row.setVisible(i != 0)

    # ---- pages
    def _page(self):
        w = QWidget()
        g = QGridLayout(w)
        g.setAlignment(Qt.AlignTop)
        self.pages.addWidget(w)
        return g

    def _build_scale_page(self, rule):
        g = self._page()
        self.scale_style = QComboBox()
        self.scale_style.addItems(["2-Color Scale", "3-Color Scale", "Data Bar"])
        g.addWidget(QLabel("Format Style:"), 0, 0)
        g.addWidget(self.scale_style, 0, 1)
        cols = ["#F8696B", "#FFEB84", "#63BE7B"]
        bar = "#638EC6"
        style = 1
        if rule is not None and rule.kind == "colorScale" and rule.scale:
            pts = [c for _, _, c in rule.scale]
            style = 1 if len(pts) == 3 else 0
            cols = [pts[0], pts[1] if len(pts) == 3 else "#FFEB84", pts[-1]]
        elif rule is not None and rule.kind == "dataBar" and rule.bar:
            style, bar = 2, rule.bar[0]
        self.c_min, self.c_mid, self.c_max = (ColorButton(c, "White") for c in cols)
        self.c_bar = ColorButton(bar, "Blue")
        self.lbl_mid = QLabel("Midpoint color:")
        self.lbl_min, self.lbl_max, self.lbl_bar = QLabel("Lowest value:"), QLabel("Highest value:"), QLabel("Bar color:")
        for i, (lab, w) in enumerate(((self.lbl_min, self.c_min), (self.lbl_mid, self.c_mid),
                                      (self.lbl_max, self.c_max), (self.lbl_bar, self.c_bar)), start=1):
            g.addWidget(lab, i, 0)
            g.addWidget(w, i, 1)
        self.scale_style.currentIndexChanged.connect(self._scale_style)
        self.scale_style.setCurrentIndex(style)
        self._scale_style(style)

    def _scale_style(self, i):
        bar = i == 2
        for w in (self.lbl_min, self.c_min, self.lbl_max, self.c_max):
            w.setVisible(not bar)
        self.lbl_mid.setVisible(i == 1)
        self.c_mid.setVisible(i == 1)
        self.lbl_bar.setVisible(bar)
        self.c_bar.setVisible(bar)

    def _build_contains_page(self, rule):
        g = self._page()
        self.contains_what = QComboBox()
        for label, data in (("Cell Value", "cell"), ("Specific Text", "text"), ("Blanks", "containsBlanks"),
                            ("No Blanks", "notContainsBlanks"), ("Errors", "containsErrors"),
                            ("No Errors", "notContainsErrors")):
            self.contains_what.addItem(label, data)
        self.cell_op = QComboBox()
        for op, label in CELL_OPS:
            self.cell_op.addItem(label, op)
        self.text_op = QComboBox()
        for op, label in TEXT_OPS:
            self.text_op.addItem(label, op)
        self.v1, self.v2, self.text_val = QLineEdit(), QLineEdit(), QLineEdit()
        self.and_lbl = QLabel("and")
        g.addWidget(QLabel("Format only cells with:"), 0, 0, 1, 4)
        g.addWidget(self.contains_what, 1, 0)
        g.addWidget(self.cell_op, 1, 1)
        g.addWidget(self.text_op, 1, 1)
        g.addWidget(self.v1, 1, 2)
        g.addWidget(self.text_val, 1, 2)
        g.addWidget(self.and_lbl, 1, 3)
        g.addWidget(self.v2, 1, 4)
        if rule is not None:
            k = rule.kind
            if k == "cellIs":
                self.contains_what.setCurrentIndex(0)
                self.cell_op.setCurrentIndex(max(0, self.cell_op.findData(rule.op)))
                vals = [CF.formula_display(f) for f in rule.formulas] + ["", ""]
                self.v1.setText(vals[0])
                self.v2.setText(vals[1])
            elif k in CF.TEXT_KINDS:
                self.contains_what.setCurrentIndex(1)
                self.text_op.setCurrentIndex(max(0, self.text_op.findData(k)))
                self.text_val.setText(rule.text or "")
            else:
                i = self.contains_what.findData(k)
                if i >= 0:
                    self.contains_what.setCurrentIndex(i)
        else:
            self.cell_op.setCurrentIndex(self.cell_op.findData("greaterThan"))
        self.contains_what.currentIndexChanged.connect(self._contains_changed)
        self.cell_op.currentIndexChanged.connect(self._contains_changed)
        self._contains_changed()

    def _contains_changed(self, *_):
        what = self.contains_what.currentData()
        cell, text = what == "cell", what == "text"
        two = cell and self.cell_op.currentData() in ("between", "notBetween")
        self.cell_op.setVisible(cell)
        self.v1.setVisible(cell)
        self.and_lbl.setVisible(two)
        self.v2.setVisible(two)
        self.text_op.setVisible(text)
        self.text_val.setVisible(text)

    def _build_top_page(self, rule):
        g = self._page()
        self.top_dir = QComboBox()
        self.top_dir.addItems(["Top", "Bottom"])
        self.top_n = QSpinBox()
        self.top_n.setRange(1, 1000)
        self.top_n.setValue(10)
        self.top_pct = QCheckBox("% of the selected range")
        if rule is not None and rule.kind == "top10":
            self.top_dir.setCurrentIndex(1 if rule.bottom else 0)
            self.top_n.setValue(int(rule.rank))
            self.top_pct.setChecked(bool(rule.percent))
        g.addWidget(QLabel("Format values that rank in the:"), 0, 0, 1, 3)
        g.addWidget(self.top_dir, 1, 0)
        g.addWidget(self.top_n, 1, 1)
        g.addWidget(self.top_pct, 1, 2)

    def _build_avg_page(self, rule):
        g = self._page()
        self.avg_dir = QComboBox()
        self.avg_dir.addItems(["above", "below"])
        if rule is not None and rule.kind == "aboveAverage" and not rule.above:
            self.avg_dir.setCurrentIndex(1)
        g.addWidget(QLabel("Format values that are:"), 0, 0)
        g.addWidget(self.avg_dir, 1, 0)
        g.addWidget(QLabel("the average for the selected range"), 1, 1)

    def _build_dup_page(self, rule):
        g = self._page()
        self.dup = QComboBox()
        self.dup.addItem("duplicate", "duplicateValues")
        self.dup.addItem("unique", "uniqueValues")
        if rule is not None and rule.kind == "uniqueValues":
            self.dup.setCurrentIndex(1)
        g.addWidget(QLabel("Format all:"), 0, 0)
        g.addWidget(self.dup, 1, 0)
        g.addWidget(QLabel("values in the selected range"), 1, 1)

    def _build_formula_page(self, rule):
        g = self._page()
        self.formula = QLineEdit()
        self.formula.setPlaceholderText("=$B2>100")
        if rule is not None and rule.kind == "expression" and rule.formulas:
            self.formula.setText("=" + rule.formulas[0].lstrip("="))
        g.addWidget(QLabel("Format values where this formula is true (written for the top-left cell "
                           f"of {_rects_text(self.rects)}; it adjusts for the other cells):"), 0, 0)
        g.addWidget(self.formula, 1, 0)

    # ---- result
    def _ok(self):
        t = self.types.currentRow()
        stop = self.rule.stop if self.rule is not None else False
        f = self.fmt.value()
        rects = self.rects
        if t == 0:
            style = self.scale_style.currentIndex()
            if style == 2:
                rule = CF.bar_rule(rects, self.c_bar.color or "#638EC6")
            else:
                cols = [self.c_min.color or "#FFFFFF"] + ([self.c_mid.color or "#FFFFFF"] if style == 1 else []) + \
                       [self.c_max.color or "#FFFFFF"]
                rule = CF.scale_rule(rects, cols)
        elif t == 1:
            what = self.contains_what.currentData()
            if what == "cell":
                op = self.cell_op.currentData()
                vals = [self.v1.text()] + ([self.v2.text()] if op in ("between", "notBetween") else [])
                if any(not x.strip() for x in vals):
                    QMessageBox.warning(self, self.windowTitle(), "Type a value (or a cell reference like =B1).")
                    return
                rule = CF.new_rule(rects, "cellIs", op=op, formulas=[CF.value_formula(x) for x in vals], **f)
            elif what == "text":
                if not self.text_val.text():
                    QMessageBox.warning(self, self.windowTitle(), "Type the text to look for.")
                    return
                rule = CF.new_rule(rects, self.text_op.currentData(), text=self.text_val.text(), **f)
            else:
                rule = CF.new_rule(rects, what, **f)
        elif t == 2:
            rule = CF.new_rule(rects, "top10", rank=self.top_n.value(), percent=self.top_pct.isChecked(),
                               bottom=self.top_dir.currentIndex() == 1, **f)
        elif t == 3:
            rule = CF.new_rule(rects, "aboveAverage", above=self.avg_dir.currentIndex() == 0, **f)
        elif t == 4:
            rule = CF.new_rule(rects, self.dup.currentData(), **f)
        else:
            text = self.formula.text().strip()
            if not text:
                QMessageBox.warning(self, self.windowTitle(), "Type a formula, for example =$B2>100.")
                return
            rule = CF.new_rule(rects, "expression", formulas=[text.lstrip("=")], **f)
        if stop:
            rule = CF.edited(rule, stop=True)
        self.result = rule
        self.accept()


# ================================================================ Manage Rules
class ManageRulesDialog(QDialog):
    def __init__(self, win):
        super().__init__(win)
        self.win = win
        self.sheet = win.sheet
        self.rules = list(self.sheet.cond_formats)
        self.sel = _sel(win)
        self.setWindowTitle("Conditional Formatting Rules Manager")
        self.resize(760, 380)
        v = QVBoxLayout(self)
        top = QHBoxLayout()
        top.addWidget(QLabel("Show formatting rules for:"))
        self.scope = QComboBox()
        self.scope.addItems(["Current Selection", "This Worksheet"])
        top.addWidget(self.scope)
        top.addStretch(1)
        v.addLayout(top)
        btns = QHBoxLayout()
        for text, slot in (("New Rule...", self._new), ("Edit Rule...", self._edit), ("Delete Rule", self._delete),
                           ("Move Up", lambda: self._move(-1)), ("Move Down", lambda: self._move(1))):
            b = QPushButton(text)
            b.clicked.connect(slot)
            btns.addWidget(b)
        btns.addStretch(1)
        v.addLayout(btns)
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["Rule (applied in order shown)", "Format", "Applies to", "Stop If True"])
        self.tree.setRootIsDecorated(False)
        self.tree.setColumnWidth(0, 300)
        self.tree.setColumnWidth(1, 120)
        self.tree.setColumnWidth(2, 170)
        self.tree.itemChanged.connect(self._item_changed)
        self.tree.itemDoubleClicked.connect(lambda *_: self._edit())
        v.addWidget(self.tree, 1)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel | QDialogButtonBox.Apply)
        bb.accepted.connect(self._ok)
        bb.rejected.connect(self.reject)
        bb.button(QDialogButtonBox.Apply).clicked.connect(self._apply)
        v.addWidget(bb)
        self.scope.currentIndexChanged.connect(lambda _: self._fill())
        self._fill()

    def _visible(self, rule):
        if self.scope.currentIndex() == 1:
            return True
        return any(not (a[2] < b[0] or a[0] > b[2] or a[3] < b[1] or a[1] > b[3])
                   for a in rule.rects for b in self.sel)

    def _fill(self, select=None):
        self.tree.blockSignals(True)
        self.tree.clear()
        for i, rule in enumerate(self.rules):
            if not self._visible(rule):
                continue
            it = QTreeWidgetItem([CF.describe(rule), "AaBbCcYyZz", _rects_text(rule.rects), ""])
            it.setData(0, Qt.UserRole, i)
            it.setFlags(it.flags() | Qt.ItemIsEditable | Qt.ItemIsUserCheckable)
            it.setCheckState(3, Qt.Checked if rule.stop else Qt.Unchecked)
            fill = rule.fill
            if rule.kind == "colorScale" and rule.scale:
                fill = rule.scale[-1][2]
            elif rule.kind == "dataBar" and rule.bar:
                fill = rule.bar[0]
            if fill:
                it.setBackground(1, QBrush(QColor(fill)))
            if rule.color:
                it.setForeground(1, QBrush(QColor(rule.color)))
            if rule.bold or rule.italic:
                f = QFont(self.tree.font())
                f.setBold(bool(rule.bold))
                f.setItalic(bool(rule.italic))
                it.setFont(1, f)
            self.tree.addTopLevelItem(it)
            if select == i:
                self.tree.setCurrentItem(it)
        self.tree.blockSignals(False)

    def _current(self):
        it = self.tree.currentItem()
        return None if it is None else it.data(0, Qt.UserRole)

    def _item_changed(self, it, col):
        i = it.data(0, Qt.UserRole)
        rule = self.rules[i]
        if col == 2:
            rects = _parse_rects(it.text(2))
            if rects is None:
                QMessageBox.warning(self, self.windowTitle(), f"'{it.text(2)}' isn't a cell range like A1:B10.")
                self._fill(i)
                return
            self.rules[i] = CF.edited(rule, rects=rects)
        elif col == 3:
            self.rules[i] = CF.edited(rule, stop=it.checkState(3) == Qt.Checked)

    def _new(self):
        ed = RuleEditor(self.win, self.sel)
        if ed.exec() == QDialog.Accepted and ed.result is not None:
            self.rules.insert(0, ed.result)
            self._fill(0)

    def _edit(self):
        i = self._current()
        if i is None:
            return
        ed = RuleEditor(self.win, self.rules[i].rects, self.rules[i])
        if ed.exec() == QDialog.Accepted and ed.result is not None:
            self.rules[i] = ed.result
            self._fill(i)

    def _delete(self):
        i = self._current()
        if i is not None:
            del self.rules[i]
            self._fill()

    def _move(self, d):
        i = self._current()
        if i is None or not (0 <= i + d < len(self.rules)):
            return
        self.rules[i], self.rules[i + d] = self.rules[i + d], self.rules[i]
        self._fill(i + d)

    def _apply(self):
        if self.rules != self.sheet.cond_formats:
            self.win.set_cond_formats(list(self.rules), "Conditional Formatting")

    def _ok(self):
        self._apply()
        self.accept()


# ================================================================ the menu
def build_cf_menu(win, menu):
    """Fill `menu` with Excel's Conditional Formatting items (rebuilt each time it opens)."""
    def fill():
        menu.clear()
        hl = menu.addMenu("&Highlight Cells Rules")
        for label, kind, op, nvals in (("&Greater Than...", "cellIs", "greaterThan", 1),
                                       ("&Less Than...", "cellIs", "lessThan", 1),
                                       ("&Between...", "cellIs", "between", 2),
                                       ("&Equal To...", "cellIs", "equal", 1),
                                       ("&Text that Contains...", "containsText", None, 1)):
            hl.addAction(label, lambda k=kind, o=op, n=nvals, t=label: quick_highlight(win, t, k, o, n))
        hl.addAction("&Duplicate Values...", lambda: quick_duplicates(win))
        hl.addSeparator()
        hl.addAction("&More Rules...", lambda: new_rule_dialog(win))
        tb = menu.addMenu("&Top/Bottom Rules")
        for label, bottom, pct in (("Top 10 &Items...", False, False), ("Top 10 &%...", False, True),
                                   ("&Bottom 10 Items...", True, False), ("B&ottom 10 %...", True, True)):
            tb.addAction(label, lambda b=bottom, p=pct, t=label: quick_top(win, t, b, p))
        tb.addAction("&Above Average...", lambda: quick_average(win, True))
        tb.addAction("Below A&verage...", lambda: quick_average(win, False))
        db = menu.addMenu("&Data Bars")
        for label, color in CF.DATA_BARS:
            db.addAction(label, lambda c=color: win.add_cond_format(CF.bar_rule(_sel(win), c), "Data Bars"))
        cs = menu.addMenu("Color &Scales")
        for label, cols in CF.COLOR_SCALES:
            cs.addAction(label, lambda c=cols: win.add_cond_format(CF.scale_rule(_sel(win), c), "Color Scale"))
        menu.addSeparator()
        menu.addAction("&New Rule...", lambda: new_rule_dialog(win))
        cl = menu.addMenu("&Clear Rules")
        cl.addAction("Clear Rules from &Selected Cells", lambda: clear_rules(win, selection=True))
        cl.addAction("Clear Rules from &Entire Sheet", lambda: clear_rules(win, selection=False))
        menu.addAction("&Manage Rules...", lambda: ManageRulesDialog(win).exec())
    menu.aboutToShow.connect(fill)
    fill()


def _sel(win):
    return [tuple(r) for r in win.grid.selected_rects()]  # whole columns stay whole, like Excel


def quick_highlight(win, title, kind, op, nvals):
    title = title.replace("&", "").rstrip(".")
    if kind == "cellIs":
        prompt = {"greaterThan": "Format cells that are GREATER THAN:", "lessThan": "Format cells that are LESS THAN:",
                  "between": "Format cells that are BETWEEN:", "equal": "Format cells that are EQUAL TO:"}[op]
        d = QuickRuleDialog(win, title, prompt, nvals, _selection_average(win))
    else:
        d = QuickRuleDialog(win, title, "Format cells that contain the text:", 1, "")
    if d.exec() != QDialog.Accepted:
        return
    f = d.fmt.value()
    if kind == "cellIs":
        rule = CF.new_rule(_sel(win), "cellIs", op=op, formulas=[CF.value_formula(e.text()) for e in d.edits], **f)
    else:
        rule = CF.new_rule(_sel(win), kind, text=d.edits[0].text(), **f)
    win.add_cond_format(rule, "Conditional Formatting")


def quick_duplicates(win):
    d = QuickRuleDialog(win, "Duplicate Values", "Format cells that contain:", 0,
                        choice=[("Duplicate", "duplicateValues"), ("Unique", "uniqueValues")])
    if d.exec() == QDialog.Accepted:
        win.add_cond_format(CF.new_rule(_sel(win), d.choice.currentData(), **d.fmt.value()), "Conditional Formatting")


def quick_top(win, title, bottom, pct):
    title = title.replace("&", "").rstrip(".")
    prompt = f"Format cells that rank in the {'BOTTOM' if bottom else 'TOP'}{' (percent)' if pct else ''}:"
    d = QuickRuleDialog(win, title, prompt, 0, spin=10)
    if d.exec() == QDialog.Accepted:
        rule = CF.new_rule(_sel(win), "top10", rank=d.spin.value(), percent=pct, bottom=bottom, **d.fmt.value())
        win.add_cond_format(rule, "Conditional Formatting")


def quick_average(win, above):
    d = QuickRuleDialog(win, "Above Average" if above else "Below Average",
                        f"Format cells that are {'ABOVE' if above else 'BELOW'} AVERAGE:", 0)
    if d.exec() == QDialog.Accepted:
        win.add_cond_format(CF.new_rule(_sel(win), "aboveAverage", above=above, **d.fmt.value()),
                            "Conditional Formatting")


def new_rule_dialog(win):
    ed = RuleEditor(win, _sel(win))
    if ed.exec() == QDialog.Accepted and ed.result is not None:
        win.add_cond_format(ed.result, "Conditional Formatting")


def clear_rules(win, selection):
    sh = win.sheet
    if not selection:
        if sh.cond_formats:
            win.set_cond_formats([], "Clear Rules")
        return
    out = []
    for rule in sh.cond_formats:
        left = rule.rects
        for cut in _sel(win):
            left = V.subtract_rects(left, cut)
        if left == rule.rects:
            out.append(rule)
        elif left:
            out.append(CF.edited(rule, rects=left))
    if len(out) != len(sh.cond_formats) or any(a is not b for a, b in zip(out, sh.cond_formats)):
        win.set_cond_formats(out, "Clear Rules")


# ================================================================ Data Validation
class ValidationDialog(QDialog):
    """Excel's Data Validation dialog: Settings, Input Message, Error Alert."""

    def __init__(self, win, dv=None):
        super().__init__(win)
        self.setWindowTitle("Data Validation")
        self.resize(460, 340)
        self.cleared = False
        v = QVBoxLayout(self)
        tabs = QTabWidget()
        v.addWidget(tabs, 1)
        kind = (dv.type if dv is not None else None) or "any"

        # Settings
        s = QWidget()
        g = QGridLayout(s)
        g.setAlignment(Qt.AlignTop)
        self.allow = QComboBox()
        for k, label in V.TYPES:
            self.allow.addItem(label, k)
        self.allow.setCurrentIndex(max(0, self.allow.findData(kind)))
        self.blank = QCheckBox("Ignore blank")
        self.blank.setChecked(bool(dv.allow_blank) if dv is not None else True)
        self.data = QComboBox()
        for k, label in V.OPERATORS:
            self.data.addItem(label, k)
        self.data.setCurrentIndex(max(0, self.data.findData((dv.operator if dv is not None else None) or "between")))
        self.lbl1, self.lbl2 = QLabel("Minimum:"), QLabel("Maximum:")
        self.f1, self.f2 = QLineEdit(), QLineEdit()
        if dv is not None:
            self.f1.setText(V.formula_text(dv.formula1, kind))
            self.f2.setText(V.formula_text(dv.formula2, kind))
        self.dropdown = QCheckBox("In-cell dropdown")
        self.dropdown.setChecked(not (dv is not None and dv.showDropDown))
        g.addWidget(QLabel("Validation criteria"), 0, 0, 1, 2)
        g.addWidget(QLabel("Allow:"), 1, 0)
        g.addWidget(self.allow, 2, 0)
        g.addWidget(self.blank, 2, 1)
        g.addWidget(self.dropdown, 3, 1)
        self.lbl_data = QLabel("Data:")
        g.addWidget(self.lbl_data, 3, 0)
        g.addWidget(self.data, 4, 0)
        g.addWidget(self.lbl1, 5, 0)
        g.addWidget(self.f1, 6, 0, 1, 2)
        g.addWidget(self.lbl2, 7, 0)
        g.addWidget(self.f2, 8, 0, 1, 2)
        tabs.addTab(s, "Settings")

        # Input Message
        m = QWidget()
        f = QFormLayout(m)
        self.show_prompt = QCheckBox("Show input message when cell is selected")
        self.show_prompt.setChecked(bool(dv.showInputMessage) if dv is not None else True)
        self.prompt_title = QLineEdit((dv.promptTitle or "") if dv is not None else "")
        self.prompt = QPlainTextEdit((dv.prompt or "") if dv is not None else "")
        f.addRow(self.show_prompt)
        f.addRow("Title:", self.prompt_title)
        f.addRow("Input message:", self.prompt)
        tabs.addTab(m, "Input Message")

        # Error Alert
        e = QWidget()
        f = QFormLayout(e)
        self.show_error = QCheckBox("Show error alert after invalid data is entered")
        self.show_error.setChecked(bool(dv.showErrorMessage) if dv is not None else True)
        self.style = QComboBox()
        for k, label in V.STYLES:
            self.style.addItem(label, k)
        self.style.setCurrentIndex(max(0, self.style.findData((dv.errorStyle if dv is not None else None) or "stop")))
        self.error_title = QLineEdit((dv.errorTitle or "") if dv is not None else "")
        self.error = QPlainTextEdit((dv.error or "") if dv is not None else "")
        f.addRow(self.show_error)
        f.addRow("Style:", self.style)
        f.addRow("Title:", self.error_title)
        f.addRow("Error message:", self.error)
        tabs.addTab(e, "Error Alert")

        bottom = QHBoxLayout()
        clear = QPushButton("Clear All")
        clear.clicked.connect(self._clear)
        bottom.addWidget(clear)
        bottom.addStretch(1)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self._ok)
        bb.rejected.connect(self.reject)
        bottom.addWidget(bb)
        v.addLayout(bottom)
        self.allow.currentIndexChanged.connect(self._refresh)
        self.data.currentIndexChanged.connect(self._refresh)
        self._refresh()
        self.result = None

    def _refresh(self, *_):
        kind = self.allow.currentData()
        op = self.data.currentData()
        compare = kind in ("whole", "decimal", "date", "time", "textLength")
        two = compare and op in V.TWO_VALUES
        self.lbl_data.setVisible(compare)
        self.data.setVisible(compare)
        self.blank.setVisible(kind != "any")
        self.dropdown.setVisible(kind == "list")
        self.lbl1.setVisible(kind != "any")
        self.f1.setVisible(kind != "any")
        self.lbl2.setVisible(two)
        self.f2.setVisible(two)
        if kind == "list":
            self.lbl1.setText("Source:")
            self.f1.setPlaceholderText("Yes,No,Maybe   or   =$A$1:$A$10")
        elif kind == "custom":
            self.lbl1.setText("Formula:")
            self.f1.setPlaceholderText("=ISNUMBER(A1)")
        else:
            word = {"date": "date", "time": "time", "textLength": "length"}.get(kind, "")
            if two:
                self.lbl1.setText(f"Start {word}:" if word in ("date", "time") else "Minimum:")
                self.lbl2.setText(f"End {word}:" if word in ("date", "time") else "Maximum:")
            else:
                self.lbl1.setText({"date": "Date:", "time": "Time:", "textLength": "Length:"}.get(kind, "Value:"))
            self.f1.setPlaceholderText("")

    def _clear(self):
        self.cleared = True
        self.result = None
        self.accept()

    def _limit(self, text, kind):
        """A typed limit -> stored formula text: numbers, =references, dates/times as serial numbers."""
        t = text.strip()
        if t.startswith("=") or not t:
            return t
        if kind in ("date", "time"):
            from ..numfmt import parse_input
            val, _fmt = parse_input(t)
            if is_num(val):
                return repr(float(val)).rstrip("0").rstrip(".") if float(val) != int(val) else str(int(val))
            return None
        try:
            float(t)
            return t
        except ValueError:
            return None

    def _ok(self):
        kind = self.allow.currentData()
        op = self.data.currentData()
        f1, f2 = self.f1.text(), self.f2.text()
        if kind in ("list", "custom") and not f1.strip():
            QMessageBox.warning(self, "Data Validation", "Enter a source list." if kind == "list" else "Enter a formula.")
            return
        if kind in ("whole", "decimal", "date", "time", "textLength"):
            f1 = self._limit(f1, kind)
            f2 = self._limit(f2, kind) if op in V.TWO_VALUES else None
            if not f1 or (op in V.TWO_VALUES and not f2):
                QMessageBox.warning(self, "Data Validation",
                                    "Enter a value (a number, a date or time, or a cell reference like =B1).")
                return
        has_text = any(x.strip() for x in (self.prompt_title.text(), self.prompt.toPlainText(),
                                           self.error_title.text(), self.error.toPlainText()))
        if kind == "any" and not has_text:
            self.cleared = True
            self.result = None
            self.accept()
            return
        self.result = V.make(kind, op, f1, f2, allow_blank=self.blank.isChecked(),
                             dropdown=self.dropdown.isChecked(),
                             prompt_title=self.prompt_title.text().strip(), prompt=self.prompt.toPlainText().strip(),
                             show_prompt=self.show_prompt.isChecked(),
                             error_title=self.error_title.text().strip(), error=self.error.toPlainText().strip(),
                             style=self.style.currentData(), show_error=self.show_error.isChecked())
        self.accept()


def validation_dialog(win):
    sh = win.sheet
    r, c = win.grid.sel.active
    dv, _ = V.rule_at(sh, r, c)
    d = ValidationDialog(win, dv)
    if d.exec() != QDialog.Accepted:
        return
    rects = _sel(win)
    win.set_validations(V.apply_to(sh, rects, d.result), "Data Validation")


def show_list(win, r, c, cell_rect):
    """The in-cell dropdown: a menu under the cell; picking an item types it into the cell."""
    sh = win.sheet
    dv, rects = V.rule_at(sh, r, c)
    if dv is None or dv.type != "list":
        return
    items = V.list_items(sh, dv, rects, r, c)
    if not items:
        return
    menu = QMenu(win)
    menu.setMinimumWidth(max(cell_rect.width() + 18, 80))
    current = sh.value(r, c)
    for t in items[:2000]:
        a = menu.addAction(t)
        a.setCheckable(True)
        a.setChecked(current is not None and str(current) == t)
    pos = win.grid.mapToGlobal(cell_rect.bottomLeft())
    chosen = menu.exec(pos)
    if chosen is not None:
        win.commit_cell(r, c, chosen.text(), False)
