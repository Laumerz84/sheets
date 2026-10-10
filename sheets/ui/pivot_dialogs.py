"""PivotTable dialogs: Value Field Settings, the filter dropdown, Label / Value / Top 10 filters,
Grouping, More Sort Options, PivotTable Options and Calculated Field."""
from PySide6.QtCore import QPoint, Qt
from PySide6.QtGui import QColor, QIcon, QPainter, QPixmap
from PySide6.QtWidgets import (QAbstractItemView, QButtonGroup, QCheckBox, QComboBox, QDialog,
                               QDialogButtonBox, QDoubleSpinBox, QFrame, QGridLayout,
                               QHBoxLayout, QLabel, QLineEdit, QListWidget, QListWidgetItem, QMenu,
                               QMessageBox, QPushButton, QRadioButton, QTabWidget, QVBoxLayout,
                               QWidget)

from .. import pivot as P


def _warn(parent, text):
    QMessageBox.warning(parent, "PivotTable", text)


def style_icon(name, size=(36, 26)):
    """A tiny picture of a PivotTable style: header band, rows, total line."""
    S = P.STYLES[name]
    w, h = size
    pm = QPixmap(w, h)
    pm.fill(QColor("#FFFFFF"))
    p = QPainter(pm)
    p.fillRect(0, 0, w, 6, QColor(S["head"] or "#FFFFFF"))
    p.setPen(QColor(S["line"]))
    p.drawLine(0, 6, w, 6)
    for i, y in enumerate(range(8, h - 6, 4)):
        if i % 2 and S["band"]:
            p.fillRect(0, y, w, 4, QColor(S["band"]))
    p.fillRect(0, h - 6, w, 6, QColor(S["grand"] or "#FFFFFF"))
    p.drawLine(0, h - 6, w, h - 6)
    p.setPen(QColor("#A0A0A0"))
    p.drawRect(0, 0, w - 1, h - 1)
    p.end()
    return QIcon(pm)


# ================================================================ Value Field Settings
class ValueFieldDialog(QDialog):
    """Excel's Value Field Settings: summarize by, show values as (with base field / item), custom
    name, number format."""

    def __init__(self, parent, wb, pv, vi, tab=0):
        super().__init__(parent)
        self.setWindowTitle("Value Field Settings")
        self.wb, self.pv, self.vi = wb, pv, vi
        entry = pv["values"][vi]
        self.field, self.agg = entry[0], entry[1]
        o = P.vopts(entry)
        self.numfmt = o.get("numfmt")
        self.is_calc = self.field in (pv.get("calc") or {})
        v = QVBoxLayout(self)
        g = QGridLayout()
        g.addWidget(QLabel("Source Name:"), 0, 0)
        g.addWidget(QLabel(self.field), 0, 1)
        g.addWidget(QLabel("Custom Name:"), 1, 0)
        self.name = QLineEdit(P.value_label(self.field, self.agg, o))
        g.addWidget(self.name, 1, 1)
        v.addLayout(g)
        self.tabs = QTabWidget()
        # summarize
        w1 = QWidget()
        l1 = QVBoxLayout(w1)
        l1.addWidget(QLabel("Summarize value field by\nChoose the type of calculation that you want to use to "
                            "summarize data from the selected field"))
        self.aggs = QListWidget()
        for k, label in P.AGGS:
            it = QListWidgetItem(label)
            it.setData(Qt.UserRole, k)
            self.aggs.addItem(it)
            if k == self.agg:
                self.aggs.setCurrentItem(it)
        if self.is_calc:
            self.aggs.setEnabled(False)
            l1.addWidget(QLabel("A calculated field always uses Sum of the fields in its formula."))
        self.aggs.currentItemChanged.connect(self._agg_changed)
        l1.addWidget(self.aggs)
        self.tabs.addTab(w1, "Summarize Values By")
        # show values as
        w2 = QWidget()
        l2 = QGridLayout(w2)
        l2.addWidget(QLabel("Show values as"), 0, 0, 1, 2)
        self.show_as = QComboBox()
        for k, label in P.SHOW_AS:
            self.show_as.addItem(label, k)
        self.show_as.setCurrentIndex(max(0, self.show_as.findData(o.get("show", "none"))))
        l2.addWidget(self.show_as, 1, 0, 1, 2)
        l2.addWidget(QLabel("Base field:"), 2, 0)
        l2.addWidget(QLabel("Base item:"), 2, 1)
        self.base_field = QListWidget()
        self.base_item = QListWidget()
        for f in [f for f in pv["rows"] + pv["cols"] if f != P.VALUES]:
            self.base_field.addItem(f)
        l2.addWidget(self.base_field, 3, 0)
        l2.addWidget(self.base_item, 3, 1)
        self.tabs.addTab(w2, "Show Values As")
        v.addWidget(self.tabs)
        self._want_item = o.get("base_item")
        bf = o.get("base_field")
        found = self.base_field.findItems(bf, Qt.MatchExactly) if bf else []
        self.base_field.currentTextChanged.connect(self._fill_items)
        if found:
            self.base_field.setCurrentItem(found[0])
        elif self.base_field.count():
            self.base_field.setCurrentRow(0)
        self.show_as.currentIndexChanged.connect(self._sync_base)
        self._sync_base()
        h = QHBoxLayout()
        nf = QPushButton("Number Format")
        nf.clicked.connect(self._number_format)
        h.addWidget(nf)
        self.fmt_lbl = QLabel()
        h.addWidget(self.fmt_lbl, 1)
        v.addLayout(h)
        self._show_fmt()
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self._ok)
        bb.rejected.connect(self.reject)
        v.addWidget(bb)
        self.tabs.setCurrentIndex(tab)
        self.result = None
        self.resize(420, 440)

    def _agg_changed(self, cur, prev):
        if cur is None or prev is None:
            return
        old_default = P.value_label(self.field, prev.data(Qt.UserRole))
        if self.name.text() == old_default:
            self.name.setText(P.value_label(self.field, cur.data(Qt.UserRole)))

    def _fill_items(self, field):
        self.base_item.clear()
        self.base_item.addItems(["(previous)", "(next)"] + P.items(self.wb, self.pv, field))
        want = self._want_item or "(previous)"
        got = self.base_item.findItems(want, Qt.MatchExactly)
        self.base_item.setCurrentItem(got[0] if got else self.base_item.item(0))

    def _sync_base(self, *_):
        need = P.SHOW_BASE.get(self.show_as.currentData())
        self.base_field.setEnabled(bool(need))
        self.base_item.setEnabled(need == "item")

    def _number_format(self):
        from .dialogs import FormatCellsDialog
        st = P.DEFAULT_STYLE.with_(numfmt=self.numfmt or "General")
        d = FormatCellsDialog(self, st, 1234.5678, 0)
        while d.tabs.count() > 1:
            d.tabs.removeTab(1)
        if d.exec() == QDialog.Accepted:
            self.numfmt = d.current_number_format()
            if self.numfmt == "General":
                self.numfmt = None
            self._show_fmt()

    def _show_fmt(self):
        self.fmt_lbl.setText(f"Format: {self.numfmt}" if self.numfmt else "Format: automatic")

    def _ok(self):
        name = self.name.text().strip()
        agg = self.aggs.currentItem().data(Qt.UserRole) if self.aggs.currentItem() else self.agg
        if self.is_calc:
            agg = "sum"
        if name and name.lower() in (n.lower() for n in P.source_fields(self.wb, self.pv)):
            _warn(self, "PivotTable field name already exists. Pick a name that isn't one of the source "
                        "columns (Excel adds a space, e.g. 'Revenue ').")
            return
        show = self.show_as.currentData()
        opts = {"show": show}
        need = P.SHOW_BASE.get(show)
        if need:
            if not self.base_field.currentItem():
                _warn(self, "This calculation needs a base field in Rows or Columns.")
                return
            opts["base_field"] = self.base_field.currentItem().text()
            if need == "item" and self.base_item.currentItem():
                opts["base_item"] = self.base_item.currentItem().text()
        if name and name != P.value_label(self.field, agg):
            opts["name"] = name
        if self.numfmt:
            opts["numfmt"] = self.numfmt
        self.result = P.value_entry(self.field, agg, opts)
        self.accept()


# ================================================================ filter dropdown
class PivotFilterPopup(QFrame):
    """The dropdown of a pivot's Row Labels / Column Labels / field header / report filter cell."""

    def __init__(self, parent, ctl, sheet, pv, area, fields, field=None):
        super().__init__(parent, Qt.Popup)
        self.ctl, self.sheet, self.pv, self.area = ctl, sheet, pv, area
        self.setFrameShape(QFrame.StyledPanel)
        self.setStyleSheet("PivotFilterPopup { background: white; border: 1px solid #A0A0A0; }")
        self.setMinimumWidth(290)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(6, 6, 6, 6)
        self.combo = QComboBox()
        self.combo.addItems(fields)
        if len(fields) > 1:
            row = QHBoxLayout()
            row.addWidget(QLabel("Select field:"))
            row.addWidget(self.combo, 1)
            lay.addLayout(row)
        self.body = QVBoxLayout()
        lay.addLayout(self.body)
        self.combo.currentTextChanged.connect(self._build)
        if field in fields:
            self.combo.setCurrentText(field)
        self._build(self.combo.currentText())

    def _flat(self, text, slot, enabled=True):
        b = QPushButton(text)
        b.setFlat(True)
        b.setStyleSheet("text-align: left; padding: 3px 6px; border: none;")
        b.setEnabled(enabled)
        if slot is not None:
            b.clicked.connect(lambda: (self.close(), slot()))
        self.body.addWidget(b)
        return b

    def _build(self, field):
        while self.body.count():
            it = self.body.takeAt(0)
            w = it.widget()
            if w is not None:
                w.deleteLater()
            elif it.layout() is not None:
                while it.layout().count():
                    x = it.layout().takeAt(0).widget()
                    if x is not None:
                        x.deleteLater()
        self.field = field
        ctl, sh, pv = self.ctl, self.sheet, self.pv
        if self.area != "report":
            self._flat("Sort A to Z", lambda: ctl.sort(sh, pv, field, "asc"))
            self._flat("Sort Z to A", lambda: ctl.sort(sh, pv, field, "desc"))
            self._flat("More Sort Options...", lambda: ctl.more_sort(sh, pv, field))
            line = QFrame()
            line.setFrameShape(QFrame.HLine)
            line.setStyleSheet("color: #E0E0E0;")
            self.body.addWidget(line)
        has = field in (pv.get("filters") or {}) or field in (pv.get("label_filters") or {}) or \
            field in (pv.get("value_filters") or {})
        self._flat(f'Clear Filter From "{field}"', lambda: ctl.clear_filter(sh, pv, field), has)
        if self.area != "report":
            for title, ops, kind in (("Label Filters", P.LABEL_OPS, "label"), ("Value Filters", P.VALUE_OPS, "value")):
                b = self._flat(title + "  ▸", None)
                m = QMenu(b)
                for op, label in ops:
                    m.addAction(label, lambda o=op, k=kind: (self.close(), ctl.cond_filter(sh, pv, field, k, o)))
                cur = (pv.get("label_filters" if kind == "label" else "value_filters") or {}).get(field)
                if cur:
                    m.addSeparator()
                    m.addAction(f"Clear {title[:-1]}", lambda k=kind: (self.close(), ctl.clear_cond(sh, pv, field, k)))
                b.setMenu(m)
        self.search = QLineEdit()
        self.search.setPlaceholderText("Search")
        self.body.addWidget(self.search)
        self.list = QListWidget()
        self.list.setMinimumHeight(200)
        self.body.addWidget(self.list)
        allowed = (pv.get("filters") or {}).get(field)
        self.all_item = QListWidgetItem("(Select All)")
        self.all_item.setFlags(self.all_item.flags() | Qt.ItemIsUserCheckable)
        self.list.addItem(self.all_item)
        self.items = []
        for t in P.items(ctl.win.wb, pv, field)[:10000]:
            it = QListWidgetItem(t)
            it.setFlags(it.flags() | Qt.ItemIsUserCheckable)
            it.setCheckState(Qt.Checked if allowed is None or t in allowed else Qt.Unchecked)
            self.list.addItem(it)
            self.items.append(it)
        self._busy = False
        self._sync_all()
        self.list.itemChanged.connect(self._item_changed)
        self.search.textChanged.connect(self._filter_list)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self._ok)
        bb.rejected.connect(self.close)
        self.body.addWidget(bb)

    def _sync_all(self):
        self._busy = True
        states = {it.checkState() for it in self.items if not it.isHidden()}
        self.all_item.setCheckState(Qt.Checked if states == {Qt.Checked} else
                                    Qt.Unchecked if states == {Qt.Unchecked} else Qt.PartiallyChecked)
        self._busy = False

    def _item_changed(self, it):
        if self._busy:
            return
        if it is self.all_item:
            st = Qt.Checked if it.checkState() != Qt.Unchecked else Qt.Unchecked
            self._busy = True
            for x in self.items:
                if not x.isHidden():
                    x.setCheckState(st)
            self._busy = False
        self._sync_all()

    def _filter_list(self, text):
        t = text.lower()
        self._busy = True
        for it in self.items:
            it.setHidden(bool(t) and t not in it.text().lower())
            if t:
                it.setCheckState(Qt.Checked if not it.isHidden() else Qt.Unchecked)
        self._busy = False
        self._sync_all()

    def allowed(self):
        ticked = [it.text() for it in self.items if it.checkState() == Qt.Checked]
        return None if len(ticked) == len(self.items) else ticked

    def _ok(self):
        allowed = self.allowed()
        if allowed == []:
            _warn(self, "Pick at least one item.")
            return
        self.close()
        self.ctl.set_items(self.sheet, self.pv, self.field, allowed)


# ================================================================ Label / Value / Top 10 filters
class CondFilterDialog(QDialog):
    def __init__(self, parent, pv, field, kind, op, current=None):
        super().__init__(parent)
        self.kind, self.op = kind, op
        vlabels = P.value_labels(pv)
        title = {"label": "Label Filter", "value": "Value Filter"}[kind]
        self.setWindowTitle(f"Top 10 Filter ({field})" if op == "top" else f"{title} ({field})")
        v = QVBoxLayout(self)
        g = QGridLayout()
        self.vcombo = QComboBox()
        self.vcombo.addItems(vlabels)
        cur = current if current and current[0] == op else None
        if op == "top":
            v.addWidget(QLabel("Show"))
            self.tb = QComboBox()
            self.tb.addItems(["Top", "Bottom"])
            self.n = QDoubleSpinBox()
            self.n.setRange(0, 1e9)
            self.n.setDecimals(0)
            self.n.setValue(10)
            self.kind_box = QComboBox()
            for k, t in (("items", "Items"), ("percent", "Percent"), ("sum", "Sum")):
                self.kind_box.addItem(t, k)
            if cur:
                self.n.setValue(P._to_float(cur[2], 10))
                self.kind_box.setCurrentIndex(max(0, self.kind_box.findData(cur[3] if len(cur) > 3 else "items")))
                self.tb.setCurrentIndex(1 if len(cur) > 4 and cur[4] == "bottom" else 0)
                if cur[1] in vlabels:
                    self.vcombo.setCurrentText(cur[1])
            g.addWidget(self.tb, 0, 0)
            g.addWidget(self.n, 0, 1)
            g.addWidget(self.kind_box, 0, 2)
            g.addWidget(QLabel("by"), 0, 3)
            g.addWidget(self.vcombo, 0, 4)
        else:
            ops = P.LABEL_OPS if kind == "label" else P.VALUE_OPS[:-1]
            v.addWidget(QLabel("Show items for which the label" if kind == "label" else "Show items for which"))
            col = 0
            if kind == "value":
                g.addWidget(self.vcombo, 0, col)
                col += 1
                if cur and cur[1] in vlabels:
                    self.vcombo.setCurrentText(cur[1])
            self.opbox = QComboBox()
            for k, t in ops:
                self.opbox.addItem(t.rstrip("."), k)
            self.opbox.setCurrentIndex(max(0, self.opbox.findData(op)))
            g.addWidget(self.opbox, 0, col)
            self.a = QLineEdit()
            self.b = QLineEdit()
            self.and_lbl = QLabel("and")
            g.addWidget(self.a, 0, col + 1)
            g.addWidget(self.and_lbl, 0, col + 2)
            g.addWidget(self.b, 0, col + 3)
            if cur:
                ia = 1 if kind == "label" else 2
                self.a.setText(str(cur[ia]) if len(cur) > ia and cur[ia] is not None else "")
                self.b.setText(str(cur[ia + 1]) if len(cur) > ia + 1 and cur[ia + 1] is not None else "")
            self.opbox.currentIndexChanged.connect(self._sync)
            self._sync()
            if kind == "label":
                v.addWidget(QLabel("Use ? to represent any single character, * for any series of characters."))
        v.addLayout(g)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        v.addWidget(bb)
        if not vlabels and kind == "value":
            v.addWidget(QLabel("Add a field to Values first."))
            bb.button(QDialogButtonBox.Ok).setEnabled(False)

    def _sync(self, *_):
        two = self.opbox.currentData() in ("between", "not_between")
        self.and_lbl.setVisible(two)
        self.b.setVisible(two)

    def spec(self):
        if self.op == "top":
            return ["top", self.vcombo.currentText(), self.n.value(), self.kind_box.currentData(),
                    "bottom" if self.tb.currentIndex() else "top"]
        op = self.opbox.currentData()
        if self.kind == "label":
            return [op, self.a.text(), self.b.text() or None]
        return [op, self.vcombo.currentText(), self.a.text(), self.b.text() or None]


# ================================================================ Grouping
class GroupDialog(QDialog):
    def __init__(self, parent, wb, pv, field):
        super().__init__(parent)
        self.setWindowTitle("Grouping")
        self.kind = P.field_kind(wb, pv, field)
        cur = (pv.get("groups") or {}).get(field) or {}
        v = QVBoxLayout(self)
        g = QGridLayout()
        lo, hi = P.number_range(wb, pv, field)
        if self.kind == "date":
            from ..numfmt import format_value
            g.addWidget(QLabel("Starting at:"), 0, 0)
            g.addWidget(QLabel(format_value(lo, "m/d/yyyy")[0]), 0, 1)
            g.addWidget(QLabel("Ending at:"), 1, 0)
            g.addWidget(QLabel(format_value(hi, "m/d/yyyy")[0]), 1, 1)
            v.addLayout(g)
            v.addWidget(QLabel("By"))
            self.levels = QListWidget()
            self.levels.setSelectionMode(QAbstractItemView.MultiSelection)
            want = cur.get("date") or ["months"]
            for k, label in reversed(P.DATE_LEVELS):
                it = QListWidgetItem(label)
                it.setData(Qt.UserRole, k)
                self.levels.addItem(it)
                it.setSelected(k in want)
            v.addWidget(self.levels)
        else:
            spec = cur.get("num") or [lo, hi, 10]
            self.start = QDoubleSpinBox()
            self.end = QDoubleSpinBox()
            self.by = QDoubleSpinBox()
            for i, (label, box, val) in enumerate((("Starting at:", self.start, spec[0]),
                                                   ("Ending at:", self.end, spec[1]), ("By:", self.by, spec[2]))):
                box.setRange(-1e12, 1e12)
                box.setDecimals(4)
                box.setValue(float(val))
                g.addWidget(QLabel(label), i, 0)
                g.addWidget(box, i, 1)
            self.by.setRange(1e-9, 1e12)
            v.addLayout(g)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self._ok)
        bb.rejected.connect(self.reject)
        v.addWidget(bb)
        self.spec = None

    def _ok(self):
        if self.kind == "date":
            picked = {self.levels.item(i).data(Qt.UserRole) for i in range(self.levels.count())
                      if self.levels.item(i).isSelected()}
            if not picked:
                _warn(self, "Pick at least one of Seconds, Minutes, Hours, Days, Months, Quarters or Years.")
                return
            self.spec = {"date": [k for k in P._LEVEL_ORDER if k in picked]}
        else:
            if self.end.value() < self.start.value():
                _warn(self, "The ending number must be greater than the starting number.")
                return
            self.spec = {"num": [self.start.value(), self.end.value(), self.by.value()]}
        self.accept()


# ================================================================ More Sort Options
class SortDialog(QDialog):
    def __init__(self, parent, pv, field):
        super().__init__(parent)
        self.setWindowTitle(f"Sort ({field})")
        cur = (pv.get("sort") or {}).get(field) or ["asc", None]
        v = QVBoxLayout(self)
        v.addWidget(QLabel("Sort options"))
        self.asc = QRadioButton("Ascending (A to Z) by:")
        self.desc = QRadioButton("Descending (Z to A) by:")
        grp = QButtonGroup(self)
        grp.addButton(self.asc)
        grp.addButton(self.desc)
        (self.desc if cur[0] == "desc" else self.asc).setChecked(True)
        self.by = QComboBox()
        self.by.addItem(field, None)
        for t in P.value_labels(pv):
            self.by.addItem(t, t)
        if len(cur) > 1 and cur[1]:
            self.by.setCurrentIndex(max(0, self.by.findData(cur[1])))
        v.addWidget(self.asc)
        v.addWidget(self.desc)
        v.addWidget(self.by)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        v.addWidget(bb)

    def spec(self):
        return ["desc" if self.desc.isChecked() else "asc", self.by.currentData()]


# ================================================================ PivotTable Options
FORMS = [("compact", "Show in Compact Form"), ("outline", "Show in Outline Form"), ("tabular", "Show in Tabular Form")]
SUBTOTALS = [("none", "Do Not Show Subtotals"), ("bottom", "Show all Subtotals at Bottom of Group"),
             ("top", "Show all Subtotals at Top of Group")]
GRANDS = [("none", "Off for Rows and Columns"), ("both", "On for Rows and Columns"), ("rows", "On for Rows Only"),
          ("cols", "On for Columns Only")]


class OptionsDialog(QDialog):
    def __init__(self, parent, pv):
        super().__init__(parent)
        self.setWindowTitle("PivotTable Options")
        self.pv = pv
        lay = P.layout(pv)
        v = QVBoxLayout(self)
        h = QHBoxLayout()
        h.addWidget(QLabel("PivotTable Name:"))
        self.name = QLineEdit(pv["name"])
        h.addWidget(self.name, 1)
        v.addLayout(h)
        tabs = QTabWidget()
        v.addWidget(tabs)
        # Layout & Format
        w = QWidget()
        g = QGridLayout(w)
        self.form = QComboBox()
        for k, t in FORMS:
            self.form.addItem(t, k)
        self.form.setCurrentIndex(self.form.findData(lay["form"]))
        self.repeat = QCheckBox("Repeat All Item Labels")
        self.repeat.setChecked(lay["repeat"])
        self.sub = QComboBox()
        for k, t in SUBTOTALS:
            self.sub.addItem(t, k)
        self.sub.setCurrentIndex(self.sub.findData(lay["subtotals"]))
        self.grand = QComboBox()
        for k, t in GRANDS:
            self.grand.addItem(t, k)
        self.grand.setCurrentIndex(self.grand.findData(lay["grand"]))
        self.blank = QCheckBox("Insert Blank Line after Each Item")
        self.blank.setChecked(lay["blank_line"])
        self.empty_on = QCheckBox("For empty cells show:")
        self.empty = QLineEdit(lay["empty"])
        self.empty_on.setChecked(bool(lay["empty"]))
        self.empty.setEnabled(bool(lay["empty"]))
        self.empty_on.toggled.connect(self.empty.setEnabled)
        g.addWidget(QLabel("Report Layout:"), 0, 0)
        g.addWidget(self.form, 0, 1)
        g.addWidget(self.repeat, 1, 1)
        g.addWidget(QLabel("Subtotals:"), 2, 0)
        g.addWidget(self.sub, 2, 1)
        g.addWidget(QLabel("Grand Totals:"), 3, 0)
        g.addWidget(self.grand, 3, 1)
        g.addWidget(self.blank, 4, 1)
        g.addWidget(self.empty_on, 5, 0)
        g.addWidget(self.empty, 5, 1)
        g.setRowStretch(6, 1)
        tabs.addTab(w, "Layout && Format")
        # Display
        w2 = QWidget()
        g2 = QGridLayout(w2)
        self.buttons = QCheckBox("Show expand/collapse buttons and filter dropdowns")
        self.buttons.setChecked(lay["buttons"])
        self.empty_items = QCheckBox("Show items with no data on rows and columns")
        self.empty_items.setChecked(lay["show_empty_items"])
        self.values_rows = QCheckBox("Show the Σ Values field in Rows (with 2+ value fields)")
        self.values_rows.setChecked(P.VALUES in pv["rows"])
        g2.addWidget(self.buttons, 0, 0)
        g2.addWidget(self.empty_items, 1, 0)
        g2.addWidget(self.values_rows, 2, 0)
        g2.setRowStretch(3, 1)
        tabs.addTab(w2, "Display")
        # Design
        w3 = QWidget()
        g3 = QGridLayout(w3)
        self.style = QComboBox()
        for n in P.STYLES:
            self.style.addItem(style_icon(n), n)
        self.style.setCurrentText(lay["style"])
        self.bands_r = QCheckBox("Banded Rows")
        self.bands_r.setChecked(lay["banded_rows"])
        self.bands_c = QCheckBox("Banded Columns")
        self.bands_c.setChecked(lay["banded_cols"])
        g3.addWidget(QLabel("PivotTable Style:"), 0, 0)
        g3.addWidget(self.style, 0, 1)
        g3.addWidget(self.bands_r, 1, 1)
        g3.addWidget(self.bands_c, 2, 1)
        g3.setRowStretch(3, 1)
        tabs.addTab(w3, "Design")
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self._ok)
        bb.rejected.connect(self.reject)
        v.addWidget(bb)
        self.result = None

    def _ok(self):
        name = self.name.text().strip()
        if not name:
            _warn(self, "Type a name for the PivotTable.")
            return
        pv = P.with_layout(self.pv, form=self.form.currentData(), repeat=self.repeat.isChecked(),
                           subtotals=self.sub.currentData(), grand=self.grand.currentData(),
                           blank_line=self.blank.isChecked(),
                           empty=self.empty.text() if self.empty_on.isChecked() else "",
                           buttons=self.buttons.isChecked(), show_empty_items=self.empty_items.isChecked(),
                           style=self.style.currentText(), banded_rows=self.bands_r.isChecked(),
                           banded_cols=self.bands_c.isChecked())
        pv = P.changed(pv, name=name)
        in_rows = P.VALUES in pv["rows"]
        if self.values_rows.isChecked() != in_rows:
            pv = values_moved(pv, "rows" if self.values_rows.isChecked() else "cols")
        self.result = pv
        self.accept()


def values_moved(pv, area):
    """pv with the Σ Values field moved to the end of Rows or Columns."""
    rows = [f for f in pv["rows"] if f != P.VALUES]
    cols = [f for f in pv["cols"] if f != P.VALUES]
    if area == "rows":
        rows.append(P.VALUES)
    else:
        cols.append(P.VALUES)
    return P.changed(pv, rows=rows, cols=cols)


# ================================================================ Calculated Field
class CalcFieldDialog(QDialog):
    def __init__(self, parent, wb, pv):
        super().__init__(parent)
        self.setWindowTitle("Insert Calculated Field")
        self.wb, self.pv = wb, pv
        self.calc = dict(pv.get("calc") or {})
        v = QVBoxLayout(self)
        g = QGridLayout()
        g.addWidget(QLabel("Name:"), 0, 0)
        self.name = QComboBox()
        self.name.setEditable(True)
        self.name.addItems(list(self.calc))
        n = 1
        while f"Field{n}" in self.calc:
            n += 1
        self.name.setEditText(f"Field{n}")
        g.addWidget(self.name, 0, 1)
        g.addWidget(QLabel("Formula:"), 1, 0)
        self.formula = QLineEdit("= 0")
        g.addWidget(self.formula, 1, 1)
        v.addLayout(g)
        v.addWidget(QLabel("Fields:"))
        self.flist = QListWidget()
        self.flist.addItems(P.source_fields(wb, pv))
        self.flist.itemDoubleClicked.connect(lambda it: self._insert())
        v.addWidget(self.flist)
        h = QHBoxLayout()
        ins = QPushButton("Insert Field")
        ins.clicked.connect(self._insert)
        self.delete = QPushButton("Delete")
        self.delete.clicked.connect(self._delete)
        h.addWidget(ins)
        h.addStretch(1)
        h.addWidget(self.delete)
        v.addLayout(h)
        self.name.currentTextChanged.connect(self._name_changed)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self._ok)
        bb.rejected.connect(self.reject)
        v.addWidget(bb)
        self.result = None   # ("set", name, formula) or ("delete", name)
        self._name_changed(self.name.currentText())
        self.resize(420, 380)

    def _name_changed(self, t):
        if t in self.calc:
            self.formula.setText(self.calc[t])
        self.delete.setEnabled(t in self.calc)

    def _insert(self):
        it = self.flist.currentItem()
        if it is None:
            return
        name = it.text()
        simple = name.replace("_", "").isalnum() and not name[0].isdigit()
        ref = name if simple else f"'{name}'"
        t = self.formula.text()
        if t.strip() in ("= 0", "=0", "="):
            t = "="
        self.formula.setText(t + ref)
        self.formula.setFocus()

    def _delete(self):
        self.result = ("delete", self.name.currentText())
        self.accept()

    def _ok(self):
        err = P.check_calc(self.wb, self.pv, self.name.currentText(), self.formula.text())
        if err:
            _warn(self, err)
            return
        f = self.formula.text().strip()
        self.result = ("set", self.name.currentText().strip(), f if f.startswith("=") else "=" + f)
        self.accept()


def popup_pos(widget, global_pt):
    """Keep a popup on screen."""
    scr = widget.screen().availableGeometry() if widget.screen() else None
    if scr is None:
        return global_pt
    x = min(global_pt.x(), scr.right() - widget.width())
    y = global_pt.y()
    if y + widget.height() > scr.bottom():
        y = max(scr.top(), y - widget.height() - 24)
    return QPoint(max(scr.left(), x), y)
