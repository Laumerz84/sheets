"""PivotTable UI: Insert > PivotTable dialog, the 'PivotTable Fields' panel (drag fields into
Filters / Columns / Rows / Values, or right-click them), value settings and filter lists."""
from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (QAbstractItemView, QButtonGroup, QCheckBox, QDialog, QDialogButtonBox,
                               QDockWidget, QGridLayout, QHBoxLayout, QLabel, QLineEdit, QListWidget,
                               QListWidgetItem, QMenu, QPushButton, QRadioButton, QVBoxLayout, QWidget)

from .. import pivot as P
from ..refs import parse_range, range_addr

AREAS = [("filter_fields", "Filters"), ("cols", "Columns"), ("rows", "Rows"), ("values", "Values")]


# ================================================================ Insert > PivotTable
class CreatePivotDialog(QDialog):
    def __init__(self, win, sheet_name, rect):
        super().__init__(win)
        self.win = win
        self.setWindowTitle("Create PivotTable")
        v = QVBoxLayout(self)
        v.addWidget(QLabel("Choose the data that you want to analyze (the first row holds the headers):"))
        g = QGridLayout()
        g.addWidget(QLabel("Table/Range:"), 0, 0)
        self.src = QLineEdit(f"{_quote(sheet_name)}!{_abs(range_addr(*rect))}")
        g.addWidget(self.src, 0, 1)
        v.addLayout(g)
        v.addWidget(QLabel("Choose where you want the PivotTable to be placed:"))
        self.new_sheet = QRadioButton("New Worksheet")
        self.existing = QRadioButton("Existing Worksheet")
        self.new_sheet.setChecked(True)
        grp = QButtonGroup(self)
        grp.addButton(self.new_sheet)
        grp.addButton(self.existing)
        v.addWidget(self.new_sheet)
        h = QHBoxLayout()
        h.addWidget(self.existing)
        h.addWidget(QLabel("Location:"))
        self.loc = QLineEdit()
        self.loc.setPlaceholderText("e.g. H3")
        self.loc.textEdited.connect(lambda _: self.existing.setChecked(True))
        h.addWidget(self.loc, 1)
        v.addLayout(h)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self._ok)
        bb.rejected.connect(self.reject)
        v.addWidget(bb)
        self.result = None

    def _ok(self):
        sheet, rect = _parse_ref(self.src.text(), self.win.sheet.name)
        if sheet is None or self.win.wb.get_sheet(sheet) is None or rect is None:
            self.win.warn("Type the data range like Sheet1!$A$1:$D$100.")
            return
        if rect[2] <= rect[0]:
            self.win.warn("The range needs a header row and at least one row of data.")
            return
        dest = None
        if self.existing.isChecked():
            ds, dr = _parse_ref(self.loc.text(), self.win.sheet.name)
            if ds is None or self.win.wb.get_sheet(ds) is None or dr is None:
                self.win.warn("Type where the PivotTable should go, like H3 or Sheet2!A3.")
                return
            dest = (ds, dr[0], dr[1])
        self.result = (sheet, rect, dest)
        self.accept()


def _quote(name):
    return f"'{name}'" if any(ch in name for ch in " -!'") else name


def _abs(a):
    out = []
    for part in a.split(":"):
        col = "".join(ch for ch in part if ch.isalpha())
        row = part[len(col):]
        out.append(f"${col}${row}")
    return ":".join(out)


def _parse_ref(text, default_sheet):
    t = text.strip().lstrip("=")
    sheet = default_sheet
    if "!" in t:
        sheet, t = t.rsplit("!", 1)
        sheet = sheet.strip().strip("'")
    b = parse_range(t.replace("$", ""))
    return (sheet, tuple(b)) if b else (None, None)


# ================================================================ filter list
class FilterDialog(QDialog):
    def __init__(self, parent, field, items, allowed):
        super().__init__(parent)
        self.setWindowTitle(f"Filter: {field}")
        v = QVBoxLayout(self)
        self.all = QCheckBox("(Select All)")
        v.addWidget(self.all)
        self.list = QListWidget()
        for t in items:
            it = QListWidgetItem(t)
            it.setFlags(it.flags() | Qt.ItemIsUserCheckable)
            it.setCheckState(Qt.Checked if allowed is None or t in allowed else Qt.Unchecked)
            self.list.addItem(it)
        v.addWidget(self.list, 1)
        self.all.setChecked(allowed is None)
        self.all.clicked.connect(self._all)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        v.addWidget(bb)
        self.resize(280, 360)

    def _all(self, on):
        for i in range(self.list.count()):
            self.list.item(i).setCheckState(Qt.Checked if on else Qt.Unchecked)

    def allowed(self):
        """None = everything, else the ticked labels."""
        ticked = [self.list.item(i).text() for i in range(self.list.count())
                  if self.list.item(i).checkState() == Qt.Checked]
        return None if len(ticked) == self.list.count() else ticked


# ================================================================ the field list panel
class _AreaList(QListWidget):
    """One of Filters / Columns / Rows / Values: accepts fields dragged from the field list or
    another area, and reorders by dragging within itself."""
    dropped = Signal(str, str, str, int)  # field, from area ('' = field list), to area, index

    def __init__(self, area, panel):
        super().__init__()
        self.area, self.panel = area, panel
        self.setDragDropMode(QAbstractItemView.DragDrop)
        self.setDefaultDropAction(Qt.MoveAction)
        self.setAcceptDrops(True)
        self.setMinimumHeight(70)
        self.setContextMenuPolicy(Qt.CustomContextMenu)

    def dropEvent(self, e):
        src = e.source()
        item = src.currentItem() if isinstance(src, QListWidget) else None
        if item is None:
            e.ignore()
            return
        field = item.data(Qt.UserRole)
        target = self.itemAt(e.position().toPoint())
        index = self.row(target) if target is not None else self.count()
        from_area = src.area if isinstance(src, _AreaList) else ""
        e.setDropAction(Qt.IgnoreAction)  # the panel rebuilds the lists itself
        e.accept()
        self.dropped.emit(field, from_area, self.area, index)


class PivotPanel(QDockWidget):
    """'PivotTable Fields': shown while the active cell is inside a PivotTable."""

    def __init__(self, win):
        super().__init__("PivotTable Fields", win)
        self.win = win
        self.setObjectName("PivotPanel")
        self.sheet = None
        self.pv = None
        body = QWidget()
        v = QVBoxLayout(body)
        v.setContentsMargins(8, 6, 8, 8)
        v.addWidget(QLabel("Choose fields to add to report:"))
        self.fields = QListWidget()
        self.fields.setDragEnabled(True)
        self.fields.setDragDropMode(QAbstractItemView.DragOnly)
        self.fields.setContextMenuPolicy(Qt.CustomContextMenu)
        self.fields.customContextMenuRequested.connect(self._field_menu)
        self.fields.itemChanged.connect(self._field_ticked)
        v.addWidget(self.fields, 2)
        v.addWidget(QLabel("Drag fields between areas below:"))
        grid = QGridLayout()
        self.areas = {}
        for i, (area, label) in enumerate(AREAS):
            box = QVBoxLayout()
            box.addWidget(QLabel(label))
            lst = _AreaList(area, self)
            lst.dropped.connect(self._dropped)
            lst.customContextMenuRequested.connect(lambda pos, l=lst: self._area_menu(l, pos))
            lst.itemDoubleClicked.connect(lambda it, l=lst: self._area_double(l, it))
            self.areas[area] = lst
            box.addWidget(lst)
            grid.addLayout(box, i // 2, i % 2)
        v.addLayout(grid, 3)
        h = QHBoxLayout()
        refresh = QPushButton("Refresh")
        refresh.setToolTip("Re-read the source data (Ctrl+Alt+F5 refreshes every PivotTable)")
        refresh.clicked.connect(lambda: self.win.refresh_pivot(self.sheet, self.pv))
        h.addStretch(1)
        h.addWidget(refresh)
        v.addLayout(h)
        self.setWidget(body)
        self.setMinimumWidth(270)

    # ---- showing a pivot
    def show_pivot(self, sheet, pv):
        self.sheet, self.pv = sheet, pv
        names = P.fields(self.win.wb, pv)
        used = set(pv["rows"]) | set(pv["cols"]) | set(pv["filter_fields"]) | {f for f, _ in pv["values"]}
        self.fields.blockSignals(True)
        self.fields.clear()
        for n in names:
            it = QListWidgetItem(n)
            it.setData(Qt.UserRole, n)
            it.setFlags(it.flags() | Qt.ItemIsUserCheckable | Qt.ItemIsDragEnabled)
            it.setCheckState(Qt.Checked if n in used else Qt.Unchecked)
            self.fields.addItem(it)
        self.fields.blockSignals(False)
        for area, lst in self.areas.items():
            lst.clear()
            for i, entry in enumerate(pv[area]):
                if area == "values":
                    f, agg = entry
                    it = QListWidgetItem(P.value_label(f, agg))
                    it.setData(Qt.UserRole + 1, i)
                else:
                    f = entry
                    text = f
                    if area == "filter_fields" and pv["filters"].get(f) is not None:
                        text += "  (filtered)"
                    it = QListWidgetItem(text)
                f_item = f
                it.setData(Qt.UserRole, f_item)
                lst.addItem(it)
        self.show()

    # ---- changes
    def _commit(self, new, text="Change PivotTable"):
        self.win.change_pivot(self.sheet, self.pv, new, text)

    def _without(self, pv, field, area):
        """pv with `field` taken out of `area` (all of its Values entries when area is 'values')."""
        if area == "values":
            return P.changed(pv, values=[x for x in pv["values"] if x[0] != field])
        out = P.changed(pv, **{area: [f for f in pv[area] if f != field]})
        if area == "filter_fields":
            out["filters"] = {k: v for k, v in pv["filters"].items() if k != field}
        return out

    def _add(self, pv, field, area, index=None):
        if area == "values":
            vals = list(pv["values"])
            entry = [field, P.default_agg(self.win.wb, pv, field)]
            vals.insert(len(vals) if index is None else index, entry)
            return P.changed(pv, values=vals)
        pv2 = pv
        for other in ("rows", "cols", "filter_fields"):  # a field is in one of these at a time
            if field in pv2[other]:
                pv2 = self._without(pv2, field, other)
        lst = list(pv2[area])
        lst.insert(len(lst) if index is None else min(index, len(lst)), field)
        return P.changed(pv2, **{area: lst})

    def _dropped(self, field, from_area, to_area, index):
        pv = self.pv
        if from_area == to_area and to_area != "values":
            lst = [f for f in pv[to_area] if f != field]
            old = pv[to_area].index(field)
            lst.insert(index - (1 if index > old else 0), field)
            self._commit(P.changed(pv, **{to_area: lst}))
            return
        if from_area == to_area == "values":
            src = self.areas["values"].currentItem()
            i = src.data(Qt.UserRole + 1)
            vals = list(pv["values"])
            entry = vals.pop(i)
            vals.insert(index - (1 if index > i else 0), entry)
            self._commit(P.changed(pv, values=vals))
            return
        if from_area == "values":
            src = self.areas["values"].currentItem()
            vals = list(pv["values"])
            vals.pop(src.data(Qt.UserRole + 1))
            pv = P.changed(pv, values=vals)
        elif from_area:
            pv = self._without(pv, field, from_area)
        self._commit(self._add(pv, field, to_area, index))

    def _field_ticked(self, it):
        field = it.data(Qt.UserRole)
        pv = self.pv
        if it.checkState() == Qt.Checked:
            # like Excel: numbers go to Values, everything else to Rows
            area = "values" if P.default_agg(self.win.wb, pv, field) == "sum" else "rows"
            self._commit(self._add(pv, field, area))
        else:
            for area in ("rows", "cols", "filter_fields", "values"):
                pv = self._without(pv, field, area)
            self._commit(pv)

    def _field_menu(self, pos):
        it = self.fields.itemAt(pos)
        if it is None:
            return
        field = it.data(Qt.UserRole)
        m = QMenu(self)
        for area, label in AREAS:
            m.addAction(f"Add to {label}", lambda a=area: self._dropped(field, self._area_of(field, a), a, 10 ** 6))
        m.exec(self.fields.mapToGlobal(pos))

    def _area_of(self, field, target):
        if target == "values":
            return ""
        for area in ("rows", "cols", "filter_fields"):
            if field in self.pv[area]:
                return area
        return ""

    def _area_menu(self, lst, pos):
        it = lst.itemAt(pos)
        if it is None:
            return
        lst.setCurrentItem(it)
        field, area = it.data(Qt.UserRole), lst.area
        m = QMenu(self)
        n = lst.count()
        i = lst.row(it)
        if i > 0:
            m.addAction("Move Up", lambda: self._dropped(field, area, area, i - 1))
        if i < n - 1:
            m.addAction("Move Down", lambda: self._dropped(field, area, area, i + 2))
        m.addSeparator()
        for other, label in AREAS:
            if other != area:
                m.addAction(f"Move to {label}", lambda o=other: self._dropped(field, area, o, 10 ** 6))
        m.addSeparator()
        if area == "values":
            sm = m.addMenu("Summarize Values By")
            vi = it.data(Qt.UserRole + 1)
            for agg, label in P.AGGS:
                a = sm.addAction(label, lambda g=agg: self._set_agg(vi, g))
                a.setCheckable(True)
                a.setChecked(self.pv["values"][vi][1] == agg)
        if area in ("filter_fields", "rows", "cols"):
            m.addAction("Filter...", lambda: self._filter(field))
        m.addAction("Remove Field", lambda: self._remove(area, it))
        m.exec(lst.mapToGlobal(pos))

    def _area_double(self, lst, it):
        if lst.area == "values":
            menu = QMenu(self)
            vi = it.data(Qt.UserRole + 1)
            for agg, label in P.AGGS:
                a = menu.addAction(label, lambda g=agg: self._set_agg(vi, g))
                a.setCheckable(True)
                a.setChecked(self.pv["values"][vi][1] == agg)
            menu.exec(lst.mapToGlobal(lst.visualItemRect(it).bottomLeft()))
        else:
            self._filter(it.data(Qt.UserRole))

    def _set_agg(self, vi, agg):
        vals = [list(x) for x in self.pv["values"]]
        vals[vi][1] = agg
        self._commit(P.changed(self.pv, values=vals))

    def _remove(self, area, it):
        if area == "values":
            vals = list(self.pv["values"])
            vals.pop(it.data(Qt.UserRole + 1))
            self._commit(P.changed(self.pv, values=vals))
        else:
            self._commit(self._without(self.pv, it.data(Qt.UserRole), area))

    def _filter(self, field):
        items = P.items(self.win.wb, self.pv, field)
        d = FilterDialog(self, field, items, self.pv["filters"].get(field))
        if d.exec() != QDialog.Accepted:
            return
        filters = dict(self.pv["filters"])
        allowed = d.allowed()
        if allowed is None:
            filters.pop(field, None)
        else:
            filters[field] = allowed
        pv = P.changed(self.pv, filters=filters)
        if field not in pv["filter_fields"] and field not in pv["rows"] and field not in pv["cols"]:
            pv = P.changed(pv, filter_fields=pv["filter_fields"] + [field])
        self._commit(pv, "Filter PivotTable")
