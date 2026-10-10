"""Dialogs for charts: Insert Chart / Change Chart Type (a gallery with previews of the user's own data),
Select Data, Edit Series and Move Chart."""
from PySide6.QtCore import QSize, Qt
from PySide6.QtGui import QIcon, QPixmap
from PySide6.QtWidgets import (QAbstractItemView, QCheckBox, QComboBox, QDialog, QDialogButtonBox, QFormLayout,
                               QGridLayout, QHBoxLayout, QLabel, QLineEdit, QListWidget,
                               QMessageBox, QPushButton, QRadioButton, QScrollArea, QTableWidget, QTableWidgetItem,
                               QToolButton, QVBoxLayout, QWidget)

from .. import charts as C
from . import chart_paint as CP

TILE_W, TILE_H = 196, 124

# a small made-up data set for icons that don't depend on the user's cells
_SAMPLE = {
    "cats": ["A", "B", "C", "D", "E"],
    "ys": [[4.0, 7.0, 5.0, 9.0, 6.0], [3.0, 4.0, 8.0, 5.0, 7.0], [2.0, 3.0, 3.0, 4.0, 5.0]],
}


def _sample_data(n_series):
    series = []
    for i in range(n_series):
        src = dict(C.SERIES_DEFAULTS)
        series.append({"name": f"S{i + 1}", "y": _SAMPLE["ys"][i], "x": [1.0, 2.0, 3.0, 4.0, 5.0], "cats": _SAMPLE["cats"],
                       "fmt": "General", "src": src, "i": i})
    return {"series": series, "cats": _SAMPLE["cats"], "sampled": 1, "xfmt": "General", "error": None}


def _bare(ch):
    """A chart dict without title / legend / axis titles, for small previews."""
    return dict(ch, title=None, legend="none", x_title="", y_title="", y2_title="", border=False)


def sample_chart_image(ctype, w, h):
    n = 1 if ctype in ("pie", "doughnut") else 2 if ctype != "combo" and ctype != "combo_sec" else 2
    ch = C.normalize({"type": ctype, "from": [0, 0, 0, 0], "to": [5, 5, 0, 0],
                      "series": [dict(C.SERIES_DEFAULTS, name=f"S{i}") for i in range(n)]})
    ch = C.apply_type(ch, ctype)
    return CP.render_image(_bare(ch), _sample_data(n), w, h, scale=max(0.3, h / 300.0))


def preview_image(wb, home, data_sheet, rect, by, ctype, w=TILE_W * 2, h=TILE_H * 2, base=None):
    """The user's own data drawn as `ctype`, small, or None when there's nothing to plot."""
    try:
        if base is not None:
            ch = C.apply_type(base, ctype)
        else:
            ch = C.make_chart(data_sheet, rect, ctype, by, anchors=([0, 0, 0, 0], [8, 6, 0, 0]), home=home)
        data = C.resolve(wb, ch, home, limit=200)
        if not data["series"]:
            return None
        return CP.render_image(_bare(ch), data, w, h, scale=1.0)
    except (ValueError, KeyError, IndexError):
        return None


class _Gallery(QDialog):
    """Left: chart type groups. Right: a tile per type showing the chart. Combo: a table of series types."""

    def __init__(self, parent, wb, home, title, series=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(940, 640)
        self.wb, self.home = wb, home
        self.chart_type = "col"
        self._series_names = series or []
        self.v = QVBoxLayout(self)
        self.top = QVBoxLayout()
        self.v.addLayout(self.top)
        body = QHBoxLayout()
        self.v.addLayout(body, 1)
        self.groups = QListWidget()
        self.groups.setFixedWidth(150)
        for g in C.GROUPS:
            self.groups.addItem(g)
        self.groups.currentRowChanged.connect(self._group_changed)
        body.addWidget(self.groups)
        right = QVBoxLayout()
        body.addLayout(right, 1)
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.tiles_host = QWidget()
        self.tiles = QGridLayout(self.tiles_host)
        self.tiles.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        self.scroll.setWidget(self.tiles_host)
        right.addWidget(self.scroll, 3)
        self.combo_box = QWidget()
        cl = QVBoxLayout(self.combo_box)
        cl.setContentsMargins(0, 0, 0, 0)
        cl.addWidget(QLabel("Choose the chart type and axis for each series:"))
        self.combo_table = QTableWidget(0, 3)
        self.combo_table.setHorizontalHeaderLabels(["Series name", "Chart type", "Secondary axis"])
        self.combo_table.horizontalHeader().setStretchLastSection(True)
        self.combo_table.setColumnWidth(0, 220)
        self.combo_table.setColumnWidth(1, 240)
        self.combo_table.setSelectionMode(QAbstractItemView.NoSelection)
        cl.addWidget(self.combo_table)
        right.addWidget(self.combo_box, 2)
        self.msg = QLabel("")
        self.msg.setWordWrap(True)
        self.v.addWidget(self.msg)
        self.bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.bb.accepted.connect(self._ok)
        self.bb.rejected.connect(self.reject)
        self.v.addWidget(self.bb)
        self.tile_buttons = {}
        self.combo_box.hide()

    # subclasses: preview(ctype) -> QImage | None
    def preview(self, ctype):
        return None

    def select_type(self, ctype):
        self.chart_type = ctype
        g = C.TYPES[ctype][1]
        self.groups.blockSignals(True)
        self.groups.setCurrentRow(C.GROUPS.index(g))
        self.groups.blockSignals(False)
        self._fill_tiles(g)
        if ctype in self.tile_buttons:
            self.tile_buttons[ctype].setChecked(True)

    def _group_changed(self, row):
        if row < 0:
            return
        g = C.GROUPS[row]
        self._fill_tiles(g)
        first = next(t for t, (_, grp) in C.TYPES.items() if grp == g)
        if C.TYPES[self.chart_type][1] != g:
            self.chart_type = first
        if self.chart_type in self.tile_buttons:
            self.tile_buttons[self.chart_type].setChecked(True)
        self._combo_visibility()

    def refresh_previews(self):
        self._fill_tiles(C.TYPES[self.chart_type][1])
        if self.chart_type in self.tile_buttons:
            self.tile_buttons[self.chart_type].setChecked(True)

    def _fill_tiles(self, group):
        while self.tiles.count():
            it = self.tiles.takeAt(0)
            wd = it.widget()
            if wd is not None:
                wd.hide()
                wd.setParent(None)
                wd.deleteLater()
        self.tile_buttons = {}
        n = 0
        for t, (label, grp) in C.TYPES.items():
            if grp != group:
                continue
            b = QToolButton()
            b.setCheckable(True)
            b.setAutoExclusive(True)
            b.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)
            b.setText(label)
            b.setFixedSize(TILE_W + 24, TILE_H + 52)
            b.setIconSize(QSize(TILE_W, TILE_H))
            img = self.preview(t) or sample_chart_image(t, TILE_W * 2, TILE_H * 2)
            pm = QPixmap.fromImage(img)
            pm.setDevicePixelRatio(2)
            b.setIcon(QIcon(pm))
            b.clicked.connect(lambda _=False, tt=t: self._tile_clicked(tt))
            b.setStyleSheet("QToolButton { background: #FFFFFF; border: 1px solid #D0D0D0; padding: 4px; }"
                            "QToolButton:checked { border: 2px solid #217346; background: #E8F3EC; }")
            self.tiles.addWidget(b, n // 3, n % 3)
            self.tile_buttons[t] = b
            n += 1

    def _tile_clicked(self, t):
        self.chart_type = t
        self._combo_visibility()
        if t in ("combo", "combo_sec"):
            self._fill_combo(preset=t)

    def _combo_visibility(self):
        is_combo = C.TYPES[self.chart_type][1] == "Combo"
        self.combo_box.setVisible(is_combo)
        if is_combo and self.combo_table.rowCount() == 0:
            self._fill_combo(preset=self.chart_type)

    def series_for_combo(self):
        return self._series_names

    def combo_defaults(self):
        return None

    def _fill_combo(self, preset="combo"):
        names = self.series_for_combo()
        self.combo_table.setRowCount(len(names))
        prev = self.combo_defaults()
        for i, name in enumerate(names):
            it = QTableWidgetItem(name)
            it.setFlags(Qt.ItemIsEnabled)
            self.combo_table.setItem(i, 0, it)
            cb = QComboBox()
            for t in C.SERIES_TYPES:
                cb.addItem(C.type_label(t) if t in C.TYPES else {"scatter_lines": "Scatter with Lines"}.get(t, t), t)
            default = "col" if i == 0 else "line"
            if preset not in ("combo", "combo_sec") and prev and i < len(prev) and prev[i][0]:
                default = prev[i][0]
            cb.setCurrentIndex(max(0, cb.findData(default)))
            self.combo_table.setCellWidget(i, 1, cb)
            ck = QCheckBox()
            ck.setChecked(preset == "combo_sec" and i > 0)
            w = QWidget()
            hl = QHBoxLayout(w)
            hl.setContentsMargins(8, 0, 0, 0)
            hl.addWidget(ck)
            self.combo_table.setCellWidget(i, 2, w)

    def combo_result(self):
        out = []
        for i in range(self.combo_table.rowCount()):
            cb = self.combo_table.cellWidget(i, 1)
            ck = self.combo_table.cellWidget(i, 2).findChild(QCheckBox)
            out.append((cb.currentData(), ck.isChecked()))
        return out

    def _ok(self):
        self.accept()


# ================================================================ Insert Chart
class InsertChartDialog(_Gallery):
    def __init__(self, parent, sheet, rect, ctype=None):
        super().__init__(parent, parent.wb, sheet, "Insert Chart")
        self.sheet = sheet
        self.data_sheet = sheet
        self.rect = rect
        self.by = None
        self.combo = None
        self.preset_series = None
        form = QHBoxLayout()
        form.addWidget(QLabel("Data range:"))
        self.range_edit = QLineEdit(C.make_ref(sheet.name, rect)[1:])
        self.range_edit.editingFinished.connect(self._range_changed)
        form.addWidget(self.range_edit, 1)
        form.addWidget(QLabel("Series in:"))
        self.by_box = QComboBox()
        self.by_box.addItem("Automatic", None)
        self.by_box.addItem("Columns", "cols")
        self.by_box.addItem("Rows", "rows")
        self.by_box.currentIndexChanged.connect(self._range_changed)
        form.addWidget(self.by_box)
        self.top.addLayout(form)
        self._names = []
        self._load_range()
        self.select_type(ctype if ctype in C.TYPES else "col")
        self.refresh_previews()
        self._combo_visibility()

    def _parse_range(self):
        text = self.range_edit.text().strip()
        got = C.ref_rect(self.wb, text if text.startswith("=") else "=" + text, self.sheet.name)
        return got

    def _load_range(self):
        got = self._parse_range()
        if got is None:
            self.msg.setText("That isn't a valid range. Type something like Sheet1!A1:D10.")
            self._names = []
            return False
        self.msg.setText("")
        self.data_sheet, rect = got
        self.rect = C.clamp_rect(self.data_sheet, rect)
        self.by = self.by_box.currentData()
        try:
            probe = C.make_chart(self.data_sheet, self.rect, "col", self.by, anchors=([0, 0, 0, 0], [5, 5, 0, 0]),
                                 home=self.sheet)
            self._names = [s["name"] if not s["name"].startswith("=") else self._name_of(s["name"])
                           for s in probe["series"]]
        except ValueError as e:
            self.msg.setText(str(e))
            self._names = []
            return False
        return True

    def _name_of(self, ref):
        got = C.ref_rect(self.wb, ref, self.sheet.name)
        if got is None:
            return ref
        v = got[0].value(got[1][0], got[1][1])
        return "" if v is None else str(v)

    def _range_changed(self, *_):
        ok = self._load_range()
        if ok:
            self.combo_table.setRowCount(0)
            self.refresh_previews()
            self._combo_visibility()

    def series_for_combo(self):
        return self._names

    def preview(self, ctype):
        if not self._names:
            return None
        return preview_image(self.wb, self.sheet, self.data_sheet, self.rect, self.by, ctype)

    def _ok(self):
        if not self._load_range():
            return
        self.combo = self.combo_result() if C.TYPES[self.chart_type][1] == "Combo" else None
        self.accept()


# ================================================================ Change Chart Type
class ChangeTypeDialog(_Gallery):
    def __init__(self, parent, wb, sheet, ch):
        self.base = ch
        names = []
        for s in ch["series"]:
            n = s["name"]
            if n.startswith("="):
                got = C.ref_rect(wb, n, sheet.name)
                v = got[0].value(got[1][0], got[1][1]) if got else None
                n = "" if v is None else str(v)
            names.append(n)
        super().__init__(parent, wb, sheet, "Change Chart Type", names)
        self.combo = None
        t = "combo" if ch["type"] == "combo" else ch["type"]
        self.select_type(t)
        self._combo_visibility()
        if ch["type"] == "combo":
            self._fill_combo(preset="custom")

    def combo_defaults(self):
        return [(s.get("type"), s.get("secondary")) for s in self.base["series"]]

    def preview(self, ctype):
        return preview_image(self.wb, self.home, self.home, None, None, ctype, base=self.base)

    def _ok(self):
        self.combo = self.combo_result() if C.TYPES[self.chart_type][1] == "Combo" else None
        self.accept()

    def apply(self, ch):
        t = self.chart_type
        if self.combo:
            out = C.apply_type(ch, "combo" if t == "combo_sec" else "combo")
            ser = []
            for s, (st, sec) in zip(out["series"], self.combo):
                ser.append(dict(s, type=st, secondary=sec))
            out = dict(out, series=ser)
            return out
        return C.apply_type(ch, t)


# ================================================================ Select Data
class SeriesDialog(QDialog):
    def __init__(self, parent, wb, home, series, scatter=False):
        super().__init__(parent)
        self.setWindowTitle("Edit Series")
        self.wb, self.home = wb, home
        form = QFormLayout(self)
        self.name = QLineEdit(series.get("name", ""))
        self.name.setToolTip("Text, or a cell like =Sheet1!$B$1")
        self.vals = QLineEdit(series.get("values", ""))
        self.vals.setToolTip("A range like =Sheet1!$B$2:$B$9")
        self.cats = QLineEdit(series.get("cats", ""))
        form.addRow("Series name:", self.name)
        form.addRow("Series values:", self.vals)
        form.addRow("X values:" if scatter else "Category labels:", self.cats)
        self.series = dict(series)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self._ok)
        bb.rejected.connect(self.reject)
        form.addRow(bb)

    def _ref_ok(self, text, what):
        if text.strip() == "":
            return True
        t = text.strip()
        got = C.ref_rect(self.wb, t if t.startswith("=") else "=" + t, self.home.name)
        if got is None:
            QMessageBox.warning(self, self.windowTitle(), f"{what} isn't a valid range. Use something like "
                                                          "=Sheet1!$B$2:$B$9.")
            return False
        return True

    def _fix(self, text):
        t = text.strip()
        if not t:
            return ""
        t = t if t.startswith("=") else "=" + t
        got = C.ref_rect(self.wb, t, self.home.name)
        return C.make_ref(got[0].name, C.clamp_rect(got[0], got[1])) if got else t

    def _ok(self):
        name = self.name.text().strip()
        if name.startswith("=") and not self._ref_ok(name, "Series name"):
            return
        if not self.vals.text().strip() or not self._ref_ok(self.vals.text(), "Series values"):
            if not self.vals.text().strip():
                QMessageBox.warning(self, self.windowTitle(), "A series needs values.")
            return
        if not self._ref_ok(self.cats.text(), "The labels range"):
            return
        self.series.update(name=self._fix(name) if name.startswith("=") else name, values=self._fix(self.vals.text()),
                           cats=self._fix(self.cats.text()))
        self.accept()


class SelectDataDialog(QDialog):
    def __init__(self, parent, wb, home, ch):
        super().__init__(parent)
        self.setWindowTitle("Select Data Source")
        self.resize(640, 480)
        self.wb, self.home = wb, home
        self.ch = dict(ch, series=[dict(s) for s in ch["series"]])
        self.scatter = ch["type"].startswith("scatter")
        v = QVBoxLayout(self)
        row = QHBoxLayout()
        row.addWidget(QLabel("Chart data range:"))
        self.range_edit = QLineEdit(ch.get("range", "")[1:] if ch.get("range") else "")
        self.range_edit.setPlaceholderText("(series were set one by one)")
        self.range_edit.editingFinished.connect(self._range_edited)
        row.addWidget(self.range_edit, 1)
        v.addLayout(row)
        self.switch_btn = QPushButton("Switch Row/Column")
        self.switch_btn.clicked.connect(self._switch)
        self.switch_btn.setEnabled(bool(ch.get("range")))
        v.addWidget(self.switch_btn, 0, Qt.AlignLeft)
        v.addWidget(QLabel("Legend Entries (Series):"))
        h = QHBoxLayout()
        self.list = QListWidget()
        self.list.itemDoubleClicked.connect(lambda _: self._edit())
        h.addWidget(self.list, 1)
        col = QVBoxLayout()
        for text, fn in (("Add...", self._add), ("Edit...", self._edit), ("Remove", self._remove),
                         ("Move Up", lambda: self._move(-1)), ("Move Down", lambda: self._move(1))):
            b = QPushButton(text)
            b.clicked.connect(fn)
            col.addWidget(b)
        col.addStretch(1)
        h.addLayout(col)
        v.addLayout(h, 1)
        row = QHBoxLayout()
        self.cat_label = QLabel()
        row.addWidget(QLabel("X values:" if self.scatter else "Horizontal (Category) Axis Labels:"))
        row.addWidget(self.cat_label, 1)
        b = QPushButton("Edit...")
        b.clicked.connect(self._edit_cats)
        row.addWidget(b)
        v.addLayout(row)
        self.msg = QLabel("")
        self.msg.setWordWrap(True)
        v.addWidget(self.msg)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self._ok)
        bb.rejected.connect(self.reject)
        v.addWidget(bb)
        self.result_chart = None
        self._fill()

    def _series_label(self, s):
        n = s["name"]
        if n.startswith("="):
            got = C.ref_rect(self.wb, n, self.home.name)
            if got:
                val = got[0].value(got[1][0], got[1][1])
                n = f"{val}" if val is not None else ""
        return f"{n or '(no name)'}    {s['values']}"

    def _fill(self, row=None):
        self.list.clear()
        for s in self.ch["series"]:
            self.list.addItem(self._series_label(s))
        if self.ch["series"]:
            self.list.setCurrentRow(row if row is not None and row < len(self.ch["series"]) else 0)
        cats = self.ch["series"][0]["cats"] if self.ch["series"] else ""
        self.cat_label.setText(cats[1:] if cats else "(1, 2, 3, ...)")

    def _range_edited(self):
        text = self.range_edit.text().strip()
        if not text:
            return
        got = C.ref_rect(self.wb, text if text.startswith("=") else "=" + text, self.home.name)
        if got is None:
            self.msg.setText("That isn't a valid range.")
            return
        try:
            self.ch = C.rebuild_from_range(got[0], self.ch, got[1], None)
            self.msg.setText("")
            self.switch_btn.setEnabled(True)
        except ValueError as e:
            self.msg.setText(str(e))
            return
        self._fill()

    def _switch(self):
        text = self.range_edit.text().strip()
        got = C.ref_rect(self.wb, text if text.startswith("=") else "=" + text, self.home.name)
        if got is None:
            return
        by = "rows" if self.ch.get("by") == "cols" else "cols"
        try:
            self.ch = C.rebuild_from_range(got[0], self.ch, got[1], by)
            self.msg.setText("")
        except ValueError as e:
            self.msg.setText(str(e))
            return
        self._fill()

    def _add(self):
        base = dict(C.SERIES_DEFAULTS)
        if self.ch["series"]:
            base["cats"] = self.ch["series"][0]["cats"]
        base["name"] = f"Series{len(self.ch['series']) + 1}"
        d = SeriesDialog(self, self.wb, self.home, base, self.scatter)
        if d.exec() == QDialog.Accepted:
            self.ch["series"].append(d.series)
            self.ch["range"] = ""
            self.range_edit.setText("")
            self.switch_btn.setEnabled(False)
            self._fill(len(self.ch["series"]) - 1)

    def _edit(self):
        i = self.list.currentRow()
        if i < 0:
            return
        d = SeriesDialog(self, self.wb, self.home, self.ch["series"][i], self.scatter)
        if d.exec() == QDialog.Accepted:
            self.ch["series"][i] = d.series
            self.ch["range"] = ""
            self.range_edit.setText("")
            self.switch_btn.setEnabled(False)
            self._fill(i)

    def _remove(self):
        i = self.list.currentRow()
        if i < 0:
            return
        del self.ch["series"][i]
        self.ch["range"] = ""
        self.range_edit.setText("")
        self.switch_btn.setEnabled(False)
        self._fill(max(0, i - 1))

    def _move(self, d):
        i = self.list.currentRow()
        j = i + d
        if i < 0 or not 0 <= j < len(self.ch["series"]):
            return
        s = self.ch["series"]
        s[i], s[j] = s[j], s[i]
        self._fill(j)

    def _edit_cats(self):
        if not self.ch["series"]:
            return
        cur = self.ch["series"][0]["cats"]
        d = SeriesDialog(self, self.wb, self.home, {"name": "x", "values": self.ch["series"][0]["values"], "cats": cur},
                         self.scatter)
        d.name.setEnabled(False)
        d.vals.setEnabled(False)
        if d.exec() == QDialog.Accepted:
            for s in self.ch["series"]:
                s["cats"] = d.series["cats"]
            self._fill(self.list.currentRow())

    def _ok(self):
        if not self.ch["series"]:
            QMessageBox.warning(self, self.windowTitle(), "Add at least one series.")
            return
        self.result_chart = self.ch
        self.accept()


# ================================================================ Move Chart
class MoveChartDialog(QDialog):
    def __init__(self, parent, wb, sheet):
        super().__init__(parent)
        self.setWindowTitle("Move Chart")
        self.wb, self.sheet = wb, sheet
        v = QVBoxLayout(self)
        v.addWidget(QLabel("Choose where you want the chart to be placed:"))
        self.r_new = QRadioButton("New sheet")
        self.r_obj = QRadioButton("Object in:")
        self.r_obj.setChecked(True)
        v.addWidget(self.r_new)
        h = QHBoxLayout()
        h.addWidget(self.r_obj)
        self.sheets = QComboBox()
        for s in wb.sheets:
            self.sheets.addItem(s.name, s)
        self.sheets.setCurrentIndex(wb.sheets.index(sheet))
        h.addWidget(self.sheets, 1)
        v.addLayout(h)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self._ok)
        bb.rejected.connect(self.reject)
        v.addWidget(bb)
        self.target_new = False
        self.target_sheet = sheet

    def _ok(self):
        self.target_new = self.r_new.isChecked()
        self.target_sheet = self.sheets.currentData()
        self.accept()
