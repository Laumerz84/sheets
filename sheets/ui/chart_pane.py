"""The Format Chart pane (docked on the right, like Excel's task pane): pick an element - chart area,
title, plot area, legend, axes, a series - and change it. Every change is one undo step."""
from PySide6.QtCore import QTimer, Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (QCheckBox, QColorDialog, QComboBox, QDockWidget, QFormLayout, QFrame,
                               QLabel, QLineEdit, QPushButton, QScrollArea, QSpinBox, QVBoxLayout, QWidget)

from .. import charts as C


class ChartPane(QDockWidget):
    def __init__(self, win, layer):
        super().__init__("Format Chart", win)
        self.setObjectName("ChartPane")
        self.win, self.layer = win, layer
        self.setAllowedAreas(Qt.LeftDockWidgetArea | Qt.RightDockWidgetArea)
        root = QWidget()
        v = QVBoxLayout(root)
        v.setContentsMargins(8, 8, 8, 8)
        self.combo = QComboBox()
        self.combo.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
        self.combo.setMinimumContentsLength(12)
        self.combo.activated.connect(self._element_picked)
        v.addWidget(self.combo)
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.NoFrame)
        self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        v.addWidget(self.scroll, 1)
        self.setWidget(root)
        self.setMinimumWidth(300)
        self._items = []
        self._timer = QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.setInterval(0)
        self._timer.timeout.connect(self.refresh)
        layer.changed.connect(self._timer.start)   # later: a change comes from this pane's own widgets

    # ------------------------------------------------------------ elements
    def _elements(self, ch):
        pie = C.family(ch["type"]) in ("pie", "doughnut")
        out = [("Chart Area", ("chart", None)), ("Chart Title", ("title", None)), ("Plot Area", ("plot", None)),
               ("Legend", ("legend", None))]
        if not pie:
            out.append(("Horizontal (Category) Axis", ("xaxis", None)))
            out.append(("Vertical (Value) Axis", ("yaxis", None)))
            if any(s.get("secondary") for s in ch["series"]):
                out.append(("Secondary Vertical (Value) Axis", ("y2axis", None)))
        for i, s in enumerate(ch["series"]):
            name = s["name"] if not s["name"].startswith("=") else self._name_of(s["name"])
            out.append((f"Series \"{name or i + 1}\"", ("series", i)))
        return out

    def _name_of(self, ref):
        got = C.ref_rect(self.win.wb, ref, self.layer.sheet.name)
        if got is None:
            return ref
        v = got[0].value(got[1][0], got[1][1])
        return "" if v is None else str(v)

    @staticmethod
    def _norm(elem):
        if elem is None:
            return ("chart", None)
        kind, key = elem
        if kind in ("xtitle",):
            return ("xaxis", None)
        if kind == "ytitle":
            return ("yaxis", None)
        if kind == "y2title":
            return ("y2axis", None)
        if kind == "series":
            return ("series", key[0] if isinstance(key, tuple) else key)
        return (kind, None)

    def _element_picked(self, idx):
        if idx < 0 or idx >= len(self._items):
            return
        kind, key = self._items[idx]
        elem = None if kind == "chart" else (kind, (key,) if kind == "series" else None)
        self.layer.select(self.layer.selected, elem)

    # ------------------------------------------------------------ building
    def refresh(self):
        if not self.isVisible():
            return
        ch = self.layer.selected_chart()
        self.combo.blockSignals(True)
        self.combo.clear()
        if ch is None:
            self.combo.blockSignals(False)
            self.combo.setEnabled(False)
            lab = QLabel("Select a chart to format it.\n\nClick a chart on the sheet, or insert one with "
                         "Insert > Chart.")
            lab.setWordWrap(True)
            lab.setMargin(8)
            self._show(lab)
            return
        self.combo.setEnabled(True)
        els = self._elements(ch)
        self._items = [e[1] for e in els]
        cur = self._norm(self.layer.elem)
        for label, e in els:
            self.combo.addItem(label)
        if cur in self._items:
            self.combo.setCurrentIndex(self._items.index(cur))
        else:
            cur = ("chart", None)
            self.combo.setCurrentIndex(0)
        self.combo.blockSignals(False)
        host = QWidget()
        form = QFormLayout(host)
        form.setLabelAlignment(Qt.AlignLeft)
        form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)
        getattr(self, "_page_" + cur[0])(ch, form, cur[1])
        wrap = QWidget()
        wl = QVBoxLayout(wrap)
        wl.setContentsMargins(0, 8, 0, 0)
        wl.addWidget(host)
        wl.addStretch(1)
        self._show(wrap)

    def _show(self, widget):
        old = self.scroll.takeWidget()
        self.scroll.setWidget(widget)
        if old is not None:
            old.hide()
            old.deleteLater()

    # ------------------------------------------------------------ small builders
    def _apply(self, text="Format Chart", **changes):
        cid = self.layer.selected
        if cid:
            self.layer.update(cid, text, **changes)

    def _apply_series(self, i, text="Format Data Series", **changes):
        cid = self.layer.selected

        def fn(c):
            ser = [dict(s) for s in c["series"]]
            if i < len(ser):
                ser[i].update(changes)
            return dict(c, series=ser)
        if cid:
            self.layer.edit(cid, fn, text)

    def _text(self, form, label, value, field, placeholder="", text="Format Chart", post=None):
        e = QLineEdit(value or "")
        e.setPlaceholderText(placeholder)

        def done():
            val = e.text()
            if post:
                val = post(val)
            self._apply(text, **{field: val})
        e.editingFinished.connect(done)
        form.addRow(label, e)
        return e

    def _check(self, form, label, value, field, text="Format Chart", inverse=False):
        c = QCheckBox(label)
        c.setChecked(bool(value))
        c.toggled.connect(lambda on: self._apply(text, **{field: on}))
        form.addRow(c)
        return c

    def _choice(self, form, label, items, value, setter):
        cb = QComboBox()
        cb.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
        cb.setMinimumContentsLength(8)
        for text, data in items:
            cb.addItem(text, data)
        cb.setCurrentIndex(max(0, cb.findData(value)))
        cb.activated.connect(lambda _: setter(cb.currentData()))
        form.addRow(label, cb)
        return cb

    def _number(self, form, label, value, field, text="Format Axis"):
        e = QLineEdit("" if value is None else f"{value:g}")
        e.setPlaceholderText("Auto")

        def done():
            t = e.text().strip()
            if t == "":
                self._apply(text, **{field: None})
                return
            try:
                self._apply(text, **{field: float(t)})
            except ValueError:
                e.setText("" if value is None else f"{value:g}")
        e.editingFinished.connect(done)
        form.addRow(label, e)
        return e

    def _color_button(self, form, label, color, setter):
        b = QPushButton()
        b.setFixedHeight(24)

        def paint(c):
            b.setText(c.upper())
            b.setStyleSheet(f"QPushButton {{ background: {c}; color: {'#FFFFFF' if C.is_dark(c) else '#000000'}; "
                            "border: 1px solid #888; }")
        paint(color)

        def pick():
            col = QColorDialog.getColor(QColor(color), self, "Color")
            if col.isValid():
                setter(col.name().upper())
        b.clicked.connect(pick)
        form.addRow(label, b)
        return b

    def _heading(self, form, text):
        lab = QLabel(f"<b>{text}</b>")
        form.addRow(lab)

    def _button(self, form, text, fn):
        b = QPushButton(text)
        b.clicked.connect(fn)
        form.addRow(b)

    # ------------------------------------------------------------ pages
    def _page_chart(self, ch, form, key):
        lay = self.layer
        self._heading(form, "Chart Area")
        form.addRow("Chart type:", QLabel(C.type_label(ch["type"])))
        self._button(form, "Change Chart Type...", lambda: lay.change_type(ch["id"]))
        self._button(form, "Select Data...", lambda: lay.select_data(ch["id"]))
        self._button(form, "Move Chart...", lambda: lay.move_chart(ch["id"]))
        self._choice(form, "Style:", [(v[0], k) for k, v in C.LOOKS.items()], ch.get("look", 1),
                     lambda k: self._apply("Chart Style", look=k, **C.LOOKS[k][1]))
        self._choice(form, "Colors:", [(v[0], k) for k, v in C.PALETTES.items()], ch.get("style", 1),
                     lambda k: self._apply("Change Colors", style=k, series=[dict(s, color=None) for s in ch["series"]]))
        self._color_button(form, "Fill:", ch.get("fill", "#FFFFFF"), lambda c: self._apply("Chart Fill", fill=c))
        self._check(form, "Border", ch.get("border", True), "border")
        fam = C.family(ch["type"])
        if fam in ("col", "bar") or any((s.get("type") or "") == "col" for s in ch["series"]):
            sp = QSpinBox()
            sp.setRange(0, 500)
            sp.setSuffix(" %")
            sp.setValue(int(ch.get("gap") or (150 if C.grouping(ch["type"]) != "clustered" else 219)))
            sp.editingFinished.connect(lambda: self._apply("Gap Width", gap=sp.value()))
            form.addRow("Gap width:", sp)
        if fam == "doughnut":
            sp = QSpinBox()
            sp.setRange(10, 90)
            sp.setSuffix(" %")
            sp.setValue(int(ch.get("hole", 60)))
            sp.editingFinished.connect(lambda: self._apply("Doughnut Hole Size", hole=sp.value()))
            form.addRow("Hole size:", sp)
        self._check(form, "Data labels", ch.get("labels"), "labels", "Data Labels")
        if fam in ("pie", "doughnut"):
            self._check(form, "Show labels as percentages", ch.get("label_pct"), "label_pct", "Data Labels")

    def _page_title(self, ch, form, key):
        self._heading(form, "Chart Title")
        on = QCheckBox("Show chart title")
        on.setChecked(ch["title"] is not None)
        on.toggled.connect(lambda v: self._apply("Chart Title", title="" if v else None))
        form.addRow(on)
        if ch["title"] is not None:
            self._text(form, "Title text:", ch["title"], "title",
                       "Automatic (series name)" if len(ch["series"]) == 1 else "Chart Title", "Chart Title")
            lab = QLabel("Leave the text empty to use the series name (or 'Chart Title').")
            lab.setWordWrap(True)
            form.addRow(lab)

    def _page_plot(self, ch, form, key):
        self._heading(form, "Plot Area")
        if C.family(ch["type"]) in ("pie", "doughnut"):
            form.addRow(QLabel("The pie has no gridlines or axes."))
            return
        self._check(form, "Horizontal gridlines", ch["grid_y"], "grid_y", "Gridlines")
        self._check(form, "Vertical gridlines", ch["grid_x"], "grid_x", "Gridlines")

    def _page_legend(self, ch, form, key):
        self._heading(form, "Legend")
        self._choice(form, "Position:", [("Right", "right"), ("Top", "top"), ("Left", "left"), ("Bottom", "bottom"),
                                         ("None (hide)", "none")], ch["legend"],
                     lambda k: self._apply("Legend", legend=k))

    def _page_xaxis(self, ch, form, key):
        scatter = C.family(ch["type"]) == "scatter"
        self._heading(form, "Horizontal Axis" if C.family(ch["type"]) != "bar" else "Category Axis")
        self._text(form, "Axis title:", ch["x_title"], "x_title", "None", "Axis Title")
        self._check(form, "Major gridlines", ch["grid_x"], "grid_x", "Gridlines")
        if scatter:
            self._number(form, "Minimum:", ch["x_min"], "x_min")
            self._number(form, "Maximum:", ch["x_max"], "x_max")
        else:
            self._check(form, "Categories in reverse order", ch["x_reverse"], "x_reverse", "Format Axis")

    def _page_yaxis(self, ch, form, key):
        self._heading(form, "Vertical (Value) Axis")
        self._text(form, "Axis title:", ch["y_title"], "y_title", "None", "Axis Title")
        self._number(form, "Minimum:", ch["y_min"], "y_min")
        self._number(form, "Maximum:", ch["y_max"], "y_max")
        self._number(form, "Major unit:", ch["y_major"], "y_major")
        self._check(form, "Major gridlines", ch["grid_y"], "grid_y", "Gridlines")
        self._check(form, "Logarithmic scale", ch["log_y"], "log_y", "Format Axis")
        self._text(form, "Number format:", ch["y_fmt"] or "", "y_fmt", "Linked to source (e.g. #,##0 or 0%)",
                   "Format Axis", post=lambda t: t.strip() or None)

    def _page_y2axis(self, ch, form, key):
        self._heading(form, "Secondary Vertical Axis")
        self._text(form, "Axis title:", ch["y2_title"], "y2_title", "None", "Axis Title")
        self._number(form, "Minimum:", ch["y2_min"], "y2_min")
        self._number(form, "Maximum:", ch["y2_max"], "y2_max")

    def _page_series(self, ch, form, i):
        if i is None or i >= len(ch["series"]):
            return
        s = ch["series"][i]
        self._heading(form, "Series Options")
        name = s["name"] if not s["name"].startswith("=") else s["name"]
        e = QLineEdit(name)
        e.editingFinished.connect(lambda: self._apply_series(i, "Series Name", name=e.text().strip()))
        form.addRow("Series name:", e)
        cur = C.palette_colors(ch.get("style", 1), len(ch["series"]))[i % max(1, len(ch["series"]))] \
            if C.family(ch["type"]) not in ("pie", "doughnut") else C.palette_colors(ch.get("style", 1), 1)[0]
        col = s.get("color") or cur
        self._color_button(form, "Color:", col, lambda c: self._apply_series(i, "Series Color", color=c))
        if s.get("color"):
            self._button(form, "Automatic color", lambda: self._apply_series(i, "Series Color", color=None))
        lab = s.get("labels")
        self._labels(form, ch, i, lab)
        if ch["type"] == "combo":
            self._choice(form, "Chart type:", [("Column", "col"), ("Line", "line"), ("Line with markers", "line_markers"),
                                               ("Area", "area")], s.get("type") or "col",
                         lambda t: self._apply_series(i, "Change Series Type", type=t))
            self._check_series(form, "Secondary axis", ch, i, "secondary")
        fam = C.family(s.get("type") or ch["type"]) if ch["type"] == "combo" else C.family(ch["type"])
        if fam in ("line", "scatter"):
            self._check_series(form, "Smoothed line", ch, i, "smooth")
        form.addRow(QLabel(f"<span style='color:#666'>Values: {s['values'][1:] or '-'}</span>"))
        self._button(form, "Select Data...", lambda: self.layer.select_data(ch["id"]))

    def _labels(self, form, ch, i, lab):
        c = QCheckBox("Data labels")
        c.setChecked(bool(ch["labels"] if lab is None else lab))
        c.toggled.connect(lambda on: self._apply_series(i, "Data Labels", labels=on))
        form.addRow(c)

    def _check_series(self, form, label, ch, i, field):
        c = QCheckBox(label)
        c.setChecked(bool(ch["series"][i].get(field)))
        c.toggled.connect(lambda on: self._apply_series(i, "Format Data Series", **{field: on}))
        form.addRow(c)
