"""PivotTable UI: Insert > PivotTable dialog, the 'PivotTable Fields' panel (search, drag fields into
Filters / Columns / Rows / Values, right-click them, Defer Layout Update), the in-sheet buttons (filter
dropdowns and +/- expand/collapse, drawn by an overlay on the grid), the pivot right-click menu,
double-click drill-down (Show Details) and the Design options. Dialogs live in pivot_dialogs.py."""
from PySide6.QtCore import QEvent, QPoint, QRect, Qt, Signal
from PySide6.QtGui import QColor, QPainter, QPen
from PySide6.QtWidgets import (QAbstractItemView, QButtonGroup, QCheckBox, QDialog, QDialogButtonBox,
                               QDockWidget, QGridLayout, QHBoxLayout, QLabel, QLineEdit, QListWidget,
                               QListWidgetItem, QMenu, QPushButton, QRadioButton, QToolButton, QVBoxLayout,
                               QWidget)

from .. import pivot as P
from ..refs import key, parse_range, range_addr
from ..workbook import DEFAULT_COL_WIDTH, DEFAULT_STYLE
from . import pivot_dialogs as PD

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


# ================================================================ column widths
def autofit_pivot(win, pv):
    """Excel autofits a pivot's columns on every update; also leave room for indents (compact rows)
    and the dropdown buttons of header cells. Runs inside change_pivot's undo macro."""
    o = pv["out"]
    win.autofit_cols(range(o[1], o[3] + 1))
    sh = win.sheet
    flt = {(pv["anchor"][0] + h[1], pv["anchor"][1] + h[2]) for h in pv.get("hot") or () if h[0] == "flt"}
    need = {}
    for r in range(o[0], o[2] + 1):
        for c in range(o[1], o[3] + 1):
            k = key(r, c)
            st = sh.styles.get(k)
            has_btn = (r, c) in flt
            if not has_btn and (st is None or not st.indent):
                continue
            w = 10.0
            if k in sh.values:
                w = win._text_width(k, st or DEFAULT_STYLE) + 10
            w += (st.indent * 9 if st is not None else 0) + (20 if has_btn else 0)
            need[c] = max(need.get(c, 0), w)
    new = dict(sh.col_widths)
    for c, w in need.items():
        if w > new.get(c, DEFAULT_COL_WIDTH):
            new[c] = int(min(1200, w + 1))
    if new != sh.col_widths:
        win._push_meta({"col_widths": (dict(sh.col_widths), new)}, "PivotTable")


# ================================================================ actions
def _relabel(old, new):
    """Sorts and value filters name their value field by its label: keep them pointing at the same
    value field when it's renamed or summarized differently (same position in Values)."""
    if len(old["values"]) != len(new["values"]) or [e[0] for e in old["values"]] != [e[0] for e in new["values"]]:
        return new
    a, b = P.value_labels(old), P.value_labels(new)
    m = {x: y for x, y in zip(a, b) if x != y}
    if not m:
        return new
    out = dict(new)
    if new.get("sort"):
        out["sort"] = {f: [s[0], m.get(s[1], s[1])] + list(s[2:]) for f, s in new["sort"].items()}
    if new.get("value_filters"):
        out["value_filters"] = {f: [s[0], m.get(s[1], s[1])] + list(s[2:]) for f, s in new["value_filters"].items()}
    return out


class PivotController:
    """Every change to a pivot made from the sheet, the panel or a dialog goes through here and ends in
    win.change_pivot (one undo step)."""

    def __init__(self, win):
        self.win = win

    @property
    def wb(self):
        return self.win.wb

    def cur(self, sheet, pv):
        """The pivot object currently on the sheet for `pv` (pv may be an older copy)."""
        if any(p is pv for p in sheet.pivots):
            return pv
        return next((p for p in sheet.pivots if p["name"] == pv["name"] and p["anchor"] == pv["anchor"]), pv)

    def commit(self, sheet, pv, new, text="Change PivotTable"):
        return self.win.change_pivot(sheet, self.cur(sheet, pv), _relabel(pv, new), text)

    # ---- sorting
    def sort(self, sheet, pv, field, order, by=None):
        self.commit(sheet, pv, P.with_key(pv, "sort", field, None if (order == "asc" and by is None)
                                          else [order, by]), "Sort PivotTable")

    def more_sort(self, sheet, pv, field):
        d = PD.SortDialog(self.win, pv, field)
        if d.exec() == QDialog.Accepted:
            o, by = d.spec()
            self.sort(sheet, pv, field, o, by)

    # ---- filtering
    def set_items(self, sheet, pv, field, allowed):
        new = P.with_key(pv, "filters", field, allowed)
        if field not in new["filter_fields"] and field not in new["rows"] and field not in new["cols"]:
            new = P.changed(new, filter_fields=new["filter_fields"] + [field])
        self.commit(sheet, pv, new, "Filter PivotTable")

    def clear_filter(self, sheet, pv, field):
        new = pv
        for k in ("filters", "label_filters", "value_filters"):
            new = P.with_key(new, k, field, None)
        self.commit(sheet, pv, new, "Clear Filter")

    def cond_filter(self, sheet, pv, field, kind, op):
        k = "label_filters" if kind == "label" else "value_filters"
        if kind == "value" and not pv["values"]:
            self.win.warn("Add a field to Values first: a value filter compares the values of a value field.")
            return
        d = PD.CondFilterDialog(self.win, pv, field, kind, op, (pv.get(k) or {}).get(field))
        if d.exec() == QDialog.Accepted:
            self.commit(sheet, pv, P.with_key(pv, k, field, d.spec()), "Filter PivotTable")

    def clear_cond(self, sheet, pv, field, kind):
        self.commit(sheet, pv, P.with_key(pv, "label_filters" if kind == "label" else "value_filters", field, None),
                    "Clear Filter")

    def keep_only(self, sheet, pv, field, labels, hide=False):
        items = P.items(self.wb, pv, field)
        allowed = (pv.get("filters") or {}).get(field) or items
        keep = [t for t in allowed if (t in labels) != hide]
        if not keep:
            self.win.warn("At least one item must stay visible.")
            return
        self.set_items(sheet, pv, field, None if len(keep) == len(items) else keep)

    # ---- expand / collapse
    def toggle(self, sheet, pv, field, label, collapse):
        self.commit(sheet, pv, P.toggled(pv, field, label, collapse), "Collapse" if collapse else "Expand")

    def field_collapse(self, sheet, pv, field, collapse):
        self.commit(sheet, pv, P.field_collapsed(pv, field, collapse),
                    "Collapse Entire Field" if collapse else "Expand Entire Field")

    # ---- grouping
    def group(self, sheet, pv, field):
        kind = P.field_kind(self.wb, pv, field)
        if kind == "virtual":
            field = P.virtual_fields(pv)[field][0]
            kind = "date"
        if kind not in ("date", "number"):
            self.win.warn("Cannot group that selection: only date or number fields can be grouped.")
            return
        d = PD.GroupDialog(self.win, self.wb, pv, field)
        if d.exec() == QDialog.Accepted:
            self.commit(sheet, pv, P.grouped(pv, field, d.spec), "Group")

    def ungroup(self, sheet, pv, field):
        if field in P.virtual_fields(pv):
            field = P.virtual_fields(pv)[field][0]
        self.commit(sheet, pv, P.grouped(pv, field, None), "Ungroup")

    def group_of(self, pv, field):
        """The field whose grouping a field belongs to (itself or a date-group field's base)."""
        virt = P.virtual_fields(pv)
        return virt[field][0] if field in virt else field

    # ---- values
    def value_settings(self, sheet, pv, vi, tab=0):
        d = PD.ValueFieldDialog(self.win, self.wb, pv, vi, tab)
        if d.exec() == QDialog.Accepted and d.result is not None:
            vals = [list(x) for x in pv["values"]]
            vals[vi] = d.result
            self.commit(sheet, pv, P.changed(pv, values=vals), "Value Field Settings")

    def set_agg(self, sheet, pv, vi, agg):
        e = pv["values"][vi]
        o = dict(P.vopts(e))
        if o.get("name") == P.value_label(e[0], agg):
            o.pop("name")
        vals = [list(x) for x in pv["values"]]
        vals[vi] = P.value_entry(e[0], agg, o)
        self.commit(sheet, pv, P.changed(pv, values=vals), "Summarize Values By")

    def set_show(self, sheet, pv, vi, show):
        if P.SHOW_BASE.get(show):
            e = pv["values"][vi]
            o = dict(P.vopts(e), show=show)
            tmp = [list(x) for x in pv["values"]]
            tmp[vi] = [e[0], e[1], o]
            self.value_settings(sheet, P.changed(pv, values=tmp), vi, tab=1)
            return
        e = pv["values"][vi]
        o = dict(P.vopts(e), show=show)
        vals = [list(x) for x in pv["values"]]
        vals[vi] = P.value_entry(e[0], e[1], o)
        self.commit(sheet, pv, P.changed(pv, values=vals), "Show Values As")

    def number_format(self, sheet, pv, vi):
        from .dialogs import FormatCellsDialog
        e = pv["values"][vi]
        o = dict(P.vopts(e))
        d = FormatCellsDialog(self.win, DEFAULT_STYLE.with_(numfmt=o.get("numfmt") or "General"), 1234.5678, 0)
        while d.tabs.count() > 1:
            d.tabs.removeTab(1)
        if d.exec() != QDialog.Accepted:
            return
        fmt = d.current_number_format()
        o["numfmt"] = None if fmt == "General" else fmt
        vals = [list(x) for x in pv["values"]]
        vals[vi] = P.value_entry(e[0], e[1], o)
        self.commit(sheet, pv, P.changed(pv, values=vals), "Number Format")

    def remove_value(self, sheet, pv, vi):
        vals = [list(x) for i, x in enumerate(pv["values"]) if i != vi]
        self.commit(sheet, pv, P.changed(pv, values=vals), "Remove Field")

    def remove_field(self, sheet, pv, field):
        new = pv
        for area in ("rows", "cols", "filter_fields"):
            if field in new[area]:
                new = P.changed(new, **{area: [f for f in new[area] if f != field]})
        new = P._drop_field_state(new, [field])
        self.commit(sheet, pv, new, "Remove Field")

    # ---- layout / design
    def set_layout(self, sheet, pv, text="PivotTable Layout", **kw):
        self.commit(sheet, pv, P.with_layout(pv, **kw), text)

    def options(self, sheet, pv):
        d = PD.OptionsDialog(self.win, pv)
        if d.exec() == QDialog.Accepted and d.result is not None:
            self.commit(sheet, pv, d.result, "PivotTable Options")

    def calc_field(self, sheet, pv):
        d = PD.CalcFieldDialog(self.win, self.wb, pv)
        if d.exec() != QDialog.Accepted or d.result is None:
            return
        calc = dict(pv.get("calc") or {})
        if d.result[0] == "delete":
            name = d.result[1]
            calc.pop(name, None)
            new = P.changed(pv, calc=calc, values=[x for x in pv["values"] if x[0] != name])
            self.commit(sheet, pv, new, "Delete Calculated Field")
            return
        _, name, formula = d.result
        calc[name] = formula
        new = P.changed(pv, calc=calc)
        if not any(x[0] == name for x in pv["values"]):
            new = P.changed(new, values=[list(x) for x in pv["values"]] + [[name, "sum"]])
        self.commit(sheet, pv, new, "Calculated Field")

    def move_values(self, sheet, pv, area):
        self.commit(sheet, pv, PD.values_moved(pv, area), "Move Σ Values")

    # ---- Show Details
    def drill(self, sheet, pv, r, c):
        win = self.win
        info, cells = P.cell_info(self.wb, pv, r, c)
        if not info or info[0] != "data":
            return
        src = self.wb.get_sheet(pv["source"])
        if src is None:
            return
        if getattr(src, "big", None) is not None:
            win.warn("Show Details isn't available yet for PivotTables built from big files.")
            return
        names, rows, srcrows = P.detail_rows(cells, info[1], info[2])
        if not rows:
            return
        if len(rows) * len(names) > 2_000_000:
            win.warn("There are too many rows behind that value to list them on a sheet.")
            return
        c1 = pv["src"][1]
        win.undo.beginMacro("Show Details")
        win.add_sheet()
        out = win.sheet
        head = DEFAULT_STYLE.with_(bold=True, fill=P.HEAD_FILL, border=(None, None, None, P.LINE))
        states = {key(0, j): (("v", n), head) for j, n in enumerate(names)}
        styles = {}
        for i, (row, sr) in enumerate(zip(rows, srcrows)):
            for j, v in enumerate(row):
                if v in (None, ""):
                    continue
                fmt = src.style(sr, c1 + j).numfmt
                st = styles.get(fmt)
                if st is None:
                    st = styles[fmt] = DEFAULT_STYLE.with_(numfmt=fmt)
                states[key(i + 1, j)] = (("v", v), st)
        win._push_states(states, "Show Details", sheet=out)
        win.autofit_cols(range(len(names)))
        win.undo.endMacro()
        win.grid.set_active(0, 0)
        win.statusBar().showMessage(f"Show Details: {len(rows)} source row{'s' if len(rows) != 1 else ''} "
                                    f"behind that value.", 5000)

    # ---- filter dropdown
    def filter_popup(self, sheet, pv, area, fields, global_pt, field=None):
        pop = PD.PivotFilterPopup(self.win, self, sheet, pv, area, fields, field)
        pop.adjustSize()
        pop.move(PD.popup_pos(pop, global_pt))
        pop.show()
        self.last_popup = pop
        return pop

    # ---- menus
    def design_menu(self, m, sheet, pv):
        """Report Layout / Subtotals / Grand Totals / Blank Rows / Styles (Excel's Design tab) and the
        rest of the PivotTable-wide commands."""
        lay = P.layout(pv)

        def check(menu, text, on, slot):
            a = menu.addAction(text, slot)
            a.setCheckable(True)
            a.setChecked(bool(on))
            return a
        rl = m.addMenu("Report Layout")
        for k, t in PD.FORMS:
            check(rl, t, lay["form"] == k, lambda k=k: self.set_layout(sheet, pv, "Report Layout", form=k))
        rl.addSeparator()
        check(rl, "Repeat All Item Labels", lay["repeat"],
              lambda: self.set_layout(sheet, pv, "Report Layout", repeat=True))
        check(rl, "Do Not Repeat Item Labels", not lay["repeat"],
              lambda: self.set_layout(sheet, pv, "Report Layout", repeat=False))
        sm = m.addMenu("Subtotals")
        for k, t in PD.SUBTOTALS:
            check(sm, t, lay["subtotals"] == k, lambda k=k: self.set_layout(sheet, pv, "Subtotals", subtotals=k))
        gm = m.addMenu("Grand Totals")
        for k, t in PD.GRANDS:
            check(gm, t, lay["grand"] == k, lambda k=k: self.set_layout(sheet, pv, "Grand Totals", grand=k))
        bm = m.addMenu("Blank Rows")
        check(bm, "Insert Blank Line after Each Item", lay["blank_line"],
              lambda: self.set_layout(sheet, pv, "Blank Rows", blank_line=True))
        check(bm, "Remove Blank Line after Each Item", not lay["blank_line"],
              lambda: self.set_layout(sheet, pv, "Blank Rows", blank_line=False))
        st = m.addMenu("PivotTable Styles")
        for n in P.STYLES:
            a = check(st, n, lay["style"] == n, lambda n=n: self.set_layout(sheet, pv, "PivotTable Style", style=n))
            a.setIcon(PD.style_icon(n))
        st.addSeparator()
        check(st, "Banded Rows", lay["banded_rows"],
              lambda: self.set_layout(sheet, pv, "Banded Rows", banded_rows=not lay["banded_rows"]))
        check(st, "Banded Columns", lay["banded_cols"],
              lambda: self.set_layout(sheet, pv, "Banded Columns", banded_cols=not lay["banded_cols"]))
        check(m, "+/- Buttons and Field Headers", lay["buttons"],
              lambda: self.set_layout(sheet, pv, "+/- Buttons", buttons=not lay["buttons"]))
        if len(pv["values"]) >= 2:
            vm = m.addMenu("Σ Values")
            in_rows = P.VALUES in pv["rows"]
            check(vm, "Show in Columns", not in_rows, lambda: self.move_values(sheet, pv, "cols"))
            check(vm, "Show in Rows", in_rows, lambda: self.move_values(sheet, pv, "rows"))
        m.addSeparator()
        m.addAction("Calculated Field...", lambda: self.calc_field(sheet, pv))
        m.addAction("PivotTable Options...", lambda: self.options(sheet, pv))

    def field_menu(self, m, sheet, pv, field, axis):
        """Sort / Filter / Group / Expand-Collapse commands for a Row or Column field."""
        srt = m.addMenu("Sort")
        srt.addAction("Sort A to Z", lambda: self.sort(sheet, pv, field, "asc"))
        srt.addAction("Sort Z to A", lambda: self.sort(sheet, pv, field, "desc"))
        srt.addAction("More Sort Options...", lambda: self.more_sort(sheet, pv, field))
        fm = m.addMenu("Filter")
        has = any(field in (pv.get(k) or {}) for k in ("filters", "label_filters", "value_filters"))
        a = fm.addAction(f'Clear Filter From "{field}"', lambda: self.clear_filter(sheet, pv, field))
        a.setEnabled(has)
        fm.addAction("Top 10...", lambda: self.cond_filter(sheet, pv, field, "value", "top"))
        lf = fm.addMenu("Label Filters")
        for op, t in P.LABEL_OPS:
            lf.addAction(t, lambda o=op: self.cond_filter(sheet, pv, field, "label", o))
        vf = fm.addMenu("Value Filters")
        for op, t in P.VALUE_OPS:
            vf.addAction(t, lambda o=op: self.cond_filter(sheet, pv, field, "value", o))
        return srt, fm

    def cell_menu(self, sheet, pv, r, c, gpos):
        win = self.win
        info, cells = P.cell_info(self.wb, pv, r, c)
        m = QMenu(win)
        m.addAction(win.a_copy)
        m.addAction(win.a_format_cells)
        kind = info[0] if info else None
        vi = None
        if kind == "data":
            vi = info[3]
        elif kind == "vhead":
            vi = info[1]
        if vi is not None and vi < len(pv["values"]):
            m.addAction("Number Format...", lambda: self.number_format(sheet, pv, vi))
        m.addAction("Refresh", lambda: win.refresh_pivot(sheet, self.cur(sheet, pv)))
        m.addSeparator()
        node = None
        field = None
        if kind in ("ritem", "rtotal", "citem", "ctotal") and info[1] is not None:
            node = info[1]
            field = node.field if node.field != P.VALUES else None
        elif kind == "caption":
            field = info[2][0]   # Excel: the first field of "Row Labels" / "Column Labels"
        if kind == "data":
            ln = info[4]
            nd = ln.get("node")
            row_field = nd.field if nd is not None and nd.field != P.VALUES else \
                (cells.ctx["rdf"][-1] if cells.ctx["rdf"] else None)
            label = cells.ctx["vlabels"][vi]
            if row_field:
                srt = m.addMenu("Sort")
                srt.addAction("Sort Smallest to Largest",
                              lambda: self.sort(sheet, pv, row_field, "asc", label))
                srt.addAction("Sort Largest to Smallest",
                              lambda: self.sort(sheet, pv, row_field, "desc", label))
                srt.addAction("More Sort Options...", lambda: self.more_sort(sheet, pv, row_field))
            m.addSeparator()
            m.addAction("Show Details", lambda: self.drill(sheet, pv, r, c))
        if field:
            axis = "rows" if kind in ("ritem", "rtotal") or (kind == "caption" and info[1] == "rows") else "cols"
            self.field_menu(m, sheet, pv, field, axis)
            if node is not None and kind in ("ritem", "citem"):
                labels = self._selected_labels(sheet, pv, cells, field) or [node.label]
                m.addAction("Keep Only Selected Items", lambda: self.keep_only(sheet, pv, field, labels))
                m.addAction("Hide Selected Items", lambda: self.keep_only(sheet, pv, field, labels, hide=True))
            m.addSeparator()
            kindf = P.field_kind(self.wb, pv, field)
            if kindf in ("date", "number", "virtual"):
                m.addAction("Group...", lambda: self.group(sheet, pv, field))
            if self.group_of(pv, field) in (pv.get("groups") or {}):
                m.addAction("Ungroup", lambda: self.ungroup(sheet, pv, field))
            flds = cells.ctx["rdf"] if axis == "rows" else cells.ctx["cdf"]
            if field in flds and flds.index(field) < len(flds) - 1:
                em = m.addMenu("Expand/Collapse")
                if node is not None and kind in ("ritem", "citem"):
                    em.addAction("Expand", lambda: self.toggle(sheet, pv, field, node.label, False)) \
                        .setEnabled(node.collapsed)
                    em.addAction("Collapse", lambda: self.toggle(sheet, pv, field, node.label, True)) \
                        .setEnabled(not node.collapsed)
                    em.addSeparator()
                em.addAction("Expand Entire Field", lambda: self.field_collapse(sheet, pv, field, False))
                em.addAction("Collapse Entire Field", lambda: self.field_collapse(sheet, pv, field, True))
            m.addAction(f'Remove "{field}"', lambda: self.remove_field(sheet, pv, field))
        if kind == "report":
            fld = info[1]
            m.addAction(f'Clear Filter From "{fld}"', lambda: self.clear_filter(sheet, pv, fld))
            m.addAction(f'Remove "{fld}"', lambda: self.remove_field(sheet, pv, fld))
        if vi is not None and vi < len(pv["values"]):
            m.addSeparator()
            sm = m.addMenu("Summarize Values By")
            for agg, t in P.AGGS:
                a = sm.addAction(t, lambda g=agg: self.set_agg(sheet, pv, vi, g))
                a.setCheckable(True)
                a.setChecked(pv["values"][vi][1] == agg)
            shm = m.addMenu("Show Values As")
            cur = P.vopts(pv["values"][vi]).get("show", "none")
            for k, t in P.SHOW_AS:
                a = shm.addAction(t + ("..." if P.SHOW_BASE.get(k) else ""), lambda k=k: self.set_show(sheet, pv, vi, k))
                a.setCheckable(True)
                a.setChecked(cur == k)
            m.addAction("Value Field Settings...", lambda: self.value_settings(sheet, pv, vi))
            m.addAction(f'Remove "{cells.ctx["vlabels"][vi]}"', lambda: self.remove_value(sheet, pv, vi))
        m.addSeparator()
        self.design_menu(m.addMenu("Design"), sheet, pv)
        m.addAction("PivotTable Options...", lambda: self.options(sheet, pv))
        panel = win.pivot_panel
        m.addAction("Hide Field List" if panel.isVisible() else "Show Field List",
                    lambda: panel.setVisible(not panel.isVisible()))
        self.last_menu = m
        m.exec(gpos)

    def _selected_labels(self, sheet, pv, cells, field):
        ar, ac = pv["anchor"]
        out = []
        for r1, c1, r2, c2 in self.win.grid.sel.rects:
            o = pv["out"]
            for r in range(max(r1, o[0]), min(r2, o[2]) + 1):
                for c in range(max(c1, o[1]), min(c2, o[3]) + 1):
                    info = cells.cmap.get((r - ar, c - ac))
                    if info and info[0] in ("ritem", "citem") and info[1] is not None and info[1].field == field:
                        out.append(info[1].label)
        return out


# ================================================================ in-sheet buttons
class PivotOverlay(QWidget):
    """Draws the pivots' dropdown and +/- buttons over the grid and handles clicks on them, plus the
    pivot right-click menu and double-click (drill down / expand-collapse). Mouse events pass through
    to the grid; the grid's events are watched with an event filter."""
    PAD = 3

    def __init__(self, ctl):
        grid = ctl.win.grid
        super().__init__(grid)
        self.ctl, self.grid = ctl, grid
        self.setAttribute(Qt.WA_TransparentForMouseEvents)
        self.setAttribute(Qt.WA_NoSystemBackground)
        self.setGeometry(grid.rect())
        self.stackUnder(grid.editor)
        self._lazy = {}
        grid.installEventFilter(self)
        self.show()

    def _hot(self, pv):
        hot = pv.get("hot")
        if hot is not None:
            return hot
        got = self._lazy.get(id(pv))
        if got is None or got[0] is not pv:   # a pivot loaded from an older file: work its buttons out once
            cells, _, err = P.build(self.ctl.wb, pv)
            got = self._lazy[id(pv)] = (pv, [] if err else list(cells.hot))
        return got[1]

    def _buttons(self):
        """[(rect, entry, pv)] for the buttons of pivots on the current sheet that are on screen."""
        g = self.grid
        sh = g.sheet
        out = []
        if sh is None or not getattr(sh, "pivots", None):
            return out
        W, H = g.width(), g.height()
        z = g.zoom
        for pv in sh.pivots:
            o = pv.get("out")
            if not o or not P.layout(pv).get("buttons", True):
                continue
            ar, ac = pv["anchor"]
            for h in self._hot(pv):
                r, c = ar + h[1], ac + h[2]
                if g.rows.size(r) <= 0 or g.cols.size(c) <= 0:
                    continue
                R = g.cell_rect(r, c)
                if R.bottom() < g.hh or R.top() > H or R.right() < g.rw or R.left() > W:
                    continue
                if h[0] == "flt":
                    out.append((g._filter_button_rect(R.x(), R.y(), R.width(), R.height()), h, pv))
                else:
                    s = max(7, int(round(9 * z)))
                    x = R.x() + 2 + int(h[6] * 9 * z)
                    y = R.y() + (R.height() - s) // 2
                    out.append((QRect(x, y, s, s), h, pv))
        return out

    def paintEvent(self, ev):
        g = self.grid
        if g.sheet is None or not getattr(g.sheet, "pivots", None):
            return
        p = QPainter(self)
        p.setClipRect(QRect(g.rw, g.hh, self.width() - g.rw, self.height() - g.hh))
        for rect, h, pv in self._buttons():
            if not rect.intersects(ev.rect()):
                continue
            if h[0] == "flt":
                g._paint_filter_button(p, rect, bool(h[5]))
            else:
                p.setPen(QPen(QColor("#7F7F7F"), 1))
                p.setBrush(QColor("#FFFFFF"))
                p.drawRect(rect.adjusted(0, 0, -1, -1))
                p.setPen(QPen(QColor("#333333"), 1))
                cx, cy = rect.center().x(), rect.center().y()
                d = max(2, rect.width() // 2 - 2)
                p.drawLine(cx - d, cy, cx + d, cy)
                if h[5]:   # collapsed: +
                    p.drawLine(cx, cy - d, cx, cy + d)
        p.end()

    def _pivot_cell(self, pos):
        g = self.grid
        if g.sheet is None or not getattr(g.sheet, "pivots", None):
            return None, -1, -1
        area, r, c = g._hit(pos)
        if area != "cell":
            return None, r, c
        return self.ctl.win.pivot_at(g.sheet, r, c), r, c

    def eventFilter(self, obj, ev):
        t = ev.type()
        if t == QEvent.Resize:
            self.setGeometry(self.grid.rect())
            return False
        if t not in (QEvent.MouseButtonPress, QEvent.MouseButtonDblClick, QEvent.ContextMenu):
            return False
        g = self.grid
        if g.editing or getattr(g, "read_only", False):
            return False
        if t == QEvent.ContextMenu:
            pv, r, c = self._pivot_cell(ev.pos())
            if pv is None:
                return False
            if not g.sel.contains(r, c):
                g.set_active(r, c)
            self.ctl.cell_menu(g.sheet, pv, r, c, ev.globalPos())
            return True
        if ev.button() != Qt.LeftButton:
            return False
        pos = ev.position().toPoint()
        pv, r, c = self._pivot_cell(pos)
        if pv is None:
            return False
        sh = g.sheet
        for rect, h, bpv in self._buttons():
            if bpv is pv and rect.adjusted(-2, -2, 2, 2).contains(pos):
                if h[0] == "flt":
                    R = g.cell_rect(pv["anchor"][0] + h[1], pv["anchor"][1] + h[2])
                    self.ctl.filter_popup(sh, pv, h[3], h[4], g.mapToGlobal(QPoint(R.x(), R.bottom() + 1)))
                else:
                    self.ctl.toggle(sh, pv, h[3], h[4], not h[5])
                return True
        if t == QEvent.MouseButtonDblClick:
            info, _ = P.cell_info(self.ctl.wb, pv, r, c)
            if info and info[0] == "data":
                self.ctl.drill(sh, pv, r, c)
            elif info and info[0] in ("ritem", "citem") and info[1] is not None and info[1].expandable:
                self.ctl.toggle(sh, pv, info[1].field, info[1].label, not info[1].collapsed)
            return True   # no in-cell editing inside a pivot
        return False


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


def _explicit(pv):
    """pv with Σ Values written into Columns when it's only there implicitly (2+ value fields)."""
    if len(pv["values"]) >= 2 and P.VALUES not in pv["rows"] and P.VALUES not in pv["cols"]:
        return P.changed(pv, cols=pv["cols"] + [P.VALUES])
    return pv


class PivotPanel(QDockWidget):
    """'PivotTable Fields': shown while the active cell is inside a PivotTable."""

    def __init__(self, win):
        super().__init__("PivotTable Fields", win)
        self.win = win
        self.ctl = PivotController(win)
        self.setObjectName("PivotPanel")
        self.sheet = None
        self.pv = None
        self.pending = None   # Defer Layout Update: the layout waiting for Update
        body = QWidget()
        v = QVBoxLayout(body)
        v.setContentsMargins(8, 6, 8, 8)
        top = QHBoxLayout()
        top.addWidget(QLabel("Choose fields to add to report:"), 1)
        self.opts = QToolButton()
        self.opts.setText("Options")
        self.opts.setToolTip("Report layout, subtotals, grand totals, styles, calculated fields, options")
        self.opts.setPopupMode(QToolButton.InstantPopup)
        self.opts_menu = QMenu(self.opts)
        self.opts_menu.aboutToShow.connect(self._fill_opts)
        self.opts.setMenu(self.opts_menu)
        top.addWidget(self.opts)
        v.addLayout(top)
        self.search = QLineEdit()
        self.search.setPlaceholderText("Search")
        self.search.setClearButtonEnabled(True)
        self.search.textChanged.connect(self._search)
        v.addWidget(self.search)
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
        self.defer = QCheckBox("Defer Layout Update")
        self.defer.setToolTip("Collect field changes and apply them all with Update")
        self.defer.toggled.connect(self._defer_toggled)
        self.update_btn = QPushButton("Update")
        self.update_btn.setEnabled(False)
        self.update_btn.clicked.connect(self._update_now)
        refresh = QPushButton("Refresh")
        refresh.setToolTip("Re-read the source data (Ctrl+Alt+F5 refreshes every PivotTable)")
        refresh.clicked.connect(lambda: self.win.refresh_pivot(self.sheet, self.pv))
        h.addWidget(self.defer)
        h.addStretch(1)
        h.addWidget(self.update_btn)
        h.addWidget(refresh)
        v.addLayout(h)
        self.setWidget(body)
        self.setMinimumWidth(280)
        self.overlay = PivotOverlay(self.ctl) if getattr(win, "grid", None) is not None else None

    # ---- showing a pivot
    def show_pivot(self, sheet, pv):
        if self.pending is not None and not (sheet is self.sheet and pv is self.pv):
            self.pending = None
        self.sheet, self.pv = sheet, pv
        self.update_btn.setEnabled(self.pending is not None)
        self._display(self.pending if self.pending is not None else pv)
        self.show()

    @property
    def view(self):
        """The layout being edited: the pending one while Defer Layout Update holds changes back."""
        return self.pending if self.pending is not None else self.pv

    def _display(self, pv):
        names = P.fields(self.win.wb, pv)
        used = set(pv["rows"]) | set(pv["cols"]) | set(pv["filter_fields"]) | {x[0] for x in pv["values"]}
        calc = pv.get("calc") or {}
        virt = P.virtual_fields(pv)
        self.fields.blockSignals(True)
        self.fields.clear()
        for n in names:
            it = QListWidgetItem(n)
            it.setData(Qt.UserRole, n)
            it.setFlags(it.flags() | Qt.ItemIsUserCheckable | Qt.ItemIsDragEnabled)
            it.setCheckState(Qt.Checked if n in used else Qt.Unchecked)
            if n in calc:
                it.setToolTip(f"Calculated field: {calc[n]}")
                f = it.font()
                f.setItalic(True)
                it.setFont(f)
            elif n in virt:
                it.setToolTip(f"{dict(P.DATE_LEVELS)[virt[n][1]]} of {virt[n][0]} (date grouping)")
            self.fields.addItem(it)
        self.fields.blockSignals(False)
        self._search(self.search.text())
        labels = P.value_labels(pv)
        filt = {k for key_ in ("filters", "label_filters", "value_filters") for k in (pv.get(key_) or {})}
        shown = _explicit(pv)
        for area, lst in self.areas.items():
            lst.clear()
            for i, entry in enumerate(shown[area]):
                if area == "values":
                    it = QListWidgetItem(labels[i])
                    it.setData(Qt.UserRole, entry[0])
                    it.setData(Qt.UserRole + 1, i)
                else:
                    if entry == P.VALUES and len(pv["values"]) < 2:
                        continue
                    it = QListWidgetItem(entry + ("  (filtered)" if entry in filt else ""))
                    it.setData(Qt.UserRole, entry)
                lst.addItem(it)

    def _search(self, text):
        t = text.strip().lower()
        for i in range(self.fields.count()):
            it = self.fields.item(i)
            it.setHidden(bool(t) and t not in it.text().lower())

    # ---- defer layout update
    def _defer_toggled(self, on):
        if not on and self.pending is not None:
            self._update_now()

    def _update_now(self):
        if self.pending is None or self.pv is None:
            return
        new, self.pending = self.pending, None
        self.update_btn.setEnabled(False)
        self.win.change_pivot(self.sheet, self.ctl.cur(self.sheet, self.pv), new, "Change PivotTable")

    # ---- changes
    def _commit(self, new, text="Change PivotTable"):
        if self.defer.isChecked():
            self.pending = new
            self.update_btn.setEnabled(True)
            self._display(new)
            return
        self.win.change_pivot(self.sheet, self.ctl.cur(self.sheet, self.pv), new, text)

    def _without(self, pv, field, area):
        """pv with `field` taken out of `area` (all of its Values entries when area is 'values')."""
        if area == "values":
            return P.changed(pv, values=[x for x in pv["values"] if x[0] != field])
        out = P.changed(pv, **{area: [f for f in pv[area] if f != field]})
        if area == "filter_fields":
            out["filters"] = {k: v for k, v in pv["filters"].items() if k != field}
        return out

    def _add(self, pv, field, area, index=None):
        wb = self.win.wb
        if area == "values":
            vals = [list(x) for x in pv["values"]]
            entry = [field, P.default_agg(wb, pv, field)]
            vals.insert(len(vals) if index is None else index, entry)
            return P.changed(pv, values=vals)
        if field in (pv.get("calc") or {}):
            self.win.warn("A calculated field can only go in Values.")
            return pv
        pv2 = pv
        for other in ("rows", "cols", "filter_fields"):  # a field is in one of these at a time
            if field in pv2[other]:
                pv2 = self._without(pv2, field, other)
        lst = list(pv2[area])
        lst.insert(len(lst) if index is None else min(index, len(lst)), field)
        out = P.changed(pv2, **{area: lst})
        if area in ("rows", "cols") and field not in (pv.get("groups") or {}):
            out = P.auto_group(wb, out, field)   # Excel 365 time grouping of date fields
        return out

    def _dropped(self, field, from_area, to_area, index):
        pv = self.view
        if field == P.VALUES:
            if to_area not in ("rows", "cols"):
                return
            pv = _explicit(pv)
            rows = [f for f in pv["rows"] if f != P.VALUES]
            cols = [f for f in pv["cols"] if f != P.VALUES]
            lst = rows if to_area == "rows" else cols
            old = pv[to_area].index(P.VALUES) if from_area == to_area else None
            index = index - (1 if old is not None and index > old else 0)
            lst.insert(min(index, len(lst)), P.VALUES)
            self._commit(P.changed(pv, rows=rows, cols=cols), "Move Σ Values")
            return
        if from_area == to_area and to_area != "values":
            pv = _explicit(pv)
            lst = [f for f in pv[to_area] if f != field]
            old = pv[to_area].index(field)
            lst.insert(index - (1 if index > old else 0), field)
            self._commit(P.changed(pv, **{to_area: lst}))
            return
        if from_area == to_area == "values":
            src = self.areas["values"].currentItem()
            i = src.data(Qt.UserRole + 1)
            vals = [list(x) for x in pv["values"]]
            entry = vals.pop(i)
            vals.insert(index - (1 if index > i else 0), entry)
            self._commit(P.changed(pv, values=vals))
            return
        if from_area == "values":
            src = self.areas["values"].currentItem()
            vals = [list(x) for x in pv["values"]]
            vals.pop(src.data(Qt.UserRole + 1))
            pv = P.changed(pv, values=vals)
        elif from_area:
            pv = self._without(pv, field, from_area)
        self._commit(self._add(pv, field, to_area, index))

    def _field_ticked(self, it):
        field = it.data(Qt.UserRole)
        pv = self.view
        if it.checkState() == Qt.Checked:
            # like Excel: numbers (and calculated fields) go to Values, everything else to Rows
            kind = P.field_kind(self.win.wb, pv, field)
            area = "values" if kind in ("number", "calc") else "rows"
            self._commit(self._add(pv, field, area))
        else:
            for area in ("rows", "cols", "filter_fields", "values"):
                pv = self._without(pv, field, area)
            self._commit(P._drop_field_state(pv, [field]))

    def _field_menu(self, pos):
        it = self.fields.itemAt(pos)
        if it is None:
            return
        field = it.data(Qt.UserRole)
        m = QMenu(self)
        for area, label in AREAS:
            m.addAction(f"Add to {label}", lambda a=area: self._dropped(field, self._area_of(field, a), a, 10 ** 6))
        if field in (self.view.get("calc") or {}):
            m.addSeparator()
            m.addAction("Edit Calculated Field...", lambda: self.ctl.calc_field(self.sheet, self.pv))
        self._menu = m
        m.exec(self.fields.mapToGlobal(pos))

    def _fill_opts(self):
        m = self.opts_menu
        m.clear()
        if self.pv is None:
            return
        self.ctl.design_menu(m, self.sheet, self.pv)

    def _area_of(self, field, target):
        if target == "values":
            return ""
        for area in ("rows", "cols", "filter_fields"):
            if field in self.view[area]:
                return area
        return ""

    def _area_menu(self, lst, pos):
        it = lst.itemAt(pos)
        if it is None:
            return
        lst.setCurrentItem(it)
        field, area = it.data(Qt.UserRole), lst.area
        pv = self.view
        sheet = self.sheet
        m = QMenu(self)
        n = lst.count()
        i = lst.row(it)
        if i > 0:
            m.addAction("Move Up", lambda: self._dropped(field, area, area, i - 1))
            m.addAction("Move to Beginning", lambda: self._dropped(field, area, area, 0))
        if i < n - 1:
            m.addAction("Move Down", lambda: self._dropped(field, area, area, i + 2))
            m.addAction("Move to End", lambda: self._dropped(field, area, area, n))
        m.addSeparator()
        if field == P.VALUES:
            other = "rows" if area == "cols" else "cols"
            m.addAction(f"Move to {dict(AREAS)[other]}", lambda: self._dropped(field, area, other, 10 ** 6))
            self._menu = m
            m.exec(lst.mapToGlobal(pos))
            return
        for other, label in AREAS:
            if other != area:
                m.addAction(f"Move to {label}", lambda o=other: self._dropped(field, area, o, 10 ** 6))
        m.addSeparator()
        if area == "values":
            vi = it.data(Qt.UserRole + 1)
            sm = m.addMenu("Summarize Values By")
            for agg, label in P.AGGS:
                a = sm.addAction(label, lambda g=agg: self._set_agg(vi, g))
                a.setCheckable(True)
                a.setChecked(pv["values"][vi][1] == agg)
            m.addAction("Value Field Settings...", lambda: self._value_settings(vi))
        else:
            if area in ("rows", "cols"):
                self.ctl.field_menu(m, sheet, pv, field, area)
                kind = P.field_kind(self.win.wb, pv, field)
                if kind in ("date", "number", "virtual"):
                    m.addAction("Group...", lambda: self.ctl.group(sheet, pv, field))
                if self.ctl.group_of(pv, field) in (pv.get("groups") or {}):
                    m.addAction("Ungroup", lambda: self.ctl.ungroup(sheet, pv, field))
                m.addAction("Expand Entire Field", lambda: self.ctl.field_collapse(sheet, pv, field, False))
                m.addAction("Collapse Entire Field", lambda: self.ctl.field_collapse(sheet, pv, field, True))
            else:
                m.addAction("Filter...", lambda: self._filter(lst, it, field))
        m.addSeparator()
        m.addAction("Remove Field", lambda: self._remove(area, it))
        self._menu = m
        m.exec(lst.mapToGlobal(pos))

    def _area_double(self, lst, it):
        field = it.data(Qt.UserRole)
        if field == P.VALUES:
            return
        if lst.area == "values":
            self._value_settings(it.data(Qt.UserRole + 1))
        else:
            self._filter(lst, it, field)

    def _filter(self, lst, it, field):
        area = "report" if lst.area == "filter_fields" else lst.area
        pt = lst.mapToGlobal(lst.visualItemRect(it).bottomLeft())
        self.ctl.filter_popup(self.sheet, self.view, area, [field], pt)

    def _value_settings(self, vi):
        self.ctl.value_settings(self.sheet, self.view, vi)

    def _set_agg(self, vi, agg):
        self.ctl.set_agg(self.sheet, self.view, vi, agg)

    def _remove(self, area, it):
        pv = self.view
        if area == "values":
            vals = [list(x) for x in pv["values"]]
            vals.pop(it.data(Qt.UserRole + 1))
            self._commit(P.changed(pv, values=vals))
        else:
            self._commit(P._drop_field_state(self._without(pv, it.data(Qt.UserRole), area), [it.data(Qt.UserRole)]))
