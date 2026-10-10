"""Workbook / sheet model with dependency-tracked recalculation."""
from bisect import bisect_left, bisect_right, insort
from dataclasses import dataclass, replace

from . import errors
from .engine import Evaluator
from .errors import XLError
from .formula import (ParseError, adjust_structure, collect_refs, func_names,
                      parse, rename_sheet)
from .functions import FUNCS, VOLATILE
from .refs import MAX_COLS, MAX_ROWS, key, unkey

DEFAULT_COL_WIDTH = 64   # px at 100% zoom (Excel's 8.43 chars)
DEFAULT_ROW_HEIGHT = 20  # px


@dataclass(frozen=True)
class Style:
    bold: bool = False
    italic: bool = False
    underline: bool = False
    strike: bool = False
    font: str = None
    size: float = None
    color: str = None       # "#RRGGBB"
    fill: str = None        # "#RRGGBB"
    halign: str = None      # None(general) / left / center / right
    valign: str = None      # None(bottom) / top / center / bottom
    wrap: bool = False
    numfmt: str = "General"
    border: tuple = (None, None, None, None)  # left, top, right, bottom: None or (style, color)
    indent: int = 0

    def with_(self, **kw):
        return intern_style(replace(self, **kw))


_STYLES = {}


def intern_style(s):
    return _STYLES.setdefault(s, s)


DEFAULT_STYLE = intern_style(Style())


class Formula:
    __slots__ = ("text", "ast", "refs", "value", "dirty", "visiting", "volatile", "fallback", "unknown")

    def __init__(self, text, fallback=None):
        self.text = text
        try:
            self.ast = parse(text)
            bad = False
        except (ParseError, RecursionError, IndexError):
            self.ast = ("err", errors.NAME)
            bad = True
        self.refs = collect_refs(self.ast)
        names = func_names(self.ast)
        self.volatile = bool(names & VOLATILE) or _has_names(self.ast)
        self.unknown = bad or any(n not in FUNCS for n in names)
        self.value = None
        self.dirty = True
        self.visiting = False
        self.fallback = fallback


def _has_names(n):
    k = n[0]
    if k == "name":
        return True
    if k == "func":
        return any(_has_names(a) for a in n[2])
    if k == "bin":
        return _has_names(n[2]) or _has_names(n[3])
    if k in ("neg", "pct"):
        return _has_names(n[1])
    return False


# cell "content" in states: None, ('v', value) or ('f', formula_text)

class Sheet:
    def __init__(self, wb, name):
        self.wb = wb
        self.name = name
        self.values = {}
        self.formulas = {}
        self.styles = {}
        self.col_widths = {}
        self.row_heights = {}
        self.hidden_rows = set()
        self.hidden_cols = set()
        self.merges = []          # [(r1, c1, r2, c2)]
        self.freeze = (0, 0)      # frozen rows, cols
        self.freeze_origin = (0, 0)  # first frozen row/col shown (when frozen while scrolled)
        self.autofilter = None    # (r1, c1, r2, c2)
        self.filters = {}         # col -> spec dict
        self.filter_hidden = set()
        self.max_row = -1
        self.max_col = -1
        self._fcols = {}          # col -> sorted rows having formulas
        self.xl = None            # backing openpyxl worksheet (xlsx round-trip)
        self.tab_color = None
        self.cond_formats = []    # condfmt.CFRule list, first = highest priority (written back to xlsx)
        self.cf_complete = True   # False when a loaded file had rules Ekxel couldn't read (left as they were)
        # xlsx round-trip data that moves with rows/cols:
        self.notes = {}           # key -> (openpyxl Comment, Hyperlink)
        self.xl_dv = []           # [[openpyxl DataValidation, [rects]]]
        self.xl_styles = {}       # key -> (Style it converted to, original StyleArray)
        self.controls = []        # form controls (sliders/spinners), see ui/controls.py
        self.charts = []          # charts floating over the cells (dicts), see charts.py
        self.pivots = []          # PivotTables drawn on this sheet (dicts, see pivot.py)
        self.show_grid = True
        self.zoom = 1.0

    def __repr__(self):
        return f"<Sheet {self.name}>"

    # ------------------------------------------------------------ reading
    def value(self, r, c):
        k = (r << 14) | c
        f = self.formulas.get(k)
        if f is None:
            return self.values.get(k)
        if f.dirty:
            if f.visiting:
                return errors.CIRC
            self.wb._ensure(self, k, f)
        return f.value

    def value_at_key(self, k):
        f = self.formulas.get(k)
        if f is None:
            return self.values.get(k)
        if f.dirty:
            if f.visiting:
                return errors.CIRC
            self.wb._ensure(self, k, f)
        return f.value

    def formula_text(self, r, c):
        f = self.formulas.get(key(r, c))
        return f.text if f else None

    def style(self, r, c):
        return self.styles.get((r << 14) | c, DEFAULT_STYLE)

    def has_content(self, r, c):
        k = (r << 14) | c
        return k in self.values or k in self.formulas

    def keys_in(self, r1, c1, r2, c2):
        lo, hi = (r1 << 14), ((r2 << 14) | (MAX_COLS - 1))
        for d in (self.values, self.formulas):
            for k in d:
                if lo <= k <= hi:
                    c = k & 0x3FFF
                    if c1 <= c <= c2:
                        yield k

    def content_keys(self):
        return set(self.values) | set(self.formulas)

    def used_extent(self, include_styles=True):
        """(max_row, max_col) of anything present, -1 when empty."""
        mr = mc = -1
        keysets = [self.values, self.formulas]
        if include_styles:
            keysets.append(self.styles)
        for d in keysets:
            for k in d:
                r, c = k >> 14, k & 0x3FFF
                if r > mr:
                    mr = r
                if c > mc:
                    mc = c
        return mr, mc

    def recompute_extent(self):
        mr = mc = -1
        for d in (self.values, self.formulas):
            for k in d:
                r, c = k >> 14, k & 0x3FFF
                if r > mr:
                    mr = r
                if c > mc:
                    mc = c
        self.max_row, self.max_col = mr, mc

    # ------------------------------------------------------------ states
    def get_state(self, k):
        f = self.formulas.get(k)
        if f is not None:
            content = ("f", f.text)
        elif k in self.values:
            content = ("v", self.values[k])
        else:
            content = None
        return content, self.styles.get(k, DEFAULT_STYLE)

    def set_state(self, k, state):
        content, style = state
        if self.xl_styles and k in self.xl_styles and self.styles.get(k, DEFAULT_STYLE) != (style or DEFAULT_STYLE):
            self.xl_styles.pop(k)
        self.set_content_key(k, content)
        if style is DEFAULT_STYLE or style is None or style == DEFAULT_STYLE:
            self.styles.pop(k, None)
        else:
            self.styles[k] = intern_style(style)

    def set_content_key(self, k, content, fallback=None):
        wb = self.wb
        old = self.formulas.pop(k, None)
        if old is not None:
            wb._unregister(self, k, old)
            r, c = k >> 14, k & 0x3FFF
            rows = self._fcols.get(c)
            if rows:
                i = bisect_left(rows, r)
                if i < len(rows) and rows[i] == r:
                    rows.pop(i)
        self.values.pop(k, None)
        if content is not None:
            kind, payload = content
            r, c = k >> 14, k & 0x3FFF
            if kind == "f":
                f = Formula(payload, fallback)
                self.formulas[k] = f
                wb._register(self, k, f)
                insort(self._fcols.setdefault(c, []), r)
            else:
                if payload is None:
                    wb._changed.append((self, k))
                    return
                self.values[k] = payload
            if r > self.max_row:
                self.max_row = r
            if c > self.max_col:
                self.max_col = c
        wb._changed.append((self, k))

    def set_value(self, r, c, v):
        self.set_content_key(key(r, c), None if v is None else ("v", v))

    def set_formula(self, r, c, text, fallback=None):
        self.set_content_key(key(r, c), ("f", text), fallback)

    def set_style(self, r, c, style):
        k = key(r, c)
        if style is None or style == DEFAULT_STYLE:
            self.styles.pop(k, None)
        else:
            self.styles[k] = intern_style(style)

    # ------------------------------------------------------------ geometry
    def col_width(self, c):
        return self.col_widths.get(c, DEFAULT_COL_WIDTH)

    def row_height(self, r):
        return self.row_heights.get(r, DEFAULT_ROW_HEIGHT)

    def row_label(self, r):
        """Number shown in the row header (big sheets show the file's row when sorted/filtered)."""
        return r + 1

    def merge_at(self, r, c):
        for m in self.merges:
            if m[0] <= r <= m[2] and m[1] <= c <= m[3]:
                return m
        return None

    # ------------------------------------------------------------ structure
    def _remap(self, axis, at, count):
        n = -count

        def move(x):
            if count > 0:
                if x >= at:
                    x += count
                    return x if x < (MAX_ROWS if axis == "row" else MAX_COLS) else None
                return x
            if x >= at + n:
                return x - n
            if x >= at:
                return None
            return x

        def remap_dict(d):
            out = {}
            for k, v in d.items():
                r, c = k >> 14, k & 0x3FFF
                if axis == "row":
                    r = move(r)
                    if r is None:
                        continue
                else:
                    c = move(c)
                    if c is None:
                        continue
                out[(r << 14) | c] = v
            return out

        def remap_index(d):
            out = {}
            for x, v in d.items():
                y = move(x)
                if y is not None:
                    out[y] = v
            return out

        def remap_set(s):
            return {y for y in (move(x) for x in s) if y is not None}

        for k, f in self.formulas.items():
            self.wb._unregister(self, k, f)
        self.values = remap_dict(self.values)
        self.formulas = remap_dict(self.formulas)
        self.styles = remap_dict(self.styles)
        self.notes = remap_dict(self.notes)
        self.xl_styles = remap_dict(self.xl_styles)
        if axis == "row":
            self.row_heights = remap_index(self.row_heights)
            self.hidden_rows = remap_set(self.hidden_rows)
            self.filter_hidden = remap_set(self.filter_hidden)
        else:
            self.col_widths = remap_index(self.col_widths)
            self.hidden_cols = remap_set(self.hidden_cols)
            self.filters = remap_index(self.filters)

        def remap_rect(rect):
            r1, c1, r2, c2 = rect
            if axis == "row":
                a, b = r1, r2
            else:
                a, b = c1, c2
            if count > 0:
                if a >= at:
                    a += count
                    b += count
                elif b >= at:
                    b += count
            else:
                lo, hi = at, at + n - 1
                if lo <= a and b <= hi:
                    return None
                na = a if a < lo else (lo if a <= hi else a - n)
                nb = b if b < lo else (lo - 1 if b <= hi else b - n)
                a, b = na, nb
            return (a, c1, b, c2) if axis == "row" else (r1, a, r2, b)

        self.merges = [m for m in (remap_rect(m) for m in self.merges)
                       if m and (m[0] != m[2] or m[1] != m[3])]

        def remap_rects(rects):
            return [x for x in (remap_rect(r) for r in rects) if x]
        for rule in self.cond_formats:
            rule.rects = remap_rects(rule.rects)
            rule._asts = {}
        self.cond_formats = [r for r in self.cond_formats if r.rects]
        self.xl_dv = [[dv, remap_rects(rects)] for dv, rects in self.xl_dv]
        self.xl_dv = [e for e in self.xl_dv if e[1]]
        if self.controls:
            from .ui.controls import remap as remap_control
            same = lambda x: x
            mr, mc = (move, same) if axis == "row" else (same, move)
            self.controls = [c for c in (remap_control(c, mr, mc, remap_rect) for c in self.controls) if c]
        if self.autofilter:
            self.autofilter = remap_rect(self.autofilter)
            if self.autofilter is None:
                self.filters = {}
                self.filter_hidden = set()
        # PivotTables: their source range (on any sheet) and, for pivots on this sheet, their position
        from .pivot import remapped
        for other in self.wb.sheets:
            if other.pivots and (other is self or any(p["source"] == self.name for p in other.pivots)):
                other.pivots = [x for x in (remapped(p, self.name, remap_rect, other is self)
                                            for p in other.pivots) if x]
        # charts: their cell references (on any sheet) and, on this sheet, where they sit
        from . import charts as _charts
        for other in self.wb.sheets:
            if other.charts:
                other.charts = [_charts.remap(c, self.name, remap_rect, other is self, axis, at, count, other.name)
                                for c in other.charts]
        self._rebuild_fcols()
        self.recompute_extent()

    def _rebuild_fcols(self):
        cols = {}
        for k in self.formulas:
            cols.setdefault(k & 0x3FFF, []).append(k >> 14)
        for rows in cols.values():
            rows.sort()
        self._fcols = cols

    def _structural(self, axis, at, count):
        self._remap(axis, at, count)
        wb = self.wb
        for sh in wb.sheets:
            for k, f in list(sh.formulas.items()):
                new = adjust_structure(f.text, sh.name, self.name, axis, at, count)
                if new != f.text:
                    if sh is not self:
                        wb._unregister(sh, k, f)
                    sh.formulas[k] = Formula(new, f.fallback)
                    if sh is not self:
                        wb._register(sh, k, sh.formulas[k])
        for k, f in self.formulas.items():
            wb._register(self, k, f)
        for name, text in list(wb.names.items()):
            wb.names[name] = adjust_structure("=" + text, "", self.name, axis, at, count)[1:]
        wb.recalc(full=True)

    def insert_rows(self, at, count=1):
        self._structural("row", at, count)

    def delete_rows(self, at, count=1):
        self._structural("row", at, -count)

    def insert_cols(self, at, count=1):
        self._structural("col", at, count)

    def delete_cols(self, at, count=1):
        self._structural("col", at, -count)


class Workbook:
    def __init__(self):
        self.sheets = []
        self.evaluator = Evaluator(self)
        self.path = None
        self.file_format = None   # 'csv' | 'xlsx' | 'xls' | None
        self.csv_options = None   # dict(encoding, delimiter, bom, newline)
        self.xl_book = None       # openpyxl workbook when loaded from xlsx
        self.xl_lost_features = []  # features the original file had that can't be kept
        self.names = {}           # defined names (upper) -> reference text
        self.active = 0
        self._changed = []
        self._dep_cell = {}
        self._dep_col = {}
        self._dep_wide = {}
        self._volatile = set()
        self._full_pending = False

    # ------------------------------------------------------------ sheets
    def get_sheet(self, name):
        u = name.upper()
        for s in self.sheets:
            if s.name.upper() == u:
                return s
        return None

    def unique_name(self, base="Sheet"):
        i = 1
        while self.get_sheet(f"{base}{i}"):
            i += 1
        return f"{base}{i}"

    def add_sheet(self, name=None, index=None):
        sh = Sheet(self, name or self.unique_name())
        if index is None:
            self.sheets.append(sh)
        else:
            self.sheets.insert(index, sh)
        self.rebuild_dependencies()
        self.recalc(full=True)
        return sh

    def remove_sheet(self, sh):
        for k, f in sh.formulas.items():
            self._unregister(sh, k, f)
        self.sheets.remove(sh)
        self.active = min(self.active, len(self.sheets) - 1)
        self.rebuild_dependencies()
        self.recalc(full=True)

    def rename_sheet(self, sh, new):
        old = sh.name
        for s in self.sheets:
            for k, f in list(s.formulas.items()):
                t = rename_sheet(f.text, old, new)
                if t != f.text:
                    s.formulas[k] = Formula(t, f.fallback)
        sh.name = new
        for s in self.sheets:  # PivotTables reading from the renamed sheet
            if any(p["source"] == old for p in s.pivots):
                s.pivots = [dict(p, source=new) if p["source"] == old else p for p in s.pivots]
        from . import charts as _charts
        for s in self.sheets:  # charts whose data is on the renamed sheet
            if s.charts:
                s.charts = [_charts.rename_sheet(c, old, new) for c in s.charts]
        for name, text in list(self.names.items()):
            self.names[name] = rename_sheet("=" + text, old, new)[1:]
        self.rebuild_dependencies()
        self.recalc(full=True)

    def move_sheet(self, sh, index):
        self.sheets.remove(sh)
        self.sheets.insert(index, sh)

    # ------------------------------------------------------------ dependencies
    def _targets(self, sh, f):
        for sname, r1, c1, r2, c2 in f.refs:
            tgt = sh if sname is None else self.get_sheet(sname)
            if tgt is not None:
                yield tgt, r1, c1, r2, c2

    def _register(self, sh, k, f):
        fk = (sh, k)
        if f.volatile:
            self._volatile.add(fk)
        for tgt, r1, c1, r2, c2 in self._targets(sh, f):
            if r1 == r2 and c1 == c2:
                self._dep_cell.setdefault((tgt, (r1 << 14) | c1), set()).add(fk)
            elif c2 - c1 < 64:
                for c in range(c1, c2 + 1):
                    self._dep_col.setdefault((tgt, c), {}).setdefault(fk, []).append((r1, r2))
            else:
                self._dep_wide.setdefault(tgt, {}).setdefault(fk, []).append((r1, c1, r2, c2))

    def _unregister(self, sh, k, f):
        fk = (sh, k)
        self._volatile.discard(fk)
        for tgt, r1, c1, r2, c2 in self._targets(sh, f):
            if r1 == r2 and c1 == c2:
                s = self._dep_cell.get((tgt, (r1 << 14) | c1))
                if s:
                    s.discard(fk)
            elif c2 - c1 < 64:
                for c in range(c1, c2 + 1):
                    d = self._dep_col.get((tgt, c))
                    if d:
                        d.pop(fk, None)
            else:
                d = self._dep_wide.get(tgt)
                if d:
                    d.pop(fk, None)

    def rebuild_dependencies(self):
        self._dep_cell = {}
        self._dep_col = {}
        self._dep_wide = {}
        self._volatile = set()
        for sh in self.sheets:
            for k, f in sh.formulas.items():
                self._register(sh, k, f)
            sh._rebuild_fcols()

    def dependents(self, sh, k):
        r, c = k >> 14, k & 0x3FFF
        out = set(self._dep_cell.get((sh, k), ()))
        d = self._dep_col.get((sh, c))
        if d:
            for fk, spans in d.items():
                for r1, r2 in spans:
                    if r1 <= r <= r2:
                        out.add(fk)
                        break
        d = self._dep_wide.get(sh)
        if d:
            for fk, rects in d.items():
                for r1, c1, r2, c2 in rects:
                    if r1 <= r <= r2 and c1 <= c <= c2:
                        out.add(fk)
                        break
        return out

    # ------------------------------------------------------------ recalculation
    def all_formulas(self):
        for sh in self.sheets:
            for k, f in sh.formulas.items():
                yield sh, k, f

    def recalc(self, full=False):
        """Recalculate formulas affected by cells changed since the last call."""
        changed = self._changed
        self._changed = []
        dirty = []
        if full or self._full_pending or len(changed) > 20000:
            self._full_pending = False
            for sh, k, f in self.all_formulas():
                f.dirty = True
                dirty.append((sh, k, f))
        else:
            queue = []
            for sh, k in changed:
                f = sh.formulas.get(k)
                if f is not None and not f.visiting:
                    f.dirty = True
                    dirty.append((sh, k, f))
                queue.append((sh, k))
            seen = set(queue)
            limit = 300000
            while queue:
                cell = queue.pop()
                for fk in self.dependents(*cell):
                    if fk in seen:
                        continue
                    seen.add(fk)
                    f = fk[0].formulas.get(fk[1])
                    if f is not None:
                        f.dirty = True
                        dirty.append((fk[0], fk[1], f))
                        queue.append(fk)
                if len(seen) > limit:
                    return self.recalc(full=True)
            for sh, k in list(self._volatile):
                f = sh.formulas.get(k)
                if f is not None:
                    f.dirty = True
                    dirty.append((sh, k, f))
        for sh, k, f in dirty:
            if f.dirty:
                self._ensure(sh, k, f)
        return dirty

    def _precedents(self, sh, f):
        for tgt, r1, c1, r2, c2 in self._targets(sh, f):
            if r1 == r2 and c1 == c2:
                k = (r1 << 14) | c1
                pf = tgt.formulas.get(k)
                if pf is not None:
                    yield tgt, k, pf
                continue
            fcols = tgt._fcols
            for c in range(c1, min(c2, tgt.max_col) + 1):
                rows = fcols.get(c)
                if not rows:
                    continue
                lo = bisect_left(rows, r1)
                hi = bisect_right(rows, r2)
                for r in rows[lo:hi]:
                    k = (r << 14) | c
                    pf = tgt.formulas.get(k)
                    if pf is not None:
                        yield tgt, k, pf

    def _ensure(self, sh, k, f):
        """Compute f after its dirty precedents (iteratively, no deep recursion)."""
        f.visiting = True
        stack = [(sh, k, f, self._precedents(sh, f))]
        run = self.evaluator.run
        try:
            while stack:
                s, kk, ff, it = stack[-1]
                pushed = False
                for ps, pk, pf in it:
                    if pf.dirty and not pf.visiting:
                        pf.visiting = True
                        stack.append((ps, pk, pf, self._precedents(ps, pf)))
                        pushed = True
                        break
                if pushed:
                    continue
                stack.pop()
                v = run(ff.ast, (s, kk >> 14, kk & 0x3FFF))
                if ff.fallback is not None and ff.unknown and isinstance(v, XLError):
                    v = ff.fallback
                ff.value = v
                ff.dirty = False
                ff.visiting = False
        finally:
            for s, kk, ff, it in stack:
                ff.visiting = False

    # ------------------------------------------------------------ undo support
    def apply_states(self, sh, states):
        """Set many cell states at once. Returns the previous states."""
        old = {}
        values, formulas, styles = sh.values, sh.formulas, sh.styles
        changed = self._changed
        mr, mc = sh.max_row, sh.max_col
        for k, st in states.items():
            f = formulas.get(k)
            if f is not None:
                prev = ("f", f.text)
            elif k in values:
                prev = ("v", values[k])
            else:
                prev = None
            prev_style = styles.get(k, DEFAULT_STYLE)
            old[k] = (prev, prev_style)
            content, style = st
            if sh.xl_styles and k in sh.xl_styles and prev_style != (style or DEFAULT_STYLE):
                sh.xl_styles.pop(k)
            if f is not None or (content is not None and content[0] == "f"):
                sh.set_content_key(k, content)
            else:
                if content is None or content[1] is None:
                    values.pop(k, None)
                else:
                    values[k] = content[1]
                    r, c = k >> 14, k & 0x3FFF
                    if r > mr:
                        mr = r
                    if c > mc:
                        mc = c
                changed.append((sh, k))
            if style is None or style is DEFAULT_STYLE or style == DEFAULT_STYLE:
                styles.pop(k, None)
            else:
                styles[k] = intern_style(style)
        sh.max_row = max(sh.max_row, mr)
        sh.max_col = max(sh.max_col, mc)
        self.recalc()
        return old

    def snapshot_sheet(self, sh):
        return {
            "values": sh.values.snapshot() if hasattr(sh.values, "snapshot") else dict(sh.values),
            "formulas": {k: (f.text, f.fallback) for k, f in sh.formulas.items()},
            "styles": dict(sh.styles),
            "col_widths": dict(sh.col_widths),
            "row_heights": dict(sh.row_heights),
            "hidden_rows": set(sh.hidden_rows),
            "hidden_cols": set(sh.hidden_cols),
            "merges": list(sh.merges),
            "freeze": sh.freeze,
            "freeze_origin": sh.freeze_origin,
            "autofilter": sh.autofilter,
            "filters": dict(sh.filters),
            "filter_hidden": set(sh.filter_hidden),
            "name": sh.name,
            "notes": dict(sh.notes),
            "xl_styles": dict(sh.xl_styles),
            "xl_dv": [(dv, list(rects)) for dv, rects in sh.xl_dv],
            "cond_formats": [(rule, list(rule.rects)) for rule in sh.cond_formats],
            "controls": list(sh.controls),
            "charts": list(sh.charts),
            "pivots": list(sh.pivots),
        }

    def snapshot(self, sheets=None):
        """Capture state for undo. `sheets`: sheets whose cells matter (default all);
        formula text of every sheet is always captured since structure changes rewrite them."""
        full = set(sheets) if sheets is not None else set(self.sheets)
        snap = {"order": list(self.sheets), "active": self.active, "full": {}, "ftext": {},
                "names": dict(self.names)}
        for sh in self.sheets:
            if sh in full:
                snap["full"][sh] = self.snapshot_sheet(sh)
            else:
                snap["ftext"][sh] = ({k: (f.text, f.fallback) for k, f in sh.formulas.items()}, sh.name,
                                     list(sh.pivots), list(sh.charts))  # a pivot / chart elsewhere may read from a changed sheet
        return snap

    def restore(self, snap):
        self.sheets = list(snap["order"])
        for sh in self.sheets:
            sh.wb = self
        for sh, d in snap["full"].items():
            sh.values = dict(d["values"]) if isinstance(d["values"], dict) else d["values"]
            sh.formulas = {k: Formula(t, fb) for k, (t, fb) in d["formulas"].items()}
            sh.styles = dict(d["styles"])
            sh.col_widths = dict(d["col_widths"])
            sh.row_heights = dict(d["row_heights"])
            sh.hidden_rows = set(d["hidden_rows"])
            sh.hidden_cols = set(d["hidden_cols"])
            sh.merges = list(d["merges"])
            sh.freeze = d["freeze"]
            sh.freeze_origin = d.get("freeze_origin", (0, 0))
            sh.autofilter = d["autofilter"]
            sh.filters = dict(d["filters"])
            sh.filter_hidden = set(d["filter_hidden"])
            sh.name = d["name"]
            sh.notes = dict(d["notes"])
            sh.xl_styles = dict(d["xl_styles"])
            sh.xl_dv = [[dv, list(rects)] for dv, rects in d["xl_dv"]]
            for rule, rects in d["cond_formats"]:
                rule.rects = list(rects)
                rule._asts = {}
            sh.cond_formats = [rule for rule, _ in d["cond_formats"]]
            sh.controls = list(d.get("controls", []))
            sh.charts = list(d.get("charts", []))
            sh.pivots = list(d.get("pivots", []))
            sh.recompute_extent()
        for sh, (ft, name, *rest) in snap["ftext"].items():
            sh.formulas = {k: Formula(t, fb) for k, (t, fb) in ft.items()}
            sh.name = name
            if rest:
                sh.pivots = list(rest[0])
            if len(rest) > 1:
                sh.charts = list(rest[1])
        self.active = min(snap["active"], len(self.sheets) - 1)
        self.names = dict(snap.get("names", self.names))
        self._changed = []
        self.rebuild_dependencies()
        self.recalc(full=True)


def new_workbook():
    wb = Workbook()
    wb.sheets.append(Sheet(wb, "Sheet1"))
    return wb
