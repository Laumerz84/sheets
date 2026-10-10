"""Workbook actions Claude can take, executed on the GUI thread of one window.

Every call returns a JSON-serialisable dict. All changes made while answering
one request are grouped into a single undo step ("Claude: ...")."""
import math
from builtins import range as _irange  # tool params are named "range"

from .. import ops
from ..errors import XLError
from ..formula import quote_sheet
from ..numfmt import format_value
from ..refs import MAX_COLS, MAX_ROWS, addr, col_index, key, parse_range, range_addr
from ..values import is_num
from ..workbook import DEFAULT_STYLE

MAX_READ_CELLS = 5000
MAX_WRITE_CELLS = 50000

BORDER_MODES = {"all", "outline", "inside", "none", "bottom", "top", "left", "right", "thick_outline"}
ALIGN = {"left", "center", "right", "general"}


class ToolError(Exception):
    pass


def _json_value(v):
    if isinstance(v, XLError):
        return v.code
    if isinstance(v, float) and v.is_integer() and abs(v) < 2 ** 53:
        return int(v)
    return v


class WorkbookTools:
    def __init__(self, win):
        self.win = win
        self.turn_label = None
        self.macro_open = False
        self.changed = []          # ranges touched this turn, for highlighting
        self.active = False        # a request is running
        self.stopped = False       # Stop was pressed: refuse anything still queued
        self.in_tool = False       # currently executing a tool (the undo guard lets these through)

    # ------------------------------------------------------------ turn / undo grouping
    def begin_turn(self, label):
        self.turn_label = label
        self.macro_open = False
        self.changed = []
        self.active = True
        self.stopped = False

    def end_turn(self):
        if self.macro_open:
            self.win.undo.endMacro()
        self.macro_open = False
        self.turn_label = None
        self.active = False

    def user_locked(self):
        """True while Claude works: the user's own edits would land inside Claude's undo step."""
        return self.active and not self.in_tool

    def _writing(self):
        if not self.macro_open and self.turn_label is not None:
            self.win.undo.beginMacro(f"Claude: {self.turn_label}"[:80])
            self.macro_open = True

    def _push(self, sh, states):
        """Apply cell states as part of this turn; no-op changes don't create an undo step."""
        clean = {k: st for k, st in states.items() if sh.get_state(k) != st}
        if not clean:
            return False
        self._writing()
        self.win._push_states(clean, "Claude", sheet=sh)
        return True

    # ------------------------------------------------------------ helpers
    @property
    def wb(self):
        return self.win.wb

    def _sheet(self, name):
        if not name:
            return self.win.grid.sheet
        sh = self.wb.get_sheet(name)
        if sh is None:
            raise ToolError(f"No sheet named {name!r}. Sheets: {', '.join(s.name for s in self.wb.sheets)}")
        return sh

    def _range(self, text, sheet=None):
        """'B2:D9' / 'Sheet2!A:A' -> (Sheet, (r1, c1, r2, c2))."""
        if not text:
            raise ToolError("A range is required, e.g. 'A1' or 'B2:D10'.")
        text = text.strip()
        if "!" in text:
            sname, text = text.rsplit("!", 1)
            sheet = sname.strip().strip("'").replace("''", "'")
        sh = self._sheet(sheet)
        rect = parse_range(text.replace("$", ""))
        if rect is None:
            raise ToolError(f"{text!r} is not a valid cell or range.")
        return sh, rect

    def _clamp(self, sh, rect):
        r1, c1, r2, c2 = rect
        ur, uc = sh.used_extent()
        if r2 >= MAX_ROWS - 1:
            r2 = max(r1, ur)
        if c2 >= MAX_COLS - 1:
            c2 = max(c1, uc)
        return (r1, c1, r2, c2)

    def _name(self, sh, rect):
        return f"{quote_sheet(sh.name)}!{range_addr(*rect)}"

    def _touched(self, sh, rect):
        self.changed.append((sh, rect))

    # ------------------------------------------------------------ tools
    def workbook_info(self):
        out = []
        for sh in self.wb.sheets:
            mr, mc = sh.used_extent(include_styles=False)
            info = {"name": sh.name, "used_range": range_addr(0, 0, mr, mc) if mr >= 0 else None,
                    "rows_used": mr + 1, "columns_used": mc + 1}
            if sh.freeze != (0, 0):
                info["frozen"] = {"rows": sh.freeze[0], "columns": sh.freeze[1]}
            if sh.autofilter:
                info["filter_range"] = range_addr(*sh.autofilter)
            if getattr(sh, "controls", None):
                from ..ui.controls import describe
                info["controls"] = [describe(c) for c in sh.controls]
            if getattr(sh, "charts", None):
                from .. import charts as _charts
                info["charts"] = [_charts.describe(sh, c) for c in sh.charts]
            out.append(info)
        g = self.win.grid
        r, c = g.sel.active
        return {"file": self.wb.path or "(unsaved new workbook)",
                "active_sheet": g.sheet.name,
                "selection": ", ".join(range_addr(*self._clamp(g.sheet, rc)) for rc in g.sel.rects),
                "active_cell": addr(r, c),
                "sheets": out}

    def read_range(self, range, sheet=None, include_formats=False):
        sh, rect = self._range(range, sheet)
        r1, c1, r2, c2 = self._clamp(sh, rect)
        n = (r2 - r1 + 1) * (c2 - c1 + 1)
        truncated = False
        if n > MAX_READ_CELLS:
            r2 = r1 + max(0, MAX_READ_CELLS // (c2 - c1 + 1)) - 1
            truncated = True
        values, shown, formulas, formats = [], [], {}, {}
        for r in _irange(r1, r2 + 1):
            row_v, row_s = [], []
            for c in _irange(c1, c2 + 1):
                k = key(r, c)
                v = sh.value_at_key(k)
                row_v.append(_json_value(v))
                st = sh.styles.get(k, DEFAULT_STYLE)
                row_s.append(format_value(v, st.numfmt)[0] if v is not None else "")
                f = sh.formulas.get(k)
                if f is not None:
                    formulas[addr(r, c)] = f.text
                if include_formats and st is not DEFAULT_STYLE:
                    formats[addr(r, c)] = {k2: v2 for k2, v2 in (
                        ("bold", st.bold), ("italic", st.italic), ("fill", st.fill), ("color", st.color),
                        ("number_format", st.numfmt if st.numfmt != "General" else None),
                        ("align", st.halign), ("wrap", st.wrap)) if v2}
            values.append(row_v)
            shown.append(row_s)
        out = {"range": self._name(sh, (r1, c1, r2, c2)), "values": values, "displayed": shown}
        if formulas:
            out["formulas"] = formulas
        if include_formats:
            out["formats"] = formats
        if truncated:
            out["note"] = f"Truncated to the first {r2 - r1 + 1} rows; read further down in separate calls."
        return out

    def write_range(self, start_cell, rows, sheet=None):
        sh, rect = self._range(start_cell, sheet)
        if not isinstance(rows, list) or not rows:
            raise ToolError("rows must be a non-empty list of rows, e.g. [[\"Name\", \"Total\"], [\"A\", 3]].")
        rows = [r if isinstance(r, list) else [r] for r in rows]
        width = max(len(r) for r in rows)
        if len(rows) * width > MAX_WRITE_CELLS:
            raise ToolError(f"Too many cells in one call (max {MAX_WRITE_CELLS}).")
        r0, c0 = rect[0], rect[1]
        if r0 + len(rows) > MAX_ROWS or c0 + width > MAX_COLS:
            raise ToolError("That would go past the edge of the sheet.")
        states = {}
        for i, row in enumerate(rows):
            for j, v in enumerate(row):
                r, c = r0 + i, c0 + j
                if v is None:
                    text = ""
                elif isinstance(v, bool):
                    text = "TRUE" if v else "FALSE"
                elif isinstance(v, (int, float)):
                    text = repr(float(v)) if isinstance(v, float) else str(v)
                else:
                    text = str(v)
                states[key(r, c)] = ops.input_state(sh, r, c, text)
        self._push(sh, states)
        written = (r0, c0, r0 + len(rows) - 1, c0 + width - 1)
        self._touched(sh, written)
        errors = {}
        for i in _irange(len(rows)):
            for j in _irange(width):
                v = sh.value(r0 + i, c0 + j)
                if isinstance(v, XLError):
                    errors[addr(r0 + i, c0 + j)] = v.code
        out = {"written": self._name(sh, written)}
        if errors:
            out["formula_errors"] = errors
        return out

    def format_range(self, range, sheet=None, bold=None, italic=None, underline=None, font_color=None,
                     fill_color=None, font_size=None, number_format=None, align=None, wrap=None,
                     borders=None):
        sh, rect = self._range(range, sheet)
        rect = self._clamp(sh, rect)
        kw = {}
        for name, val in (("bold", bold), ("italic", italic), ("underline", underline), ("wrap", wrap)):
            if val is not None:
                kw[name] = bool(val)
        for name, val in (("color", font_color), ("fill", fill_color)):
            if val is not None:
                kw[name] = _color(val)
        if font_size is not None:
            size = float(font_size)
            if not math.isfinite(size) or not 1 <= size <= 409:
                raise ToolError("font_size must be between 1 and 409.")
            kw["size"] = None if size == 11 else size
        if number_format is not None:
            kw["numfmt"] = number_format or "General"
        if align is not None:
            if align not in ALIGN:
                raise ToolError(f"align must be one of {sorted(ALIGN)}")
            kw["halign"] = None if align == "general" else align
        if borders is not None and borders not in BORDER_MODES:
            raise ToolError(f"borders must be one of {sorted(BORDER_MODES)}")
        if not kw and borders is None:
            raise ToolError("Nothing to change: pass at least one formatting option.")
        if kw:
            self._push(sh, ops.restyle(sh, [rect], lambda s: s.with_(**kw)))
        if borders is not None:
            self._push(sh, ops.borders_states(sh, rect, borders))
        self._touched(sh, rect)
        return {"formatted": self._name(sh, rect)}

    def clear_range(self, range, sheet=None, what="contents"):
        sh, rect = self._range(range, sheet)
        rect = self._clamp(sh, rect)
        fn = {"contents": ops.clear_contents, "formats": ops.clear_formats, "all": ops.clear_all}.get(what)
        if fn is None:
            raise ToolError("what must be 'contents', 'formats' or 'all'.")
        self._push(sh, fn(sh, [rect]))
        self._touched(sh, rect)
        return {"cleared": self._name(sh, rect), "what": what}

    def add_sheet(self, name, position=None):
        from ..ui.commands import SnapshotCommand
        name = (name or "").strip()
        if not name or len(name) > 31 or any(ch in name for ch in "[]:*?/\\"):
            raise ToolError("Sheet names need 1-31 characters and can't contain [ ] : * ? / \\")
        if self.wb.get_sheet(name):
            raise ToolError(f"A sheet named {name!r} already exists.")
        idx = len(self.wb.sheets) if position is None else max(0, min(int(position), len(self.wb.sheets)))
        current = self.win.grid.sheet
        self._writing()

        def action():
            self.wb.add_sheet(name, idx)
        self.win.undo.push(SnapshotCommand(self.win, current, "Claude", action, sheets=[]))
        self.win.show_sheet(current)  # keep the user where they were
        return {"added_sheet": name, "position": idx}

    def sort_range(self, range, column, sheet=None, ascending=True, has_header=True):
        sh, rect = self._range(range, sheet)
        rect = self._clamp(sh, rect)
        try:
            col = col_index(column.strip().upper().replace("$", "")) if isinstance(column, str) else int(column)
        except (ValueError, AttributeError):
            raise ToolError("column must be a column letter like 'C'.")
        if not rect[1] <= col <= rect[3]:
            raise ToolError("The sort column must be inside the range.")
        states = ops.sort_states(sh, rect, [(col, bool(ascending))], bool(has_header))
        if states:
            self._push(sh, states)
            self._touched(sh, rect)
        return {"sorted": self._name(sh, rect), "by_column": column, "ascending": bool(ascending),
                "changed": bool(states)}

    def insert_or_delete(self, action, at, count=1, sheet=None):
        from ..ui.commands import SnapshotCommand
        sh = self._sheet(sheet)
        count = int(count)
        if count < 1 or count > 10000:
            raise ToolError("count must be between 1 and 10000.")
        if action in ("insert_rows", "delete_rows"):
            try:
                pos = int(str(at).strip()) - 1
            except ValueError:
                b = parse_range(str(at))
                pos = b[0] if b else -1
            if not 0 <= pos < MAX_ROWS:
                raise ToolError("For rows, 'at' is a row number like 5.")
            fn = (lambda: sh.insert_rows(pos, count)) if action == "insert_rows" else (lambda: sh.delete_rows(pos, count))
            label = f"rows {pos + 1}-{pos + count}"
        elif action in ("insert_columns", "delete_columns"):
            try:
                pos = col_index(str(at).strip().upper())
            except Exception:
                raise ToolError("For columns, 'at' is a column letter like 'C'.")
            fn = (lambda: sh.insert_cols(pos, count)) if action == "insert_columns" else (lambda: sh.delete_cols(pos, count))
            label = f"columns starting at {str(at).upper()}"
        else:
            raise ToolError("action must be insert_rows, delete_rows, insert_columns or delete_columns.")
        self._writing()
        self.win.undo.push(SnapshotCommand(self.win, sh, "Claude", fn, sheets=[sh]))
        return {"done": action, "where": label, "sheet": sh.name}

    def find(self, text, sheet=None, match_case=False, whole_cell=False, max_results=200):
        sheets = [self._sheet(sheet)] if sheet else self.wb.sheets
        hits = []
        for sh in sheets:
            for r, c in ops.find_matches(sh, str(text), bool(match_case), bool(whole_cell), True):
                hits.append({"cell": f"{quote_sheet(sh.name)}!{addr(r, c)}",
                             "value": _json_value(sh.value(r, c)),
                             **({"formula": sh.formula_text(r, c)} if sh.formula_text(r, c) else {})})
                if len(hits) >= max_results:
                    return {"matches": hits, "note": f"Stopped at {max_results} matches."}
        return {"matches": hits}

    def set_column_width(self, columns, width=None, sheet=None):
        sh, rect = self._range(columns if ":" in columns else f"{columns}:{columns}", sheet)
        cols = list(_irange(rect[1], rect[3] + 1))
        if len(cols) > 500:
            raise ToolError("Too many columns at once.")
        if width not in (None, "auto"):
            w = float(width)
            if not math.isfinite(w) or not 0 <= w <= 255:
                raise ToolError("width must be between 0 and 255 characters.")
        current = self.win.grid.sheet
        self._writing()
        if current is not sh:
            self.win.show_sheet(sh)
        try:
            if width is None or width == "auto":
                self.win.autofit_cols(cols)
            else:
                px = int(round(w * 7 + 5)) if w > 0 else 0
                new = dict(sh.col_widths)
                for c in cols:
                    new[c] = px
                if new != sh.col_widths:
                    self.win._push_meta({"col_widths": (dict(sh.col_widths), new)}, "Claude")
        finally:
            if current is not sh:
                self.win.show_sheet(current)  # keep the user where they were
        return {"columns": range_addr(0, rect[1], MAX_ROWS - 1, rect[3]), "width": width or "auto"}

    def add_control(self, kind, linked_cell, place=None, min=0, max=100, step=1, sheet=None):
        """A slider ('scrollbar') or 'spinner' over `place` that drives the number in linked_cell."""
        from ..ui.commands import MetaCommand
        from ..ui.controls import default_place, describe, make_control
        sh, link_rect = self._range(linked_cell, sheet)
        if link_rect[0] != link_rect[2] or link_rect[1] != link_rect[3]:
            raise ToolError("linked_cell must be a single cell, e.g. 'B4'.")
        link = (link_rect[0], link_rect[1])
        if place:
            sh2, prect = self._range(place, sh.name)
            if sh2 is not sh:
                raise ToolError("The control must be on the same sheet as its linked cell.")
            if prect[2] >= MAX_ROWS - 1 or prect[3] >= MAX_COLS - 1:
                raise ToolError("place must be a cell or small range like 'C4:F4'.")
        else:
            prect = default_place(kind, link)
        try:
            ctl = make_control(kind, prect, link, min, max, step)
        except ValueError as e:
            raise ToolError(str(e))
        self._writing()
        self.win.undo.push(MetaCommand(self.win, sh, {"controls": (list(sh.controls), sh.controls + [ctl])},
                                       "Claude", relayout=False))
        self.win.grid.update()
        return {"added": describe(ctl), "sheet": sh.name}

    def add_chart(self, data_range, chart_type="column", title=None, place=None, series_in=None,
                  x_axis_title=None, y_axis_title=None, legend=None, data_labels=False, sheet=None):
        """An Excel-style chart linked to the cells in data_range (it updates when they change)."""
        from .. import charts as C
        from ..ui.commands import MetaCommand
        sh, rect = self._range(data_range, sheet)
        rect = C.clamp_rect(sh, self._clamp(sh, rect))
        key_ = str(chart_type or "column").strip().lower().replace(" ", "_").replace("-", "_")
        ctype = CHART_TYPE_NAMES.get(key_)
        if ctype is None:
            raise ToolError(f"Unknown chart_type {chart_type!r}. Use one of: {', '.join(sorted(CHART_TYPE_NAMES))}.")
        by = {None: None, "": None, "columns": "cols", "cols": "cols", "rows": "rows"}.get(
            (series_in or "").strip().lower() if series_in else None, "?")
        if by == "?":
            raise ToolError("series_in must be 'columns' or 'rows' (or left out for Excel's automatic choice).")
        home = sh
        if place:
            home, prect = self._range(place, sh.name)
            if prect[2] >= MAX_ROWS - 1 or prect[3] >= MAX_COLS - 1:
                raise ToolError("place must be a cell like 'H2' or a small range like 'H2:N18'.")
            if prect[0] == prect[2] and prect[1] == prect[3]:
                anchors = C.place_at_cell(home, prect[0], prect[1])
            else:
                anchors = ([prect[0], prect[1], 0, 0], [prect[2] + 1, prect[3] + 1, 0, 0])
        else:
            anchors = C.free_place(sh, rect[0], rect[3] + 2)
        try:
            ch = C.make_chart(sh, rect, ctype, by, anchors=anchors, home=home)
        except ValueError as e:
            raise ToolError(str(e))
        if title is not None:
            ch["title"] = title
        if x_axis_title:
            ch["x_title"] = x_axis_title
        if y_axis_title:
            ch["y_title"] = y_axis_title
        if legend:
            if legend not in C.LEGENDS:
                raise ToolError(f"legend must be one of {', '.join(C.LEGENDS)}.")
            ch["legend"] = legend
        ch["labels"] = bool(data_labels)
        self._writing()
        self.win.undo.push(MetaCommand(self.win, home, {"charts": (list(home.charts), home.charts + [ch])},
                                       "Claude", relayout=False))
        self.win.grid.update()
        return {"added": C.describe(home, ch), "sheet": home.name}

    def select_range(self, range, sheet=None):
        sh, rect = self._range(range, sheet)
        self.win.show_sheet(sh)
        self.win.grid.set_selection([rect], active=(rect[0], rect[1]))
        self.win.grid.ensure_visible(rect[0], rect[1])
        return {"selected": self._name(sh, rect)}

    # ------------------------------------------------------------ dispatch
    def call(self, name, args):
        fn = TOOLS.get(name)
        if fn is None:
            return {"error": f"Unknown tool {name!r}."}
        if not self.active or self.stopped:
            return {"error": "The user stopped this request; make no further changes."}
        self.in_tool = True
        try:
            return getattr(self, fn)(**(args or {}))
        except ToolError as e:
            return {"error": str(e)}
        except TypeError as e:
            return {"error": f"Bad arguments for {name}: {e}"}
        except (ValueError, OverflowError) as e:
            return {"error": f"Bad value for {name}: {e}"}
        except Exception as e:  # never let a tool bug take the app down
            return {"error": f"{type(e).__name__}: {e}"}
        finally:
            self.in_tool = False


TOOLS = {n: n for n in ("workbook_info", "read_range", "write_range", "format_range", "clear_range",
                        "add_sheet", "sort_range", "insert_or_delete", "find", "set_column_width",
                        "select_range", "add_control", "add_chart")}

# names Claude may use for chart_type -> charts.TYPES keys
CHART_TYPE_NAMES = {
    "column": "col", "clustered_column": "col", "stacked_column": "col_stacked", "percent_column": "col_pct",
    "100%_stacked_column": "col_pct", "bar": "bar", "clustered_bar": "bar", "stacked_bar": "bar_stacked",
    "percent_bar": "bar_pct", "line": "line", "line_markers": "line_markers", "line_with_markers": "line_markers",
    "pie": "pie", "doughnut": "doughnut", "area": "area", "stacked_area": "area_stacked",
    "percent_area": "area_pct", "scatter": "scatter", "scatter_lines": "scatter_lines",
    "scatter_smooth": "scatter_smooth", "combo": "combo", "combo_secondary": "combo_sec",
}

_NAMED = {"red": "#FF0000", "green": "#00B050", "blue": "#0070C0", "yellow": "#FFFF00", "orange": "#FFC000",
          "purple": "#7030A0", "black": "#000000", "white": "#FFFFFF", "gray": "#808080", "grey": "#808080",
          "lightgray": "#D9D9D9", "lightgrey": "#D9D9D9", "lightblue": "#DDEBF7", "lightgreen": "#E2EFDA",
          "lightyellow": "#FFF2CC", "lightred": "#FFC7CE", "darkblue": "#1F4E79", "darkgreen": "#375623"}


def _color(val):
    if val in ("", "none", None):
        return None
    v = str(val).strip()
    if v.lower().replace(" ", "") in _NAMED:
        return _NAMED[v.lower().replace(" ", "")]
    if not v.startswith("#"):
        v = "#" + v
    if len(v) != 7 or any(ch not in "0123456789abcdefABCDEF" for ch in v[1:]):
        raise ToolError(f"Colors are '#RRGGBB' (e.g. '#FFF2CC') or a basic name; got {val!r}.")
    return v.upper()
