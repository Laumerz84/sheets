"""Charts, like Excel's: the data model (no Qt in this file).

A chart floats over the grid of one sheet and is stored in `Sheet.charts` as a plain dict
(never mutated in place: edits replace the list, which MetaCommand undoes). The chart is linked
to cells by references (series name / values / categories, e.g. "=Sheet1!$B$2:$B$9"), so it
follows the cells: edits, undo/redo, inserted/deleted rows and renamed sheets all show up in it.

Where it sits is a two-cell anchor like Excel's "move and size with cells":
`from` / `to` = [row, col, dx, dy] (the offset is in pixels at 100% zoom inside that cell).

Drawing is in ui/chart_paint.py, the mouse/UI in ui/charts_ui.py, xlsx in chart_xlsx.py.
Big-file sheets (bigdata.BigSheet) are never read cell by cell: a long range is sampled to a
couple of thousand points (see MAX_POINTS)."""
import copy
import math
import re
import uuid

from .formula import quote_sheet
from .numfmt import format_value
from .refs import MAX_COLS, MAX_ROWS, parse_range, range_addr
from .values import is_num

MAX_POINTS = 2000        # points per series that are drawn; longer ranges are sampled
MAX_PIE_POINTS = 120     # a pie / doughnut with more slices than this is refused
DEFAULT_SIZE = (480, 288)   # px at 100% zoom, like Excel's 5 x 3 inch default
MIN_SIZE = (60, 40)

# key -> (label, group). Order = gallery order.
TYPES = {
    "col": ("Clustered Column", "Column"),
    "col_stacked": ("Stacked Column", "Column"),
    "col_pct": ("100% Stacked Column", "Column"),
    "line": ("Line", "Line"),
    "line_markers": ("Line with Markers", "Line"),
    "pie": ("Pie", "Pie"),
    "doughnut": ("Doughnut", "Pie"),
    "bar": ("Clustered Bar", "Bar"),
    "bar_stacked": ("Stacked Bar", "Bar"),
    "bar_pct": ("100% Stacked Bar", "Bar"),
    "area": ("Area", "Area"),
    "area_stacked": ("Stacked Area", "Area"),
    "area_pct": ("100% Stacked Area", "Area"),
    "scatter": ("Scatter", "X Y (Scatter)"),
    "scatter_smooth": ("Scatter with Smooth Lines and Markers", "X Y (Scatter)"),
    "scatter_lines": ("Scatter with Straight Lines and Markers", "X Y (Scatter)"),
    "combo": ("Clustered Column - Line", "Combo"),
    "combo_sec": ("Clustered Column - Line on Secondary Axis", "Combo"),
}
GROUPS = ["Column", "Line", "Pie", "Bar", "Area", "X Y (Scatter)", "Combo"]
SERIES_TYPES = ("col", "line", "line_markers", "area", "scatter", "scatter_lines")  # allowed per series in a combo
LEGENDS = ("right", "left", "top", "bottom", "none")

# ---------------------------------------------------------------- palettes ("chart styles > color")
PALETTES = {
    1: ("Colorful 1", ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47"]),
    2: ("Colorful 2", ["#ED7D31", "#FFC000", "#70AD47", "#4472C4", "#A5A5A5", "#5B9BD5"]),
    3: ("Colorful 3", ["#70AD47", "#5B9BD5", "#FFC000", "#ED7D31", "#4472C4", "#A5A5A5"]),
    4: ("Colorful 4", ["#7030A0", "#C00000", "#00B0F0", "#92D050", "#FFC000", "#A5A5A5"]),
    5: ("Monochromatic Blue", "#4472C4"),
    6: ("Monochromatic Orange", "#ED7D31"),
    7: ("Monochromatic Green", "#70AD47"),
    8: ("Monochromatic Gray", "#7F7F7F"),
}
# "Chart Styles" (looks): which options each one sets
LOOKS = {
    1: ("Style 1: default", {"grid_y": True, "labels": False, "fill": "#FFFFFF"}),
    2: ("Style 2: data labels, no gridlines", {"grid_y": False, "labels": True, "fill": "#FFFFFF"}),
    3: ("Style 3: gridlines on both axes", {"grid_y": True, "grid_x": True, "labels": False, "fill": "#FFFFFF"}),
    4: ("Style 4: dark", {"grid_y": True, "labels": False, "fill": "#262626"}),
    5: ("Style 5: light gray", {"grid_y": True, "labels": False, "fill": "#F2F2F2"}),
}


def _hex_rgb(h):
    h = h.lstrip("#")
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def _rgb_hex(r, g, b):
    return "#%02X%02X%02X" % (max(0, min(255, int(round(r)))), max(0, min(255, int(round(g)))),
                              max(0, min(255, int(round(b)))))


def mix(a, b, t):
    """Colour a blended t of the way towards b (both '#RRGGBB')."""
    ra, ga, ba = _hex_rgb(a)
    rb, gb, bb = _hex_rgb(b)
    return _rgb_hex(ra + (rb - ra) * t, ga + (gb - ga) * t, ba + (bb - ba) * t)


def palette_colors(style, n):
    """n series / slice colours for a chart style number."""
    name, spec = PALETTES.get(style, PALETTES[1])
    if isinstance(spec, list):
        return [spec[i % len(spec)] if i < len(spec) else mix(spec[i % len(spec)], "#000000", 0.25 * (i // len(spec)))
                for i in range(n)]
    out = []
    for i in range(n):          # shades of one colour, dark to light like Excel's monochromatic palettes
        t = 0.0 if n == 1 else i / (n - 1)
        out.append(mix(mix(spec, "#000000", 0.35), mix(spec, "#FFFFFF", 0.65), t))
    return out


def is_dark(color):
    r, g, b = _hex_rgb(color)
    return (0.299 * r + 0.587 * g + 0.114 * b) < 128


# ---------------------------------------------------------------- types
def family(t):
    if t.startswith("col"):
        return "col"
    for f in ("bar", "line", "pie", "doughnut", "area", "scatter", "combo"):
        if t.startswith(f):
            return f
    return "col"


def grouping(t):
    if t.endswith("_stacked"):
        return "stacked"
    if t.endswith("_pct"):
        return "percent"
    return "clustered"


def type_label(t):
    return TYPES.get(t, (t, ""))[0]


# ---------------------------------------------------------------- defaults / normalising
DEFAULTS = {
    "type": "col", "title": "", "legend": "right", "labels": False, "label_pct": False,
    "x_title": "", "y_title": "", "y2_title": "", "grid_y": True, "grid_x": False,
    "y_min": None, "y_max": None, "y_major": None, "x_min": None, "x_max": None,
    "y2_min": None, "y2_max": None, "y_fmt": None, "x_reverse": False, "log_y": False,
    "style": 1, "look": 1, "fill": "#FFFFFF", "border": True, "gap": None, "hole": 60,
    "range": "", "by": "cols", "name": "",
}
SERIES_DEFAULTS = {"name": "", "values": "", "cats": "", "color": None, "type": None, "secondary": False,
                   "labels": None, "smooth": False, "marker": None}


def new_id():
    return uuid.uuid4().hex[:10]


def normalize(d):
    """A complete, valid chart dict from a partial / saved one (raises ValueError when hopeless)."""
    if not isinstance(d, dict):
        raise ValueError("chart must be an object")
    ch = dict(DEFAULTS)
    ch.update({k: v for k, v in d.items() if k in DEFAULTS})
    if ch["type"] not in TYPES:
        raise ValueError(f"unknown chart type {ch['type']!r}")
    if ch["type"] == "combo_sec":
        ch["type"] = "combo"
    if ch["legend"] not in LEGENDS:
        ch["legend"] = "right"
    ch["id"] = str(d.get("id") or new_id())
    for k in ("from", "to"):
        v = d.get(k)
        if not v or len(v) != 4:
            raise ValueError("chart needs from / to anchors")
        r, c, dx, dy = v
        ch[k] = [max(0, min(MAX_ROWS - 1, int(r))), max(0, min(MAX_COLS - 1, int(c))),
                 max(0, int(dx)), max(0, int(dy))]
    ch["style"] = ch["style"] if ch["style"] in PALETTES else 1
    ch["hole"] = max(10, min(90, int(ch["hole"] or 60)))
    for k in ("y_min", "y_max", "y_major", "x_min", "x_max", "y2_min", "y2_max"):
        v = ch[k]
        ch[k] = float(v) if v not in (None, "") else None
    ch["series"] = []
    for s in d.get("series", []):
        sd = dict(SERIES_DEFAULTS)
        sd.update({k: v for k, v in s.items() if k in SERIES_DEFAULTS})
        if sd["type"] is not None and sd["type"] not in SERIES_TYPES:
            sd["type"] = None
        ch["series"].append(sd)
    return ch


def chart_name(sheet_charts, base="Chart"):
    used = {c.get("name") for c in sheet_charts}
    i = 1
    while f"{base} {i}" in used:
        i += 1
    return f"{base} {i}"


# ---------------------------------------------------------------- references
_REF_RE = re.compile(r"^=?\s*(?:(?:'((?:[^']|'')+)'|([^!'=:]+))!)?\s*(.+?)\s*$")


def make_ref(sheet_name, rect):
    r1, c1, r2, c2 = rect
    a = f"${_col(c1)}${r1 + 1}" + ("" if (r1 == r2 and c1 == c2) else f":${_col(c2)}${r2 + 1}")
    return f"={quote_sheet(sheet_name)}!{a}"


def _col(c):
    from .refs import col_name
    return col_name(c)


def parse_ref(text, default_sheet=None):
    """'=Sheet1!$B$2:$B$9' -> (sheet name or default, rect) or None."""
    if not text or not isinstance(text, str):
        return None
    m = _REF_RE.match(text)
    if not m:
        return None
    name = (m.group(1).replace("''", "'") if m.group(1) else m.group(2)) or default_sheet
    body = m.group(3).replace("$", "")
    rect = parse_range(body)
    if rect is None or name is None:
        return None
    return name.strip() if name else name, rect


def ref_rect(wb, text, default_sheet=None):
    """-> (Sheet, rect) or None (bad reference / missing sheet)."""
    p = parse_ref(text, default_sheet)
    if p is None:
        return None
    sh = wb.get_sheet(p[0])
    return (sh, p[1]) if sh is not None else None


def ref_names_sheet(text, name):
    p = parse_ref(text)
    return p is not None and p[0].upper() == name.upper()


def rename_in_ref(text, old, new):
    p = parse_ref(text)
    if not text or not text.startswith("=") or p is None or p[0].upper() != old.upper():
        return text
    return make_ref(new, p[1])


# ---------------------------------------------------------------- pixel anchors
def _sizes(sh, axis):
    if axis == "col":
        return sh.col_widths, sh.hidden_cols, 64
    return sh.row_heights, set(sh.hidden_rows) | set(sh.filter_hidden), 20


def axis_size(sh, axis, i):
    sizes, hidden, default = _sizes(sh, axis)
    return 0 if i in hidden else sizes.get(i, default)


def axis_pos(sh, axis, idx):
    """Pixel position (100% zoom) of the start of row / column idx."""
    sizes, hidden, default = _sizes(sh, axis)
    pos = idx * default
    for i, s in sizes.items():
        if i < idx and i not in hidden:
            pos += s - default
    for i in hidden:
        if i < idx:
            pos -= sizes.get(i, default)
    return pos


def px_to_anchor(sh, axis, p):
    """Pixel position -> (index, offset in that cell)."""
    sizes, hidden, default = _sizes(sh, axis)
    p = max(0, p)
    idx = int(p // default)
    # correct for custom sizes: walk from a lower guess until the position is inside the cell
    lo, hi = 0, (MAX_ROWS if axis == "row" else MAX_COLS) - 1
    idx = max(lo, min(hi, idx))
    while idx > 0 and axis_pos(sh, axis, idx) > p:
        idx = max(0, idx - max(1, int((axis_pos(sh, axis, idx) - p) // max(default, 1))))
    while idx < hi and axis_pos(sh, axis, idx) + axis_size(sh, axis, idx) <= p:
        idx += max(1, int((p - axis_pos(sh, axis, idx)) // max(default, 1)))
    idx = min(idx, hi)
    return idx, max(0, int(p - axis_pos(sh, axis, idx)))


def anchor_px(sh, anchor):
    """[r, c, dx, dy] -> (x, y) at 100% zoom."""
    r, c, dx, dy = anchor
    return (axis_pos(sh, "col", c) + min(dx, max(axis_size(sh, "col", c), 0)),
            axis_pos(sh, "row", r) + min(dy, max(axis_size(sh, "row", r), 0)))


def chart_px_rect(sh, ch):
    x1, y1 = anchor_px(sh, ch["from"])
    x2, y2 = anchor_px(sh, ch["to"])
    return x1, y1, max(MIN_SIZE[0], x2 - x1), max(MIN_SIZE[1], y2 - y1)


def anchors_for(sh, x, y, w, h):
    """Pixel rect (100% zoom, sheet coordinates) -> (from, to) anchors."""
    c1, dx1 = px_to_anchor(sh, "col", x)
    r1, dy1 = px_to_anchor(sh, "row", y)
    c2, dx2 = px_to_anchor(sh, "col", x + w)
    r2, dy2 = px_to_anchor(sh, "row", y + h)
    return [r1, c1, dx1, dy1], [r2, c2, dx2, dy2]


def place_at_cell(sh, row, col, w=DEFAULT_SIZE[0], h=DEFAULT_SIZE[1]):
    x, y = axis_pos(sh, "col", col), axis_pos(sh, "row", row)
    return anchors_for(sh, x, y, w, h)


def overlapping(sh, x, y, w, h):
    """The first existing chart on `sh` that the pixel rect (100% zoom) overlaps, as (x, y, w, h), or None."""
    for c in sh.charts:
        cx, cy, cw, chh = chart_px_rect(sh, c)
        if x < cx + cw and x + w > cx and y < cy + chh and y + h > cy:
            return cx, cy, cw, chh
    return None


def free_place(sh, row, col, w=DEFAULT_SIZE[0], h=DEFAULT_SIZE[1]):
    """Anchors for a new chart at a cell, moved down past any chart already there."""
    x, y = axis_pos(sh, "col", col), axis_pos(sh, "row", row)
    for _ in range(60):
        hit = overlapping(sh, x, y, w, h)
        if hit is None:
            break
        y = hit[1] + hit[3] + 16
    return anchors_for(sh, x, y, w, h)


# ---------------------------------------------------------------- the default chart for a selection
def _val(sh, r, c):
    return sh.value(r, c)


_DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}")


def _read_values(sh, cells):
    """Values of the (row, col) cells. On a big-file sheet a long run down one column is read with a
    single pyarrow `take` of the sampled rows (plus the user's edits on top), never cell by cell."""
    big = getattr(sh, "big", None)
    if big is None or len(cells) < 24 or len({c for _, c in cells}) != 1 or cells[0][1] >= big.ncols             or cells[0][1] in big.calc_defs:
        return [sh.value(r, c) for r, c in cells]
    from . import bigdata as BD
    pa, pc = BD.arrow()
    c = cells[0][1]
    nv = big.nview
    out = [None] * len(cells)
    inside = [i for i, (r, _) in enumerate(cells) if r < nv]
    if inside:
        drs = [int(big.drow(cells[i][0])) for i in inside]
        with big.lock:
            arr = pc.take(big.cols[c], pa.array(drs, pa.int64()))
            if isinstance(arr, pa.ChunkedArray):
                arr = arr.combine_chunks()
            nums = BD._to_numbers(arr).to_pylist()
        raws = arr.to_pylist()
        edits = big.edits
        for k, i in enumerate(inside):
            raw, num = raws[k], nums[k]
            if raw is None or raw == "":
                v = None
            elif num is not None and not _DATE_RE.match(raw):
                v = num
            else:
                v = raw
            e = edits.get((drs[k] << 14) | c, _MISSING)
            if e is not _MISSING:
                v = None if e is BD.CLEARED else e
            out[i] = v
    fm = sh.formulas
    for i, (r, cc) in enumerate(cells):
        if r >= nv or (fm and ((r << 14) | cc) in fm):
            out[i] = sh.value(r, cc)
    return out


_MISSING = object()


def _is_text(v):
    return isinstance(v, str) and v != ""


def _is_dateish(sh, r, c, v):
    if not is_num(v):
        return False
    fmt = sh.style(r, c).numfmt or "General"
    return any(ch in fmt.lower() for ch in ("yy", "dd", "mmm")) and "0" not in fmt


def _yearish(vals):
    return bool(vals) and all(is_num(v) and float(v).is_integer() and 1800 <= v <= 2200 for v in vals) \
        and len(vals) >= 2 and all(vals[i] < vals[i + 1] for i in range(len(vals) - 1))


def clamp_rect(sh, rect):
    r1, c1, r2, c2 = rect
    ur, uc = sh.max_row, sh.max_col
    if ur < 0 or uc < 0:
        ur, uc = sh.used_extent(include_styles=False)
    big = getattr(sh, "big", None)
    if big is not None:
        ur, uc = max(ur, big.nview - 1), max(uc, big.ncols - 1)
    return r1, c1, max(r1, min(r2, max(ur, r1))), max(c1, min(c2, max(uc, c1)))


def plan(sh, rect, by=None, scatter=False):
    """Excel's rules for turning a block of cells into series. Returns
    {"by", "series": [(name_cell|None, values_rect, name_text)], "cats": rect|None, "header_row", "cat_col"}
    or None when there's nothing to plot. Looks at the first rows only (cheap on huge sheets)."""
    r1, c1, r2, c2 = clamp_rect(sh, rect)
    nr, nc = r2 - r1 + 1, c2 - c1 + 1
    if by not in ("cols", "rows"):
        by = "rows" if nc > nr else "cols"
    rows_mode = by == "rows"

    def cell(i, j):        # i = line index (series), j = point index; in the oriented matrix
        return (r1 + i, c1 + j) if rows_mode else (r1 + j, c1 + i)
    n_lines = nr if rows_mode else nc      # candidate series
    n_pts = nc if rows_mode else nr        # candidate points
    sample = min(n_pts, 60)

    def v(i, j):
        r, c = cell(i, j)
        return r, c, _val(sh, r, c)

    # oriented: matrix M[j][i] -> point j, line i. header = point 0, categories = line 0
    def col_vals(i, j0, j1):
        return [v(i, j) for j in range(j0, min(j1, n_pts - 1) + 1)]

    tl_blank = v(0, 0)[2] in (None, "")
    header = cats = False
    if n_lines >= 2 and n_pts >= 2:
        first_line = col_vals(0, 1, sample)
        textish = [x for x in first_line if x[2] not in (None, "")]
        dates = textish and all(_is_text(x[2]) or _is_dateish(sh, x[0], x[1], x[2]) for x in textish)
        yearish = _yearish([x[2] for x in textish])
        cats = bool(textish) and (all(_is_text(x[2]) for x in textish) or bool(dates) or yearish or scatter
                                  and all(is_num(x[2]) for x in textish))
        if tl_blank and textish:
            cats = True
        start = 1 if cats else 0
        hdr_vals = [v(i, 0)[2] for i in range(start, n_lines)]
        header = all(_is_text(x) for x in hdr_vals) and bool(hdr_vals)
        if tl_blank:
            header = True
    elif n_lines >= 1 and n_pts >= 2:
        # one line only: header if its first cell is text and the rest numbers
        first = v(0, 0)[2]
        rest = [x[2] for x in col_vals(0, 1, sample)]
        header = _is_text(first) and any(is_num(x) for x in rest)
    elif n_pts == 1 and n_lines >= 2:
        # a single point across: a header cell may lead
        header = False
    j0 = 1 if header else 0
    i0 = 1 if cats else 0
    if j0 >= n_pts or i0 >= n_lines:
        return None
    series = []
    for i in range(i0, n_lines):
        vals_rect = _line_rect(cell, i, j0, n_pts - 1)
        name_cell = cell(i, 0) if header else None
        series.append((name_cell, vals_rect))
    cats_rect = _line_rect(cell, 0, j0, n_pts - 1) if cats else None
    # drop completely empty series
    keep = []
    for name_cell, vr in series:
        probe = _probe(sh, vr)
        if probe:
            keep.append((name_cell, vr))
    if not keep:
        return None
    return {"by": by, "series": keep, "cats": cats_rect, "header": header, "cat_col": cats}


def _line_rect(cell, i, j0, j1):
    a, b = cell(i, j0), cell(i, j1)
    return (a[0], a[1], b[0], b[1])


def _probe(sh, rect, limit=400):
    """Does the range hold any content (looks at up to `limit` cells)?"""
    r1, c1, r2, c2 = rect
    n = 0
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            if sh.value(r, c) not in (None, ""):
                return True
            n += 1
            if n >= limit:
                return True      # long and not obviously empty: assume yes
    return False


def build_series(sh, pl, ctype):
    """Series dicts for a plan (scatter: first category column is X)."""
    out = []
    cats = make_ref(sh.name, pl["cats"]) if pl["cats"] else ""
    for n, (name_cell, vr) in enumerate(pl["series"]):
        s = dict(SERIES_DEFAULTS)
        s["name"] = make_ref(sh.name, (name_cell[0], name_cell[1], name_cell[0], name_cell[1])) if name_cell \
            else f"Series{n + 1}"
        s["values"] = make_ref(sh.name, vr)
        s["cats"] = cats
        out.append(s)
    return out


def apply_type(ch, ctype):
    """Chart dict with its type changed (combo presets fill in the series types)."""
    ch = dict(ch)
    preset_sec = ctype == "combo_sec"
    if preset_sec:
        ctype = "combo"
    ch["type"] = ctype
    series = [dict(s) for s in ch["series"]]
    if ctype == "combo":
        for i, s in enumerate(series):
            if preset_sec or s.get("type") is None:
                s["type"] = "col" if i == 0 else "line"
            if preset_sec:
                s["secondary"] = i > 0
        if len(series) == 1:
            series[0]["type"] = "col"
            series[0]["secondary"] = False
    else:
        for s in series:
            s["type"] = None
            s["secondary"] = False
    if ctype.startswith("scatter"):
        for s in series:
            s["smooth"] = ctype == "scatter_smooth"
    ch["series"] = series
    return ch


def make_chart(sh, rect, ctype="col", by=None, anchors=None, title=None, home=None):
    """A new chart for the cells in `rect` of sheet `sh` (the Excel default rules).
    Raises ValueError when the cells hold nothing to plot."""
    if ctype not in TYPES:
        raise ValueError(f"unknown chart type {ctype!r}")
    scatter = ctype.startswith("scatter")
    pl = plan(sh, rect, by, scatter)
    if pl is None:
        raise ValueError("There's no data to plot in that range. Select cells with numbers (and optionally "
                         "a header row and a first column of labels).")
    if scatter and pl["cats"] is None and len(pl["series"]) >= 2:
        # Excel: first column of numbers = X values
        first, rest = pl["series"][0], pl["series"][1:]
        pl = dict(pl, series=rest, cats=first[1])
    ch = dict(DEFAULTS)
    ch["type"] = "combo" if ctype == "combo_sec" else ctype
    ch["by"] = pl["by"]
    ch["series"] = build_series(sh, pl, ctype)
    r1, c1, r2, c2 = clamp_rect(sh, rect)
    ch["range"] = make_ref(sh.name, (r1, c1, r2, c2))
    ch["id"] = new_id()
    ch["from"], ch["to"] = anchors if anchors else free_place(home or sh, r1, c2 + 2)
    if title is not None:
        ch["title"] = title
    ch = apply_type(ch, ctype)
    fam = family(ctype)
    if fam in ("pie", "doughnut"):
        ch["legend"] = "right"
        ch["grid_y"] = False
    else:
        ch["legend"] = "none" if len(ch["series"]) == 1 else "right"
    ch["name"] = chart_name((home or sh).charts)
    return normalize(ch)


def switch_by(sh, ch):
    """Select Data > Switch Row/Column: rebuild the series from the chart's data range the other way."""
    got = parse_ref(ch.get("range", ""))
    if got is None:
        raise ValueError("This chart has no single data range to switch.")
    src = sh.wb.get_sheet(got[0])
    if src is None:
        raise ValueError("The chart's data sheet is missing.")
    new_by = "rows" if ch.get("by") == "cols" else "cols"
    return rebuild_from_range(src, ch, got[1], new_by)


def rebuild_from_range(src, ch, rect, by=None):
    """The chart's series rebuilt from a data range (colours of series that keep their position stay)."""
    scatter = ch["type"].startswith("scatter")
    pl = plan(src, rect, by, scatter)
    if pl is None:
        raise ValueError("There's no data to plot in that range.")
    if scatter and pl["cats"] is None and len(pl["series"]) >= 2:
        first, rest = pl["series"][0], pl["series"][1:]
        pl = dict(pl, series=rest, cats=first[1])
    series = build_series(src, pl, ch["type"])
    old = ch["series"]
    for i, s in enumerate(series):
        if i < len(old):
            for k in ("color", "type", "secondary", "labels", "smooth", "marker"):
                s[k] = old[i].get(k)
    out = dict(ch, series=series, by=pl["by"], range=make_ref(src.name, clamp_rect(src, rect)))
    return apply_type(out, ch["type"]) if ch["type"] == "combo" else out


# ---------------------------------------------------------------- data for drawing
def _hidden_lines(sh, axis, lo, hi):
    _, hidden, _ = _sizes(sh, axis)
    return sorted(i for i in hidden if lo <= i <= hi)


def _cells_of(sh, rect, limit):
    """([(r, c), ...] to read, step) for a range; hidden rows / columns are skipped and a long
    range is sampled evenly down to `limit` cells (never reads the whole of a huge range)."""
    r1, c1, r2, c2 = rect
    if r1 == r2 or c1 == c2:
        vertical = c1 == c2 and r1 != r2
        lo, hi = (r1, r2) if vertical else (c1, c2)
        hid = _hidden_lines(sh, "row" if vertical else "col", lo, hi)
        hs = set(hid) if len(hid) <= 200_000 else set()
        total = hi - lo + 1
        if hs and total <= 2_000_000:
            idx = [i for i in range(lo, hi + 1) if i not in hs]     # exact: skip the hidden lines
            n = len(idx)
        else:
            idx = range(lo, hi + 1)                                 # lazy; never a list of millions
            n = total
        step = 1
        if n > limit:
            step = -(-n // limit)
            idx = idx[::step]
        if hs and isinstance(idx, range):
            idx = [i for i in idx if i not in hs]                   # only the sampled few
        cells = [(i, c1) for i in idx] if vertical else [(r1, i) for i in idx]
        return cells, step
    cells = []
    hr = set(_hidden_lines(sh, "row", r1, r2))
    hc = set(_hidden_lines(sh, "col", c1, c2))
    for r in range(r1, r2 + 1):
        if r in hr:
            continue
        for c in range(c1, c2 + 1):
            if c in hc:
                continue
            cells.append((r, c))
            if len(cells) >= limit:
                return cells, 1
    return cells, 1


def _text_of(sh, r, c, v):
    if v is None:
        return ""
    if is_num(v):
        if getattr(sh, "big", None) is not None:
            return format_value(v, "General")[0]     # (no style lookup per sampled cell)
        return format_value(v, sh.style(r, c).numfmt or "General")[0]
    return v if isinstance(v, str) else format_value(v)[0]


def resolve(wb, ch, home=None, limit=MAX_POINTS):
    """Read the cells a chart points at -> dict for the painter:
    {"series": [{"name", "y": [float|None], "x": [float]|None, "cats": [str], "fmt", "src": series dict, "i"}],
     "cats": [str], "sampled": step, "xfmt", "error": str|None}"""
    out = {"series": [], "cats": [], "sampled": 1, "xfmt": "General", "error": None}
    fam = family(ch["type"])
    scatter = fam == "scatter" or any((s.get("type") or "").startswith("scatter") for s in ch["series"])
    home = home or (wb.sheets[0] if wb.sheets else None)
    sampled = 1
    for i, s in enumerate(ch["series"]):
        got = ref_rect(wb, s.get("values", ""), home.name if home else None)
        if got is None:
            continue
        sh, rect = got
        cells, step = _cells_of(sh, rect, limit)
        sampled = max(sampled, step)
        ys, fmt, probed = [], "General", False
        for (r, c), v in zip(cells, _read_values(sh, cells)):
            if is_num(v) and not math.isfinite(v):
                ys.append(None)            # inf / nan can't be drawn
            elif is_num(v):
                ys.append(float(v))
                if not probed:       # the axis shows numbers the way the first number cell does
                    probed = True
                    fmt = sh.style(r, c).numfmt or "General"
            elif isinstance(v, bool):
                ys.append(1.0 if v else 0.0)
            elif isinstance(v, str) and v != "":
                ys.append(0.0)            # Excel plots text as zero
            else:
                ys.append(None)            # blank / error: a gap
        xs, cats = None, []
        cg = ref_rect(wb, s.get("cats", ""), home.name if home else None)
        if cg is not None:
            csh, crect = cg
            ccells, _ = _cells_of(csh, crect, limit)
            if len(ccells) != len(cells) and len(ccells) > len(cells):
                ccells = ccells[:len(cells)]
            vals = [(r, c, v) for (r, c), v in zip(ccells, _read_values(csh, ccells))]
            if scatter:
                xs = [float(v) if is_num(v) and math.isfinite(v) else None for _, _, v in vals]
                if any(x is None for x in xs[:50]) and not any(x is not None for x in xs):
                    xs = None
                if vals and is_num(vals[0][2]):
                    out["xfmt"] = csh.style(vals[0][0], vals[0][1]).numfmt or "General"
            cats = [_text_of(csh, r, c, v) for r, c, v in vals]
        n = len(ys)
        if len(cats) < n:
            base = (sampled if sampled > 1 else 1)
            cats = cats + [str(j * base + 1) for j in range(len(cats), n)]
        if scatter and (xs is None or len(xs) < n):
            xs = (xs or []) + [float(j + 1) for j in range(len(xs or []), n)]
        nm = s.get("name") or ""
        if nm.startswith("="):
            ng = ref_rect(wb, nm, home.name if home else None)
            if ng is not None:
                nsh, nr = ng
                nv = nsh.value(nr[0], nr[1])
                nm = _text_of(nsh, nr[0], nr[1], nv) if nv is not None else ""
            else:
                nm = "#REF!"
        out["series"].append({"name": nm or f"Series{i + 1}", "y": ys, "x": xs, "cats": cats[:n], "fmt": fmt,
                              "src": s, "i": i})
    out["sampled"] = sampled
    for sd in out["series"]:
        if sd["cats"]:
            out["cats"] = sd["cats"]
            break
    if not out["cats"] and out["series"]:
        out["cats"] = [str(j + 1) for j in range(max(len(s["y"]) for s in out["series"]))]
    if fam in ("pie", "doughnut") and out["series"] and max(len(s["y"]) for s in out["series"]) > MAX_PIE_POINTS:
        n = max(len(s["y"]) for s in out["series"])
        out["error"] = (f"A {type_label(ch['type']).lower()} can't show {n:,} slices. "
                        "Pick a smaller range, or use a column or line chart.")
        out["series"] = []
    return out


# ---------------------------------------------------------------- structure changes / renames
def remap(ch, sheet_name, remap_rect, on_this_sheet, axis, at, count, home_name):
    """Chart `ch` (living on sheet `home_name`) after rows / columns were inserted or deleted on
    `sheet_name`: refs into that sheet move, and (on that sheet) so does the chart."""
    out = dict(ch)

    def fix(text):
        p = parse_ref(text, home_name)
        if not text or not text.startswith("=") or p is None or p[0].upper() != sheet_name.upper():
            return text
        rect = remap_rect(p[1])
        return make_ref(p[0], rect) if rect else "=#REF!"
    out["series"] = []
    for s in ch["series"]:
        s = dict(s, name=fix(s["name"]), values=fix(s["values"]), cats=fix(s["cats"]))
        if s["values"] == "=#REF!":
            continue
        if s["cats"] == "=#REF!":
            s["cats"] = ""
        out["series"].append(s)
    if ch.get("range"):
        out["range"] = fix(ch["range"])
    if on_this_sheet:
        for k in ("from", "to"):
            r, c, dx, dy = ch[k]
            rect = remap_rect((r, c, r, c))
            if rect is None:
                # the anchor cell itself was deleted: it lands on the row / column that followed
                if axis == "row":
                    r, dy = at, 0
                else:
                    c, dx = at, 0
            else:
                r, c = rect[0], rect[1]
            out[k] = [r, c, dx, dy]
    return out


def rename_sheet(ch, old, new):
    out = dict(ch)
    out["series"] = [dict(s, name=rename_in_ref(s["name"], old, new), values=rename_in_ref(s["values"], old, new),
                          cats=rename_in_ref(s["cats"], old, new)) for s in ch["series"]]
    if ch.get("range"):
        out["range"] = rename_in_ref(ch["range"], old, new)
    return out


# ---------------------------------------------------------------- saving
_SAVED_KEYS = tuple(DEFAULTS) + ("id", "from", "to", "series")


def to_json(ch):
    d = {k: ch[k] for k in _SAVED_KEYS if k in ch}
    d["series"] = [dict(s) for s in ch["series"]]
    return d


def from_json(d):
    return normalize(d)


def describe(sh, ch):
    """One line for Claude: what the chart is, where it sits and what it shows."""
    r1, c1, _, _ = ch["from"]
    r2, c2, dx2, dy2 = ch["to"]
    if dx2 == 0 and c2 > c1:
        c2 -= 1            # the chart ends exactly at a column / row border
    if dy2 == 0 and r2 > r1:
        r2 -= 1
    t = ch["title"] if ch["title"] not in (None, "") else "(auto title)"
    names = []
    for s in ch["series"]:
        n = s["name"]
        if n.startswith("="):
            got = ref_rect(sh.wb, n, sh.name)
            v = got[0].value(got[1][0], got[1][1]) if got else None
            n = "" if v is None else str(v)
        names.append(f"{n or 'Series'} = {s['values'][1:]}")
    kind = type_label(ch["type"]) + (" (secondary axis)" if any(s.get("secondary") for s in ch["series"]) else "")
    return (f"{ch.get('name') or 'Chart'}: {kind} chart '{t}' at {range_addr(r1, c1, r2, c2)} "
            f"showing {'; '.join(names) or 'nothing'}")


def copy_chart(ch):
    return copy.deepcopy(ch)
