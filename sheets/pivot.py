"""PivotTables: summarize a source range by Row / Column fields with Sum, Count, Average, Min,
Max or Distinct Count of Value fields, optionally filtered by Filter fields.

A pivot is a plain dict, kept in Sheet.pivots of the sheet it's drawn on, and never changed in
place (undo snapshots keep references): updates build a new dict.
    {"name", "source": sheet name, "src": [r1, c1, r2, c2] (row r1 = headers),
     "anchor": [r, c], "rows": [field], "cols": [field], "values": [[field, agg]],
     "filters": {field: [allowed item labels]}  (a filter field with no entry shows all),
     "filter_fields": [field], "out": [r1, c1, r2, c2] or None (where it was last drawn)}
Its output is ordinary cells (Excel sees the numbers); Ekxel stores the definition in its hidden
sheet so it can be refreshed and changed later."""
from .errors import XLError
from .values import is_num
from .workbook import DEFAULT_STYLE

AGGS = [("sum", "Sum"), ("count", "Count"), ("average", "Average"), ("max", "Max"), ("min", "Min"),
        ("distinct", "Distinct Count")]
AGG_LABEL = dict(AGGS)
HEAD_FILL = "#DDEBF7"   # Excel's default "Light 9"-ish pivot style
LINE = ("thin", "#9BC2E6")
BLANK = "(blank)"


def new(name, source, src, anchor):
    return {"name": name, "source": source, "src": list(src), "anchor": list(anchor), "rows": [], "cols": [],
            "values": [], "filters": {}, "filter_fields": [], "out": None}


def changed(pv, **kw):
    out = dict(pv)
    out.update(kw)
    return out


def value_label(field, agg):
    return f"{AGG_LABEL.get(agg, agg)} of {field}"


# ---------------------------------------------------------------- reading the source
def fields(wb, pv):
    """Header names of the source range (blank or repeated headers get Excel-style names)."""
    sh = wb.get_sheet(pv["source"])
    if sh is None:
        return []
    r1, c1, _, c2 = pv["src"]
    c2 = min(c2, max(sh.max_col, c1))
    out, seen = [], {}
    for c in range(c1, c2 + 1):
        v = sh.value(r1, c)
        name = str(v).strip() if v not in (None, "") and not isinstance(v, XLError) else f"Column{c - c1 + 1}"
        n = seen.get(name.lower(), 0)
        seen[name.lower()] = n + 1
        out.append(name if n == 0 else f"{name}{n + 1}")
    return out


def records(wb, pv):
    """(field names, list of row tuples, numfmt per field) from the source range; blank rows skipped."""
    sh = wb.get_sheet(pv["source"])
    if sh is None:
        return [], [], []
    names = fields(wb, pv)
    r1, c1, r2, _ = pv["src"]
    r2 = min(r2, sh.max_row)
    cols = range(c1, c1 + len(names))
    rows = []
    for r in range(r1 + 1, r2 + 1):
        row = tuple(sh.value(r, c) for c in cols)
        if any(v not in (None, "") for v in row):
            rows.append(row)
    fmts = [sh.style(r1 + 1, c).numfmt if r2 > r1 else "General" for c in cols]
    return names, rows, fmts


_MONTHS = ["january", "february", "march", "april", "may", "june", "july", "august", "september",
           "october", "november", "december"]
_DAYS = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]
_CALENDAR = {**{m: i for i, m in enumerate(_MONTHS)}, **{m[:3]: i for i, m in enumerate(_MONTHS)},
             "sept": 8, **{d: 20 + i for i, d in enumerate(_DAYS)}, **{d[:3]: 20 + i for i, d in enumerate(_DAYS)}}


def _item_key(v):
    """Grouping/sort key: numbers by value, text case-insensitively with month and weekday names in
    calendar order (Excel's built-in custom lists), blanks last."""
    if v is None or v == "":
        return (2, 0, "")
    if isinstance(v, XLError):
        return (3, 0, str(v))
    if is_num(v) and not isinstance(v, bool):
        return (0, float(v), "")
    t = str(v).strip().lower()
    return (1, _CALENDAR.get(t, 99), t)


def item_label(v):
    if v is None or v == "":
        return BLANK
    return str(v)


def items(wb, pv, field):
    """Distinct item labels of a field, sorted, for the filter list."""
    names, rows, _ = records(wb, pv)
    if field not in names:
        return []
    i = names.index(field)
    seen = {}
    for row in rows:
        seen.setdefault(_item_key(row[i]), item_label(row[i]))
    return [seen[k] for k in sorted(seen)]


def default_agg(wb, pv, field):
    """Sum when every value of the field is a number, otherwise Count (what Excel picks)."""
    names, rows, _ = records(wb, pv)
    if field not in names:
        return "count"
    i = names.index(field)
    vals = [row[i] for row in rows if row[i] not in (None, "")]
    return "sum" if vals and all(is_num(v) and not isinstance(v, bool) for v in vals) else "count"


# ---------------------------------------------------------------- computing
def _aggregate(vals, agg):
    if agg == "count":
        return float(sum(1 for v in vals if v not in (None, "")))
    if agg == "distinct":
        return float(len({_item_key(v) for v in vals if v not in (None, "")}))
    nums = [float(v) for v in vals if is_num(v) and not isinstance(v, bool)]
    if agg == "sum":
        return sum(nums)
    if not nums:
        return None
    if agg == "average":
        return sum(nums) / len(nums)
    return max(nums) if agg == "max" else min(nums)


def build(wb, pv):
    """-> (cells {(dr, dc): (value, style)}, (height, width), error or None), relative to the anchor."""
    names, rows, fmts = records(wb, pv)
    if not names:
        return {}, (0, 0), f"The source sheet '{pv['source']}' isn't there any more."
    idx = {n: i for i, n in enumerate(names)}
    rf = [f for f in pv["rows"] if f in idx]
    cf = [f for f in pv["cols"] if f in idx]
    vf = [(f, a) for f, a in pv["values"] if f in idx]
    ff = [f for f in pv.get("filter_fields", []) if f in idx]
    if not (rf or cf or vf or ff):  # a new, empty pivot: Excel's placeholder
        tip = DEFAULT_STYLE.with_(italic=True, color="#595959", fill=HEAD_FILL)
        return {(0, 0): (f"{pv['name']}: tick fields in the PivotTable Fields panel to build the report", tip)}, \
            (1, 1), None

    # filters (report filters, and item filters on Row/Column fields)
    for f, allowed in pv["filters"].items():
        if f in idx and allowed is not None and (f in ff or f in rf or f in cf):
            ok = set(allowed)
            rows = [row for row in rows if item_label(row[idx[f]]) in ok]

    cells = {}
    head = DEFAULT_STYLE.with_(bold=True, fill=HEAD_FILL, border=(None, None, None, LINE))
    sub = DEFAULT_STYLE.with_(bold=True)
    grand = DEFAULT_STYLE.with_(bold=True, fill=HEAD_FILL, border=(None, LINE, None, None))
    lab = DEFAULT_STYLE

    # report filters, above the table
    y = 0
    for f in ff:
        allowed = pv["filters"].get(f)
        shown = "(All)" if allowed is None else (allowed[0] if len(allowed) == 1 else "(Multiple Items)")
        cells[(y, 0)] = (f, sub)
        cells[(y, 1)] = (shown, lab)
        y += 1
    if ff:
        y += 1

    # group keys
    def key_of(row, flds):
        return tuple(_item_key(row[idx[f]]) for f in flds)
    labels = {}
    groups = {}
    for row in rows:
        rk, ck = key_of(row, rf), key_of(row, cf)
        for f, k in zip(rf + cf, rk + ck):
            labels.setdefault((f, k), item_label(row[idx[f]]))
        groups.setdefault((rk, ck), []).append(row)
    row_keys = sorted({rk for rk, _ in groups})
    col_keys = sorted({ck for _, ck in groups}) if cf else [()]
    nvf = max(1, len(vf))

    def agg_cells(rpred, cpred):
        """Aggregates over records whose row key matches rpred and column key cpred."""
        out = []
        for f, a in vf:
            vals = []
            for (rk, ck), rs in groups.items():
                if rpred(rk) and cpred(ck):
                    vals.extend(r[idx[f]] for r in rs)
            out.append((_aggregate(vals, a) if vals or a in ("count", "distinct") else None,
                        fmts[idx[f]] if a in ("sum", "average", "max", "min") else "General"))
        return out

    left = max(1, len(rf))
    data_cols = []  # (column key or None for grand total, value index)
    for ck in col_keys:
        for vi in range(len(vf)):
            data_cols.append((ck, vi))
    if cf and vf:
        for vi in range(len(vf)):
            data_cols.append((None, vi))

    # header rows
    hr = len(cf) + (1 if (len(vf) > 1 or not cf) else 0)
    hr = max(hr, 1)
    top = y
    for r in range(hr):
        for c in range(left + len(data_cols)):
            cells[(top + r, c)] = ("", head)
    if cf and len(vf) == 1:
        cells[(top, 0)] = (value_label(*vf[0]), head)
    for i, f in enumerate(rf):
        cells[(top + hr - 1, i)] = (f, head)
    for c, (ck, vi) in enumerate(data_cols):
        x = left + c
        if ck is None:
            text = "Grand Total" if len(vf) == 1 else f"Total {value_label(*vf[vi])}"
            cells[(top + (0 if len(vf) == 1 else hr - 1), x)] = (text, head)
            continue
        for lvl, f in enumerate(cf):
            prev = data_cols[c - 1][0] if c else None
            if prev is None or prev[:lvl + 1] != ck[:lvl + 1] or (len(vf) > 1 and vi == 0 and lvl == len(cf) - 1):
                cells[(top + lvl, x)] = (labels[(f, ck[lvl])], head)
        if len(vf) > 1 or not cf:
            cells[(top + hr - 1, x)] = (value_label(*vf[vi]), head)
    y = top + hr

    def put_values(yy, rpred, style):
        for c, (ck, vi) in enumerate(data_cols):
            cpred = (lambda k: True) if ck is None else (lambda k, ck=ck: k == ck)
            val, fmt = agg_cells(rpred, cpred)[vi]
            cells[(yy, left + c)] = (val, style.with_(numfmt=fmt) if val is not None else style)

    # body: row groups with subtotals for the outer fields (tabular layout)
    prev = None
    for n, rk in enumerate(row_keys):
        for lvl in range(len(rf)):
            if prev is None or prev[:lvl + 1] != rk[:lvl + 1]:
                cells[(y, lvl)] = (labels[(rf[lvl], rk[lvl])], lab)
        for lvl in range(len(rf)):
            cells.setdefault((y, lvl), ("", lab))
        if vf:
            put_values(y, lambda k, rk=rk: k == rk, lab)
        y += 1
        prev = rk
        nxt = row_keys[n + 1] if n + 1 < len(row_keys) else None
        for lvl in range(len(rf) - 2, -1, -1):  # close outer groups that end here
            if nxt is None or nxt[:lvl + 1] != rk[:lvl + 1]:
                if vf:
                    cells[(y, lvl)] = (f"{labels[(rf[lvl], rk[lvl])]} Total", sub)
                    for c in range(len(rf)):
                        if c != lvl:
                            cells[(y, c)] = ("", sub)
                    put_values(y, lambda k, p=rk[:lvl + 1], L=lvl + 1: k[:L] == p, sub)
                    y += 1
    if not rf and vf:  # values only: one row of totals
        put_values(y, lambda k: True, lab)
        cells[(y, 0)] = cells.get((y, 0), ("", lab))
        y += 1
    if rf and vf:
        for c in range(left):
            cells[(y, c)] = ("", grand)
        cells[(y, 0)] = ("Grand Total", grand)
        put_values(y, lambda k: True, grand)
        y += 1
    width = left + len(data_cols)
    return cells, (y, max(width, 2 if ff else 1)), None


def out_rect(pv, size):
    h, w = size
    r, c = pv["anchor"]
    return [r, c, r + max(h, 1) - 1, c + max(w, 1) - 1]


def contains(rect, r, c):
    return rect is not None and rect[0] <= r <= rect[2] and rect[1] <= c <= rect[3]


# ---------------------------------------------------------------- structure changes
def remapped(pv, sheet_name, remap_rect, on_this_sheet):
    """`pv` after rows/columns were inserted or deleted on sheet `sheet_name`."""
    out = dict(pv)
    if pv["source"] == sheet_name:
        src = remap_rect(tuple(pv["src"]))
        if src is not None:
            out["src"] = list(src)
    if on_this_sheet:
        a = remap_rect((pv["anchor"][0], pv["anchor"][1], pv["anchor"][0], pv["anchor"][1]))
        if a is None:
            return None  # its top-left cell was deleted
        out["anchor"] = [a[0], a[1]]
        if pv.get("out"):
            o = remap_rect(tuple(pv["out"]))
            out["out"] = list(o) if o else None
    return out


def to_json(pv):
    return {k: pv[k] for k in ("name", "source", "src", "anchor", "rows", "cols", "values", "filters",
                               "filter_fields", "out")}


def from_json(d):
    pv = new(d.get("name", "PivotTable1"), d["source"], d["src"], d["anchor"])
    for k in ("rows", "cols", "values", "filter_fields"):
        pv[k] = list(d.get(k, []))
    pv["values"] = [list(x) for x in pv["values"]]
    pv["filters"] = {k: list(v) for k, v in (d.get("filters") or {}).items()}
    pv["out"] = list(d["out"]) if d.get("out") else None
    return pv
