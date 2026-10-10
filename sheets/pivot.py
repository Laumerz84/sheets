"""PivotTables: summarize a source range by Row / Column fields with Value fields (Sum, Count,
Average, ... with "Show Values As" calculations), filtered by report filters, item lists, label
filters and value filters; date fields grouped by years/quarters/months/days, numbers by ranges.

A pivot is a plain dict, kept in Sheet.pivots of the sheet it's drawn on, and never changed in
place (undo snapshots keep references): updates build a new dict.
    {"name", "source": sheet name, "src": [r1, c1, r2, c2] (row r1 = headers),
     "anchor": [r, c], "rows": [field], "cols": [field], "filter_fields": [field],
     "values": [[field, agg]] or [[field, agg, {"name", "numfmt", "show", "base_field", "base_item"}]],
     "filters": {field: [allowed item labels]}  (a field with no entry shows all),
     "label_filters": {field: [op, a, b]}, "value_filters": {field: [op, value label, a, b]},
     "sort": {field: ["asc"|"desc", value label or None]},
     "groups": {field: {"date": [levels]} or {"num": [start, end, by]}},
     "calc": {name: formula}, "collapsed": {field: [labels, "*" = all, "!label" = except]},
     "layout": {see LAYOUT}, "out": [r1, c1, r2, c2] or None (where it was last drawn),
     "hot": [[kind, dr, dc, ...]] (in-sheet buttons of the last drawing)}
Rows/Columns may hold VALUES ("Σ Values"): where the value fields go when there are 2 or more.
Every key except name/source/src/anchor is optional (Phase 1 files lack most of them).
Its output is ordinary cells (Excel sees the numbers); Ekxel stores the definition in its hidden
sheet so it can be refreshed and changed later."""
import fnmatch
import math

from . import errors
from .errors import XLError
from .values import is_num
from .workbook import DEFAULT_STYLE

VALUES = "Σ Values"
AGGS = [("sum", "Sum"), ("count", "Count"), ("average", "Average"), ("max", "Max"), ("min", "Min"),
        ("product", "Product"), ("countnums", "Count Numbers"), ("stdev", "StdDev"), ("stdevp", "StdDevp"),
        ("var", "Var"), ("varp", "Varp"), ("distinct", "Distinct Count")]
AGG_LABEL = dict(AGGS)
SHOW_AS = [("none", "No Calculation"), ("pct_grand", "% of Grand Total"), ("pct_col", "% of Column Total"),
           ("pct_row", "% of Row Total"), ("pct_of", "% Of"), ("pct_parent_row", "% of Parent Row Total"),
           ("pct_parent_col", "% of Parent Column Total"), ("pct_parent", "% of Parent Total"),
           ("diff", "Difference From"), ("pct_diff", "% Difference From"), ("running", "Running Total In"),
           ("pct_running", "% Running Total In"), ("rank_asc", "Rank Smallest to Largest"),
           ("rank_desc", "Rank Largest to Smallest"), ("index", "Index")]
SHOW_LABEL = dict(SHOW_AS)
SHOW_BASE = {"pct_of": "item", "diff": "item", "pct_diff": "item", "running": "field", "pct_running": "field",
             "rank_asc": "field", "rank_desc": "field", "pct_parent": "field"}
_PCT_SHOWS = {"pct_grand", "pct_col", "pct_row", "pct_of", "pct_parent_row", "pct_parent_col", "pct_parent",
              "pct_diff", "pct_running"}
DATE_LEVELS = [("years", "Years"), ("quarters", "Quarters"), ("months", "Months"), ("days", "Days"),
               ("hours", "Hours"), ("minutes", "Minutes"), ("seconds", "Seconds")]
_LEVEL_ORDER = [k for k, _ in DATE_LEVELS]
LABEL_OPS = [("equals", "Equals..."), ("not_equals", "Does Not Equal..."), ("begins", "Begins With..."),
             ("not_begins", "Does Not Begin With..."), ("ends", "Ends With..."), ("not_ends", "Does Not End With..."),
             ("contains", "Contains..."), ("not_contains", "Does Not Contain..."), ("gt", "Is Greater Than..."),
             ("ge", "Is Greater Than Or Equal To..."), ("lt", "Is Less Than..."),
             ("le", "Is Less Than Or Equal To..."), ("between", "Between..."), ("not_between", "Not Between...")]
VALUE_OPS = [("equals", "Equals..."), ("not_equals", "Does Not Equal..."), ("gt", "Is Greater Than..."),
             ("ge", "Is Greater Than Or Equal To..."), ("lt", "Is Less Than..."),
             ("le", "Is Less Than Or Equal To..."), ("between", "Between..."), ("not_between", "Not Between..."),
             ("top", "Top 10...")]
LAYOUT = {"form": "compact",        # compact / outline / tabular
          "repeat": False,          # repeat all item labels (outline / tabular)
          "subtotals": "top",       # none / top / bottom (tabular always shows them at the bottom)
          "grand": "both",          # both / rows (the Grand Total column) / cols (the Grand Total row) / none
          "blank_line": False,      # a blank line after each item of the outer row fields
          "empty": "",              # text shown in empty value cells
          "show_empty_items": False,
          "buttons": True,          # +/- buttons and filter dropdowns in the sheet
          "style": "Light 9", "banded_rows": False, "banded_cols": False}
PHASE1_LAYOUT = {"form": "tabular", "subtotals": "bottom"}   # how pivots saved before these options looked

# PivotTable styles: header fill/text, the line under headers and above the grand total, subtotal and
# grand total fills, banding fill
STYLES = {
    "None":      {"head": None, "head_color": None, "line": "#000000", "sub": None, "grand": None, "band": "#F2F2F2"},
    "Light 1":   {"head": None, "head_color": None, "line": "#8EA9DB", "sub": None, "grand": None, "band": "#D9E1F2"},
    "Light 9":   {"head": "#DDEBF7", "head_color": None, "line": "#9BC2E6", "sub": None, "grand": "#DDEBF7",
                  "band": "#F2F7FC"},
    "Light 10":  {"head": "#FCE4D6", "head_color": None, "line": "#F4B084", "sub": None, "grand": "#FCE4D6",
                  "band": "#FDF2EB"},
    "Light 11":  {"head": "#EDEDED", "head_color": None, "line": "#C9C9C9", "sub": None, "grand": "#EDEDED",
                  "band": "#F6F6F6"},
    "Light 13":  {"head": "#FFF2CC", "head_color": None, "line": "#FFD966", "sub": None, "grand": "#FFF2CC",
                  "band": "#FFF9E6"},
    "Light 14":  {"head": "#E2EFDA", "head_color": None, "line": "#A9D08E", "sub": None, "grand": "#E2EFDA",
                  "band": "#F0F7EC"},
    "Medium 2":  {"head": "#4472C4", "head_color": "#FFFFFF", "line": "#4472C4", "sub": "#D9E1F2",
                  "grand": "#B4C6E7", "band": "#D9E1F2"},
    "Medium 3":  {"head": "#ED7D31", "head_color": "#FFFFFF", "line": "#ED7D31", "sub": "#FCE4D6",
                  "grand": "#F8CBAD", "band": "#FCE4D6"},
    "Medium 4":  {"head": "#7F7F7F", "head_color": "#FFFFFF", "line": "#7F7F7F", "sub": "#EDEDED",
                  "grand": "#DBDBDB", "band": "#EDEDED"},
    "Medium 7":  {"head": "#70AD47", "head_color": "#FFFFFF", "line": "#70AD47", "sub": "#E2EFDA",
                  "grand": "#C6E0B4", "band": "#E2EFDA"},
    "Dark 2":    {"head": "#203764", "head_color": "#FFFFFF", "line": "#203764", "sub": "#B4C6E7",
                  "grand": "#8EA9DB", "band": "#D9E1F2"},
}
HEAD_FILL = "#DDEBF7"
LINE = ("thin", "#9BC2E6")
BLANK = "(blank)"


def new(name, source, src, anchor):
    return {"name": name, "source": source, "src": list(src), "anchor": list(anchor), "rows": [], "cols": [],
            "values": [], "filters": {}, "filter_fields": [], "out": None, "layout": dict(LAYOUT)}


def changed(pv, **kw):
    out = dict(pv)
    out.update(kw)
    return out


def layout(pv):
    lay = dict(LAYOUT)
    lay.update(pv.get("layout") or PHASE1_LAYOUT)
    return lay


def with_layout(pv, **kw):
    lay = dict(pv.get("layout") or PHASE1_LAYOUT)
    lay.update(kw)
    return changed(pv, layout=lay)


def with_key(pv, key, field, value):
    """pv with pv[key][field] = value (None removes it): for filters, sort, groups, ..."""
    d = dict(pv.get(key) or {})
    if value is None:
        d.pop(field, None)
    else:
        d[field] = value
    return changed(pv, **{key: d})


# ---------------------------------------------------------------- value fields
def vopts(entry):
    return entry[2] if len(entry) > 2 and isinstance(entry[2], dict) else {}


def value_label(field, agg, opts=None):
    if opts and opts.get("name"):
        return opts["name"]
    return f"{AGG_LABEL.get(agg, agg)} of {field}"


def value_labels(pv):
    """Display names of the value fields, made unique like Excel ("Sum of Units2")."""
    out, seen = [], {}
    for e in pv["values"]:
        t = value_label(e[0], e[1], vopts(e))
        n = seen.get(t.lower(), 0)
        seen[t.lower()] = n + 1
        out.append(t if n == 0 else f"{t}{n + 1}")
    return out


def value_entry(field, agg, opts=None):
    opts = {k: v for k, v in (opts or {}).items() if v not in (None, "", "none")}
    return [field, agg, opts] if opts else [field, agg]


# ---------------------------------------------------------------- reading the source
def source_fields(wb, pv):
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


def group_field_name(level, base):
    return f"{dict(DATE_LEVELS)[level]} ({base})"


def virtual_fields(pv):
    """{name: (base field, date level)} for the extra fields date grouping makes (Years (Date), ...)."""
    out = {}
    for base, g in (pv.get("groups") or {}).items():
        lv = g.get("date") or []
        for level in lv[:-1]:
            out[group_field_name(level, base)] = (base, level)
    return out


def fields(wb, pv):
    """Every field the field list shows: source fields (each followed by its date-group fields),
    then calculated fields."""
    out = []
    virt = virtual_fields(pv)
    for n in source_fields(wb, pv):
        out.append(n)
        out.extend(v for v, (b, _) in virt.items() if b == n)
    out.extend(n for n in (pv.get("calc") or {}) if n not in out)
    return out


def records(wb, pv):
    """(field names, list of row tuples, numfmt per field) from the source range; blank rows skipped."""
    names, rows, fmts, _ = _records(wb, pv)
    return names, rows, fmts


def _records(wb, pv):
    sh = wb.get_sheet(pv["source"])
    if sh is None:
        return [], [], [], []
    names = source_fields(wb, pv)
    r1, c1, r2, _ = pv["src"]
    r2 = min(r2, sh.max_row)
    cols = range(c1, c1 + len(names))
    rows, srcrows = [], []
    for r in range(r1 + 1, r2 + 1):
        row = tuple(sh.value(r, c) for c in cols)
        if any(v not in (None, "") for v in row):
            rows.append(row)
            srcrows.append(r)
    fmts = [sh.style(r1 + 1, c).numfmt if r2 > r1 else "General" for c in cols]
    return names, rows, fmts, srcrows


_MONTHS = ["january", "february", "march", "april", "may", "june", "july", "august", "september",
           "october", "november", "december"]
_MON = [m[:3].title() for m in _MONTHS]
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
    if isinstance(v, float) and v.is_integer() and abs(v) < 1e15:
        return str(int(v))
    return str(v)


def _num(x):
    return str(int(x)) if float(x).is_integer() else f"{x:g}"


def _date_part(v, level):
    from .numfmt import serial_to_datetime
    try:
        d = serial_to_datetime(round(float(v) * 86400) / 86400)
    except (ValueError, OverflowError):
        return _item_key(v), item_label(v)
    if level == "years":
        return (0, float(d.year), ""), str(d.year)
    if level == "quarters":
        q = (d.month - 1) // 3 + 1
        return (0, float(q), ""), f"Qtr{q}"
    if level == "months":
        return (0, float(d.month), ""), _MON[d.month - 1]
    if level == "days":
        return (0, float(d.month * 100 + d.day), ""), f"{d.day}-{_MON[d.month - 1]}"
    if level == "hours":
        h = d.hour
        return (0, float(h), ""), f"{(h % 12) or 12} {'AM' if h < 12 else 'PM'}"
    if level == "minutes":
        return (0, float(d.minute), ""), f":{d.minute:02d}"
    return (0, float(d.second), ""), f":{d.second:02d}"


def _num_bucket(v, spec):
    start, end, by = (float(x) for x in spec)
    if v < start:
        return (0, -math.inf, ""), f"<{_num(start)}"
    if v > end:
        return (0, math.inf, ""), f">{_num(end)}"
    b = start + math.floor((v - start) / by + 1e-9) * by
    if float(start).is_integer() and float(by).is_integer():
        return (0, b, ""), f"{_num(b)}-{_num(b + by - 1)}"
    return (0, b, ""), f"{_num(b)}-{_num(b + by)}"


# ---------------------------------------------------------------- calculated fields
class CalcError(ValueError):
    pass


_CALC_FUNCS = {"IF", "ROUND", "ABS", "MIN", "MAX", "SQRT", "AND", "OR", "NOT", "INT"}


def compile_calc(text, names):
    """A calculated field formula like =Revenue/Units or ='Unit Price'*Units -> (fn(sums) -> value,
    referenced field names). Raises CalcError for a bad formula."""
    s = text.strip()
    if s.startswith("="):
        s = s[1:]
    by_len = sorted(names, key=len, reverse=True)
    toks = []
    i = 0
    while i < len(s):
        ch = s[i]
        if ch.isspace():
            i += 1
            continue
        if ch == "'":
            j = s.find("'", i + 1)
            if j < 0:
                raise CalcError("A field name in quotes isn't closed.")
            nm = s[i + 1:j]
            match = next((n for n in names if n.lower() == nm.lower()), None)
            if match is None:
                raise CalcError(f"There's no field named '{nm}'.")
            toks.append(("field", match))
            i = j + 1
            continue
        if ch.isdigit() or (ch == "." and i + 1 < len(s) and s[i + 1].isdigit()):
            j = i
            while j < len(s) and (s[j].isdigit() or s[j] in ".eE" or (s[j] in "+-" and s[j - 1] in "eE")):
                j += 1
            try:
                toks.append(("num", float(s[i:j])))
            except ValueError:
                raise CalcError(f"'{s[i:j]}' isn't a number.")
            i = j
            continue
        two = s[i:i + 2]
        if two in ("<=", ">=", "<>"):
            toks.append(("op", two))
            i += 2
            continue
        if ch in "+-*/^(),=<>&%":
            toks.append(("op", ch))
            i += 1
            continue
        if ch == '"':
            j = s.find('"', i + 1)
            if j < 0:
                raise CalcError("A text in quotes isn't closed.")
            toks.append(("str", s[i + 1:j]))
            i = j + 1
            continue
        low = s[i:].lower()
        match = next((n for n in by_len if low.startswith(n.lower()) and
                      (i + len(n) >= len(s) or not (s[i + len(n)].isalnum() or s[i + len(n)] == "_"))), None)
        j = i
        while j < len(s) and (s[j].isalnum() or s[j] in "_."):
            j += 1
        word = s[i:j]
        if word.upper() in _CALC_FUNCS and s[j:].lstrip().startswith("("):
            toks.append(("func", word.upper()))
            i = j
            continue
        if match is not None:
            toks.append(("field", match))
            i += len(match)
            continue
        if word.upper() in ("TRUE", "FALSE"):
            toks.append(("num", 1.0 if word.upper() == "TRUE" else 0.0))
            i = j
            continue
        raise CalcError(f"'{word or ch}' isn't a field name.")
    pos = [0]
    refs = set()

    def peek():
        return toks[pos[0]] if pos[0] < len(toks) else (None, None)

    def take():
        t = peek()
        pos[0] += 1
        return t

    prec = {"=": 1, "<>": 1, "<": 1, ">": 1, "<=": 1, ">=": 1, "&": 2, "+": 3, "-": 3, "*": 4, "/": 4, "^": 5}

    def expr(minp=1):
        left = unary()
        while True:
            k, op = peek()
            if k != "op" or op not in prec or prec[op] < minp:
                return left
            take()
            right = expr(prec[op] + (0 if op == "^" else 1))
            left = _binop(op, left, right)

    def unary():
        k, op = peek()
        if k == "op" and op in "+-":
            take()
            inner = unary()
            return inner if op == "+" else (lambda env, f=inner: _neg(f(env)))
        node = primary()
        if peek() == ("op", "%"):
            take()
            return lambda env, f=node: _arith("/", f(env), 100.0)
        return node

    def primary():
        k, v = take()
        if k == "num":
            return lambda env, v=v: v
        if k == "str":
            return lambda env, v=v: v
        if k == "field":
            refs.add(v)
            return lambda env, v=v: env.get(v, 0.0)
        if k == "op" and v == "(":
            e = expr()
            if take() != ("op", ")"):
                raise CalcError("A bracket isn't closed.")
            return e
        if k == "func":
            if take() != ("op", "("):
                raise CalcError(f"{v} needs brackets.")
            args = []
            if peek() != ("op", ")"):
                while True:
                    args.append(expr())
                    if peek() == ("op", ","):
                        take()
                        continue
                    break
            if take() != ("op", ")"):
                raise CalcError("A bracket isn't closed.")
            return _func(v, args)
        raise CalcError("The formula isn't complete." if k is None else f"Unexpected '{v}'.")

    if not toks:
        raise CalcError("Type a formula, like =Revenue/Units.")
    fn = expr()
    if pos[0] != len(toks):
        raise CalcError(f"Unexpected '{toks[pos[0]][1]}'.")
    return fn, refs


def _neg(a):
    return a if isinstance(a, XLError) else -_to_n(a)


def _to_n(a):
    if isinstance(a, XLError):
        raise _Err(a)
    if isinstance(a, str):
        try:
            return float(a)
        except ValueError:
            raise _Err(errors.VALUE)
    return float(a)


class _Err(Exception):
    def __init__(self, e):
        super().__init__(str(e))
        self.e = e


def _arith(op, a, b):
    a, b = _to_n(a), _to_n(b)
    if op == "+":
        return a + b
    if op == "-":
        return a - b
    if op == "*":
        return a * b
    if op == "/":
        if b == 0:
            raise _Err(errors.DIV0)
        return a / b
    try:
        return float(a ** b)
    except (OverflowError, ZeroDivisionError, TypeError):
        raise _Err(errors.NUM)


def _binop(op, l, r):
    if op in "+-*/^":
        return lambda env: _arith(op, l(env), r(env))
    if op == "&":
        return lambda env: f"{item_label(l(env))}{item_label(r(env))}"

    def cmp(env):
        a, b = l(env), r(env)
        if isinstance(a, str) or isinstance(b, str):
            a, b = str(a).lower(), str(b).lower()
        return 1.0 if {"=": a == b, "<>": a != b, "<": a < b, ">": a > b, "<=": a <= b, ">=": a >= b}[op] else 0.0
    return cmp


def _func(name, args):
    def need(n):
        if len(args) < n:
            raise CalcError(f"{name} needs {n} argument{'s' if n > 1 else ''}.")
    if name == "IF":
        need(2)
        a, b, c = args[0], args[1], (args[2] if len(args) > 2 else (lambda env: 0.0))
        return lambda env: b(env) if _to_n(a(env)) else c(env)
    if name == "ROUND":
        need(2)
        return lambda env: float(round(_to_n(args[0](env)), int(_to_n(args[1](env)))))
    if name in ("ABS", "SQRT", "INT", "NOT"):
        need(1)
        f = {"ABS": abs, "INT": lambda x: float(math.floor(x)), "NOT": lambda x: 0.0 if x else 1.0,
             "SQRT": lambda x: math.sqrt(x) if x >= 0 else (_ for _ in ()).throw(_Err(errors.NUM))}[name]
        return lambda env: f(_to_n(args[0](env)))
    need(1)
    if name == "MIN":
        return lambda env: min(_to_n(a(env)) for a in args)
    if name == "MAX":
        return lambda env: max(_to_n(a(env)) for a in args)
    if name == "AND":
        return lambda env: 1.0 if all(_to_n(a(env)) for a in args) else 0.0
    return lambda env: 1.0 if any(_to_n(a(env)) for a in args) else 0.0


def check_calc(wb, pv, name, formula):
    """Error text for a calculated field definition, or None when it's fine."""
    name = name.strip()
    if not name:
        return "Type a name for the field."
    if name.lower() in (n.lower() for n in source_fields(wb, pv)) or name == VALUES:
        return f"'{name}' is already a field of the source data. Pick another name."
    try:
        compile_calc(formula, source_fields(wb, pv))
    except CalcError as e:
        return str(e)
    return None


# ---------------------------------------------------------------- aggregation
def _aggregate(vals, agg):
    if agg == "count":
        return float(sum(1 for v in vals if v not in (None, "")))
    if agg == "distinct":
        return float(len({_item_key(v) for v in vals if v not in (None, "")}))
    nums = [float(v) for v in vals if is_num(v) and not isinstance(v, bool)]
    if agg == "countnums":
        return float(len(nums))
    if agg == "sum":
        return sum(nums)
    if not nums:
        return None
    if agg == "average":
        return sum(nums) / len(nums)
    if agg == "product":
        return float(math.prod(nums))
    if agg in ("stdev", "var", "stdevp", "varp"):
        n = len(nums)
        sample = agg in ("stdev", "var")
        if n < (2 if sample else 1):
            return errors.DIV0
        m = sum(nums) / n
        var = sum((x - m) ** 2 for x in nums) / (n - 1 if sample else n)
        return math.sqrt(var) if agg.startswith("stdev") else var
    return max(nums) if agg == "max" else min(nums)


class _Data:
    """The source records plus per-field (key, label) columns (with grouping applied) and aggregation."""

    def __init__(self, wb, pv):
        self.partials = None
        sh = wb.get_sheet(pv["source"])
        if sh is not None and getattr(sh, "big", None) is not None:
            self._init_big(sh, pv)
        else:
            self.names, self.rows, self.fmts, self.srcrows = _records(wb, pv)
        self.idx = {n: i for i, n in enumerate(self.names)}
        self.groups = {f: g for f, g in (pv.get("groups") or {}).items() if f in self.idx}
        self.virt = {v: bl for v, bl in virtual_fields(pv).items() if bl[0] in self.idx}
        self.calc = {}
        for n, text in (pv.get("calc") or {}).items():
            if n in self.idx:
                continue
            try:
                self.calc[n] = compile_calc(text, self.names)
            except CalcError:
                pass
        self._keys = {}

    def _init_big(self, sh, pv):
        """A big-file sheet (sheets/bigdata.py, millions of rows in a pyarrow table): group the rows by
        the fields the pivot shows or filters on with pyarrow, keeping per-group partial aggregates
        (count, sum, min, max, sum of squares, distinct values) of the value fields. Each group then
        acts as one record. The user's cell edits on a big sheet (big.edits) are ignored for now: the
        pivot reads the file's data (honouring the current sort/filter view)."""
        from .bigdata import arrow
        from .io_csv import convert_field
        pa, pc = arrow()
        big = sh.big
        names = source_fields(sh.wb, pv)
        self.names = names
        r1, c1, r2, _ = pv["src"]
        first = r1 + 1
        last = min(r2, sh.max_row, big.nview - 1)
        n = max(0, last - first + 1)
        view = big.view.slice(first, n) if big.view is not None else None

        def part(arr):
            return arr.slice(first, n) if view is None else pc.take(arr, view)
        virt = virtual_fields(pv)

        def src_of(f):
            return virt[f][0] if f in virt else f
        dims = []
        used = list(pv.get("rows", [])) + list(pv.get("cols", [])) + list(pv.get("filter_fields", [])) + \
            list(pv.get("filters") or {}) + list(pv.get("label_filters") or {})
        for f in used:
            f = src_of(f)
            if f in names and f not in dims:
                dims.append(f)
        calc = pv.get("calc") or {}
        vcols, distinct = [], set()
        for e in pv.get("values", []):
            refs = [src_of(e[0])]
            if e[0] in calc:
                try:
                    refs = list(compile_calc(calc[e[0]], names)[1])
                except CalcError:
                    refs = []
            for f in refs:
                if f in names and f not in vcols:
                    vcols.append(f)
                if e[1] == "distinct" and f in names:
                    distinct.add(f)
        cols, aggs = {}, []
        for j, f in enumerate(dims):
            cols[f"d{j}"] = pc.fill_null(part(big.cols[c1 + names.index(f)]), "")
        for j, f in enumerate(vcols):
            c = c1 + names.index(f)
            x = part(big.numeric(c))
            t = part(big.cols[c])
            cols[f"s{j}"] = x
            cols[f"q{j}"] = pc.multiply(x, x)
            cols[f"e{j}"] = pc.cast(pc.and_(pc.is_valid(t), pc.not_equal(pc.fill_null(t, ""), "")), pa.int64())
            aggs += [(f"s{j}", "sum"), (f"s{j}", "count"), (f"s{j}", "min"), (f"s{j}", "max"),
                     (f"q{j}", "sum"), (f"e{j}", "sum")]
            if f in distinct:
                aggs.append((f"t{j}", "distinct"))
                cols[f"t{j}"] = t
        keys = [f"d{j}" for j in range(len(dims))]
        if not keys:    # one group: a constant (null) key
            cols["_k"] = pa.nulls(n, pa.int8())
            keys = ["_k"]
        res = pa.table(cols).group_by(keys).aggregate(aggs)
        rt = {}
        dimvals = []
        for j in range(len(dims)):
            conv = {}
            out = []
            for t in res.column(f"d{j}").to_pylist():
                if t not in conv:
                    got = convert_field(t, rt) if t else None
                    conv[t] = got[0] if got else None
                out.append(conv[t])
            dimvals.append(out)
        m = res.num_rows
        rows = []
        for i in range(m):
            row = [None] * len(names)
            for j, f in enumerate(dims):
                row[names.index(f)] = dimvals[j][i]
            rows.append(tuple(row))
        partials = [dict() for _ in range(m)]
        for j, f in enumerate(vcols):
            got = {k: res.column(f"{col}_{op}").to_pylist() for k, (col, op) in
                   {"sum": (f"s{j}", "sum"), "n": (f"s{j}", "count"), "min": (f"s{j}", "min"),
                    "max": (f"s{j}", "max"), "sq": (f"q{j}", "sum"), "ne": (f"e{j}", "sum")}.items()}
            dl = res.column(f"t{j}_distinct").to_pylist() if f in distinct else None
            for i in range(m):
                partials[i][f] = {k: v[i] for k, v in got.items()}
                if dl is not None:
                    partials[i][f]["distinct"] = dl[i]
        self.rows = rows
        self.partials = partials
        self.srcrows = []
        self.fmts = [sh.style(first, c).numfmt if n else "General" for c in range(c1, c1 + len(names))]

    def _agg_big(self, f, a, idxs):
        ps = [self.partials[k].get(f) for k in idxs]
        ps = [p for p in ps if p]
        if not ps:
            return None
        nn = sum(p["n"] or 0 for p in ps)
        if a == "count":
            return float(sum(p["ne"] or 0 for p in ps))
        if a == "countnums":
            return float(nn)
        if a == "distinct":
            vals = set()
            for p in ps:
                vals.update(x for x in (p.get("distinct") or []) if x)
            return float(len(vals))
        tot = sum(p["sum"] or 0.0 for p in ps)
        if a == "sum":
            return tot
        if not nn:
            return None
        if a == "average":
            return tot / nn
        if a == "max":
            return max(p["max"] for p in ps if p["max"] is not None)
        if a == "min":
            return min(p["min"] for p in ps if p["min"] is not None)
        if a in ("stdev", "var", "stdevp", "varp"):
            sample = a in ("stdev", "var")
            if nn < (2 if sample else 1):
                return errors.DIV0
            q = sum(p["sq"] or 0.0 for p in ps)
            var = max(0.0, (q - tot * tot / nn) / (nn - 1 if sample else nn))
            return math.sqrt(var) if a.startswith("stdev") else var
        return errors.NA    # Product isn't available for big files

    def known(self, f):
        return f in self.idx or f in self.virt

    def valuable(self, f):
        return f in self.idx or f in self.virt or f in self.calc

    def keys(self, f):
        """[(key, label)] per record for an axis / filter field."""
        got = self._keys.get(f)
        if got is not None:
            return got
        if f in self.virt:
            base, level = self.virt[f]
            i = self.idx[base]
            out = [_date_part(r[i], level) if is_num(r[i]) and not isinstance(r[i], bool)
                   else (_item_key(r[i]), item_label(r[i])) for r in self.rows]
        else:
            i = self.idx[f]
            g = self.groups.get(f)
            if g and g.get("date"):
                level = g["date"][-1]
                out = [_date_part(r[i], level) if is_num(r[i]) and not isinstance(r[i], bool)
                       else (_item_key(r[i]), item_label(r[i])) for r in self.rows]
            elif g and g.get("num"):
                out = [_num_bucket(float(r[i]), g["num"]) if is_num(r[i]) and not isinstance(r[i], bool)
                       else (_item_key(r[i]), item_label(r[i])) for r in self.rows]
            else:
                fmt = self.fmts[i] if i < len(self.fmts) else "General"
                if fmt and fmt != "General":   # numbers and dates labelled the way the source shows them
                    from .numfmt import format_value
                    out = [(_item_key(r[i]), format_value(r[i], fmt)[0] if is_num(r[i]) and not
                            isinstance(r[i], bool) else item_label(r[i])) for r in self.rows]
                else:
                    out = [(_item_key(r[i]), item_label(r[i])) for r in self.rows]
        self._keys[f] = out
        return out

    def raw(self, f):
        if f in self.idx:
            i = self.idx[f]
            return [r[i] for r in self.rows]
        return [lab for _, lab in self.keys(f)]

    def fmt(self, f):
        return self.fmts[self.idx[f]] if f in self.idx else "General"

    def agg(self, entry, idxs):
        if not idxs:
            return None
        f, a = entry[0], entry[1]
        if f in self.calc:
            fn, refs = self.calc[f]
            env = {}
            for ref in refs:
                i = self.idx[ref]
                if self.partials is not None:
                    env[ref] = sum((self.partials[k].get(ref) or {}).get("sum") or 0.0 for k in idxs)
                    continue
                env[ref] = sum(float(self.rows[k][i]) for k in idxs
                               if is_num(self.rows[k][i]) and not isinstance(self.rows[k][i], bool))
            try:
                v = fn(env)
            except _Err as e:
                return e.e
            except (ValueError, OverflowError, ZeroDivisionError, TypeError):
                return errors.NUM
            return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else v
        if self.partials is not None:
            return self._agg_big(self.virt[f][0] if f in self.virt else f, a, idxs)
        if f in self.idx:
            i = self.idx[f]
            return _aggregate([self.rows[k][i] for k in idxs], a)
        col = self.raw(f)
        return _aggregate([col[k] for k in idxs], a)


def _label_ok(label, spec):
    op, a = spec[0], str(spec[1] if len(spec) > 1 and spec[1] is not None else "")
    b = str(spec[2] if len(spec) > 2 and spec[2] is not None else "")
    t = label.lower()
    al, bl = a.lower(), b.lower()

    def cmp(x, y):
        try:
            fx, fy = float(x), float(y)
            return (fx > fy) - (fx < fy)
        except ValueError:
            return (x > y) - (x < y)
    if op in ("equals", "not_equals"):
        hit = fnmatch.fnmatchcase(t, al) if any(ch in al for ch in "*?") else t == al
        return hit if op == "equals" else not hit
    if op in ("begins", "not_begins"):
        return t.startswith(al) == (op == "begins")
    if op in ("ends", "not_ends"):
        return t.endswith(al) == (op == "ends")
    if op in ("contains", "not_contains"):
        return (al in t) == (op == "contains")
    if op == "gt":
        return cmp(t, al) > 0
    if op == "ge":
        return cmp(t, al) >= 0
    if op == "lt":
        return cmp(t, al) < 0
    if op == "le":
        return cmp(t, al) <= 0
    lo, hi = (al, bl) if cmp(al, bl) <= 0 else (bl, al)
    inside = cmp(t, lo) >= 0 and cmp(t, hi) <= 0
    return inside if op == "between" else not inside


def _to_float(x, default=0.0):
    try:
        return float(x)
    except (TypeError, ValueError):
        return default


def _value_choose(vals, spec):
    """Item keys kept by a value filter; vals = {key: aggregate}."""
    op = spec[0]
    nums = {k: v for k, v in vals.items() if is_num(v)}
    if op == "top":
        n = _to_float(spec[2], 10)
        kind = spec[3] if len(spec) > 3 and spec[3] else "items"
        bottom = len(spec) > 4 and spec[4] == "bottom"
        order = sorted(nums, key=lambda k: nums[k], reverse=not bottom)
        if kind == "items":
            return set(order[:max(0, int(n))])
        total = sum(nums.values())
        goal = total * n / 100 if kind == "percent" else n
        out, run = set(), 0.0
        for k in order:
            if (run >= goal) if not bottom else (run >= goal and out):
                break
            out.add(k)
            run += nums[k]
        return out
    a = _to_float(spec[2] if len(spec) > 2 else 0)
    b = _to_float(spec[3] if len(spec) > 3 else 0)
    lo, hi = min(a, b), max(a, b)
    test = {"equals": lambda v: v == a, "not_equals": lambda v: v != a, "gt": lambda v: v > a,
            "ge": lambda v: v >= a, "lt": lambda v: v < a, "le": lambda v: v <= a,
            "between": lambda v: lo <= v <= hi, "not_between": lambda v: not lo <= v <= hi}.get(op)
    if test is None:
        return set(vals)
    return {k for k, v in nums.items() if test(v)}


# ---------------------------------------------------------------- the report
class Cells(dict):
    """{(dr, dc): (value, style)} plus .hot (in-sheet buttons) and, with want_map, .cmap / .ctx."""
    hot = ()
    cmap = None
    ctx = None


class _Node:
    __slots__ = ("field", "key", "label", "dpath", "vi", "level", "children", "collapsed", "expandable",
                 "parent")

    def __init__(self, field, key, label, dpath, vi, level, parent):
        self.field, self.key, self.label, self.dpath, self.vi, self.level = field, key, label, dpath, vi, level
        self.children = []
        self.collapsed = False
        self.expandable = False
        self.parent = parent

    def ancestors(self):
        out, n = [], self
        while n is not None:
            out.append(n)
            n = n.parent
        return out[::-1]


def is_collapsed(pv, field, label):
    c = (pv.get("collapsed") or {}).get(field)
    if not c:
        return False
    return ("*" in c and "!" + label not in c) or label in c


def toggled(pv, field, label, collapse):
    """pv with one item of `field` collapsed (True) or expanded (False)."""
    c = list((pv.get("collapsed") or {}).get(field) or [])
    if "*" in c:
        c = [x for x in c if x != "!" + label]
        if not collapse:
            c.append("!" + label)
    else:
        c = [x for x in c if x != label]
        if collapse:
            c.append(label)
    return with_key(pv, "collapsed", field, c or None)


def field_collapsed(pv, field, collapse):
    """Expand / Collapse Entire Field."""
    return with_key(pv, "collapsed", field, ["*"] if collapse else None)


def _style_set(lay):
    S = STYLES.get(lay.get("style"), STYLES["Light 9"])
    line = ("thin", S["line"])
    head = DEFAULT_STYLE.with_(bold=True, fill=S["head"], color=S["head_color"], border=(None, None, None, line))
    sub = DEFAULT_STYLE.with_(bold=True, fill=S["sub"])
    total = DEFAULT_STYLE.with_(bold=True, fill=S["sub"], border=(None, ("thin", S["line"]), None, None)) \
        if S["sub"] else DEFAULT_STYLE.with_(bold=True)
    grand = DEFAULT_STYLE.with_(bold=True, fill=S["grand"], border=(None, line, None, None))
    return {"head": head, "sub": sub, "total": total, "grand": grand, "body": DEFAULT_STYLE,
            "band": S["band"], "report": DEFAULT_STYLE.with_(bold=True)}


def build(wb, pv, want_map=False):
    """-> (cells {(dr, dc): (value, style)}, (height, width), error or None), relative to the anchor.
    cells.hot lists the in-sheet buttons: ["flt", dr, dc, area, [fields], active] filter dropdowns and
    ["tog", dr, dc, field, label, collapsed, indent] expand/collapse buttons."""
    D = _Data(wb, pv)
    if not D.names:
        return Cells(), (0, 0), f"The source sheet '{pv['source']}' isn't there any more."
    lay = layout(pv)
    form = lay["form"]
    ST = _style_set(lay)
    rax = [f for f in pv["rows"] if D.known(f) or f == VALUES]
    cax = [f for f in pv["cols"] if D.known(f) or f == VALUES]
    ff = [f for f in pv.get("filter_fields", []) if D.known(f)]
    vf = [e for e in pv["values"] if D.valuable(e[0])]
    vlabels = value_labels(changed(pv, values=vf))
    nv = len(vf)
    if nv >= 2 and VALUES not in rax and VALUES not in cax:
        cax = cax + [VALUES]
    if nv < 2:
        rax = [f for f in rax if f != VALUES]
        cax = [f for f in cax if f != VALUES]
    rdf = [f for f in rax if f != VALUES]
    cdf = [f for f in cax if f != VALUES]
    cells = Cells()
    hot = []
    cells.hot = hot
    cmap = {} if want_map else None
    cells.cmap = cmap
    if not (rdf or cdf or vf or ff):  # a new, empty pivot: Excel's placeholder
        tip = DEFAULT_STYLE.with_(italic=True, color="#595959", fill=HEAD_FILL)
        cells[(0, 0)] = (f"{pv['name']}: tick fields in the PivotTable Fields panel to build the report", tip)
        return cells, (1, 1), None
    show_btn = lay.get("buttons", True)

    # ---- filtering
    n = len(D.rows)
    keep = list(range(n))
    axis_set = set(rdf) | set(cdf)
    filters = pv.get("filters") or {}
    lfilters = pv.get("label_filters") or {}
    vfilters = pv.get("value_filters") or {}

    def item_ok(f, label):
        allowed = filters.get(f)
        if allowed is not None and (f in ff or f in axis_set) and label not in set(allowed):
            return False
        spec = lfilters.get(f)
        return not (spec and f in axis_set and not _label_ok(label, spec))
    for f in set(filters) | set(lfilters):
        if D.known(f) and (f in ff or f in axis_set):
            col = D.keys(f)
            ok_labels = {}
            keep = [i for i in keep if ok_labels.setdefault(col[i][1], item_ok(f, col[i][1]))]

    def vi_of(label):
        if label is None:
            return None
        return vlabels.index(label) if label in vlabels else None
    for flds in (rdf, cdf):
        for L, f in enumerate(flds):
            spec = vfilters.get(f)
            vi = vi_of(spec[1]) if spec else None
            if vi is None:
                continue
            pk = [D.keys(g) for g in flds[:L]]
            fk = D.keys(f)
            grp = {}
            for i in keep:
                grp.setdefault(tuple(k[i][0] for k in pk), {}).setdefault(fk[i][0], []).append(i)
            ok = set()
            for parent, its in grp.items():
                vals = {k: D.agg(vf[vi], idxs) for k, idxs in its.items()}
                ok |= {(parent, k) for k in _value_choose(vals, spec)}
            keep = [i for i in keep if (tuple(k[i][0] for k in pk), fk[i][0]) in ok]

    # ---- buckets
    labels = {}
    rcols = [D.keys(f) for f in rdf]
    ccols = [D.keys(f) for f in cdf]
    bucket = {}
    kids = {"r": {}, "c": {}}
    seen_r, seen_c = set(), set()
    nr, nc = len(rdf), len(cdf)
    for i in keep:
        rk = tuple(c[i][0] for c in rcols)
        ck = tuple(c[i][0] for c in ccols)
        for a in range(nr + 1):
            rp = rk[:a]
            for b in range(nc + 1):
                k = (rp, ck[:b])
                lst = bucket.get(k)
                if lst is None:
                    bucket[k] = [i]
                else:
                    lst.append(i)
        if rk not in seen_r:
            seen_r.add(rk)
            for a in range(nr):
                kids["r"].setdefault(rk[:a], set()).add(rk[a])
                labels.setdefault((rdf[a], rk[a]), rcols[a][i][1])
        if ck not in seen_c:
            seen_c.add(ck)
            for b in range(nc):
                kids["c"].setdefault(ck[:b], set()).add(ck[b])
                labels.setdefault((cdf[b], ck[b]), ccols[b][i][1])
    allkeys = {}

    def all_keys(f):
        if f not in allkeys:
            s = set()
            for k, lab in D.keys(f):
                if item_ok(f, lab):
                    s.add(k)
                    labels.setdefault((f, k), lab)
            allkeys[f] = s
        return allkeys[f]

    vcache = {}

    def val(rp, cp, vi):
        k = (rp, cp, vi)
        if k not in vcache:
            vcache[k] = D.agg(vf[vi], bucket.get((rp, cp)))
        return vcache[k]

    sorts = pv.get("sort") or {}
    ocache = {}

    def ordered(axis, prefix):
        k = (axis, prefix)
        if k in ocache:
            return ocache[k]
        flds = rdf if axis == "r" else cdf
        f = flds[len(prefix)]
        ks = all_keys(f) if lay["show_empty_items"] else kids[axis].get(prefix, set())
        spec = sorts.get(f) or ["asc", None]
        desc = spec[0] == "desc"
        vi = vi_of(spec[1] if len(spec) > 1 else None)
        if vi is not None and nv:
            def v_of(key):
                v = val(prefix + (key,), (), vi) if axis == "r" else val((), prefix + (key,), vi)
                return v if is_num(v) else None
            have = [x for x in ks if v_of(x) is not None]
            rest = sorted(x for x in ks if v_of(x) is None)
            out = sorted(have, key=lambda x: (v_of(x), x), reverse=desc) + rest
        else:
            out = sorted(ks, reverse=desc)
        ocache[k] = out
        return out

    # ---- axis trees
    def tree(ax, axis):
        root = _Node(None, None, None, (), None, -1, None)

        def fill(node, level):
            if level >= len(ax):
                return
            f = ax[level]
            if f == VALUES:
                for vi in range(nv):
                    ch = _Node(VALUES, ("v", vi), vlabels[vi], node.dpath, vi, level, node)
                    node.children.append(ch)
                    fill(ch, level + 1)
                return
            for key in ordered(axis, node.dpath):
                ch = _Node(f, key, labels.get((f, key), ""), node.dpath + (key,), node.vi, level, node)
                ch.expandable = any(g != VALUES for g in ax[level + 1:])
                node.children.append(ch)
                if ch.expandable and is_collapsed(pv, f, ch.label):
                    ch.collapsed = True
                    if VALUES in ax[level + 1:]:
                        vl = ax.index(VALUES)
                        for vi in range(nv):
                            ch.children.append(_Node(VALUES, ("v", vi), vlabels[vi], ch.dpath, vi, vl, ch))
                else:
                    fill(ch, level + 1)
        fill(root, 0)
        return root

    sub_mode = lay["subtotals"]
    if form == "tabular" and sub_mode == "top":
        sub_mode = "bottom"

    # ---- data columns
    ccolumns = []   # dict(kind, node, dp, vi, label)
    if cdf:
        croot = tree(cax, "c")

        def walk_c(node):
            for ch in node.children:
                if not ch.children:
                    ccolumns.append({"kind": "item", "node": ch, "dp": ch.dpath, "vi": ch.vi})
                    continue
                walk_c(ch)
                if ch.field == VALUES or sub_mode == "none" or not ch.expandable or ch.collapsed:
                    continue
                if VALUES in cax[ch.level + 1:] and not ch.collapsed:
                    for vi in range(nv):
                        ccolumns.append({"kind": "total", "node": ch, "dp": ch.dpath, "vi": vi,
                                         "label": f"{ch.label} {vlabels[vi]}"})
                else:
                    ccolumns.append({"kind": "total", "node": ch, "dp": ch.dpath, "vi": ch.vi,
                                     "label": f"{ch.label} Total"})
        walk_c(croot)
        if lay["grand"] in ("both", "rows") and nv:
            if VALUES in cax:
                for vi in range(nv):
                    ccolumns.append({"kind": "grand", "node": None, "dp": (), "vi": vi,
                                     "label": f"Total {vlabels[vi]}"})
            else:
                ccolumns.append({"kind": "grand", "node": None, "dp": (), "vi": None, "label": "Grand Total"})
    elif VALUES in cax:
        for vi in range(nv):
            ccolumns.append({"kind": "value", "node": None, "dp": (), "vi": vi, "label": vlabels[vi]})
    elif nv:
        ccolumns.append({"kind": "value", "node": None, "dp": (), "vi": None, "label": vlabels[0]})

    # ---- row lines
    rlines = []     # dict(kind, node, dp, vi, labels: [(col, text, indent, node, kind)])
    left = 1 if form == "compact" else max(1, len(rax))
    btn_rows = show_btn and len(rdf) >= 2
    rcol = {f: i for i, f in enumerate(rax)}

    def add_line(kind, node, dp, vi, data=True):
        ln = {"kind": kind, "node": node, "dp": dp, "vi": vi, "data": data, "labels": []}
        rlines.append(ln)
        return ln

    def item_col(node):
        return 0 if form == "compact" else rcol.get(node.field, 0)

    def item_indent(node, depth):
        if form == "compact":
            return depth + (1 if btn_rows else 0)
        return 1 if (node.expandable and show_btn) else 0

    def walk_r(node, depth):
        for ch in node.children:
            start = len(rlines)
            if not ch.children:
                ln = add_line("item", ch, ch.dpath, ch.vi)
                ln["labels"].append((item_col(ch), ch.label, item_indent(ch, depth), ch, "item"))
                continue
            vbelow = VALUES in rax[ch.level + 1:] and not ch.collapsed and ch.field != VALUES and ch.expandable
            if form == "tabular":
                walk_r(ch, depth + 1)
                if len(rlines) > start:
                    rlines[start]["labels"].append((item_col(ch), ch.label, item_indent(ch, depth), ch, "item"))
            else:
                top = sub_mode == "top" and not vbelow and ch.expandable and not ch.collapsed
                ln = add_line("head", ch, ch.dpath, ch.vi, data=top)
                ln["labels"].append((item_col(ch), ch.label, item_indent(ch, depth), ch, "item"))
                walk_r(ch, depth + 1)
            if lay["repeat"] and form != "compact":
                for ln in rlines[start + 1:]:
                    if ln["kind"] in ("item", "head") and not any(lb[3] is ch for lb in ln["labels"]):
                        ln["labels"].append((item_col(ch), ch.label, item_indent(ch, depth), ch, "repeat"))
            if ch.expandable and not ch.collapsed and sub_mode != "none" and (sub_mode == "bottom" or vbelow):
                tind = depth + (1 if btn_rows else 0) if form == "compact" else 0
                if vbelow:
                    for vi in range(nv):
                        ln = add_line("total", ch, ch.dpath, vi)
                        ln["labels"].append((item_col(ch), f"{ch.label} {vlabels[vi]}", tind, ch, "total"))
                else:
                    ln = add_line("total", ch, ch.dpath, ch.vi)
                    ln["labels"].append((item_col(ch), f"{ch.label} Total", tind, ch, "total"))
            if lay["blank_line"] and ch.field != VALUES and ch.expandable:
                add_line("blank", ch, (), None, data=False)

    if rdf:
        rroot = tree(rax, "r")
        walk_r(rroot, 0)
        while rlines and rlines[-1]["kind"] == "blank":
            rlines.pop()
        if lay["grand"] in ("both", "cols") and nv:
            if VALUES in rax:
                for vi in range(nv):
                    ln = add_line("grand", None, (), vi)
                    ln["labels"].append((0, f"Total {vlabels[vi]}", 0, None, "grand"))
            else:
                ln = add_line("grand", None, (), None)
                ln["labels"].append((0, "Grand Total", 0, None, "grand"))
    elif VALUES in rax:
        for vi in range(nv):
            ln = add_line("item", None, (), vi)
            ln["labels"].append((0, vlabels[vi], 0, None, "value"))
    elif nv:
        ln = add_line("item", None, (), None)
        ln["labels"].append((0, vlabels[0] if (cdf and nv == 1) else "", 0, None, "value"))

    # ---- report filters, above the table
    y = 0
    for f in ff:
        allowed = filters.get(f)
        shown = "(All)" if allowed is None else (allowed[0] if len(allowed) == 1 else "(Multiple Items)")
        cells[(y, 0)] = (f, ST["report"])
        cells[(y, 1)] = (shown, DEFAULT_STYLE)
        if show_btn:
            hot.append(["flt", y, 1, "report", [f], allowed is not None])
        if cmap is not None:
            cmap[(y, 0)] = ("report", f)
            cmap[(y, 1)] = ("report", f)
        y += 1
    if ff:
        y += 1

    width = left + len(ccolumns)
    top = y
    hr = 1 + len(cax) if cdf else 1
    for r in range(hr):
        for c in range(width):
            cells[(top + r, c)] = ("", ST["head"])

    def active(flds):
        return any(f in filters or f in lfilters or f in vfilters for f in flds)

    # captions
    if cdf:
        if nv == 1 and (rdf or form != "compact"):
            cells[(top, 0)] = (vlabels[0], ST["head"])
        if form == "compact":
            cells[(top, left)] = ("Column Labels", ST["head"])
            if show_btn:
                hot.append(["flt", top, left, "cols", list(cdf), active(cdf)])
            if cmap is not None:
                cmap[(top, left)] = ("caption", "cols", list(cdf))
        else:
            for i, f in enumerate(cax):
                if left + i >= width:
                    break
                cells[(top, left + i)] = ("Values" if f == VALUES else f, ST["head"])
                if f != VALUES:
                    if show_btn:
                        hot.append(["flt", top, left + i, "cols", [f], active([f])])
                    if cmap is not None:
                        cmap[(top, left + i)] = ("caption", "cols", [f])
    cap_row = top + hr - 1
    if form == "compact":
        if rdf:
            cells[(cap_row, 0)] = ("Row Labels", ST["head"])
            if show_btn:
                hot.append(["flt", cap_row, 0, "rows", list(rdf), active(rdf)])
            if cmap is not None:
                cmap[(cap_row, 0)] = ("caption", "rows", list(rdf))
        elif VALUES in rax:
            cells[(cap_row, 0)] = ("Values", ST["head"])
    else:
        for i, f in enumerate(rax):
            cells[(cap_row, i)] = ("Values" if f == VALUES else f, ST["head"])
            if f != VALUES:
                if show_btn:
                    hot.append(["flt", cap_row, i, "rows", [f], active([f])])
                if cmap is not None:
                    cmap[(cap_row, i)] = ("caption", "rows", [f])
    # column headers
    btn_cols = show_btn and len(cdf) >= 2
    for c, col in enumerate(ccolumns):
        x = left + c
        kind = col["kind"]
        if kind == "value" or not cdf:
            cells[(cap_row, x)] = (col["label"], ST["head"])
            if cmap is not None:
                cmap[(cap_row, x)] = ("vhead", col["vi"] if col["vi"] is not None else 0)
            continue
        if kind == "grand":
            cells[(top + 1, x)] = (col["label"], ST["head"])
            if cmap is not None:
                cmap[(top + 1, x)] = ("grand", "c")
            continue
        node = col["node"]
        if kind == "total":
            cells[(top + 1 + node.level, x)] = (col["label"], ST["head"])
            if cmap is not None:
                cmap[(top + 1 + node.level, x)] = ("ctotal", node)
            continue
        prev = ccolumns[c - 1] if c else None
        prev_anc = prev["node"].ancestors()[1:] if prev and prev["node"] is not None and prev["kind"] == "item" \
            else []
        for nd in node.ancestors()[1:]:
            if nd in prev_anc:
                continue
            st = ST["head"]
            ind = 1 if (btn_cols and nd.expandable) else 0
            if ind:
                st = st.with_(indent=ind)
            cells[(top + 1 + nd.level, x)] = (nd.label, st)
            if cmap is not None:
                cmap[(top + 1 + nd.level, x)] = ("citem", nd)
            if ind:
                hot.append(["tog", top + 1 + nd.level, x, nd.field, nd.label, nd.collapsed, 0])
    y = top + hr

    # ---- body
    empty = lay.get("empty") or ""
    try:
        empty_v = float(empty) if empty.strip() else None
    except ValueError:
        empty_v = None
    S = STYLES.get(lay.get("style"), STYLES["Light 9"])
    band_n = 0

    def fmt_for(vi):
        e = vf[vi]
        o = vopts(e)
        if o.get("numfmt"):
            return o["numfmt"]
        t = o.get("show", "none")
        if t in _PCT_SHOWS:
            return "0.00%"
        if t == "index":
            return "0.00"
        if t in ("rank_asc", "rank_desc"):
            return "0"
        if e[0] in D.calc:
            return "General"
        return D.fmt(e[0]) if e[1] in ("sum", "average", "max", "min", "product") else "General"

    def ratio(a, b):
        if a is None or isinstance(a, XLError):
            return a
        if not is_num(b) or b == 0:
            return errors.DIV0
        return a / b

    base_keys = {}

    def shown(rp, cp, vi):
        v = val(rp, cp, vi)
        o = vopts(vf[vi])
        t = o.get("show", "none")
        if t == "none" or isinstance(v, XLError):
            return v
        if t == "pct_grand":
            return ratio(v, val((), (), vi))
        if t == "pct_col":
            return ratio(v, val((), cp, vi))
        if t == "pct_row":
            return ratio(v, val(rp, (), vi))
        if t == "pct_parent_row":
            return ratio(v, val(rp[:-1], cp, vi)) if rp else (None if v is None else 1.0)
        if t == "pct_parent_col":
            return ratio(v, val(rp, cp[:-1], vi)) if cp else (None if v is None else 1.0)
        if t == "index":
            g, r_, c_ = val((), (), vi), val(rp, (), vi), val((), cp, vi)
            if v is None:
                return None
            if not (is_num(g) and is_num(r_) and is_num(c_)) or r_ * c_ == 0:
                return errors.DIV0
            return v * g / (r_ * c_)
        bf = o.get("base_field")
        if bf in rdf:
            axis, L, p = "r", rdf.index(bf), rp
        elif bf in cdf:
            axis, L, p = "c", cdf.index(bf), cp
        else:
            return errors.NA

        def at(key):
            p2 = p[:L] + (key,) + p[L + 1:]
            return val(p2, cp, vi) if axis == "r" else val(rp, p2, vi)
        if len(p) <= L:
            return None
        if t == "pct_parent":
            return ratio(v, val(p[:L + 1], cp, vi) if axis == "r" else val(rp, p[:L + 1], vi))
        sibs = ordered(axis, p[:L])
        if t in ("pct_of", "diff", "pct_diff"):
            bi = o.get("base_item") or "(previous)"
            if bi in ("(previous)", "(next)"):
                pos = sibs.index(p[L]) if p[L] in sibs else -1
                j = pos + (-1 if bi == "(previous)" else 1)
                bk = sibs[j] if pos >= 0 and 0 <= j < len(sibs) else p[L]
            else:
                if (bf, bi) not in base_keys:
                    base_keys[(bf, bi)] = next((k for (f, k), lab in labels.items() if f == bf and lab == bi), None)
                bk = base_keys[(bf, bi)]
                if bk is None:
                    return errors.NA
            if bk == p[L]:
                return (None if v is None else 1.0) if t == "pct_of" else None
            b = at(bk)
            if t == "pct_of":
                return errors.NULL if not is_num(b) else ratio(v, b)
            if t == "diff":
                if v is None and b is None:
                    return None
                return (v if is_num(v) else 0.0) - (b if is_num(b) else 0.0)
            if not is_num(b) or b == 0:
                return errors.NULL
            return ((v if is_num(v) else 0.0) - b) / b
        if p[L] not in sibs:
            return None
        pos = sibs.index(p[L])
        if t in ("running", "pct_running"):
            run = sum(x for x in (at(k) for k in sibs[:pos + 1]) if is_num(x))
            if t == "running":
                return run
            tot = sum(x for x in (at(k) for k in sibs) if is_num(x))
            return ratio(run, tot)
        if v is None or not is_num(v):
            return None
        others = [x for x in (at(k) for k in sibs) if is_num(x)]
        if t == "rank_asc":
            return float(1 + sum(1 for x in others if x < v))
        return float(1 + sum(1 for x in others if x > v))

    for ln in rlines:
        kind = ln["kind"]
        if kind == "blank":
            y += 1
            continue
        if kind == "grand":
            base = ST["grand"]
        elif kind == "total":
            base = ST["total"]
        elif kind == "head" or (ln["node"] is not None and ln["node"].collapsed):
            base = ST["sub"]
        else:
            base = ST["body"]
            if lay["banded_rows"] and kind == "item":
                if band_n % 2 == 1:
                    base = base.with_(fill=S["band"])
                band_n += 1
        for c in range(left):
            cells[(y, c)] = ("", base)
        for col, text, ind, node, lk in ln["labels"]:
            lst = base.with_(indent=ind) if ind else base
            cells[(y, col)] = (text, lst)
            if cmap is not None:
                cmap[(y, col)] = ("ritem" if lk in ("item", "repeat") else "rtotal" if lk == "total" else
                                  "grand" if lk == "grand" else "value", node if lk != "grand" else "r")
            if lk == "item" and node is not None and node.expandable and show_btn and \
                    (form == "compact" and btn_rows or form != "compact"):
                hot.append(["tog", y, col, node.field, node.label, node.collapsed,
                            max(0, ind - 1) if form == "compact" else 0])
        for c, colinfo in enumerate(ccolumns):
            x = left + c
            st = base
            if lay["banded_cols"] and colinfo["kind"] in ("item", "value") and c % 2 == 1 and kind == "item":
                st = st.with_(fill=S["band"])
            if not ln["data"]:
                cells[(y, x)] = ("", st)
                continue
            vi = ln["vi"] if ln["vi"] is not None else colinfo["vi"]
            if vi is None:
                vi = 0
            v = shown(ln["dp"], colinfo["dp"], vi)
            if v is None:
                cells[(y, x)] = (empty_v if empty_v is not None else empty, st)
            else:
                cells[(y, x)] = (v, st.with_(numfmt=fmt_for(vi)) if is_num(v) else st)
            if cmap is not None:
                cmap[(y, x)] = ("data", ln["dp"], colinfo["dp"], vi, ln, colinfo)
        y += 1
    if cmap is not None:
        cells.ctx = {"D": D, "bucket": bucket, "rdf": rdf, "cdf": cdf, "vf": vf, "vlabels": vlabels,
                     "labels": labels, "rax": rax, "cax": cax, "left": left, "top": top, "hr": hr}
    return cells, (y, max(width, 2 if ff else 1)), None


def out_rect(pv, size):
    h, w = size
    r, c = pv["anchor"]
    return [r, c, r + max(h, 1) - 1, c + max(w, 1) - 1]


def contains(rect, r, c):
    return rect is not None and rect[0] <= r <= rect[2] and rect[1] <= c <= rect[3]


# ---------------------------------------------------------------- field helpers for the UI
def items(wb, pv, field):
    """Distinct item labels of a field (after grouping), sorted, for filter lists."""
    D = _Data(wb, changed(pv, filter_fields=list(pv.get("filter_fields", [])) + [field]))
    if not D.known(field):
        return []
    seen = {}
    for k, lab in D.keys(field):
        seen.setdefault(k, lab)
    return [seen[k] for k in sorted(seen)]


def col_stats(wb, pv, field):
    """(non-blank count, number count, min, max, numfmt of the first data row) of a source field.
    Uses pyarrow on a big-file sheet instead of reading every cell."""
    sh = wb.get_sheet(pv["source"])
    names = source_fields(wb, pv)
    if sh is None or field not in names:
        return 0, 0, None, None, "General"
    r1, c1, r2, _ = pv["src"]
    c = c1 + names.index(field)
    fmt = sh.style(r1 + 1, c).numfmt
    big = getattr(sh, "big", None)
    if big is not None:
        from .bigdata import arrow
        pa, pc = arrow()
        first = r1 + 1
        n = max(0, min(r2, sh.max_row, big.nview - 1) - first + 1)
        view = big.view.slice(first, n) if big.view is not None else None

        def part(arr):
            return arr.slice(first, n) if view is None else pc.take(arr, view)
        t = part(big.cols[c])
        x = part(big.numeric(c))
        nonblank = pc.sum(pc.cast(pc.and_(pc.is_valid(t), pc.not_equal(pc.fill_null(t, ""), "")),
                                  pa.int64())).as_py() or 0
        mm = pc.min_max(x).as_py()
        return nonblank, pc.count(x).as_py() or 0, mm["min"], mm["max"], fmt
    _, rows, _ = records(wb, pv)
    i = names.index(field)
    vals = [row[i] for row in rows if row[i] not in (None, "")]
    nums = [float(v) for v in vals if is_num(v) and not isinstance(v, bool)]
    return len(vals), len(nums), (min(nums) if nums else None), (max(nums) if nums else None), fmt


def default_agg(wb, pv, field):
    """Sum when every value of the field is a number, otherwise Count (what Excel picks)."""
    if field in (pv.get("calc") or {}):
        return "sum"
    nonblank, nnum, _, _, _ = col_stats(wb, pv, field)
    return "sum" if nonblank and nnum == nonblank else "count"


def field_kind(wb, pv, field):
    """'date', 'number', 'text', 'virtual' (a date-group field) or 'calc'."""
    if field in virtual_fields(pv):
        return "virtual"
    if field in (pv.get("calc") or {}):
        return "calc"
    from .numfmt import is_date_format
    nonblank, nnum, _, _, fmt = col_stats(wb, pv, field)
    if not nonblank or nnum != nonblank:
        return "text"
    try:
        return "date" if is_date_format(fmt) else "number"
    except Exception:
        return "number"


def number_range(wb, pv, field):
    _, _, lo, hi, _ = col_stats(wb, pv, field)
    return (lo or 0.0, hi or 0.0)


def _replace_in_areas(pv, old, new_list):
    """pv with field `old` replaced by the fields in new_list (in place, in whichever area holds it)."""
    out = dict(pv)
    for area in ("rows", "cols", "filter_fields"):
        lst = list(pv.get(area, []))
        if old in lst:
            i = lst.index(old)
            lst[i:i + 1] = new_list
            out[area] = lst
    return out


def _drop_field_state(pv, names):
    out = dict(pv)
    for key in ("filters", "label_filters", "value_filters", "sort", "collapsed"):
        d = out.get(key)
        if d and any(n in d for n in names):
            out[key] = {k: v for k, v in d.items() if k not in names}
    return out


def grouped(pv, field, spec):
    """pv with `field` grouped ({"date": levels} / {"num": [start, end, by]}) or ungrouped (spec None).
    Date-group fields (Years (Date), ...) are added to / removed from the area the field is in."""
    old_virt = [v for v, (b, _) in virtual_fields(pv).items() if b == field]
    out = with_key(pv, "groups", field, spec)
    out = _drop_field_state(out, [field] + old_virt)
    new_virt = [v for v, (b, _) in virtual_fields(out).items() if b == field]
    for area in ("rows", "cols", "filter_fields"):
        lst = [f for f in out.get(area, []) if f not in old_virt or f in new_virt]
        if field in lst:
            i = lst.index(field)
            lst = [f for f in lst if f not in new_virt]
            i = lst.index(field)
            lst[i:i] = new_virt
        out[area] = lst
    out["values"] = [e for e in out["values"] if e[0] not in old_virt or e[0] in new_virt]
    return out


def auto_group(wb, pv, field):
    """Excel 365 time grouping: a date field dropped into Rows/Columns becomes Years > Quarters > Months,
    collapsed to years when the data spans more than one year."""
    if field in (pv.get("groups") or {}) or field_kind(wb, pv, field) != "date":
        return pv
    _, _, lo, hi, _ = col_stats(wb, pv, field)
    if lo is None or int(lo) == int(hi):
        return pv
    out = grouped(pv, field, {"date": ["years", "quarters", "months"]})
    if _date_part(lo, "years")[1] != _date_part(hi, "years")[1]:
        out = field_collapsed(out, group_field_name("years", field), True)
    return out


# ---------------------------------------------------------------- lookups for the UI and GETPIVOTDATA
def cell_info(wb, pv, r, c):
    """What the sheet cell (r, c) shows in the pivot: (info tuple or None, cells) - see build's cmap."""
    cells, _, err = build(wb, pv, want_map=True)
    if err:
        return None, cells
    ar, ac = pv["anchor"]
    return cells.cmap.get((r - ar, c - ac)), cells


def detail_rows(cells, rp, cp):
    """Source rows (as tuples) and their sheet row numbers behind one value cell (Show Details)."""
    ctx = cells.ctx
    D = ctx["D"]
    idxs = ctx["bucket"].get((rp, cp), [])
    return D.names, [D.rows[i] for i in idxs], [D.srcrows[i] for i in idxs]


_LOOKUPS = {}


def _lookup_table(wb, pv):
    hit = _LOOKUPS.get(id(pv))
    if hit is not None and hit[0] is pv:
        return hit[1]
    cells, _, err = build(wb, pv, want_map=True)
    table = {}
    if not err:
        ctx = cells.ctx
        labels, rdf, cdf = ctx["labels"], ctx["rdf"], ctx["cdf"]
        for (dr, dc), info in cells.cmap.items():
            if info[0] != "data":
                continue
            _, rp, cp, vi = info[:4]
            pairs = frozenset([(rdf[i].lower(), labels[(rdf[i], k)].lower()) for i, k in enumerate(rp)] +
                              [(cdf[i].lower(), labels[(cdf[i], k)].lower()) for i, k in enumerate(cp)])
            table.setdefault((vi, pairs), (dr, dc))
        table["__values__"] = (ctx["vlabels"], [e[0] for e in ctx["vf"]])
    if len(_LOOKUPS) > 32:
        _LOOKUPS.clear()
    _LOOKUPS[id(pv)] = (pv, table)
    return table


def getpivotdata(sheet, r, c, data_field, pairs):
    """GETPIVOTDATA(data_field, pivot_table, [field, item], ...): the value the pivot at sheet!(r, c)
    shows for that value field and those items (#REF! when it doesn't show one)."""
    pv = next((p for p in getattr(sheet, "pivots", []) if contains(p.get("out"), r, c)), None)
    if pv is None:
        return errors.REF
    table = _lookup_table(sheet.wb, pv)
    if "__values__" not in table:
        return errors.REF
    vlabels, vfields = table["__values__"]
    want = str(data_field).strip().lower()
    vi = next((i for i, t in enumerate(vlabels) if t.lower() == want), None)
    if vi is None:
        vi = next((i for i, f in enumerate(vfields) if f.lower() == want), None)
    if vi is None:
        return errors.REF
    if len(pairs) % 2:
        return errors.REF
    key = frozenset((str(pairs[i]).strip().lower(), item_label(pairs[i + 1]).strip().lower())
                    for i in range(0, len(pairs), 2))
    pos = table.get((vi, key))
    if pos is None:
        return errors.REF
    v = sheet.value(pv["anchor"][0] + pos[0], pv["anchor"][1] + pos[1])
    if v in (None, ""):
        return errors.REF
    return v


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


_SAVED = ("name", "source", "src", "anchor", "rows", "cols", "values", "filters", "filter_fields", "out",
          "label_filters", "value_filters", "sort", "groups", "calc", "collapsed", "layout", "hot")


def to_json(pv):
    return {k: pv[k] for k in _SAVED if k in pv}


def from_json(d):
    pv = new(d.get("name", "PivotTable1"), d["source"], d["src"], d["anchor"])
    for k in ("rows", "cols", "values", "filter_fields"):
        pv[k] = list(d.get(k, []))
    pv["values"] = [list(x) for x in pv["values"]]
    pv["filters"] = {k: list(v) for k, v in (d.get("filters") or {}).items()}
    pv["out"] = list(d["out"]) if d.get("out") else None
    for k in ("label_filters", "value_filters", "sort", "groups", "calc", "collapsed"):
        if d.get(k):
            pv[k] = dict(d[k])
    if d.get("layout"):
        pv["layout"] = dict(d["layout"])
    else:
        pv.pop("layout", None)      # saved before layouts: keep the Phase 1 look (see PHASE1_LAYOUT)
    if d.get("hot") is not None:
        pv["hot"] = [list(h) for h in d["hot"]]
    return pv
