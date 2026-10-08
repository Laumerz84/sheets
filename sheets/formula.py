"""Formula tokenizer, parser and reference rewriting.

AST nodes are tuples:
  ('num', float) ('str', s) ('bool', b) ('err', XLError) ('missing',)
  ('cell', sheet, r, c) ('range', sheet, r1, c1, r2, c2)
  ('func', NAME, [args]) ('bin', op, a, b) ('neg', a) ('pct', a)
  ('name', text) ('array', [[...]])
`sheet` is the sheet name as written (unquoted) or None for the formula's own sheet.
"""
import re

from . import errors
from .errors import XLError
from .refs import MAX_COLS, MAX_ROWS, col_index, col_name

_SHEET = r"(?:'(?:[^']|'')+'|[A-Za-z_À-￿][\w.À-￿]*)!"
_CELL = r"\$?[A-Za-z]{1,3}\$?[0-9]+"
_COLR = r"\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}"
_ROWR = r"\$?[0-9]+:\$?[0-9]+"

_TOKEN_RE = re.compile(
    r"(?P<ws>\s+)"
    r"|(?P<str>\"(?:[^\"]|\"\")*\"?)"
    r"|(?P<err>#(?:NULL!|DIV/0!|VALUE!|REF!|NAME\?|NUM!|N/A|CIRC!|SPILL!|CALC!|GETTING_DATA))"
    rf"|(?P<ref>(?:{_SHEET})?(?:{_CELL}(?::{_CELL})?|{_COLR}|{_ROWR})(?![\w(!]))"
    r"|(?P<func>(?:_xlfn\.|_xlws\.)?[A-Za-z_][A-Za-z0-9_.]*(?=\s*\())"
    r"|(?P<num>(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?)"
    rf"|(?P<badref>(?:{_SHEET})?#REF!)"
    r"|(?P<name>[A-Za-z_\\][\w.]*)"
    r"|(?P<op><>|<=|>=|[-+*/^&=<>%])"
    r"|(?P<punct>[(),;{}:])",
    re.IGNORECASE,
)

_REFPART_RE = re.compile(
    r"^(?:(?P<sheet>'(?:[^']|'')+'|[^!]+)!)?(?P<body>.+)$")
_END_RE = re.compile(r"^(\$?)([A-Za-z]{1,3})?(\$?)([0-9]+)?$")


class Token:
    __slots__ = ("kind", "text", "pos")

    def __init__(self, kind, text, pos):
        self.kind = kind
        self.text = text
        self.pos = pos

    def __repr__(self):
        return f"Token({self.kind},{self.text!r})"


def tokenize(src):
    """Tokenize formula text (with or without the leading '=')."""
    toks = []
    i = 1 if src.startswith("=") else 0
    n = len(src)
    while i < n:
        m = _TOKEN_RE.match(src, i)
        if not m:
            toks.append(Token("bad", src[i], i))
            i += 1
            continue
        kind = m.lastgroup
        text = m.group()
        if kind == "badref":
            kind = "err"
        toks.append(Token(kind, text, i))
        i = m.end()
    return toks


# ---------------------------------------------------------------- reference pieces

class RefInfo:
    """A parsed reference token: optional sheet and one or two endpoints.

    Endpoint = [col, col_abs, row, row_abs]; col or row may be None for
    whole-row / whole-column references."""
    __slots__ = ("sheet", "sheet_text", "ends")

    def __init__(self, sheet, sheet_text, ends):
        self.sheet = sheet
        self.sheet_text = sheet_text
        self.ends = ends

    @classmethod
    def parse(cls, text):
        m = _REFPART_RE.match(text)
        sheet_text = m.group("sheet")
        sheet = None
        if sheet_text:
            sheet = sheet_text[1:-1].replace("''", "'") if sheet_text.startswith("'") else sheet_text
        ends = []
        for part in m.group("body").split(":"):
            e = _END_RE.match(part)
            col = col_index(e.group(2)) if e.group(2) else None
            row = int(e.group(4)) - 1 if e.group(4) else None
            ends.append([col, bool(e.group(1)) if col is not None else False, row,
                         bool(e.group(3) if col is not None else e.group(1))])
        return cls(sheet, sheet_text, ends)

    def bounds(self):
        """-> (r1, c1, r2, c2) normalized; whole rows/cols expanded."""
        e1 = self.ends[0]
        e2 = self.ends[-1]
        c1, c2 = e1[0], e2[0]
        r1, r2 = e1[2], e2[2]
        if c1 is None:
            c1, c2 = 0, MAX_COLS - 1
        if r1 is None:
            r1, r2 = 0, MAX_ROWS - 1
        return min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)

    def is_single(self):
        return len(self.ends) == 1

    def text(self):
        parts = []
        for col, cabs, row, rabs in self.ends:
            s = ""
            if col is not None:
                s += ("$" if cabs else "") + col_name(col)
            if row is not None:
                s += ("$" if rabs else "") + str(row + 1)
            parts.append(s)
        body = ":".join(parts)
        return (self.sheet_text + "!" if self.sheet_text else "") + body


def quote_sheet(name):
    if re.fullmatch(r"[A-Za-z_À-￿][\w.À-￿]*", name) and not re.fullmatch(
            r"[A-Za-z]{1,3}\d+", name) and name.upper() not in ("TRUE", "FALSE"):
        return name
    return "'" + name.replace("'", "''") + "'"


# ---------------------------------------------------------------- parser

class ParseError(Exception):
    pass


_BIN_PREC = {"=": 1, "<>": 1, "<": 1, ">": 1, "<=": 1, ">=": 1, "&": 2,
             "+": 3, "-": 3, "*": 4, "/": 4, "^": 5}


class _Parser:
    def __init__(self, toks):
        self.toks = [t for t in toks if t.kind != "ws"]
        self.i = 0

    def peek(self):
        return self.toks[self.i] if self.i < len(self.toks) else None

    def take(self):
        t = self.peek()
        self.i += 1
        return t

    def expect(self, text):
        t = self.take()
        if t is None or t.text != text:
            raise ParseError(f"expected {text!r}")
        return t

    def parse(self):
        if not self.toks:
            raise ParseError("empty formula")
        node = self.expr(0)
        if self.peek() is not None:
            raise ParseError(f"unexpected {self.peek().text!r}")
        return node

    def expr(self, min_prec):
        left = self.unary()
        while True:
            t = self.peek()
            if t is None or t.kind != "op" or t.text not in _BIN_PREC:
                break
            prec = _BIN_PREC[t.text]
            if prec < min_prec:
                break
            self.take()
            right = self.expr(prec + 1)
            left = ("bin", t.text, left, right)
        return left

    def unary(self):
        t = self.peek()
        if t is not None and t.kind == "op" and t.text in "+-":
            self.take()
            operand = self.unary()
            return ("neg", operand) if t.text == "-" else operand
        return self.postfix(self.primary())

    def postfix(self, node):
        while True:
            t = self.peek()
            if t is not None and t.kind == "op" and t.text == "%":
                self.take()
                node = ("pct", node)
            elif t is not None and t.text == ":" and node[0] in ("cell", "range", "func"):
                # range operator between references, e.g. A1:INDEX(...)
                self.take()
                other = self.primary()
                node = ("func", "_RANGEOP", [node, other])
            else:
                return node

    def primary(self):
        t = self.take()
        if t is None:
            raise ParseError("unexpected end of formula")
        k = t.kind
        if k == "num":
            return ("num", float(t.text))
        if k == "str":
            if len(t.text) < 2 or not t.text.endswith('"'):
                raise ParseError("unterminated string")
            return ("str", t.text[1:-1].replace('""', '"'))
        if k == "err":
            code = t.text.split("!", 1)[-1] if "!" in t.text[:-1] and not t.text.startswith("#") else t.text
            if not code.startswith("#"):
                code = "#REF!"
            return ("err", errors.from_code(code) or errors.REF)
        if k == "ref":
            info = RefInfo.parse(t.text)
            if info.is_single():
                col, _, row, _ = info.ends[0]
                if row >= MAX_ROWS or col >= MAX_COLS or row < 0:
                    return ("err", errors.REF)
                return ("cell", info.sheet, row, col)
            r1, c1, r2, c2 = info.bounds()
            if r2 >= MAX_ROWS or c2 >= MAX_COLS or r1 < 0:
                return ("err", errors.REF)
            return ("range", info.sheet, r1, c1, r2, c2)
        if k == "func":
            name = t.text.upper()
            for prefix in ("_XLFN.", "_XLWS."):
                if name.startswith(prefix):
                    name = name[len(prefix):]
            self.expect("(")
            args = []
            if self.peek() is not None and self.peek().text == ")":
                self.take()
                return ("func", name, args)
            while True:
                nt = self.peek()
                if nt is not None and nt.text in (",", ";", ")"):
                    args.append(("missing",))
                else:
                    args.append(self.expr(0))
                nt = self.take()
                if nt is None:
                    raise ParseError("missing )")
                if nt.text == ")":
                    break
                if nt.text not in (",", ";"):
                    raise ParseError(f"unexpected {nt.text!r}")
            return ("func", name, args)
        if k == "name":
            up = t.text.upper()
            if up == "TRUE":
                return ("bool", True)
            if up == "FALSE":
                return ("bool", False)
            return ("name", t.text)
        if t.text == "(":
            node = self.expr(0)
            self.expect(")")
            return node
        if t.text == "{":
            return self.array()
        raise ParseError(f"unexpected {t.text!r}")

    def array(self):
        rows, row = [], []
        while True:
            t = self.take()
            if t is None:
                raise ParseError("missing }")
            neg = False
            if t.kind == "op" and t.text in "+-":
                neg = t.text == "-"
                t = self.take()
            if t.kind == "num":
                v = float(t.text)
                row.append(-v if neg else v)
            elif t.kind == "str":
                row.append(t.text[1:-1].replace('""', '"'))
            elif t.kind == "name" and t.text.upper() in ("TRUE", "FALSE"):
                row.append(t.text.upper() == "TRUE")
            elif t.kind == "err":
                row.append(errors.from_code(t.text) or errors.VALUE)
            else:
                raise ParseError("bad array constant")
            sep = self.take()
            if sep is None:
                raise ParseError("missing }")
            if sep.text == "}":
                rows.append(row)
                break
            if sep.text == ",":
                continue
            if sep.text == ";":
                rows.append(row)
                row = []
                continue
            raise ParseError("bad array constant")
        if len({len(r) for r in rows}) != 1:
            raise ParseError("ragged array")
        return ("array", rows)


def parse(src):
    return _Parser(tokenize(src)).parse()


def collect_refs(node, out=None):
    """All references in an AST -> list of (sheet_or_None, r1, c1, r2, c2)."""
    if out is None:
        out = []
    kind = node[0]
    if kind == "cell":
        out.append((node[1], node[2], node[3], node[2], node[3]))
    elif kind == "range":
        out.append((node[1], node[2], node[3], node[4], node[5]))
    elif kind == "func":
        for a in node[2]:
            collect_refs(a, out)
    elif kind == "bin":
        collect_refs(node[2], out)
        collect_refs(node[3], out)
    elif kind in ("neg", "pct"):
        collect_refs(node[1], out)
    return out


def func_names(node, out=None):
    if out is None:
        out = set()
    kind = node[0]
    if kind == "func":
        out.add(node[1])
        for a in node[2]:
            func_names(a, out)
    elif kind == "bin":
        func_names(node[2], out)
        func_names(node[3], out)
    elif kind in ("neg", "pct"):
        func_names(node[1], out)
    return out


# ---------------------------------------------------------------- rewriting

def _rebuild(src, toks, replace):
    """Rebuild formula text, replacing ref tokens through `replace(info)->str`."""
    out = ["="]
    for t in toks:
        if t.kind == "ref":
            out.append(replace(RefInfo.parse(t.text)))
        else:
            out.append(t.text)
    return "".join(out)


def normalize(src):
    """Upper-case function names and references; keep everything else."""
    toks = tokenize(src)
    out = ["="]
    for t in toks:
        if t.kind == "func":
            out.append(t.text.upper())
        elif t.kind == "ref":
            info = RefInfo.parse(t.text)
            out.append(info.text())
        elif t.kind == "name" and t.text.upper() in ("TRUE", "FALSE"):
            out.append(t.text.upper())
        else:
            out.append(t.text)
    return "".join(out)


def shift_formula(src, dr, dc):
    """Adjust relative references for a formula copied by (dr, dc)."""
    if dr == 0 and dc == 0:
        return src
    toks = tokenize(src)

    def rep(info):
        for e in info.ends:
            col, cabs, row, rabs = e
            if col is not None and not cabs:
                e[0] = col + dc
            if row is not None and not rabs:
                e[2] = row + dr
        for col, _, row, _ in info.ends:
            if (col is not None and not 0 <= col < MAX_COLS) or (row is not None and not 0 <= row < MAX_ROWS):
                return (info.sheet_text + "!" if info.sheet_text else "") + "#REF!"
        return info.text()

    return _rebuild(src, toks, rep)


def adjust_structure(src, own_sheet, target_sheet, axis, at, count):
    """Rewrite references after inserting (count>0) or deleting (count<0)
    rows/cols on `target_sheet`.  axis is 'row' or 'col'."""
    toks = tokenize(src)
    tgt = target_sheet.upper()
    own = own_sheet.upper()
    changed = False

    def rep(info):
        nonlocal changed
        sheet = (info.sheet or own_sheet).upper()
        if sheet != tgt:
            return info.text()
        idx = 2 if axis == "row" else 0
        vals = [e[idx] for e in info.ends]
        if any(v is None for v in vals):
            return info.text()  # whole-row ref on col change (or vice versa)
        if count > 0:
            for e in info.ends:
                if e[idx] >= at:
                    e[idx] += count
                    changed = True
            return info.text()
        n = -count
        lo, hi = at, at + n - 1
        if len(info.ends) == 1:
            v = vals[0]
            if lo <= v <= hi:
                changed = True
                return (info.sheet_text + "!" if info.sheet_text else "") + "#REF!"
            if v > hi:
                info.ends[0][idx] = v - n
                changed = True
            return info.text()
        a, b = sorted(vals)
        if lo <= a and b <= hi:
            changed = True
            return (info.sheet_text + "!" if info.sheet_text else "") + "#REF!"
        na = a if a < lo else (lo if a <= hi else a - n)
        nb = b if b < lo else (lo - 1 if b <= hi else b - n)
        first = 0 if info.ends[0][idx] <= info.ends[1][idx] else 1
        info.ends[first][idx] = na
        info.ends[1 - first][idx] = nb
        if (na, nb) != (a, b):
            changed = True
        return info.text()

    out = _rebuild(src, toks, rep)
    return out if changed else src


def rename_sheet(src, old, new):
    toks = tokenize(src)
    old_u = old.upper()
    hit = False

    def rep(info):
        nonlocal hit
        if info.sheet is not None and info.sheet.upper() == old_u:
            info.sheet_text = quote_sheet(new)
            hit = True
        return info.text()

    out = _rebuild(src, toks, rep)
    return out if hit else src


def ref_spans(src):
    """(start, end, RefInfo) of each reference in formula text (for highlighting)."""
    out = []
    for t in tokenize(src):
        if t.kind == "ref":
            out.append((t.pos, t.pos + len(t.text), RefInfo.parse(t.text)))
    return out


def toggle_absolute(src, cursor):
    """F4: cycle $ on the reference at/before the cursor. Returns (text, cursor)."""
    for start, end, info in ref_spans(src):
        if start <= cursor <= end:
            for e in info.ends:
                col, cabs, row, rabs = e
                if col is not None and row is not None:
                    state = (cabs, rabs)
                    nxt = {(False, False): (True, True), (True, True): (False, True),
                           (False, True): (True, False), (True, False): (False, False)}[state]
                    e[1], e[3] = nxt
                elif col is not None:
                    e[1] = not cabs
                else:
                    e[3] = not rabs
            new = info.text()
            text = src[:start] + new + src[end:]
            return text, start + len(new)
    return src, cursor


NEW_FUNCTIONS = {
    "XLOOKUP", "XMATCH", "IFS", "SWITCH", "TEXTJOIN", "CONCAT", "MAXIFS", "MINIFS",
    "IFNA", "DAYS", "UNIQUE", "FILTER", "SORT", "SEQUENCE", "STDEV.S", "STDEV.P",
    "VAR.S", "VAR.P", "CEILING.MATH", "FLOOR.MATH", "ISOWEEKNUM", "XOR", "RANK.EQ",
    "PERCENTILE.INC", "QUARTILE.INC", "MODE.SNGL", "NUMBERVALUE", "TEXTBEFORE", "TEXTAFTER",
}


def to_file_formula(src):
    """Add the _xlfn. prefix Excel needs on newer functions when saving xlsx."""
    toks = tokenize(src)
    if not any(t.kind == "func" and t.text.upper() in NEW_FUNCTIONS for t in toks):
        return src
    out = ["="]
    for t in toks:
        if t.kind == "func" and t.text.upper() in NEW_FUNCTIONS:
            out.append("_xlfn." + t.text.upper())
        elif t.kind == "ref":
            out.append(t.text)
        else:
            out.append(t.text)
    return "".join(out)


def from_file_formula(src):
    if "_xl" not in src:
        return src
    return re.sub(r"_xl(?:fn|ws)\.", "", src, flags=re.IGNORECASE)
