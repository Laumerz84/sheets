"""Calculated columns for big-file sheets.

Filling a formula (or a value) down millions of rows of a BigSheet doesn't create millions of
cells: the formula is stored once per column and computed for every row at once with pyarrow,
like an Excel Table's calculated column. The results become an ordinary column of the big
sheet (big.cols / big.numeric), so sorting, filtering, totals, Find and Save just work.

Formulas may use cells of the same row (E2, $E2), fixed cells ($H$1, other sheets), constants,
+ - * / ^ & %, comparisons and the functions in FUNCS below. Anything else raises Unsupported.
When an input cell is edited later, BigData.set_edit recomputes that row's value cell by cell
(bigcalc.recompute_row)."""
from . import errors
from .errors import XLError
from .formula import parse, shift_formula

ERR_CODES = {1: errors.VALUE, 2: errors.DIV0, 3: errors.NUM, 4: errors.NA, 5: errors.REF}
_ERR_OF = {id(v): k for k, v in ERR_CODES.items()}
K_NUM, K_TXT, K_BOOL, K_BLANK, K_FILE = 0, 1, 2, 3, 4  # K_FILE: not computed here, use the file's cell


class Unsupported(Exception):
    pass


def _arrow():
    import pyarrow as pa
    import pyarrow.compute as pc
    return pa, pc


class V:
    """A value per data row: kind (K_*), num (float64), txt (string), err (int8 code, 0 = none).
    Each field is an Arrow array of the column's length or an Arrow scalar (same for all rows)."""
    __slots__ = ("kind", "num", "txt", "err")

    def __init__(self, kind, num=None, txt=None, err=None):
        pa, pc = _arrow()
        self.kind = kind if not isinstance(kind, int) else pa.scalar(kind, pa.int8())
        self.num = num if num is not None else pa.scalar(None, pa.float64())
        self.txt = txt if txt is not None else pa.scalar(None, pa.string())
        self.err = err if err is not None else pa.scalar(0, pa.int8())


def _const(v):
    pa, pc = _arrow()
    if isinstance(v, XLError):
        return V(K_NUM, err=pa.scalar(_ERR_OF.get(id(v), 1), pa.int8()))
    if v is None or v == "":
        return V(K_BLANK)
    if isinstance(v, bool):
        return V(K_BOOL, num=pa.scalar(1.0 if v else 0.0), txt=pa.scalar("TRUE" if v else "FALSE"))
    if isinstance(v, (int, float)):
        return V(K_NUM, num=pa.scalar(float(v)))
    if isinstance(v, list):
        raise Unsupported("Array constants")
    return V(K_TXT, txt=pa.scalar(str(v)))


def _column(big, c):
    """The current values of column c of the big sheet as a V over data rows."""
    pa, pc = _arrow()
    if c >= big.ncols:
        return V(K_BLANK)
    num = big.numeric(c)
    txt = big.cols[c]
    kind = pc.if_else(pc.is_valid(num), pa.scalar(K_NUM, pa.int8()),
                      pc.if_else(pc.equal(txt, ""), pa.scalar(K_BLANK, pa.int8()), pa.scalar(K_TXT, pa.int8())))
    # TRUE/FALSE in the file are logical values
    upper = pc.utf8_upper(txt)
    isb = pc.or_(pc.equal(upper, "TRUE"), pc.equal(upper, "FALSE"))
    kind = pc.if_else(isb, pa.scalar(K_BOOL, pa.int8()), kind)
    num = pc.if_else(isb, pc.if_else(pc.equal(upper, "TRUE"), pa.scalar(1.0), pa.scalar(0.0)), num)
    return V(kind, num, txt)


def _errs(*vs):
    pa, pc = _arrow()
    out = None
    for v in vs:
        e = v.err if isinstance(v, V) else v
        out = e if out is None else pc.if_else(pc.not_equal(out, 0), out, e)
    return out


def _eq(a, b):
    _, pc = _arrow()
    return pc.equal(a, b)


def to_num(v):
    """(float64, err) the way Excel's arithmetic coerces: blank 0, TRUE 1, text #VALUE!."""
    pa, pc = _arrow()
    k = v.kind
    num = pc.if_else(_eq(k, K_BLANK), pa.scalar(0.0), v.num)
    is_txt = _eq(k, K_TXT)
    err = pc.if_else(is_txt, pa.scalar(1, pa.int8()), v.err)
    err = pc.if_else(pc.not_equal(v.err, 0), v.err, err)
    return num, err


def to_str(v):
    """Text of each value (numbers like Excel's General, rounded to 10 decimals)."""
    pa, pc = _arrow()
    k = v.kind
    numtxt = pc.cast(pc.round(v.num, 10), pa.string())  # 404.0 -> "404", like General
    s = pc.if_else(_eq(k, K_NUM), numtxt,
                   pc.if_else(_eq(k, K_BLANK), pa.scalar(""), v.txt))
    return s


def to_bool(v):
    pa, pc = _arrow()
    k = v.kind
    b = pc.if_else(_eq(k, K_TXT), pc.equal(pc.utf8_upper(v.txt), "TRUE"), pc.not_equal(pc.fill_null(v.num, 0.0), 0.0))
    bad_txt = pc.and_(_eq(k, K_TXT), pc.invert(pc.or_(pc.equal(pc.utf8_upper(v.txt), "TRUE"),
                                                      pc.equal(pc.utf8_upper(v.txt), "FALSE"))))
    err = pc.if_else(pc.fill_null(bad_txt, False), pa.scalar(1, pa.int8()), v.err)
    err = pc.if_else(pc.not_equal(v.err, 0), v.err, err)
    return pc.fill_null(b, False), err


def _numv(num, err):
    pa, pc = _arrow()
    bad = pc.or_(pc.is_null(num), pc.is_nan(pc.fill_null(num, 0.0)))
    err = pc.if_else(pc.and_(pc.equal(err, 0), bad), pa.scalar(3, pa.int8()), err)
    return V(K_NUM, num, None, err)


def _boolv(b, err):
    pa, pc = _arrow()
    return V(K_BOOL, pc.if_else(b, pa.scalar(1.0), pa.scalar(0.0)),
             pc.if_else(b, pa.scalar("TRUE"), pa.scalar("FALSE")), err)


# ---------------------------------------------------------------- operators
def _arith(op, a, b):
    pa, pc = _arrow()
    x, ex = to_num(a)
    y, ey = to_num(b)
    err = _errs(ex, ey)
    if op == "+":
        r = pc.add(x, y)
    elif op == "-":
        r = pc.subtract(x, y)
    elif op == "*":
        r = pc.multiply(x, y)
    elif op == "/":
        zero = pc.equal(y, 0.0)
        err = pc.if_else(pc.and_(pc.equal(err, 0), pc.fill_null(zero, False)), pa.scalar(2, pa.int8()), err)
        r = pc.divide(x, pc.if_else(zero, pa.scalar(1.0), y))
    elif op == "^":
        r = pc.power(x, y)
    else:
        raise Unsupported(op)
    return _numv(r, err)


def _compare(op, a, b):
    """Excel order: numbers < text < logicals; text compares case-insensitively; blank acts as 0 or ""."""
    pa, pc = _arrow()

    def side(v, other):
        k = v.kind
        # a blank takes the other side's kind (0 next to numbers, "" next to text, FALSE next to logicals)
        k2 = pc.if_else(_eq(k, K_BLANK), pc.if_else(_eq(other.kind, K_BLANK), pa.scalar(K_NUM, pa.int8()), other.kind), k)
        num = pc.if_else(_eq(k, K_BLANK), pa.scalar(0.0), v.num)
        txt = pc.utf8_lower(pc.if_else(_eq(k, K_BLANK), pa.scalar(""), v.txt))
        rank = pc.if_else(_eq(k2, K_TXT), pa.scalar(1, pa.int8()),
                          pc.if_else(_eq(k2, K_BOOL), pa.scalar(2, pa.int8()), pa.scalar(0, pa.int8())))
        return rank, num, txt
    ra, na, ta = side(a, b)
    rb, nb, tb = side(b, a)
    fns = {"=": pc.equal, "<>": pc.not_equal, "<": pc.less, ">": pc.greater, "<=": pc.less_equal,
           ">=": pc.greater_equal}
    f = fns[op]
    same_rank = pc.equal(ra, rb)
    by_num = f(pc.fill_null(na, 0.0), pc.fill_null(nb, 0.0))
    by_txt = f(pc.fill_null(ta, ""), pc.fill_null(tb, ""))
    within = pc.if_else(pc.equal(ra, 1), by_txt, by_num)
    res = pc.if_else(same_rank, within, f(ra, rb))
    return _boolv(pc.fill_null(res, False), _errs(a, b))


def _concat(a, b):
    pa, pc = _arrow()
    return V(K_TXT, None, pc.binary_join_element_wise(to_str(a), to_str(b), ""), _errs(a, b))


# ---------------------------------------------------------------- functions
def _fn_if(args):
    pa, pc = _arrow()
    if len(args) < 2 or len(args) > 3:
        raise Unsupported("IF with that many arguments")
    cond, ec = to_bool(args[0])
    a = args[1]
    b = args[2] if len(args) == 3 else _boolv(pa.scalar(False), pa.scalar(0, pa.int8()))
    pick = lambda x, y: pc.if_else(cond, x, y)
    v = V(pick(a.kind, b.kind), pick(a.num, b.num), pick(a.txt, b.txt), pick(a.err, b.err))
    v.err = pc.if_else(pc.not_equal(ec, 0), ec, v.err)
    return v


def _fn_iferror(args):
    pa, pc = _arrow()
    x, alt = args[0], args[1]
    bad = pc.not_equal(x.err, 0)
    pick = lambda p, q: pc.if_else(bad, q, p)
    return V(pick(x.kind, alt.kind), pick(x.num, alt.num), pick(x.txt, alt.txt), pick(x.err, alt.err))


def _num1(fn):
    def run(args):
        x, e = to_num(args[0])
        return _numv(fn(x, args[1:]), e)
    return run


def _round(mode):
    def run(args):
        pa, pc = _arrow()
        x, e = to_num(args[0])
        nd = 0
        if len(args) > 1:
            n = args[1]
            if not isinstance(n.num, pa.Scalar) or not n.num.is_valid:
                raise Unsupported("ROUND with a per-row number of digits")
            nd = int(n.num.as_py())
        # same steps as functions._round_half_away: strip binary noise (2.675 -> 267.5) first
        m = 10.0 ** nd
        y = pc.round(pc.multiply(pc.abs(x), m), 9, round_mode="half_to_even")
        if mode == "round":
            y = pc.floor(pc.add(y, 0.5))
        elif mode == "up":
            y = pc.ceil(y)
        else:
            y = pc.floor(y)
        r = pc.multiply(pc.divide(y, m), pc.sign(x))
        return _numv(r, e)
    return run


def _agg(kind):
    def run(args):
        pa, pc = _arrow()
        nums, errs = [], []
        count = None
        for a in args:
            k = a.kind
            n = pc.if_else(_eq(k, K_NUM), a.num, pa.scalar(None, pa.float64()))
            nums.append(n)
            errs.append(a.err)
            c = pc.cast(_eq(k, K_NUM), pa.float64())
            count = c if count is None else pc.add(count, c)
        err = _errs(*errs)
        if kind == "SUM":
            r = None
            for n in nums:
                n0 = pc.fill_null(n, 0.0)
                r = n0 if r is None else pc.add(r, n0)
        elif kind == "AVERAGE":
            s = None
            for n in nums:
                n0 = pc.fill_null(n, 0.0)
                s = n0 if s is None else pc.add(s, n0)
            zero = pc.equal(count, 0.0)
            err = pc.if_else(pc.and_(pc.equal(err, 0), zero), pa.scalar(2, pa.int8()), err)
            r = pc.divide(s, pc.if_else(zero, pa.scalar(1.0), count))
        elif kind == "MIN":
            r = pc.fill_null(pc.min_element_wise(*nums, skip_nulls=True), 0.0)
        elif kind == "MAX":
            r = pc.fill_null(pc.max_element_wise(*nums, skip_nulls=True), 0.0)
        elif kind == "COUNT":
            r = count
        else:
            raise Unsupported(kind)
        return _numv(r, err)
    return run


def _txt1(fn):
    def run(args):
        return V(K_TXT, None, fn(to_str(args[0]), args[1:]), args[0].err)
    return run


def _int_arg(v, default):
    pa, pc = _arrow()
    if v is None:
        return default
    if not isinstance(v.num, pa.Scalar) or not v.num.is_valid:
        raise Unsupported("a per-row text length")
    return int(v.num.as_py())


def _left(s, rest):
    _, pc = _arrow()
    n = _int_arg(rest[0] if rest else None, 1)
    return pc.utf8_slice_codeunits(s, 0, max(0, n))


def _right(s, rest):
    _, pc = _arrow()
    n = _int_arg(rest[0] if rest else None, 1)
    return pc.utf8_slice_codeunits(s, -n) if n > 0 else pc.utf8_slice_codeunits(s, 0, 0)


def _mid(s, rest):
    _, pc = _arrow()
    a = _int_arg(rest[0], 1)
    n = _int_arg(rest[1], 0)
    return pc.utf8_slice_codeunits(s, max(0, a - 1), max(0, a - 1) + max(0, n))


def _fn_len(args):
    pa, pc = _arrow()
    return _numv(pc.cast(pc.utf8_length(to_str(args[0])), pa.float64()), args[0].err)


def _fn_concat(args):
    v = args[0]
    for a in args[1:]:
        v = _concat(v, a)
    return v


def _is(test):
    def run(args):
        pa, pc = _arrow()
        v = args[0]
        if test == "ISBLANK":
            b = _eq(v.kind, K_BLANK)
        elif test == "ISNUMBER":
            b = pc.and_(_eq(v.kind, K_NUM), pc.equal(v.err, 0))
        elif test == "ISTEXT":
            b = pc.and_(_eq(v.kind, K_TXT), pc.equal(v.err, 0))
        else:  # ISERROR
            b = pc.not_equal(v.err, 0)
        return _boolv(b, pa.scalar(0, pa.int8()))
    return run


def _logic(kind):
    def run(args):
        pa, pc = _arrow()
        out, errs = None, []
        for a in args:
            b, e = to_bool(a)
            errs.append(e)
            out = b if out is None else (pc.and_(out, b) if kind == "AND" else pc.or_(out, b))
        return _boolv(out, _errs(*errs))
    return run


def _fn_not(args):
    _, pc = _arrow()
    b, e = to_bool(args[0])
    return _boolv(pc.invert(b), e)


def _datepart(part):
    def run(args):
        pa, pc = _arrow()
        x, e = to_num(args[0])
        secs = pc.cast(pc.multiply(pc.subtract(pc.floor(x), 25569.0), 86400.0), pa.int64())
        ts = pc.cast(secs, pa.timestamp("s"))
        r = {"YEAR": pc.year, "MONTH": pc.month, "DAY": pc.day}[part](ts)
        return _numv(pc.cast(r, pa.float64()), e)
    return run


def _fn_value(args):
    pa, pc = _arrow()
    v = args[0]
    s = pc.utf8_trim_whitespace(to_str(v))
    ok = pc.match_substring_regex(s, r"^-?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$")
    num = pc.cast(pc.if_else(ok, s, pa.scalar(None, pa.string())), pa.float64())
    err = pc.if_else(pc.fill_null(ok, False), v.err, pa.scalar(1, pa.int8()))
    return _numv(num, err)


def _mod(args):
    pa, pc = _arrow()
    x, ex = to_num(args[0])
    y, ey = to_num(args[1])
    err = _errs(ex, ey)
    zero = pc.equal(y, 0.0)
    err = pc.if_else(pc.and_(pc.equal(err, 0), pc.fill_null(zero, False)), pa.scalar(2, pa.int8()), err)
    ys = pc.if_else(zero, pa.scalar(1.0), y)
    r = pc.subtract(x, pc.multiply(ys, pc.floor(pc.divide(x, ys))))
    return _numv(r, err)


FUNCS = {
    "IF": _fn_if, "IFERROR": _fn_iferror,
    "ROUND": _round("round"), "ROUNDUP": _round("up"), "ROUNDDOWN": _round("down"),
    "INT": _num1(lambda x, r: _arrow()[1].floor(x)), "ABS": _num1(lambda x, r: _arrow()[1].abs(x)),
    "SQRT": _num1(lambda x, r: _arrow()[1].sqrt(x)), "MOD": _mod,
    "SUM": _agg("SUM"), "AVERAGE": _agg("AVERAGE"), "MIN": _agg("MIN"), "MAX": _agg("MAX"), "COUNT": _agg("COUNT"),
    "UPPER": _txt1(lambda s, r: _arrow()[1].utf8_upper(s)), "LOWER": _txt1(lambda s, r: _arrow()[1].utf8_lower(s)),
    "TRIM": _txt1(lambda s, r: _arrow()[1].utf8_trim_whitespace(s)),
    "LEFT": _txt1(_left), "RIGHT": _txt1(_right), "MID": _txt1(_mid), "LEN": _fn_len,
    "CONCAT": _fn_concat, "CONCATENATE": _fn_concat,
    "ISBLANK": _is("ISBLANK"), "ISNUMBER": _is("ISNUMBER"), "ISTEXT": _is("ISTEXT"), "ISERROR": _is("ISERROR"),
    "AND": _logic("AND"), "OR": _logic("OR"), "NOT": _fn_not,
    "YEAR": _datepart("YEAR"), "MONTH": _datepart("MONTH"), "DAY": _datepart("DAY"), "VALUE": _fn_value,
}


# ---------------------------------------------------------------- compiling
def _refs_info(text, anchor_r):
    """Parse text and work out, for every cell/range node, whether its row is relative:
    re-parse the formula copied one row down and see which rows moved."""
    a = parse(text)
    b = parse(shift_formula(text, 1, 0))
    rel = {}

    def walk(x, y):
        if not isinstance(x, tuple) or not isinstance(y, tuple):
            return
        k = x[0]
        if k == "cell":
            rel[id(x)] = (y[2] - x[2] == 1)
        elif k == "range":
            rel[id(x)] = (y[2] - x[2] == 1, y[4] - x[4] == 1)
        elif k == "func":
            for p, q in zip(x[2], y[2]):
                walk(p, q)
        elif k == "bin":
            walk(x[2], y[2])
            walk(x[3], y[3])
        elif k in ("neg", "pct"):
            walk(x[1], y[1])
    walk(a, b)
    return a, rel


def _row_relative(node, rel):
    k = node[0]
    if k in ("cell", "range"):
        r = rel.get(id(node))
        return any(r) if isinstance(r, tuple) else bool(r)
    if k == "func":
        return any(_row_relative(n, rel) for n in node[2])
    if k == "bin":
        return _row_relative(node[2], rel) or _row_relative(node[3], rel)
    if k in ("neg", "pct"):
        return _row_relative(node[1], rel)
    return False


def compile_column(sheet, text, anchor_r, target_c):
    """Return (evaluate, inputs): evaluate(big) -> V over all data rows of the formula `text`
    as written at sheet row anchor_r; inputs = columns read from the same row."""
    if not text.startswith("="):
        v = _const(sheet_value_of(text, sheet, anchor_r, target_c))
        return (lambda big: v), set()
    try:
        ast, rel = _refs_info(text, anchor_r)
    except Exception as e:  # noqa: BLE001
        raise Unsupported(f"that formula ({e})")
    inputs = set()

    def scalar_of(node):
        """A sub-formula with no same-row references: computed once by the normal engine."""
        wb = sheet.wb
        v = wb.evaluator.run(node, (sheet, anchor_r, target_c))
        return _const(v)

    def build(node):
        k = node[0]
        if not _row_relative(node, rel):
            v = scalar_of(node)
            return lambda big: v
        if k == "cell":
            if node[1] is not None and node[1].upper() != sheet.name.upper():
                raise Unsupported("same-row references to another sheet")
            if node[2] != anchor_r:
                raise Unsupported("references to other rows (like the row above)")
            c = node[3]
            if c == target_c:
                raise Unsupported("a column that refers to itself")
            inputs.add(c)
            return lambda big: _column(big, c)
        if k == "range":
            rr = rel.get(id(node))
            if node[1] is not None or node[2] != anchor_r or node[4] != anchor_r or not all(rr):
                raise Unsupported("ranges over several rows")
            cs = list(range(node[3], node[5] + 1))
            if target_c in cs:
                raise Unsupported("a column that refers to itself")
            inputs.update(cs)
            return ("cols", cs)
        if k == "bin":
            fa, fb = build(node[2]), build(node[3])
            op = node[1]
            if isinstance(fa, tuple) or isinstance(fb, tuple):
                raise Unsupported("a range used as a single value")
            if op in ("+", "-", "*", "/", "^"):
                return lambda big: _arith(op, fa(big), fb(big))
            if op == "&":
                return lambda big: _concat(fa(big), fb(big))
            return lambda big: _compare(op, fa(big), fb(big))
        if k == "neg":
            f = build(node[1])
            return lambda big: _arith("-", _const(0.0), f(big))
        if k == "pct":
            f = build(node[1])
            return lambda big: _arith("/", f(big), _const(100.0))
        if k == "func":
            fn = FUNCS.get(node[1])
            if fn is None:
                raise Unsupported(f"the function {node[1]}")
            parts = [build(n) for n in node[2]]

            def run(big, parts=parts, fn=fn):
                args = []
                for p in parts:
                    if isinstance(p, tuple):
                        args.extend(_column(big, c) for c in p[1])
                    else:
                        args.append(p(big))
                return fn(args)
            return run
        raise Unsupported("that formula")
    f = build(ast)
    if isinstance(f, tuple):
        raise Unsupported("a range as the whole formula")
    return f, inputs


def sheet_value_of(text, sheet, r, c):
    from .ops import input_state
    st = input_state(sheet, r, c, text)
    content = st[0]
    return None if content is None else content[1]


# ---------------------------------------------------------------- applying
def calc_state(sheet, src_r, c, dst_r1, dst_r2):
    """New big_calc state with column c computed for sheet rows dst_r1..dst_r2 from the cell at
    (src_r, c): its formula copied down (same-row references follow), or its value."""
    pa, pc = _arrow()
    from .bigdata import arange
    big = sheet.big
    ftext = sheet.formula_text(src_r, c)
    if ftext is None:
        v = sheet.value(src_r, c)
        text = None
        val = _const(v)
        f, inputs = (lambda b: val), set()
    else:
        text = ftext
        f, inputs = compile_column(sheet, ftext, src_r, c)
    nv = big.nview
    lo, hi = max(0, dst_r1), min(dst_r2, nv - 1)
    if hi < lo:
        return None
    # data rows covered by the fill
    idx = arange(big.N)
    if big.view is None:
        cover = pc.and_(pc.greater_equal(idx, lo), pc.less_equal(idx, hi))
    elif len(big.view) == big.N and lo + (big.N - 1 - hi) <= 5_000_000:
        # sorted, not filtered: everything except the few rows above/below the fill
        excl = pa.concat_arrays([big.view.slice(0, lo), big.view.slice(hi + 1)])
        cover = pc.invert(pc.is_in(idx, value_set=excl))
    else:
        cover = pc.is_in(idx, value_set=big.view.slice(lo, hi - lo + 1))
    if isinstance(cover, pa.ChunkedArray):
        cover = cover.combine_chunks()
    v = f(big)

    def full(x, typ):
        if isinstance(x, pa.Scalar):
            return pa.repeat(x if x.type == typ else pa.scalar(x.as_py(), typ), big.N)
        return x.combine_chunks() if isinstance(x, pa.ChunkedArray) else x
    kind = full(v.kind, pa.int8())
    err = full(v.err, pa.int8())
    num = full(v.num, pa.float64())
    txt_all = to_str(v)
    txt_all = full(txt_all, pa.string()) if isinstance(txt_all, pa.Scalar) else (
        txt_all.combine_chunks() if isinstance(txt_all, pa.ChunkedArray) else txt_all)
    # rows whose inputs the user edited: compute those cells the normal way
    if text is not None and inputs:
        from .bigdata import _position_mask
        from .workbook import Formula
        fix = sorted({dk >> 14 for dk in big.edits if (dk & 0x3FFF) in inputs and (dk >> 14) < big.N})
        fix = [d for d in fix if cover[d].as_py()]
        if fix:
            ks, ns, ts, es = [], [], [], []
            for d in fix:
                r = big.srow(d)
                val = sheet.wb.evaluator.run(Formula(shift_formula(text, r - src_r, 0)).ast, (sheet, r, c)) \
                    if r is not None else None
                one = _const(val)
                ks.append(one.kind.as_py())
                ns.append(one.num.as_py())
                es.append(one.err.as_py())
                ts.append(to_str(one).as_py() if one.kind.as_py() != K_BLANK else "")
            mask = _position_mask(big.N, fix)
            kind = pc.replace_with_mask(kind, mask, pa.array(ks, pa.int8()))
            num = pc.replace_with_mask(num, mask, pa.array(ns, pa.float64()))
            err = pc.replace_with_mask(err, mask, pa.array(es, pa.int8()))
            txt_all = pc.replace_with_mask(txt_all, mask, pa.array(ts, pa.string()))
    bad = pc.not_equal(err, 0)
    if pc.any(bad).as_py():
        errtxt = pc.if_else(pc.equal(err, 1), pa.scalar("#VALUE!"),
                            pc.if_else(pc.equal(err, 2), pa.scalar("#DIV/0!"),
                                       pc.if_else(pc.equal(err, 3), pa.scalar("#NUM!"),
                                                  pc.if_else(pc.equal(err, 4), pa.scalar("#N/A"), pa.scalar("#REF!")))))
        txt_all = pc.if_else(bad, errtxt, txt_all)
    # outside the filled rows the column keeps what it had
    if c < big.ncols:
        old_txt = big.cols[c].combine_chunks() if isinstance(big.cols[c], pa.ChunkedArray) else big.cols[c]
        old_num = big.numeric(c)
        old_num = old_num.combine_chunks() if isinstance(old_num, pa.ChunkedArray) else old_num
    else:
        old_txt = pa.repeat("", big.N)
        old_num = pa.nulls(big.N, pa.float64())
    new_txt = pc.if_else(cover, txt_all, old_txt)
    numeric = pc.if_else(pc.and_(cover, pc.and_(pc.equal(kind, K_NUM), pc.invert(bad))), num,
                         pc.if_else(cover, pa.scalar(None, pa.float64()), old_num))
    kind_store = pc.if_else(cover, pc.if_else(bad, pa.scalar(5, pa.int8()), kind), pa.scalar(K_FILE, pa.int8()))
    st = state(sheet)
    cols = list(st["cols"])
    while len(cols) < c:
        cols.append(pa.chunked_array([pa.repeat("", big.N)]))
    if c < len(cols):
        cols[c] = pa.chunked_array([new_txt])
    else:
        cols.append(pa.chunked_array([new_txt]))
    nums = dict(st["num"])
    nums[c] = pa.chunked_array([numeric])
    vals = dict(st["vals"])
    vals[c] = (kind_store, num, err)
    defs = dict(st["defs"])
    defs[c] = {"text": text, "anchor": src_r, "inputs": sorted(inputs)}
    return {"cols": cols, "num": nums, "vals": vals, "defs": defs}


def state(sheet):
    big = sheet.big
    return {"cols": list(big.cols), "num": dict(big._num), "vals": dict(getattr(big, "calc_vals", {})),
            "defs": dict(getattr(big, "calc_defs", {}))}


def set_state(sheet, st):
    big = sheet.big
    big.cols = list(st["cols"])
    big.ncols = len(big.cols)
    big._num = dict(st["num"])
    big.calc_vals = dict(st["vals"])
    big.calc_defs = dict(st["defs"])
    big._rows.clear()
    sheet.recompute_extent()
    sheet.wb.recalc(full=True)


def cell(big, d, c):
    """Value of a computed cell (data row d, column c), or _NOT_CALC when the file's text applies."""
    vals = getattr(big, "calc_vals", None)
    if not vals or c not in vals or d >= big.N:
        return NOT_CALC
    kind, num, err = vals[c]
    k = kind[d].as_py()
    if k == K_FILE:
        return NOT_CALC
    if k == 5:
        return ERR_CODES.get(err[d].as_py(), errors.VALUE)
    if k == K_NUM:
        return num[d].as_py()
    if k == K_BLANK:
        return None
    if k == K_BOOL:
        return bool(num[d].as_py())
    return big.cols[c][d].as_py()


NOT_CALC = object()


def recompute_row(sheet, d, changed_c):
    """An input of a calculated column was edited at data row d: recompute that row's cells of the
    columns that read changed_c, cell by cell, as edits flagged 'auto'."""
    big = sheet.big
    defs = getattr(big, "calc_defs", None)
    if not defs:
        return
    for c, info in defs.items():
        if changed_c not in info["inputs"] or info["text"] is None:
            continue
        if cell(big, d, c) is NOT_CALC:
            continue
        dk = (d << 14) | c
        if dk in big.edits and dk not in big.calc_auto:
            continue  # the user typed over this cell
        r = big.srow(d)
        if r is None:
            continue
        from .workbook import Formula
        f = Formula(shift_formula(info["text"], r - info["anchor"], 0))
        v = sheet.wb.evaluator.run(f.ast, (sheet, r, c))
        big.edits[dk] = v
        big.calc_auto.add(dk)
        sheet.wb._changed.append((sheet, (r << 14) | c))
