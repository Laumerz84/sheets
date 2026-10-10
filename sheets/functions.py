"""Spreadsheet function library.

Each function is registered with @fn(NAME, kind=...):
  eager - receives evaluated args (cell refs arrive as single-cell RangeRef)
  safe  - like eager, but errors in args arrive as XLError values (ISERROR etc.)
  lazy  - receives (ev, ctx, nodes) and evaluates arguments itself
Signature for eager/safe: f(ctx, *args); ctx = (sheet, row, col).
"""
import datetime as dt
import math
import random
import re
import statistics
from functools import lru_cache

from . import errors
from .errors import XLError
from .numfmt import (datetime_to_serial, format_value, parse_input,
                     serial_to_datetime)
from .refs import MAX_COLS, MAX_ROWS, parse_range
from .values import (MISSING, RangeRef, compare, flat, is_err, is_num, scalar,
                     to_2d, to_bool, to_int, to_num, to_str)

FUNCS = {}
VOLATILE = {"RAND", "RANDBETWEEN", "NOW", "TODAY", "OFFSET", "INDIRECT"}

# Short signatures shown as hints while typing a formula.
SIGNATURES = {}


def fn(name, kind="eager", sig=None, minargs=0, maxargs=None):
    def deco(f):
        for n in name.split():
            FUNCS[n] = (f, kind, minargs, maxargs)
            if sig:
                SIGNATURES[n] = f"{n}({sig})"
        return f
    return deco


def _given(a):
    return a is not MISSING and a is not None


# ================================================================ helpers

def nums(args, literal_coerce=True):
    """Numbers for SUM-like functions: ranges ignore text/bools, literals coerce."""
    for a in args:
        if isinstance(a, RangeRef):
            if a.is_cell():
                v = a.sheet.value(a.r1, a.c1)
                if is_num(v):
                    yield v
                elif isinstance(v, XLError):
                    raise v
                continue
            for v in a.nonblank():
                if is_num(v):
                    yield v
                elif isinstance(v, XLError):
                    raise v
        elif isinstance(a, list):
            for row in a:
                for v in row:
                    if is_num(v):
                        yield v
                    elif isinstance(v, XLError):
                        raise v
        elif a is MISSING or a is None:
            continue
        else:
            if isinstance(a, XLError):
                raise a
            if literal_coerce:
                yield to_num(a)


def nums_a(args):
    """Like nums() but text counts as 0 and booleans as 1/0 inside ranges (xxxA functions)."""
    for a in args:
        if isinstance(a, (RangeRef, list)):
            for v in flat(a):
                if v is None:
                    continue
                if isinstance(v, XLError):
                    raise v
                if is_num(v):
                    yield v
                elif isinstance(v, bool):
                    yield 1.0 if v else 0.0
                else:
                    yield 0.0
        elif a is MISSING:
            continue
        else:
            yield to_num(a)


def _num_list(args):
    return list(nums(args))


def _wild_re(pattern):
    return _wild_compile(pattern)


@lru_cache(maxsize=1024)
def _wild_compile(pattern):
    out = []
    i = 0
    while i < len(pattern):
        ch = pattern[i]
        if ch == "~" and i + 1 < len(pattern) and pattern[i + 1] in "*?~":
            out.append(re.escape(pattern[i + 1]))
            i += 2
            continue
        if ch == "*":
            out.append(".*")
        elif ch == "?":
            out.append(".")
        else:
            out.append(re.escape(ch))
        i += 1
    return re.compile("".join(out), re.IGNORECASE | re.DOTALL)


def _wild_segments(pattern):
    """Split an Excel wildcard pattern on '*' into segments; '?' becomes None (any char).
    '~*', '~?' and '~~' are literal."""
    segs, cur, i = [], [], 0
    while i < len(pattern):
        ch = pattern[i]
        if ch == "~" and i + 1 < len(pattern) and pattern[i + 1] in "*?~":
            cur.append(pattern[i + 1])
            i += 2
            continue
        if ch == "*":
            segs.append(cur)
            cur = []
        else:
            cur.append(None if ch == "?" else ch)
        i += 1
    segs.append(cur)
    return segs


def wild_match(pattern, text, whole=True, match_case=False):
    """Excel wildcard match without regex backtracking (worst case ~ len(text) * len(pattern))."""
    if not match_case:
        pattern, text = pattern.lower(), text.lower()
    segs = _wild_segments(pattern)
    n = len(text)

    def at(seg, i):
        if i < 0 or i + len(seg) > n:
            return False
        for j, ch in enumerate(seg):
            if ch is not None and text[i + j] != ch:
                return False
        return True

    def find(seg, start):
        for i in range(start, n - len(seg) + 1):
            if at(seg, i):
                return i
        return -1

    if len(segs) == 1:
        return (len(segs[0]) == n and at(segs[0], 0)) if whole else find(segs[0], 0) >= 0
    first, middle, last = segs[0], segs[1:-1], segs[-1]
    if whole:
        if not at(first, 0):
            return False
        pos = len(first)
    else:
        i = find(first, 0)
        if i < 0:
            return False
        pos = i + len(first)
    for seg in middle:
        i = find(seg, pos)
        if i < 0:
            return False
        pos = i + len(seg)
    if whole:
        return n - len(last) >= pos and at(last, n - len(last))
    return find(last, pos) >= 0


def _has_wild(s):
    return "*" in s or "?" in s or "~" in s


def make_criteria(crit):
    crit = scalar(crit)
    if isinstance(crit, XLError):
        raise crit
    if crit is None or crit is MISSING:
        return lambda v: v is None or v == ""
    if isinstance(crit, bool):
        return lambda v: isinstance(v, bool) and v == crit
    if is_num(crit):
        target = crit
        return lambda v: (is_num(v) and v == target) or (
            isinstance(v, str) and _text_num_eq(v, target))
    s = str(crit)
    op = "="
    for p in (">=", "<=", "<>", ">", "<", "="):
        if s.startswith(p):
            op, s = p, s[len(p):]
            break
    num, _ = parse_input(s) if s.strip() else (None, None)
    numeric = is_num(num)
    upper = s.upper()
    if op in ("=", "<>"):
        if s == "":
            pred = lambda v: v is None or v == ""
        elif numeric:
            pred = lambda v: (is_num(v) and v == num) or (isinstance(v, str) and v.strip().lower() == s.strip().lower())
        elif upper in ("TRUE", "FALSE"):
            b = upper == "TRUE"
            pred = lambda v: isinstance(v, bool) and v == b
        elif _has_wild(s):
            pred = lambda v: isinstance(v, str) and wild_match(s, v)
        else:
            low = s.lower()
            pred = lambda v: isinstance(v, str) and v.lower() == low
        if op == "<>":
            return lambda v: not pred(v)
        return pred
    if numeric:
        cmp = {">": lambda a: a > num, "<": lambda a: a < num,
               ">=": lambda a: a >= num, "<=": lambda a: a <= num}[op]
        return lambda v: is_num(v) and cmp(v)
    low = s.lower()
    cmp = {">": lambda a: a > low, "<": lambda a: a < low,
           ">=": lambda a: a >= low, "<=": lambda a: a <= low}[op]
    return lambda v: isinstance(v, str) and cmp(v.lower())


def _text_num_eq(s, target):
    x, _ = parse_input(s)
    return is_num(x) and x == target


def _ifs_mask(pairs, ctx):
    """pairs: [(range, criteria)...] -> (shape, list of bool rows)."""
    grids = []
    shape = None
    for rng, crit in pairs:
        g = to_2d(rng)
        sh = (len(g), len(g[0]) if g else 0)
        if shape is None:
            shape = sh
        elif sh != shape:
            raise errors.VALUE
        grids.append((g, make_criteria(crit)))
    mask = [[True] * shape[1] for _ in range(shape[0])]
    for g, pred in grids:
        for i, row in enumerate(g):
            mrow = mask[i]
            for j, v in enumerate(row):
                if mrow[j] and not pred(v):
                    mrow[j] = False
    return shape, mask


def _aligned(rng, shape):
    """Values of rng resized to shape (anchored at its top-left)."""
    if isinstance(rng, RangeRef):
        sh = rng.sheet
        return [[sh.value(rng.r1 + i, rng.c1 + j) for j in range(shape[1])] for i in range(shape[0])]
    g = to_2d(rng)
    return [[g[i][j] if i < len(g) and j < len(g[i]) else None for j in range(shape[1])]
            for i in range(shape[0])]


def _date(v):
    x = to_num(v)
    if x < 0:
        raise errors.NUM
    try:
        return serial_to_datetime(x)
    except (OverflowError, ValueError):
        raise errors.NUM


def _serial(d):
    return datetime_to_serial(d)


# ================================================================ math

def _big_aggregate(name, args):
    """Whole-column SUM/COUNT/... over a big-file sheet, computed column-wise (see bigdata.py)."""
    if not any(isinstance(a, RangeRef) and getattr(a.sheet, "big", None) is not None for a in args):
        return None
    from . import bigdata
    r = bigdata.aggregate(name, args)
    return None if r is bigdata._MISSING else r


@fn("SUM", sig="number1, [number2], ...")
def _sum(ctx, *args):
    fast = _big_aggregate("SUM", args)
    if fast is not None:
        return fast
    return math.fsum(nums(args))


@fn("PRODUCT", sig="number1, [number2], ...")
def _product(ctx, *args):
    p = 1.0
    seen = False
    for v in nums(args):
        p *= v
        seen = True
    return p if seen else 0.0


@fn("AVERAGE", sig="number1, [number2], ...")
def _average(ctx, *args):
    fast = _big_aggregate("AVERAGE", args)
    if fast is not None:
        return fast
    xs = _num_list(args)
    if not xs:
        raise errors.DIV0
    return math.fsum(xs) / len(xs)


@fn("AVERAGEA", sig="value1, [value2], ...")
def _averagea(ctx, *args):
    xs = list(nums_a(args))
    if not xs:
        raise errors.DIV0
    return math.fsum(xs) / len(xs)


@fn("MIN", sig="number1, [number2], ...")
def _min(ctx, *args):
    fast = _big_aggregate("MIN", args)
    if fast is not None:
        return fast
    xs = _num_list(args)
    return min(xs) if xs else 0.0


@fn("MAX", sig="number1, [number2], ...")
def _max(ctx, *args):
    fast = _big_aggregate("MAX", args)
    if fast is not None:
        return fast
    xs = _num_list(args)
    return max(xs) if xs else 0.0


@fn("MINA", sig="value1, [value2], ...")
def _mina(ctx, *args):
    xs = list(nums_a(args))
    return min(xs) if xs else 0.0


@fn("MAXA", sig="value1, [value2], ...")
def _maxa(ctx, *args):
    xs = list(nums_a(args))
    return max(xs) if xs else 0.0


@fn("COUNT", kind="safe", sig="value1, [value2], ...")
def _count(ctx, *args):
    fast = _big_aggregate("COUNT", args)
    if fast is not None:
        return fast
    n = 0
    for a in args:
        if isinstance(a, RangeRef) or isinstance(a, list):
            src = a.nonblank() if isinstance(a, RangeRef) else flat(a)
            n += sum(1 for v in src if is_num(v))
        elif a is MISSING:
            continue
        elif is_num(a) or isinstance(a, bool):
            n += 1
        elif isinstance(a, str):
            x, _ = parse_input(a)
            n += 1 if is_num(x) else 0
    return float(n)


@fn("COUNTA", kind="safe", sig="value1, [value2], ...")
def _counta(ctx, *args):
    fast = _big_aggregate("COUNTA", args)
    if fast is not None:
        return fast
    n = 0
    for a in args:
        if isinstance(a, RangeRef):
            n += sum(1 for _ in a.nonblank())
        elif isinstance(a, list):
            n += sum(1 for v in flat(a) if v is not None)
        elif a is not MISSING:
            n += 1
    return float(n)


@fn("COUNTBLANK", sig="range")
def _countblank(ctx, rng):
    if isinstance(rng, RangeRef):
        total = rng.nrows * rng.ncols
        filled = sum(1 for v in rng.nonblank() if v != "")
        return float(total - filled)
    return float(sum(1 for v in flat(rng) if v is None or v == ""))


@fn("ABS", sig="number")
def _abs(ctx, x):
    return abs(to_num(x))


def _round_half_away(x, digits):
    m = 10.0 ** digits
    y = round(abs(x) * m, 9)  # strip binary noise (2.675 -> 267.5)
    return math.copysign(math.floor(y + 0.5) / m, x)

@fn("ROUND", sig="number, num_digits")
def _round(ctx, x, d=0.0):
    return _round_half_away(to_num(x), to_int(d) if _given(d) else 0)


@fn("ROUNDUP", sig="number, num_digits")
def _roundup(ctx, x, d=0.0):
    x = to_num(x)
    m = 10 ** (to_int(d) if _given(d) else 0)
    y = round(abs(x) * m, 9)
    return math.copysign(math.ceil(y) / m, x)


@fn("ROUNDDOWN TRUNC", sig="number, [num_digits]")
def _rounddown(ctx, x, d=0.0):
    x = to_num(x)
    m = 10 ** (to_int(d) if _given(d) else 0)
    y = round(abs(x) * m, 9)
    return math.copysign(math.floor(y) / m, x)


@fn("INT", sig="number")
def _int(ctx, x):
    return float(math.floor(to_num(x)))


@fn("MOD", sig="number, divisor")
def _mod(ctx, x, d):
    x, d = to_num(x), to_num(d)
    if d == 0:
        raise errors.DIV0
    return x - d * math.floor(x / d)


@fn("QUOTIENT", sig="numerator, denominator")
def _quotient(ctx, x, d):
    x, d = to_num(x), to_num(d)
    if d == 0:
        raise errors.DIV0
    return float(math.trunc(x / d))


@fn("POWER", sig="number, power")
def _power(ctx, x, p):
    return _pow(to_num(x), to_num(p))


def _pow(x, p):
    if x == 0 and p < 0:
        raise errors.DIV0
    if x < 0 and not float(p).is_integer():
        raise errors.NUM
    try:
        return float(x ** p)
    except OverflowError:
        raise errors.NUM


@fn("SQRT", sig="number")
def _sqrt(ctx, x):
    x = to_num(x)
    if x < 0:
        raise errors.NUM
    return math.sqrt(x)


@fn("EXP", sig="number")
def _exp(ctx, x):
    return math.exp(to_num(x))


@fn("LN", sig="number")
def _ln(ctx, x):
    x = to_num(x)
    if x <= 0:
        raise errors.NUM
    return math.log(x)


@fn("LOG", sig="number, [base]")
def _log(ctx, x, b=10.0):
    x = to_num(x)
    b = to_num(b) if _given(b) else 10.0
    if x <= 0 or b <= 0 or b == 1:
        raise errors.NUM
    return math.log(x, b)


@fn("LOG10", sig="number")
def _log10(ctx, x):
    x = to_num(x)
    if x <= 0:
        raise errors.NUM
    return math.log10(x)


@fn("PI", sig="")
def _pi(ctx):
    return math.pi


@fn("SIGN", sig="number")
def _sign(ctx, x):
    x = to_num(x)
    return float((x > 0) - (x < 0))


@fn("CEILING CEILING.MATH", sig="number, [significance]")
def _ceiling(ctx, x, s=MISSING, *rest):
    x = to_num(x)
    s = to_num(s) if _given(s) else 1.0
    if s == 0:
        return 0.0
    if x > 0 > s:
        raise errors.NUM
    s = abs(s)
    return math.ceil(round(x / s, 9)) * s

@fn("FLOOR FLOOR.MATH", sig="number, [significance]")
def _floor(ctx, x, s=MISSING, *rest):
    x = to_num(x)
    s = abs(to_num(s)) if _given(s) else 1.0
    if s == 0:
        raise errors.DIV0
    return math.floor(round(x / s, 9)) * s


@fn("MROUND", sig="number, multiple")
def _mround(ctx, x, m):
    x, m = to_num(x), to_num(m)
    if m == 0:
        return 0.0
    if (x > 0 > m) or (x < 0 < m):
        raise errors.NUM
    return _round_half_away(x / m, 0) * m


@fn("EVEN", sig="number")
def _even(ctx, x):
    x = to_num(x)
    n = math.ceil(abs(x))
    if n % 2:
        n += 1
    return math.copysign(n, x)


@fn("ODD", sig="number")
def _odd(ctx, x):
    x = to_num(x)
    n = math.ceil(abs(x))
    if n % 2 == 0:
        n += 1
    return math.copysign(n, x)


@fn("FACT", sig="number")
def _fact(ctx, x):
    x = to_num(x)
    if x < 0:
        raise errors.NUM
    return float(math.factorial(int(x)))


@fn("GCD", sig="number1, [number2], ...")
def _gcd(ctx, *args):
    g = 0
    for v in nums(args):
        if v < 0:
            raise errors.NUM
        g = math.gcd(g, int(v))
    return float(g)


@fn("LCM", sig="number1, [number2], ...")
def _lcm(ctx, *args):
    m = 1
    for v in nums(args):
        if v < 0:
            raise errors.NUM
        m = math.lcm(m, int(v))
    return float(m)


@fn("RAND", sig="")
def _rand(ctx):
    return random.random()


@fn("RANDBETWEEN", sig="bottom, top")
def _randbetween(ctx, a, b):
    a, b = math.ceil(to_num(a)), math.floor(to_num(b))
    if a > b:
        raise errors.NUM
    return float(random.randint(a, b))


@fn("SUMPRODUCT", sig="array1, [array2], ...")
def _sumproduct(ctx, *arrays):
    grids = [to_2d(a) for a in arrays]
    if not grids:
        raise errors.VALUE
    shape = (len(grids[0]), len(grids[0][0]) if grids[0] else 0)
    for g in grids:
        if (len(g), len(g[0]) if g else 0) != shape:
            raise errors.VALUE
    total = 0.0
    for i in range(shape[0]):
        for j in range(shape[1]):
            p = 1.0
            for g in grids:
                v = g[i][j]
                if isinstance(v, XLError):
                    raise v
                p *= v if is_num(v) else 0.0
            total += p
    return total


@fn("SUMSQ", sig="number1, [number2], ...")
def _sumsq(ctx, *args):
    return math.fsum(v * v for v in nums(args))


for _name, _f in (("SIN", math.sin), ("COS", math.cos), ("TAN", math.tan), ("ATAN", math.atan),
                  ("DEGREES", math.degrees), ("RADIANS", math.radians), ("SINH", math.sinh),
                  ("COSH", math.cosh), ("TANH", math.tanh)):
    def _mk(f):
        return lambda ctx, x: f(to_num(x))
    fn(_name, sig="number")(_mk(_f))


@fn("ASIN", sig="number")
def _asin(ctx, x):
    x = to_num(x)
    if not -1 <= x <= 1:
        raise errors.NUM
    return math.asin(x)


@fn("ACOS", sig="number")
def _acos(ctx, x):
    x = to_num(x)
    if not -1 <= x <= 1:
        raise errors.NUM
    return math.acos(x)


@fn("ATAN2", sig="x_num, y_num")
def _atan2(ctx, x, y):
    x, y = to_num(x), to_num(y)
    if x == 0 and y == 0:
        raise errors.DIV0
    return math.atan2(y, x)


# ---------------------------------------------------------------- conditional aggregates

@fn("SUMIF", sig="range, criteria, [sum_range]")
def _sumif(ctx, rng, crit, sum_range=MISSING):
    shape, mask = _ifs_mask([(rng, crit)], ctx)
    vals = _aligned(sum_range if _given(sum_range) else rng, shape)
    total = 0.0
    for i, mrow in enumerate(mask):
        for j, ok in enumerate(mrow):
            if ok:
                v = vals[i][j]
                if is_num(v):
                    total += v
                elif isinstance(v, XLError):
                    raise v
    return total


def _pairs(args):
    if len(args) % 2:
        raise errors.VALUE
    return [(args[i], args[i + 1]) for i in range(0, len(args), 2)]


@fn("SUMIFS", sig="sum_range, criteria_range1, criteria1, ...")
def _sumifs(ctx, sum_range, *args):
    shape, mask = _ifs_mask(_pairs(args), ctx)
    vals = _aligned(sum_range, shape)
    total = 0.0
    for i, mrow in enumerate(mask):
        for j, ok in enumerate(mrow):
            if ok and is_num(vals[i][j]):
                total += vals[i][j]
    return total


@fn("COUNTIF", sig="range, criteria")
def _countif(ctx, rng, crit):
    pred = make_criteria(crit)
    if isinstance(rng, RangeRef) and not pred(None):
        return float(sum(1 for v in rng.nonblank() if pred(v)))
    shape, mask = _ifs_mask([(rng, crit)], ctx)
    n = sum(sum(r) for r in mask)
    if isinstance(rng, RangeRef):
        # blanks outside the used area also match
        used = shape[0] * shape[1]
        n += rng.nrows * rng.ncols - used
    return float(n)


@fn("COUNTIFS", sig="criteria_range1, criteria1, ...")
def _countifs(ctx, *args):
    _, mask = _ifs_mask(_pairs(args), ctx)
    return float(sum(sum(r) for r in mask))


def _masked_values(target, args, ctx):
    shape, mask = _ifs_mask(_pairs(args), ctx)
    vals = _aligned(target, shape)
    return [vals[i][j] for i, mrow in enumerate(mask) for j, ok in enumerate(mrow)
            if ok and is_num(vals[i][j])]


@fn("AVERAGEIF", sig="range, criteria, [average_range]")
def _averageif(ctx, rng, crit, avg_range=MISSING):
    xs = _masked_values(avg_range if _given(avg_range) else rng, [rng, crit], ctx)
    if not xs:
        raise errors.DIV0
    return math.fsum(xs) / len(xs)


@fn("AVERAGEIFS", sig="average_range, criteria_range1, criteria1, ...")
def _averageifs(ctx, avg_range, *args):
    xs = _masked_values(avg_range, list(args), ctx)
    if not xs:
        raise errors.DIV0
    return math.fsum(xs) / len(xs)


@fn("MAXIFS", sig="max_range, criteria_range1, criteria1, ...")
def _maxifs(ctx, rng, *args):
    xs = _masked_values(rng, list(args), ctx)
    return max(xs) if xs else 0.0


@fn("MINIFS", sig="min_range, criteria_range1, criteria1, ...")
def _minifs(ctx, rng, *args):
    xs = _masked_values(rng, list(args), ctx)
    return min(xs) if xs else 0.0


_SUBTOTAL = {1: "AVERAGE", 2: "COUNT", 3: "COUNTA", 4: "MAX", 5: "MIN", 6: "PRODUCT",
             7: "STDEV", 8: "STDEVP", 9: "SUM", 10: "VAR", 11: "VARP"}


@fn("SUBTOTAL", sig="function_num, ref1, [ref2], ...")
def _subtotal(ctx, num, *refs):
    n = to_int(num)
    base = _SUBTOTAL.get(n % 100)
    if base is None:
        raise errors.VALUE
    vis_args = []
    for r in refs:
        if isinstance(r, RangeRef):
            sh = r.sheet
            hidden = sh.filter_hidden | (sh.hidden_rows if n > 100 else set())
            rows = []
            for i in range(r.r1, r.eff_r2() + 1):
                if i in hidden:
                    continue
                rows.append([sh.value(i, c) for c in range(r.c1, r.eff_c2() + 1)])
            # nested SUBTOTAL results are ignored
            vis_args.append(rows)
        else:
            vis_args.append(r)
    f = FUNCS[base][0]
    return f(ctx, *vis_args)


# ================================================================ statistics

@fn("MEDIAN", sig="number1, [number2], ...")
def _median(ctx, *args):
    xs = _num_list(args)
    if not xs:
        raise errors.NUM
    return float(statistics.median(xs))


@fn("MODE MODE.SNGL", sig="number1, [number2], ...")
def _mode(ctx, *args):
    xs = _num_list(args)
    counts = {}
    for x in xs:
        counts[x] = counts.get(x, 0) + 1
    best = max(counts.values(), default=0)
    if best < 2:
        raise errors.NA
    for x in xs:
        if counts[x] == best:
            return x


def _var(xs, sample):
    n = len(xs)
    if n < (2 if sample else 1):
        raise errors.DIV0
    m = math.fsum(xs) / n
    return math.fsum((x - m) ** 2 for x in xs) / (n - 1 if sample else n)


@fn("STDEV STDEV.S", sig="number1, [number2], ...")
def _stdev(ctx, *args):
    return math.sqrt(_var(_num_list(args), True))


@fn("STDEVP STDEV.P", sig="number1, [number2], ...")
def _stdevp(ctx, *args):
    return math.sqrt(_var(_num_list(args), False))


@fn("VAR VAR.S", sig="number1, [number2], ...")
def _var_s(ctx, *args):
    return _var(_num_list(args), True)


@fn("VARP VAR.P", sig="number1, [number2], ...")
def _var_p(ctx, *args):
    return _var(_num_list(args), False)


@fn("LARGE", sig="array, k")
def _large(ctx, arr, k):
    xs = sorted(nums([arr]), reverse=True)
    k = to_int(k)
    if not 1 <= k <= len(xs):
        raise errors.NUM
    return xs[k - 1]


@fn("SMALL", sig="array, k")
def _small(ctx, arr, k):
    xs = sorted(nums([arr]))
    k = to_int(k)
    if not 1 <= k <= len(xs):
        raise errors.NUM
    return xs[k - 1]


@fn("RANK RANK.EQ", sig="number, ref, [order]")
def _rank(ctx, x, ref, order=MISSING):
    x = to_num(x)
    xs = list(nums([ref]))
    if x not in xs:
        raise errors.NA
    asc = _given(order) and to_num(order) != 0
    return float(1 + sum(1 for v in xs if (v < x if asc else v > x)))


def _percentile(xs, p):
    xs = sorted(xs)
    if not xs or not 0 <= p <= 1:
        raise errors.NUM
    pos = p * (len(xs) - 1)
    lo = math.floor(pos)
    hi = min(lo + 1, len(xs) - 1)
    return xs[lo] + (xs[hi] - xs[lo]) * (pos - lo)


@fn("PERCENTILE PERCENTILE.INC", sig="array, k")
def _percentile_f(ctx, arr, p):
    return _percentile(list(nums([arr])), to_num(p))


@fn("QUARTILE QUARTILE.INC", sig="array, quart")
def _quartile(ctx, arr, q):
    q = to_int(q)
    if not 0 <= q <= 4:
        raise errors.NUM
    return _percentile(list(nums([arr])), q / 4)


@fn("CORREL", sig="array1, array2")
def _correl(ctx, a, b):
    xa, xb = [v for v in flat(a)], [v for v in flat(b)]
    if len(xa) != len(xb):
        raise errors.NA
    pairs = [(x, y) for x, y in zip(xa, xb) if is_num(x) and is_num(y)]
    if len(pairs) < 2:
        raise errors.DIV0
    mx = sum(p[0] for p in pairs) / len(pairs)
    my = sum(p[1] for p in pairs) / len(pairs)
    sxy = sum((x - mx) * (y - my) for x, y in pairs)
    sxx = sum((x - mx) ** 2 for x, _ in pairs)
    syy = sum((y - my) ** 2 for _, y in pairs)
    if sxx == 0 or syy == 0:
        raise errors.DIV0
    return sxy / math.sqrt(sxx * syy)


# ================================================================ logic

@fn("IF", kind="lazy", sig="logical_test, [value_if_true], [value_if_false]", minargs=1, maxargs=3)
def _if(ev, ctx, nodes):
    cond = to_bool(scalar(ev.ev(nodes[0], ctx), ctx))
    if cond:
        return ev.ev(nodes[1], ctx) if len(nodes) > 1 else True
    if len(nodes) > 2:
        return ev.ev(nodes[2], ctx)
    return False


@fn("IFS", kind="lazy", sig="logical_test1, value_if_true1, ...", minargs=2)
def _ifs(ev, ctx, nodes):
    if len(nodes) % 2:
        raise errors.VALUE
    for i in range(0, len(nodes), 2):
        if to_bool(scalar(ev.ev(nodes[i], ctx), ctx)):
            return ev.ev(nodes[i + 1], ctx)
    raise errors.NA


@fn("IFERROR", kind="lazy", sig="value, value_if_error", minargs=2, maxargs=2)
def _iferror(ev, ctx, nodes):
    try:
        v = ev.ev(nodes[0], ctx)
        s = scalar(v, ctx) if isinstance(v, RangeRef) else v
        if isinstance(s, XLError):
            return ev.ev(nodes[1], ctx)
        return v
    except XLError as e:
        e.__traceback__ = None
        return ev.ev(nodes[1], ctx)


@fn("IFNA", kind="lazy", sig="value, value_if_na", minargs=2, maxargs=2)
def _ifna(ev, ctx, nodes):
    try:
        v = ev.ev(nodes[0], ctx)
        s = scalar(v, ctx) if isinstance(v, RangeRef) else v
        if s == errors.NA:
            return ev.ev(nodes[1], ctx)
        return v
    except XLError as e:
        e.__traceback__ = None
        if e == errors.NA:
            return ev.ev(nodes[1], ctx)
        raise


def _bools(args):
    out = []
    for a in args:
        if isinstance(a, (RangeRef, list)):
            for v in flat(a):
                if isinstance(v, XLError):
                    raise v
                if isinstance(v, bool):
                    out.append(v)
                elif is_num(v):
                    out.append(v != 0)
        elif a is MISSING:
            continue
        else:
            out.append(to_bool(a))
    if not out:
        raise errors.VALUE
    return out


@fn("AND", sig="logical1, [logical2], ...")
def _and(ctx, *args):
    return all(_bools(args))


@fn("OR", sig="logical1, [logical2], ...")
def _or(ctx, *args):
    return any(_bools(args))


@fn("XOR", sig="logical1, [logical2], ...")
def _xor(ctx, *args):
    return sum(_bools(args)) % 2 == 1


@fn("NOT", sig="logical")
def _not(ctx, x):
    return not to_bool(x)


@fn("TRUE", sig="")
def _true(ctx):
    return True


@fn("FALSE", sig="")
def _false(ctx):
    return False


@fn("CHOOSE", kind="lazy", sig="index_num, value1, [value2], ...", minargs=2)
def _choose(ev, ctx, nodes):
    i = to_int(scalar(ev.ev(nodes[0], ctx), ctx))
    if not 1 <= i < len(nodes):
        raise errors.VALUE
    return ev.ev(nodes[i], ctx)


@fn("SWITCH", kind="lazy", sig="expression, value1, result1, ..., [default]", minargs=3)
def _switch(ev, ctx, nodes):
    x = scalar(ev.ev(nodes[0], ctx), ctx)
    rest = nodes[1:]
    for i in range(0, len(rest) - 1, 2):
        v = scalar(ev.ev(rest[i], ctx), ctx)
        try:
            if compare(x, v) == 0:
                return ev.ev(rest[i + 1], ctx)
        except XLError:
            raise
    if len(rest) % 2:
        return ev.ev(rest[-1], ctx)
    raise errors.NA


# ================================================================ information

def _info(name, pred, sig="value"):
    fn(name, kind="safe", sig=sig)(lambda ctx, v=None: pred(scalar(v, ctx)))


_info("ISBLANK", lambda v: v is None)
_info("ISNUMBER", is_num)
_info("ISTEXT", lambda v: isinstance(v, str))
_info("ISNONTEXT", lambda v: not isinstance(v, str))
_info("ISLOGICAL", lambda v: isinstance(v, bool))
_info("ISERROR", is_err)
_info("ISERR", lambda v: is_err(v) and v != errors.NA)
_info("ISNA", lambda v: v == errors.NA)


@fn("ISREF", kind="safe", sig="value")
def _isref(ctx, v=None):
    return isinstance(v, RangeRef)


@fn("ISEVEN", sig="number")
def _iseven(ctx, x):
    return math.floor(to_num(x)) % 2 == 0


@fn("ISODD", sig="number")
def _isodd(ctx, x):
    return math.floor(to_num(x)) % 2 == 1


@fn("NA", sig="")
def _na(ctx):
    raise errors.NA


@fn("N", sig="value")
def _n(ctx, v):
    v = scalar(v, ctx)
    if is_num(v):
        return v
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if isinstance(v, XLError):
        raise v
    return 0.0


@fn("T", sig="value")
def _t(ctx, v):
    v = scalar(v, ctx)
    if isinstance(v, XLError):
        raise v
    return v if isinstance(v, str) else ""


@fn("TYPE", kind="safe", sig="value")
def _type(ctx, v):
    if isinstance(v, list) or (isinstance(v, RangeRef) and not v.is_cell()):
        return 64.0
    v = scalar(v, ctx)
    if is_num(v) or v is None:
        return 1.0
    if isinstance(v, str):
        return 2.0
    if isinstance(v, bool):
        return 4.0
    return 16.0


@fn("ERROR.TYPE", kind="safe", sig="error_val")
def _error_type(ctx, v):
    v = scalar(v, ctx)
    order = ["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"]
    if isinstance(v, XLError) and v.code in order:
        return float(order.index(v.code) + 1)
    raise errors.NA


# ================================================================ text

@fn("LEN", sig="text")
def _len(ctx, s):
    return float(len(to_str(s)))


@fn("LEFT", sig="text, [num_chars]")
def _left(ctx, s, n=MISSING):
    n = to_int(n) if _given(n) else 1
    if n < 0:
        raise errors.VALUE
    return to_str(s)[:n]


@fn("RIGHT", sig="text, [num_chars]")
def _right(ctx, s, n=MISSING):
    n = to_int(n) if _given(n) else 1
    if n < 0:
        raise errors.VALUE
    s = to_str(s)
    return s[len(s) - n:] if n else ""


@fn("MID", sig="text, start_num, num_chars")
def _mid(ctx, s, start, n):
    start, n = to_int(start), to_int(n)
    if start < 1 or n < 0:
        raise errors.VALUE
    return to_str(s)[start - 1:start - 1 + n]


@fn("UPPER", sig="text")
def _upper(ctx, s):
    return to_str(s).upper()


@fn("LOWER", sig="text")
def _lower(ctx, s):
    return to_str(s).lower()


@fn("PROPER", sig="text")
def _proper(ctx, s):
    return re.sub(r"[A-Za-z]+", lambda m: m.group(0).capitalize(), to_str(s))


@fn("TRIM", sig="text")
def _trim(ctx, s):
    return re.sub(r" +", " ", to_str(s)).strip(" ")


@fn("CLEAN", sig="text")
def _clean(ctx, s):
    return "".join(ch for ch in to_str(s) if ord(ch) >= 32)


@fn("CONCATENATE", sig="text1, [text2], ...")
def _concatenate(ctx, *args):
    return "".join(to_str(scalar(a, ctx)) for a in args if a is not MISSING)


@fn("CONCAT", sig="text1, [text2], ...")
def _concat(ctx, *args):
    out = []
    for a in args:
        if isinstance(a, (RangeRef, list)):
            out.extend(to_str(v) for v in flat(a))
        elif a is not MISSING:
            out.append(to_str(a))
    return "".join(out)


@fn("TEXTJOIN", sig="delimiter, ignore_empty, text1, ...")
def _textjoin(ctx, delim, ignore_empty, *args):
    d = to_str(delim)
    skip = to_bool(ignore_empty)
    parts = []
    for a in args:
        src = flat(a) if isinstance(a, (RangeRef, list)) else [a]
        for v in src:
            if v is MISSING:
                v = None
            s = to_str(v)
            if s == "" and skip:
                continue
            parts.append(s)
    return d.join(parts)


@fn("TEXT", sig="value, format_text")
def _text(ctx, v, fmt):
    v = scalar(v, ctx)
    if isinstance(v, XLError):
        raise v
    if isinstance(v, str):
        x, _ = parse_input(v)
        if is_num(x):
            v = x
    if v is None:
        v = 0.0
    return format_value(v, to_str(fmt))[0]


@fn("VALUE NUMBERVALUE", sig="text")
def _value(ctx, s, *rest):
    v = scalar(s, ctx)
    if is_num(v):
        return v
    if isinstance(v, XLError):
        raise v
    x, _ = parse_input(to_str(v))
    if is_num(x):
        return x
    raise errors.VALUE


@fn("FIXED", sig="number, [decimals], [no_commas]")
def _fixed(ctx, x, d=MISSING, no_commas=MISSING):
    x = to_num(x)
    d = to_int(d) if _given(d) else 2
    x = _round_half_away(x, d)
    if _given(no_commas) and to_bool(no_commas):
        return f"{x:.{max(d, 0)}f}"
    return f"{x:,.{max(d, 0)}f}"


@fn("DOLLAR", sig="number, [decimals]")
def _dollar(ctx, x, d=MISSING):
    d = to_int(d) if _given(d) else 2
    fmt = "$#,##0" + ("." + "0" * d if d > 0 else "") + "_);($#,##0" + ("." + "0" * d if d > 0 else "") + ")"
    return format_value(_round_half_away(to_num(x), d), fmt)[0].strip()


@fn("FIND", sig="find_text, within_text, [start_num]")
def _find(ctx, needle, hay, start=MISSING):
    start = to_int(start) if _given(start) else 1
    hay = to_str(hay)
    if start < 1 or start > len(hay) + 1:
        raise errors.VALUE
    i = hay.find(to_str(needle), start - 1)
    if i < 0:
        raise errors.VALUE
    return float(i + 1)


@fn("SEARCH", sig="find_text, within_text, [start_num]")
def _search(ctx, needle, hay, start=MISSING):
    start = to_int(start) if _given(start) else 1
    hay = to_str(hay)
    if start < 1 or start > len(hay) + 1:
        raise errors.VALUE
    rx = _wild_compile(to_str(needle))
    m = re.compile(rx.pattern, re.IGNORECASE | re.DOTALL).search(hay, start - 1)
    if not m:
        raise errors.VALUE
    return float(m.start() + 1)


@fn("SUBSTITUTE", sig="text, old_text, new_text, [instance_num]")
def _substitute(ctx, s, old, new, inst=MISSING):
    s, old, new = to_str(s), to_str(old), to_str(new)
    if old == "":
        return s
    if not _given(inst):
        return s.replace(old, new)
    n = to_int(inst)
    if n < 1:
        raise errors.VALUE
    idx = -1
    for _ in range(n):
        idx = s.find(old, idx + 1)
        if idx < 0:
            return s
    return s[:idx] + new + s[idx + len(old):]


@fn("REPLACE", sig="old_text, start_num, num_chars, new_text")
def _replace(ctx, s, start, n, new):
    s = to_str(s)
    start, n = to_int(start), to_int(n)
    if start < 1 or n < 0:
        raise errors.VALUE
    return s[:start - 1] + to_str(new) + s[start - 1 + n:]


@fn("REPT", sig="text, number_times")
def _rept(ctx, s, n):
    n = to_int(n)
    if n < 0:
        raise errors.VALUE
    return to_str(s) * n


@fn("EXACT", sig="text1, text2")
def _exact(ctx, a, b):
    return to_str(a) == to_str(b)


@fn("CHAR", sig="number")
def _char(ctx, n):
    n = to_int(n)
    if not 1 <= n <= 255:
        raise errors.VALUE
    return bytes([n]).decode("cp1252", errors="replace")


@fn("UNICHAR", sig="number")
def _unichar(ctx, n):
    return chr(to_int(n))


@fn("CODE UNICODE", sig="text")
def _code(ctx, s):
    s = to_str(s)
    if not s:
        raise errors.VALUE
    return float(ord(s[0]))


@fn("TEXTBEFORE", sig="text, delimiter, [instance_num]")
def _textbefore(ctx, s, d, inst=MISSING, *rest):
    s, d = to_str(s), to_str(d)
    n = to_int(inst) if _given(inst) else 1
    parts = s.split(d)
    if len(parts) <= abs(n):
        raise errors.NA
    return d.join(parts[:n]) if n > 0 else d.join(parts[:len(parts) + n])


@fn("TEXTAFTER", sig="text, delimiter, [instance_num]")
def _textafter(ctx, s, d, inst=MISSING, *rest):
    s, d = to_str(s), to_str(d)
    n = to_int(inst) if _given(inst) else 1
    parts = s.split(d)
    if len(parts) <= abs(n):
        raise errors.NA
    return d.join(parts[n:]) if n > 0 else d.join(parts[len(parts) + n:])


# ================================================================ date & time

@fn("TODAY", sig="")
def _today(ctx):
    return _serial(dt.date.today())


@fn("NOW", sig="")
def _now(ctx):
    return _serial(dt.datetime.now())


@fn("DATE", sig="year, month, day")
def _date_f(ctx, y, m, d):
    y, m, d = to_int(y), to_int(m), to_int(d)
    if 0 <= y < 1900:
        y += 1900
    y += (m - 1) // 12
    m = (m - 1) % 12 + 1
    try:
        base = dt.datetime(y, m, 1)
    except ValueError:
        raise errors.NUM
    s = _serial(base) + d - 1
    if s < 0:
        raise errors.NUM
    return float(s)


@fn("TIME", sig="hour, minute, second")
def _time(ctx, h, m, s):
    total = to_int(h) * 3600 + to_int(m) * 60 + to_int(s)
    if total < 0:
        raise errors.NUM
    return (total % 86400) / 86400


@fn("YEAR", sig="serial_number")
def _year(ctx, v):
    return float(_date(_dateish(v)).year)


@fn("MONTH", sig="serial_number")
def _month(ctx, v):
    return float(_date(_dateish(v)).month)


@fn("DAY", sig="serial_number")
def _day(ctx, v):
    x = to_num(_dateish(v))
    if 60 <= x < 61:
        return 29.0
    return float(_date(x).day)


def _time_parts(v):
    x = to_num(_dateish(v))
    secs = round((x - math.floor(x)) * 86400)
    return secs // 3600 % 24, secs // 60 % 60, secs % 60


@fn("HOUR", sig="serial_number")
def _hour(ctx, v):
    return float(_time_parts(v)[0])


@fn("MINUTE", sig="serial_number")
def _minute(ctx, v):
    return float(_time_parts(v)[1])


@fn("SECOND", sig="serial_number")
def _second(ctx, v):
    return float(_time_parts(v)[2])


def _dateish(v):
    v = scalar(v)
    if isinstance(v, str):
        x, _ = parse_input(v)
        if not is_num(x):
            raise errors.VALUE
        return x
    return v


@fn("WEEKDAY", sig="serial_number, [return_type]")
def _weekday(ctx, v, t=MISSING):
    d = _date(_dateish(v))
    t = to_int(t) if _given(t) else 1
    wd = d.weekday()  # Monday=0
    if t == 1:
        return float((wd + 1) % 7 + 1)
    if t == 2:
        return float(wd + 1)
    if t == 3:
        return float(wd)
    if 11 <= t <= 17:
        return float((wd - (t - 11)) % 7 + 1)
    raise errors.NUM


@fn("WEEKNUM", sig="serial_number, [return_type]")
def _weeknum(ctx, v, t=MISSING):
    d = _date(_dateish(v))
    t = to_int(t) if _given(t) else 1
    if t == 21:
        return float(d.isocalendar()[1])
    start = 6 if t == 1 else 0  # week starts Sunday (1) or Monday (2)
    jan1 = dt.datetime(d.year, 1, 1)
    offset = (jan1.weekday() - start) % 7
    return float((d - jan1).days + offset) // 7 + 1


@fn("ISOWEEKNUM", sig="date")
def _isoweeknum(ctx, v):
    return float(_date(_dateish(v)).isocalendar()[1])


def _add_months(d, n):
    m = d.month - 1 + n
    y = d.year + m // 12
    m = m % 12 + 1
    import calendar
    last = calendar.monthrange(y, m)[1]
    return d.replace(year=y, month=m, day=min(d.day, last)), last


@fn("EDATE", sig="start_date, months")
def _edate(ctx, v, n):
    d = _date(_dateish(v))
    nd, _ = _add_months(d.replace(hour=0, minute=0, second=0, microsecond=0), to_int(n))
    return float(_serial(nd))


@fn("EOMONTH", sig="start_date, months")
def _eomonth(ctx, v, n):
    d = _date(_dateish(v)).replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    nd, last = _add_months(d, to_int(n))
    return float(_serial(nd.replace(day=last)))


@fn("DAYS", sig="end_date, start_date")
def _days(ctx, end, start):
    return float(math.floor(to_num(_dateish(end))) - math.floor(to_num(_dateish(start))))


@fn("DATEDIF", sig="start_date, end_date, unit")
def _datedif(ctx, a, b, unit):
    d1, d2 = _date(_dateish(a)), _date(_dateish(b))
    if d1 > d2:
        raise errors.NUM
    u = to_str(unit).upper()
    months = (d2.year - d1.year) * 12 + d2.month - d1.month - (1 if d2.day < d1.day else 0)
    if u == "D":
        return float((d2.date() - d1.date()).days)
    if u == "M":
        return float(months)
    if u == "Y":
        return float(months // 12)
    if u == "YM":
        return float(months % 12)
    if u == "MD":
        if d2.day >= d1.day:
            return float(d2.day - d1.day)
        prev, _ = _add_months(d2.replace(day=1), -1)
        import calendar
        return float(calendar.monthrange(prev.year, prev.month)[1] - d1.day + d2.day)
    if u == "YD":
        try:
            anchor = d1.replace(year=d2.year)
        except ValueError:
            anchor = d1.replace(year=d2.year, day=28)
        if anchor > d2:
            anchor = anchor.replace(year=d2.year - 1)
        return float((d2 - anchor).days)
    raise errors.NUM


@fn("DATEVALUE", sig="date_text")
def _datevalue(ctx, s):
    x, f = parse_input(to_str(s))
    if not is_num(x) or not f:
        raise errors.VALUE
    return float(math.floor(x))


@fn("TIMEVALUE", sig="time_text")
def _timevalue(ctx, s):
    x, f = parse_input(to_str(s))
    if not is_num(x) or not f:
        raise errors.VALUE
    return x - math.floor(x)


def _holidays(h):
    if not _given(h):
        return set()
    return {math.floor(v) for v in nums([h])}


@fn("NETWORKDAYS", sig="start_date, end_date, [holidays]")
def _networkdays(ctx, a, b, hol=MISSING):
    s, e = math.floor(to_num(_dateish(a))), math.floor(to_num(_dateish(b)))
    sign = 1
    if s > e:
        s, e, sign = e, s, -1
    hs = _holidays(hol)
    n = 0
    for x in range(int(s), int(e) + 1):
        if serial_to_datetime(x).weekday() < 5 and x not in hs:
            n += 1
    return float(n * sign)


@fn("WORKDAY", sig="start_date, days, [holidays]")
def _workday(ctx, a, days, hol=MISSING):
    x = math.floor(to_num(_dateish(a)))
    n = to_int(days)
    hs = _holidays(hol)
    step = 1 if n >= 0 else -1
    while n:
        x += step
        if serial_to_datetime(x).weekday() < 5 and x not in hs:
            n -= step
    return float(x)


@fn("YEARFRAC", sig="start_date, end_date, [basis]")
def _yearfrac(ctx, a, b, basis=MISSING):
    d1, d2 = sorted((_date(_dateish(a)), _date(_dateish(b))))
    basis = to_int(basis) if _given(basis) else 0
    if basis == 0:
        dd1, dd2 = min(d1.day, 30), d2.day
        if dd2 == 31 and dd1 >= 30:
            dd2 = 30
        return ((d2.year - d1.year) * 360 + (d2.month - d1.month) * 30 + dd2 - dd1) / 360
    days = (d2 - d1).days
    return days / {1: 365.25, 2: 360, 3: 365, 4: 360}.get(basis, 365)


# ================================================================ lookup & reference

def _lookup_eq(a, b):
    if isinstance(a, str) and isinstance(b, str):
        return a.lower() == b.lower()
    if is_num(a) and is_num(b):
        return a == b
    if isinstance(a, bool) and isinstance(b, bool):
        return a == b
    return False


def _find_exact(vec, x, wildcard=True, reverse=False):
    idx = range(len(vec) - 1, -1, -1) if reverse else range(len(vec))
    if isinstance(x, str) and wildcard and _has_wild(x):
        for i in idx:
            v = vec[i]
            if isinstance(v, str) and wild_match(x, v):
                return i
        return -1
    if isinstance(x, str):
        xl = x.lower()
        for i in idx:
            v = vec[i]
            if isinstance(v, str) and v.lower() == xl:
                return i
        return -1
    for i in idx:
        if _lookup_eq(vec[i], x):
            return i
    return -1


def _same_kind(a, b):
    return (is_num(a) and is_num(b)) or (isinstance(a, str) and isinstance(b, str)) or (
        isinstance(a, bool) and isinstance(b, bool))


def _find_approx(vec, x, descending=False):
    """Binary search like Excel: largest value <= x (or smallest >= x if descending)."""
    lo, hi = 0, len(vec) - 1
    best = -1
    while lo <= hi:
        mid = (lo + hi) // 2
        # skip blanks / other types by scanning to a comparable neighbour
        m = mid
        while m <= hi and not _same_kind(vec[m], x):
            m += 1
        if m > hi:
            hi = mid - 1
            continue
        c = compare(vec[m], x)
        if (c <= 0) if not descending else (c >= 0):
            best = m
            lo = m + 1
        else:
            hi = mid - 1
    return best


def _vector(v):
    """1-D list from a range/array (row or column)."""
    if isinstance(v, RangeRef):
        if v.ncols == 1:
            return v.column(0)
        if v.nrows == 1:
            return v.row(0)
        return v.column(0)
    if isinstance(v, list):
        if len(v) == 1:
            return list(v[0])
        return [r[0] for r in v]
    return [v]


@fn("VLOOKUP", sig="lookup_value, table_array, col_index_num, [range_lookup]")
def _vlookup(ctx, x, table, col, approx=MISSING):
    x = scalar(x, ctx)
    if isinstance(x, XLError):
        raise x
    col = to_int(col)
    approx = to_bool(approx) if _given(approx) else True
    if isinstance(table, RangeRef):
        if col < 1:
            raise errors.VALUE
        if col > table.ncols:
            raise errors.REF
        keys = table.column(0)
        i = _find_approx(keys, x) if approx else _find_exact(keys, x)
        if i < 0:
            raise errors.NA
        return table.value(i, col - 1)
    g = to_2d(table)
    if col < 1:
        raise errors.VALUE
    if col > len(g[0]):
        raise errors.REF
    keys = [r[0] for r in g]
    i = _find_approx(keys, x) if approx else _find_exact(keys, x)
    if i < 0:
        raise errors.NA
    return g[i][col - 1]


@fn("HLOOKUP", sig="lookup_value, table_array, row_index_num, [range_lookup]")
def _hlookup(ctx, x, table, row, approx=MISSING):
    x = scalar(x, ctx)
    if isinstance(x, XLError):
        raise x
    row = to_int(row)
    approx = to_bool(approx) if _given(approx) else True
    g = to_2d(table) if not isinstance(table, RangeRef) else None
    keys = table.row(0) if g is None else g[0]
    nrows = table.nrows if g is None else len(g)
    if row < 1:
        raise errors.VALUE
    if row > nrows:
        raise errors.REF
    i = _find_approx(keys, x) if approx else _find_exact(keys, x)
    if i < 0:
        raise errors.NA
    return table.value(row - 1, i) if g is None else g[row - 1][i]


@fn("LOOKUP", sig="lookup_value, lookup_vector, [result_vector]")
def _lookup(ctx, x, vec, res=MISSING):
    x = scalar(x, ctx)
    keys = _vector(vec)
    i = _find_approx(keys, x)
    if i < 0:
        raise errors.NA
    out = _vector(res) if _given(res) else keys
    if i >= len(out):
        raise errors.NA
    return out[i]


@fn("MATCH", sig="lookup_value, lookup_array, [match_type]")
def _match(ctx, x, arr, mt=MISSING):
    x = scalar(x, ctx)
    if isinstance(x, XLError):
        raise x
    mt = to_int(mt) if _given(mt) else 1
    vec = _vector(arr)
    if mt == 0:
        i = _find_exact(vec, x)
    else:
        i = _find_approx(vec, x, descending=mt < 0)
    if i < 0:
        raise errors.NA
    return float(i + 1)


@fn("XMATCH", sig="lookup_value, lookup_array, [match_mode], [search_mode]")
def _xmatch(ctx, x, arr, mm=MISSING, sm=MISSING):
    i = _xfind(scalar(x, ctx), _vector(arr), mm, sm)
    if i < 0:
        raise errors.NA
    return float(i + 1)


def _xfind(x, vec, mm, sm):
    if isinstance(x, XLError):
        raise x
    mm = to_int(mm) if _given(mm) else 0
    sm = to_int(sm) if _given(sm) else 1
    rev = sm in (-1, -2)
    if mm in (0, 2):
        return _find_exact(vec, x, wildcard=(mm == 2), reverse=rev)
    i = _find_exact(vec, x, wildcard=False, reverse=rev)
    if i >= 0:
        return i
    best, bi = None, -1
    for j, v in enumerate(vec):
        if not _same_kind(v, x):
            continue
        c = compare(v, x)
        if (mm == -1 and c < 0 and (best is None or compare(v, best) > 0)) or \
           (mm == 1 and c > 0 and (best is None or compare(v, best) < 0)):
            best, bi = v, j
    return bi


@fn("XLOOKUP", sig="lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode]")
def _xlookup(ctx, x, look, ret, not_found=MISSING, mm=MISSING, sm=MISSING):
    vec = _vector(look)
    i = _xfind(scalar(x, ctx), vec, mm, sm)
    if i < 0:
        if not_found is not MISSING:
            return not_found
        raise errors.NA
    vertical = not (isinstance(look, RangeRef) and look.nrows == 1 and look.ncols > 1) and not (
        isinstance(look, list) and len(look) == 1 and len(look[0]) > 1)
    if isinstance(ret, RangeRef):
        if vertical:
            if ret.ncols == 1:
                return RangeRef(ret.sheet, ret.r1 + i, ret.c1, ret.r1 + i, ret.c1)
            return RangeRef(ret.sheet, ret.r1 + i, ret.c1, ret.r1 + i, ret.c2)
        return RangeRef(ret.sheet, ret.r1, ret.c1 + i, ret.r2, ret.c1 + i)
    g = to_2d(ret)
    try:
        return g[i][0] if vertical else g[0][i]
    except IndexError:
        raise errors.VALUE


@fn("INDEX", sig="array, row_num, [column_num]")
def _index(ctx, arr, row, col=MISSING, *rest):
    r = to_int(row) if _given(row) else 0
    if isinstance(arr, RangeRef):
        if not _given(col) and (arr.nrows == 1 and arr.ncols > 1):
            c, r = r, 1
        else:
            c = to_int(col) if _given(col) else (1 if arr.ncols == 1 else 0)
        if r < 0 or c < 0 or r > arr.nrows or c > arr.ncols:
            raise errors.REF
        if r == 0 and c == 0:
            return arr
        if r == 0:
            return RangeRef(arr.sheet, arr.r1, arr.c1 + c - 1, arr.r2, arr.c1 + c - 1)
        if c == 0:
            return RangeRef(arr.sheet, arr.r1 + r - 1, arr.c1, arr.r1 + r - 1, arr.c2)
        return RangeRef(arr.sheet, arr.r1 + r - 1, arr.c1 + c - 1, arr.r1 + r - 1, arr.c1 + c - 1)
    g = to_2d(arr)
    if not _given(col) and len(g) == 1:
        c, r = r, 1
    else:
        c = to_int(col) if _given(col) else 1
    if not (1 <= r <= len(g)) or not (1 <= c <= len(g[0])):
        raise errors.REF
    return g[r - 1][c - 1]


@fn("ROW", kind="safe", sig="[reference]")
def _row(ctx, ref=MISSING):
    if isinstance(ref, RangeRef):
        return float(ref.r1 + 1)
    if ref is MISSING:
        return float(ctx[1] + 1)
    raise errors.VALUE


@fn("COLUMN", kind="safe", sig="[reference]")
def _column(ctx, ref=MISSING):
    if isinstance(ref, RangeRef):
        return float(ref.c1 + 1)
    if ref is MISSING:
        return float(ctx[2] + 1)
    raise errors.VALUE


@fn("ROWS", sig="array")
def _rows(ctx, a):
    if isinstance(a, RangeRef):
        return float(a.nrows)
    return float(len(to_2d(a)))


@fn("COLUMNS", sig="array")
def _columns(ctx, a):
    if isinstance(a, RangeRef):
        return float(a.ncols)
    g = to_2d(a)
    return float(len(g[0]) if g else 0)


@fn("OFFSET", sig="reference, rows, cols, [height], [width]")
def _offset(ctx, ref, rows, cols, h=MISSING, w=MISSING):
    if not isinstance(ref, RangeRef):
        raise errors.VALUE
    r1 = ref.r1 + to_int(rows)
    c1 = ref.c1 + to_int(cols)
    hh = to_int(h) if _given(h) else ref.nrows
    ww = to_int(w) if _given(w) else ref.ncols
    if hh < 1 or ww < 1 or r1 < 0 or c1 < 0 or r1 + hh > MAX_ROWS or c1 + ww > MAX_COLS:
        raise errors.REF
    return RangeRef(ref.sheet, r1, c1, r1 + hh - 1, c1 + ww - 1)


@fn("INDIRECT", sig="ref_text, [a1]")
def _indirect(ctx, text, a1=MISSING):
    s = to_str(text).strip()
    sheet = ctx[0]
    if "!" in s:
        name, s = s.rsplit("!", 1)
        if name.startswith("'") and name.endswith("'"):
            name = name[1:-1].replace("''", "'")
        sheet = sheet.wb.get_sheet(name)
        if sheet is None:
            raise errors.REF
    b = parse_range(s)
    if b is None:
        raise errors.REF
    return RangeRef(sheet, *b)


@fn("_RANGEOP", sig=None)
def _rangeop(ctx, a, b):
    if not (isinstance(a, RangeRef) and isinstance(b, RangeRef)) or a.sheet is not b.sheet:
        raise errors.VALUE
    return RangeRef(a.sheet, min(a.r1, b.r1), min(a.c1, b.c1), max(a.r2, b.r2), max(a.c2, b.c2))


@fn("TRANSPOSE", sig="array")
def _transpose(ctx, a):
    g = to_2d(a)
    return [list(r) for r in zip(*g)]


# ================================================================ financial

@fn("PMT", sig="rate, nper, pv, [fv], [type]")
def _pmt(ctx, rate, nper, pv, fv=MISSING, typ=MISSING):
    r, n, p = to_num(rate), to_num(nper), to_num(pv)
    f = to_num(fv) if _given(fv) else 0.0
    t = to_num(typ) if _given(typ) else 0.0
    if n == 0:
        raise errors.NUM
    if r == 0:
        return -(p + f) / n
    q = (1 + r) ** n
    return -(r * (p * q + f)) / ((1 + r * t) * (q - 1))


@fn("FV", sig="rate, nper, pmt, [pv], [type]")
def _fv(ctx, rate, nper, pmt, pv=MISSING, typ=MISSING):
    r, n, m = to_num(rate), to_num(nper), to_num(pmt)
    p = to_num(pv) if _given(pv) else 0.0
    t = to_num(typ) if _given(typ) else 0.0
    if r == 0:
        return -(p + m * n)
    q = (1 + r) ** n
    return -(p * q + m * (1 + r * t) * (q - 1) / r)


@fn("PV", sig="rate, nper, pmt, [fv], [type]")
def _pv(ctx, rate, nper, pmt, fv=MISSING, typ=MISSING):
    r, n, m = to_num(rate), to_num(nper), to_num(pmt)
    f = to_num(fv) if _given(fv) else 0.0
    t = to_num(typ) if _given(typ) else 0.0
    if r == 0:
        return -(f + m * n)
    q = (1 + r) ** n
    return -(f + m * (1 + r * t) * (q - 1) / r) / q


@fn("NPV", sig="rate, value1, [value2], ...")
def _npv(ctx, rate, *args):
    r = to_num(rate)
    return sum(v / (1 + r) ** (i + 1) for i, v in enumerate(nums(args)))


@fn("IRR", sig="values, [guess]")
def _irr(ctx, values, guess=MISSING):
    cfs = list(nums([values]))
    r = to_num(guess) if _given(guess) else 0.1
    for _ in range(100):
        f = sum(c / (1 + r) ** i for i, c in enumerate(cfs))
        df = sum(-i * c / (1 + r) ** (i + 1) for i, c in enumerate(cfs))
        if df == 0:
            break
        nr = r - f / df
        if abs(nr - r) < 1e-10:
            return nr
        r = nr
    raise errors.NUM


@fn("GETPIVOTDATA", sig="data_field, pivot_table, [field1, item1], ...", minargs=2)
def _getpivotdata(ctx, data_field, table, *pairs):
    if not isinstance(table, RangeRef):
        raise errors.REF
    from .pivot import getpivotdata
    vals = [scalar(a) for a in pairs]
    return getpivotdata(table.sheet, table.r1, table.c1, scalar(data_field), vals)


VOLATILE.add("GETPIVOTDATA")   # reads what the pivot shows, which changes when it's refreshed

ALL_NAMES = sorted(n for n in FUNCS if not n.startswith("_"))
