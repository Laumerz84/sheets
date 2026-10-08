"""Runtime values used by the evaluator: range references, coercions, comparisons."""
import math

from . import errors
from .errors import XLError
from .numfmt import parse_input


class _Missing:
    """An omitted function argument, e.g. the middle of IF(A1,,2)."""

    def __repr__(self):
        return "MISSING"


MISSING = _Missing()


class RangeRef:
    __slots__ = ("sheet", "r1", "c1", "r2", "c2")

    def __init__(self, sheet, r1, c1, r2, c2):
        self.sheet = sheet
        self.r1, self.c1, self.r2, self.c2 = r1, c1, r2, c2

    def __repr__(self):
        return f"RangeRef({self.sheet.name},{self.r1},{self.c1},{self.r2},{self.c2})"

    @property
    def nrows(self):
        return self.r2 - self.r1 + 1

    @property
    def ncols(self):
        return self.c2 - self.c1 + 1

    def is_cell(self):
        return self.r1 == self.r2 and self.c1 == self.c2

    def eff_r2(self):
        return min(self.r2, self.sheet.max_row)

    def eff_c2(self):
        return min(self.c2, self.sheet.max_col)

    def value(self, i=0, j=0):
        r, c = self.r1 + i, self.c1 + j
        if r > self.r2 or c > self.c2:
            raise errors.REF
        return self.sheet.value(r, c)

    def values(self):
        """Row-major values of the used part of the range (blanks included)."""
        val = self.sheet.value
        er2, ec2 = self.eff_r2(), self.eff_c2()
        for r in range(self.r1, er2 + 1):
            for c in range(self.c1, ec2 + 1):
                yield val(r, c)

    def nonblank(self):
        """Faster iteration over non-blank values only (order not guaranteed)."""
        sh = self.sheet
        er2, ec2 = self.eff_r2(), self.eff_c2()
        area = (er2 - self.r1 + 1) * (ec2 - self.c1 + 1)
        if area <= 0:
            return
        if area < 4 * (len(sh.values) + len(sh.formulas)):
            for v in self.values():
                if v is not None:
                    yield v
            return
        for k in sh.keys_in(self.r1, self.c1, er2, ec2):
            v = sh.value_at_key(k)
            if v is not None:
                yield v

    def rows(self):
        """2-D list of the used part (rows clamped to the sheet's used extent)."""
        val = self.sheet.value
        er2, ec2 = self.eff_r2(), self.eff_c2()
        return [[val(r, c) for c in range(self.c1, ec2 + 1)] for r in range(self.r1, er2 + 1)]

    def column(self, j):
        val = self.sheet.value
        c = self.c1 + j
        return [val(r, c) for r in range(self.r1, self.eff_r2() + 1)]

    def row(self, i):
        val = self.sheet.value
        r = self.r1 + i
        return [val(r, c) for c in range(self.c1, self.eff_c2() + 1)]

    def shape(self):
        return self.nrows, self.ncols


def is_num(v):
    return (type(v) is float or type(v) is int)


def is_err(v):
    return isinstance(v, XLError)


def to_2d(v):
    """RangeRef / array / scalar -> list of lists."""
    if isinstance(v, RangeRef):
        return v.rows()
    if isinstance(v, list):
        return v
    return [[v]]


def flat(v):
    if isinstance(v, RangeRef):
        return v.values()
    if isinstance(v, list):
        return (x for row in v for x in row)
    return iter((v,))


def scalar(v, ctx=None):
    """Dereference a range/array to a single value (implicit intersection)."""
    if isinstance(v, RangeRef):
        if v.is_cell():
            return v.sheet.value(v.r1, v.c1)
        if ctx is not None:
            _, row, col = ctx
            if v.c1 == v.c2 and v.r1 <= row <= v.r2:
                return v.sheet.value(row, v.c1)
            if v.r1 == v.r2 and v.c1 <= col <= v.c2:
                return v.sheet.value(v.r1, col)
        return v.sheet.value(v.r1, v.c1)
    if isinstance(v, list):
        return v[0][0] if v and v[0] else None
    if v is MISSING:
        return None
    return v


def to_num(v):
    v = scalar(v)
    t = type(v)
    if t is float or t is int:
        return v
    if v is None:
        return 0.0
    if t is bool:
        return 1.0 if v else 0.0
    if t is str:
        if v.strip() == "":
            raise errors.VALUE
        x, _ = parse_input(v)
        if is_num(x):
            return x
        raise errors.VALUE
    if isinstance(v, XLError):
        raise v
    raise errors.VALUE


def to_int(v):
    return int(math.floor(to_num(v)))


def to_str(v):
    v = scalar(v)
    if v is None:
        return ""
    if isinstance(v, str):
        return v
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, XLError):
        raise v
    return num_to_str(v)


def num_to_str(v):
    if float(v).is_integer() and abs(v) < 1e15:
        return str(int(v))
    s = f"{v:.15g}"
    if "e" in s:
        m, e = s.split("e")
        return f"{m}E{int(e):+03d}"
    return s


def to_bool(v):
    v = scalar(v)
    if isinstance(v, bool):
        return v
    if v is None:
        return False
    if is_num(v):
        return v != 0
    if isinstance(v, str):
        u = v.strip().upper()
        if u == "TRUE":
            return True
        if u == "FALSE":
            return False
        raise errors.VALUE
    if isinstance(v, XLError):
        raise v
    raise errors.VALUE


def _type_rank(v):
    if is_num(v):
        return 0
    if isinstance(v, str):
        return 1
    if isinstance(v, bool):
        return 2
    return 3


def compare(a, b):
    """Excel comparison -> -1/0/1. Blank behaves like 0, "" or FALSE."""
    if isinstance(a, XLError):
        raise a
    if isinstance(b, XLError):
        raise b
    if a is None:
        a = 0.0 if is_num(b) else ("" if isinstance(b, str) else (False if isinstance(b, bool) else 0.0))
    if b is None:
        b = 0.0 if is_num(a) else ("" if isinstance(a, str) else (False if isinstance(a, bool) else 0.0))
    ra, rb = _type_rank(a), _type_rank(b)
    if ra != rb:
        return -1 if ra < rb else 1
    if ra == 1:
        a, b = a.lower(), b.lower()
    if a == b:
        return 0
    return -1 if a < b else 1


def sort_key(v):
    """Key used by sorting: numbers < text < booleans < errors (blanks handled by caller)."""
    if is_num(v):
        return (0, v, "")
    if isinstance(v, str):
        return (1, 0, v.lower())
    if isinstance(v, bool):
        return (2, int(v), "")
    return (3, 0, "")
