"""Formula evaluation."""
import math

from . import errors
from .errors import XLError
from .functions import FUNCS
from .values import (MISSING, RangeRef, compare, scalar, to_2d, to_num,
                     to_str)


class Evaluator:
    def __init__(self, wb):
        self.wb = wb

    def sheet_for(self, name, ctx):
        if name is None:
            return ctx[0]
        sh = self.wb.get_sheet(name)
        if sh is None:
            raise errors.REF
        return sh

    def run(self, ast, ctx):
        """Evaluate a formula AST to a single cell value."""
        try:
            v = self.ev(ast, ctx)
            if isinstance(v, RangeRef):
                v = scalar(v, ctx)
            elif isinstance(v, list):
                v = v[0][0] if v and v[0] else None
            if v is None or v is MISSING:
                return 0.0
            t = type(v)
            if t is int:
                return float(v)
            if t is float and (math.isnan(v) or math.isinf(v)):
                return errors.NUM
            return v
        except XLError as e:
            e.__traceback__ = None
            return e
        except ZeroDivisionError:
            return errors.DIV0
        except (OverflowError, ValueError):
            return errors.NUM
        except RecursionError:
            return errors.CALC
        except Exception:  # a bug in a function must not break recalculation or painting
            return errors.VALUE

    def ev(self, node, ctx):
        kind = node[0]
        if kind == "num" or kind == "str" or kind == "bool":
            return node[1]
        if kind == "cell":
            sh = self.sheet_for(node[1], ctx)
            return sh.value(node[2], node[3])
        if kind == "range":
            sh = self.sheet_for(node[1], ctx)
            return RangeRef(sh, node[2], node[3], node[4], node[5])
        if kind == "func":
            return self.call(node, ctx)
        if kind == "bin":
            return self.binop(node[1], self.ev(node[2], ctx), self.ev(node[3], ctx), ctx)
        if kind == "neg":
            v = self.ev(node[1], ctx)
            if isinstance(v, (RangeRef, list)):
                return [[_safe(lambda x: -to_num(x), x) for x in row] for row in to_2d(v)]
            return -to_num(v)
        if kind == "pct":
            v = self.ev(node[1], ctx)
            if isinstance(v, (RangeRef, list)):
                return [[_safe(lambda x: to_num(x) / 100, x) for x in row] for row in to_2d(v)]
            return to_num(v) / 100
        if kind == "err":
            raise node[1]
        if kind == "missing":
            return MISSING
        if kind == "array":
            return node[1]
        if kind == "name":
            return self.name_ref(node[1], ctx)
        raise errors.VALUE

    def name_ref(self, name, ctx):
        """Defined name (from an xlsx file) -> RangeRef or constant."""
        text = getattr(self.wb, "names", {}).get(name.upper())
        if text is None:
            raise errors.NAME
        from .functions import FUNCS as _F
        if "!" in text or ":" in text or "$" in text:
            parts = [p.strip() for p in text.split(",")]
            return _F["INDIRECT"][0](ctx, parts[0].replace("$", ""))
        from .numfmt import parse_input
        v, _ = parse_input(text.lstrip("="))
        return v

    def arg(self, node, ctx):
        """Evaluate a function argument; cell references stay references."""
        if node[0] == "cell":
            sh = self.sheet_for(node[1], ctx)
            return RangeRef(sh, node[2], node[3], node[2], node[3])
        return self.ev(node, ctx)

    def call(self, node, ctx):
        entry = FUNCS.get(node[1])
        if entry is None:
            raise errors.NAME
        f, kind, minargs, maxargs = entry
        nodes = node[2]
        if kind == "lazy":
            if len(nodes) < minargs or (maxargs is not None and len(nodes) > maxargs):
                raise errors.VALUE
            return f(self, ctx, nodes)
        if kind == "safe":
            args = []
            for a in nodes:
                try:
                    args.append(self.arg(a, ctx))
                except XLError as e:
                    e.__traceback__ = None
                    args.append(e)
        else:
            args = [self.arg(a, ctx) for a in nodes]
        try:
            return f(ctx, *args)
        except TypeError as e:
            if "positional argument" in str(e) or "required" in str(e):
                raise errors.VALUE
            raise

    def binop(self, op, a, b, ctx):
        if isinstance(a, (RangeRef, list)) or isinstance(b, (RangeRef, list)):
            A, B = to_2d(a), to_2d(b)
            ra, ca = len(A), (len(A[0]) if A else 0)
            rb, cb = len(B), (len(B[0]) if B else 0)
            rows = max(ra, rb)
            cols = max(ca, cb)
            out = []
            for i in range(rows):
                row = []
                for j in range(cols):
                    ia = 0 if ra == 1 else i
                    ja = 0 if ca == 1 else j
                    ib = 0 if rb == 1 else i
                    jb = 0 if cb == 1 else j
                    if ia >= ra or ja >= ca or ib >= rb or jb >= cb:
                        row.append(errors.NA)
                        continue
                    try:
                        row.append(scalar_op(op, A[ia][ja], B[ib][jb]))
                    except XLError as e:
                        e.__traceback__ = None
                        row.append(e)
                    except ZeroDivisionError:
                        row.append(errors.DIV0)
                out.append(row)
            return out
        return scalar_op(op, a, b)


def _safe(f, x):
    try:
        return f(x)
    except XLError as e:
        e.__traceback__ = None
        return e


def scalar_op(op, a, b):
    if op == "+":
        return to_num(a) + to_num(b)
    if op == "-":
        return to_num(a) - to_num(b)
    if op == "*":
        return to_num(a) * to_num(b)
    if op == "/":
        x, y = to_num(a), to_num(b)
        if y == 0:
            raise errors.DIV0
        return x / y
    if op == "^":
        x, y = to_num(a), to_num(b)
        if x == 0 and y < 0:
            raise errors.DIV0
        if x < 0 and not float(y).is_integer():
            raise errors.NUM
        try:
            return float(x ** y)
        except OverflowError:
            raise errors.NUM
    if op == "&":
        return to_str(a) + to_str(b)
    a = scalar(a)
    b = scalar(b)
    c = compare(a, b)
    if op == "=":
        return c == 0
    if op == "<>":
        return c != 0
    if op == "<":
        return c < 0
    if op == ">":
        return c > 0
    if op == "<=":
        return c <= 0
    if op == ">=":
        return c >= 0
    raise errors.VALUE
