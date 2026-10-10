"""Data validation: which rule covers a cell, whether a typed value passes it, the items of an
in-cell dropdown list, and building/removing rules.

Rules stay as openpyxl DataValidation objects in Sheet.xl_dv ([[dv, rects]]), the same objects
read from and written back to xlsx, so files round-trip and Excel sees exactly what Ekxel made.
Never change a dv in place: build a new one (undo snapshots keep references to the old ones)."""
from .errors import XLError
from .formula import parse, shift_formula
from .values import RangeRef, compare, is_num

TYPES = [("any", "Any value"), ("whole", "Whole number"), ("decimal", "Decimal"), ("list", "List"),
         ("date", "Date"), ("time", "Time"), ("textLength", "Text length"), ("custom", "Custom")]
OPERATORS = [("between", "between"), ("notBetween", "not between"), ("equal", "equal to"),
             ("notEqual", "not equal to"), ("greaterThan", "greater than"), ("lessThan", "less than"),
             ("greaterThanOrEqual", "greater than or equal to"), ("lessThanOrEqual", "less than or equal to")]
STYLES = [("stop", "Stop"), ("warning", "Warning"), ("information", "Information")]
TWO_VALUES = ("between", "notBetween")


# ---------------------------------------------------------------- rectangles
def subtract_rects(rects, cut):
    """`rects` minus the rectangle `cut`, as non-overlapping rectangles."""
    cr1, cc1, cr2, cc2 = cut
    out = []
    for r1, c1, r2, c2 in rects:
        if r2 < cr1 or r1 > cr2 or c2 < cc1 or c1 > cc2:
            out.append((r1, c1, r2, c2))
            continue
        if r1 < cr1:
            out.append((r1, c1, cr1 - 1, c2))
        if r2 > cr2:
            out.append((cr2 + 1, c1, r2, c2))
        mr1, mr2 = max(r1, cr1), min(r2, cr2)
        if c1 < cc1:
            out.append((mr1, c1, mr2, cc1 - 1))
        if c2 > cc2:
            out.append((mr1, cc2 + 1, mr2, c2))
    return out


def _contains(rects, r, c):
    return any(r1 <= r <= r2 and c1 <= c <= c2 for r1, c1, r2, c2 in rects)


# ---------------------------------------------------------------- lookup
def rule_at(sheet, r, c):
    """(dv, rects) covering cell (r, c), or (None, None). Later rules win, like Excel's last-applied."""
    for dv, rects in reversed(sheet.xl_dv):
        if _contains(rects, r, c):
            return dv, rects
    return None, None


def _anchor(rects):
    return min(x[0] for x in rects), min(x[1] for x in rects)


def _eval(sheet, text, rects, r, c, as_range=False):
    """Evaluate a rule formula (Excel stores them without '=') relative to the rule's top-left cell."""
    if text is None or str(text).strip() == "":
        return None
    text = str(text)
    if not text.startswith("="):
        text = "=" + text
    ar, ac = _anchor(rects)
    try:
        ast = parse(shift_formula(text, r - ar, c - ac))
    except Exception:  # noqa: BLE001 - a bad rule formula just doesn't validate
        return None
    ev = sheet.wb.evaluator
    if as_range:
        try:
            return ev.ev(ast, (sheet, r, c))
        except XLError as e:
            return e
        except Exception:  # noqa: BLE001
            return None
    return ev.run(ast, (sheet, r, c))


# ---------------------------------------------------------------- dropdown lists
def list_items(sheet, dv, rects, r, c):
    """Display texts for a List rule: '"a,b,c"' literals, a cell range, or a name/formula giving one."""
    src = str(dv.formula1 or "").strip()
    if not src:
        return []
    if src.startswith('"') and src.endswith('"'):
        return [t.strip() for t in src[1:-1].split(",") if t.strip() != ""]
    got = _eval(sheet, src, rects, r, c, as_range=True)
    from .numfmt import format_value
    out = []
    if isinstance(got, RangeRef):
        sh = got.sheet
        for rr in range(got.r1, got.eff_r2() + 1):
            for cc in range(got.c1, got.eff_c2() + 1):
                v = sh.value(rr, cc)
                if v is None or v == "" or isinstance(v, XLError):
                    continue
                out.append(format_value(v, sh.style(rr, cc).numfmt)[0] if is_num(v) else str(v))
    elif isinstance(got, list):
        for row in got:
            for v in row:
                if v not in (None, "") and not isinstance(v, XLError):
                    out.append(str(v))
    elif got is not None and not isinstance(got, XLError):
        out.extend(t.strip() for t in str(got).split(",") if t.strip())
    return out


def shows_dropdown(dv):
    # OOXML's showDropDown is inverted: True means "hide the in-cell arrow"
    return dv is not None and dv.type == "list" and not dv.showDropDown


# ---------------------------------------------------------------- checking
def check(sheet, dv, rects, r, c, value):
    """True if `value` (the cell's new value) passes rule `dv` at (r, c)."""
    kind = dv.type or "any"
    if kind == "any":
        return True
    blank = value is None or (isinstance(value, str) and value == "")
    if blank:
        return bool(dv.allow_blank)
    if isinstance(value, XLError):
        return False
    if kind == "custom":
        res = _eval(sheet, dv.formula1, rects, r, c)
        if isinstance(res, XLError) or res is None:
            return False
        from .values import to_bool
        try:
            return to_bool(res)
        except XLError:
            return False
    if kind == "list":
        items = [t.lower() for t in list_items(sheet, dv, rects, r, c)]
        from .numfmt import format_value
        cands = {str(value).lower()}
        if is_num(value):
            cands.add(format_value(value, "General")[0].lower())
        return any(x in items for x in cands)
    if kind == "textLength":
        v = float(len(str(value)))
    else:
        if not is_num(value) or isinstance(value, bool):
            return False
        v = float(value)
        if kind == "whole" and v != int(v):
            return False
        if kind == "time" and not (0 <= v < 1):
            return False
    lo = _eval(sheet, dv.formula1, rects, r, c)
    hi = _eval(sheet, dv.formula2, rects, r, c) if (dv.operator or "between") in TWO_VALUES else None
    op = dv.operator or "between"
    if not is_num(lo) or (op in TWO_VALUES and not is_num(hi)):
        return True  # an unreadable limit: don't block typing over it
    if op in TWO_VALUES:
        a, b = sorted((lo, hi))
        inside = a <= v <= b
        return inside if op == "between" else not inside
    x = compare(v, lo)
    return {"equal": x == 0, "notEqual": x != 0, "greaterThan": x > 0, "lessThan": x < 0,
            "greaterThanOrEqual": x >= 0, "lessThanOrEqual": x <= 0}.get(op, True)


def describe_failure(dv):
    """Title and text for the error box: the rule's own, or Excel's default wording."""
    title = (dv.errorTitle or "").strip() or "Microsoft Excel"
    text = (dv.error or "").strip() or ("This value doesn't match the data validation restrictions "
                                         "defined for this cell.")
    return title, text


# ---------------------------------------------------------------- building
def make(kind, operator=None, formula1=None, formula2=None, allow_blank=True, dropdown=True,
         prompt_title="", prompt="", show_prompt=True, error_title="", error="", style="stop",
         show_error=True):
    """A new openpyxl DataValidation. Formulas are given as typed ('10', '=$A$1', 'a,b,c')."""
    from openpyxl.worksheet.datavalidation import DataValidation

    def clean(f):
        if f is None:
            return None
        f = str(f).strip()
        if f == "":
            return None
        return f[1:] if f.startswith("=") else f

    f1 = formula1
    if kind == "list" and f1 is not None:
        s = str(f1).strip()
        f1 = clean(s) if s.startswith("=") else '"' + ",".join(t.strip() for t in s.split(",")) + '"'
    else:
        f1 = clean(f1)
    two = kind not in ("any", "list", "custom") and (operator or "between") in TWO_VALUES
    dv = DataValidation(type=None if kind == "any" else kind,
                        operator=None if kind in ("any", "list", "custom") else (operator or "between"),
                        formula1=f1, formula2=clean(formula2) if two else None,
                        allow_blank=bool(allow_blank),
                        showDropDown=(kind == "list" and not dropdown),
                        showInputMessage=bool(show_prompt), promptTitle=prompt_title or None,
                        prompt=prompt or None, showErrorMessage=bool(show_error),
                        errorTitle=error_title or None, error=error or None, errorStyle=style or "stop")
    return dv


def formula_text(f, kind):
    """A stored rule formula back into what the dialog shows."""
    if f is None:
        return ""
    s = str(f)
    if kind == "list" and s.startswith('"') and s.endswith('"'):
        return s[1:-1]
    return "=" + s if not _looks_literal(s) else s


def _looks_literal(s):
    try:
        float(s)
        return True
    except ValueError:
        return False


def apply_to(sheet, rects, dv):
    """New xl_dv list: `rects` taken out of every existing rule, then `dv` (None = clear) added."""
    out = []
    for old, old_rects in sheet.xl_dv:
        left = old_rects
        for cut in rects:
            left = subtract_rects(left, cut)
        if left:
            out.append([old, left])
    if dv is not None:
        out.append([dv, [tuple(x) for x in rects]])
    return out
