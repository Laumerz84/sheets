"""Display of conditional formatting read from xlsx files (rules are kept in the
file untouched; Sheets only renders them)."""
import colorsys

from .errors import XLError
from .formula import ParseError, parse, shift_formula
from .values import compare, is_num, to_bool


class CFRule:
    __slots__ = ("rects", "kind", "op", "formulas", "fill", "color", "bold", "italic",
                 "scale", "bar", "text", "rank", "percent", "bottom", "above", "stop", "priority", "_asts",
                 "xl_rule")

    def __init__(self, rects, kind, priority=0):
        self.rects = rects
        self.kind = kind
        self.op = None
        self.formulas = []
        self.fill = self.color = None
        self.bold = self.italic = None
        self.scale = None      # [(cfvo_type, value, color)]
        self.bar = None        # (color, cfvo_min, cfvo_max)
        self.text = None
        self.rank = 10
        self.percent = False
        self.bottom = False
        self.above = True
        self.stop = False
        self.priority = priority
        self._asts = {}
        self.xl_rule = None     # the openpyxl Rule, written back on save

    def contains(self, r, c):
        for r1, c1, r2, c2 in self.rects:
            if r1 <= r <= r2 and c1 <= c <= c2:
                return True
        return False

    def anchor(self):
        return min(x[0] for x in self.rects), min(x[1] for x in self.rects)


def _interp(c1, c2, t):
    a = [int(c1[i:i + 2], 16) for i in (1, 3, 5)]
    b = [int(c2[i:i + 2], 16) for i in (1, 3, 5)]
    return "#%02X%02X%02X" % tuple(round(x + (y - x) * t) for x, y in zip(a, b))


class CFEngine:
    """Evaluates rules for cells with per-paint caching."""

    def __init__(self, sheet):
        self.sheet = sheet
        self.cache = {}
        self.range_stats = {}

    def clear(self):
        self.cache = {}
        self.range_stats = {}

    def _eval(self, rule, idx, r, c):
        """Evaluate rule.formulas[idx] relative to the rule's anchor for cell (r, c)."""
        ar, ac = rule.anchor()
        dr, dc = r - ar, c - ac
        key = (idx, dr, dc)
        ast = rule._asts.get(key)
        if ast is None:
            text = rule.formulas[idx]
            if not text.startswith("="):
                text = "=" + text
            try:
                ast = parse(shift_formula(text, dr, dc))
            except (ParseError, Exception):
                ast = ("err", None)
            if len(rule._asts) < 20000:
                rule._asts[key] = ast
        if ast[0] == "err" and ast[1] is None:
            return None
        return self.sheet.wb.evaluator.run(ast, (self.sheet, r, c))

    def _numbers(self, rule):
        st = self.range_stats.get(id(rule))
        if st is None:
            sh = self.sheet
            nums = []
            texts = {}
            for r1, c1, r2, c2 in rule.rects:
                r2 = min(r2, sh.max_row)
                c2 = min(c2, sh.max_col)
                for r in range(r1, r2 + 1):
                    for c in range(c1, c2 + 1):
                        v = sh.value(r, c)
                        if is_num(v):
                            nums.append(v)
                        if v is not None and v != "":
                            t = str(v).lower()
                            texts[t] = texts.get(t, 0) + 1
            nums.sort()
            st = (nums, texts)
            self.range_stats[id(rule)] = st
        return st

    def _cfvo(self, rule, kind, val, nums, default):
        if not nums:
            return default
        try:
            if kind == "min":
                return nums[0]
            if kind == "max":
                return nums[-1]
            if kind == "num":
                return float(val)
            if kind == "percent":
                lo, hi = nums[0], nums[-1]
                return lo + (hi - lo) * float(val) / 100
            if kind == "percentile":
                p = float(val) / 100
                pos = p * (len(nums) - 1)
                i = int(pos)
                j = min(i + 1, len(nums) - 1)
                return nums[i] + (nums[j] - nums[i]) * (pos - i)
        except (TypeError, ValueError):
            pass
        return default

    def style_for(self, rules, r, c):
        """-> dict of overrides (fill, color, bold, italic, bar) or None."""
        k = (r, c)
        if k in self.cache:
            return self.cache[k]
        out = None
        v = None
        got_v = False
        for rule in rules:
            if not rule.contains(r, c):
                continue
            if not got_v:
                v = self.sheet.value(r, c)
                got_v = True
            try:
                hit = self._match(rule, v, r, c)
            except XLError:
                hit = False
            if not hit:
                continue
            if out is None:
                out = {}
            if rule.kind == "colorScale":
                out.setdefault("fill", hit)
            elif rule.kind == "dataBar":
                out.setdefault("bar", hit)
            else:
                for f in ("fill", "color", "bold", "italic"):
                    val = getattr(rule, f)
                    if val is not None:
                        out.setdefault(f, val)
            if rule.stop:
                break
        if len(self.cache) > 50000:
            self.cache = {}
        self.cache[k] = out
        return out

    def _match(self, rule, v, r, c):
        kind = rule.kind
        if kind == "cellIs":
            if v is None:
                v = 0.0
            if isinstance(v, XLError):
                return False
            vals = [self._eval(rule, i, r, c) for i in range(len(rule.formulas))]
            if not vals or any(isinstance(x, XLError) or x is None for x in vals):
                return False
            op = rule.op
            if op in ("between", "notBetween") and len(vals) >= 2:
                lo, hi = sorted(vals[:2], key=lambda x: (0, x) if is_num(x) else (1, str(x)))
                inside = compare(v, lo) >= 0 and compare(v, hi) <= 0
                return inside if op == "between" else not inside
            x = compare(v, vals[0])
            return {"equal": x == 0, "notEqual": x != 0, "greaterThan": x > 0, "lessThan": x < 0,
                    "greaterThanOrEqual": x >= 0, "lessThanOrEqual": x <= 0}.get(op, False)
        if kind == "expression":
            res = self._eval(rule, 0, r, c)
            if res is None or isinstance(res, XLError):
                return False
            try:
                return to_bool(res)
            except XLError:
                return False
        if kind in ("containsText", "notContainsText", "beginsWith", "endsWith"):
            if v is None or rule.text is None:
                return kind == "notContainsText"
            t, n = str(v).lower(), rule.text.lower()
            return {"containsText": n in t, "notContainsText": n not in t,
                    "beginsWith": t.startswith(n), "endsWith": t.endswith(n)}[kind]
        if kind in ("containsBlanks", "notContainsBlanks"):
            blank = v is None or (isinstance(v, str) and not v.strip())
            return blank if kind == "containsBlanks" else not blank
        if kind in ("containsErrors", "notContainsErrors"):
            return isinstance(v, XLError) == (kind == "containsErrors")
        if kind == "colorScale" and is_num(v) and rule.scale:
            nums, _ = self._numbers(rule)
            pts = [(self._cfvo(rule, t, val, nums, 0.0), col) for t, val, col in rule.scale]
            if len(pts) < 2:
                return False
            if v <= pts[0][0]:
                return pts[0][1]
            for (a, ca), (b, cb) in zip(pts, pts[1:]):
                if v <= b:
                    return _interp(ca, cb, 0 if b == a else (v - a) / (b - a))
            return pts[-1][1]
        if kind == "dataBar" and is_num(v) and rule.bar:
            nums, _ = self._numbers(rule)
            color, lo_t, hi_t = rule.bar
            lo = self._cfvo(rule, lo_t[0], lo_t[1], nums, 0.0)
            hi = self._cfvo(rule, hi_t[0], hi_t[1], nums, 1.0)
            if hi == lo:
                return (1.0, color)
            return (max(0.0, min(1.0, (v - lo) / (hi - lo))), color)
        if kind == "top10" and is_num(v):
            nums, _ = self._numbers(rule)
            n = rule.rank
            if rule.percent:
                n = max(1, int(len(nums) * rule.rank / 100))
            if not nums:
                return False
            if rule.bottom:
                return v <= nums[min(n, len(nums)) - 1]
            return v >= nums[-min(n, len(nums))]
        if kind == "aboveAverage" and is_num(v):
            nums, _ = self._numbers(rule)
            if not nums:
                return False
            avg = sum(nums) / len(nums)
            return v > avg if rule.above else v < avg
        if kind in ("duplicateValues", "uniqueValues") and v not in (None, ""):
            _, texts = self._numbers(rule)
            cnt = texts.get(str(v).lower(), 0)
            return cnt > 1 if kind == "duplicateValues" else cnt == 1
        return False


def load_rules(ws, color_fn):
    """openpyxl worksheet -> ([CFRule] sorted by priority, complete).
    `complete` is False when some rule couldn't be read; the file's own
    formatting is then left untouched on save."""
    from .refs import parse_range
    out = []
    complete = True
    try:
        entries = list(ws.conditional_formatting)
    except Exception:
        return out, False
    for cf in entries:
        rects = []
        for rng in str(cf.sqref).split():
            b = parse_range(rng)
            if b:
                rects.append(b)
            else:
                complete = False
        if not rects:
            complete = False
            continue
        for rule in cf.rules:
            kind = rule.type
            cr = CFRule(list(rects), kind, rule.priority or 0)
            cr.xl_rule = rule
            cr.op = rule.operator
            cr.formulas = [str(f) for f in (rule.formula or [])]
            cr.text = rule.text
            cr.stop = bool(rule.stopIfTrue)
            if rule.rank is not None:
                cr.rank = int(rule.rank)
            cr.percent = bool(rule.percent)
            cr.bottom = bool(rule.bottom)
            cr.above = rule.aboveAverage is not False
            dxf = rule.dxf
            if dxf is not None:
                if dxf.fill is not None:
                    fill = dxf.fill
                    col = color_fn(getattr(fill, "bgColor", None)) or color_fn(getattr(fill, "fgColor", None))
                    if getattr(fill, "fill_type", None) not in (None, "none") or col:
                        cr.fill = col
                if dxf.font is not None:
                    cr.color = color_fn(dxf.font.color)
                    if dxf.font.b is not None:
                        cr.bold = bool(dxf.font.b)
                    if dxf.font.i is not None:
                        cr.italic = bool(dxf.font.i)
            if kind == "colorScale" and rule.colorScale is not None:
                cs = rule.colorScale
                cols = [color_fn(c) or "#FFFFFF" for c in cs.color]
                cr.scale = [(v.type, v.val, col) for v, col in zip(cs.cfvo, cols)]
            if kind == "dataBar" and rule.dataBar is not None:
                db = rule.dataBar
                cfvo = list(db.cfvo) + [None, None]
                lo = (cfvo[0].type, cfvo[0].val) if cfvo[0] is not None else ("min", None)
                hi = (cfvo[1].type, cfvo[1].val) if cfvo[1] is not None else ("max", None)
                cr.bar = (color_fn(db.color) or "#638EC6", lo, hi)
            if kind in ("containsText", "notContainsText", "beginsWith", "endsWith") and not cr.text and cr.formulas:
                cr.text = None
            out.append(cr)
    out.sort(key=lambda x: x.priority)
    return out, complete
