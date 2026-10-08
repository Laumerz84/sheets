"""Excel-style number formats: rendering values and parsing typed input."""
import datetime as dt
import math
import re
from functools import lru_cache

from .errors import XLError

EPOCH = dt.datetime(1899, 12, 30)

MONTHS = ["January", "February", "March", "April", "May", "June", "July",
          "August", "September", "October", "November", "December"]
DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]

COLORS = {
    "black": "#000000", "white": "#FFFFFF", "red": "#FF0000", "green": "#008000",
    "blue": "#0000FF", "yellow": "#FFFF00", "magenta": "#FF00FF", "cyan": "#00FFFF",
}

# Built-in Excel number format ids -> format strings (used by xlsx/xls readers)
BUILTIN_FORMATS = {
    0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00",
    5: '"$"#,##0_);("$"#,##0)', 6: '"$"#,##0_);[Red]("$"#,##0)',
    7: '"$"#,##0.00_);("$"#,##0.00)', 8: '"$"#,##0.00_);[Red]("$"#,##0.00)',
    9: "0%", 10: "0.00%", 11: "0.00E+00", 12: "# ?/?", 13: "# ??/??",
    14: "m/d/yyyy", 15: "d-mmm-yy", 16: "d-mmm", 17: "mmm-yy", 18: "h:mm AM/PM",
    19: "h:mm:ss AM/PM", 20: "h:mm", 21: "h:mm:ss", 22: "m/d/yyyy h:mm",
    37: "#,##0_);(#,##0)", 38: "#,##0_);[Red](#,##0)", 39: "#,##0.00_);(#,##0.00)",
    40: "#,##0.00_);[Red](#,##0.00)", 45: "mm:ss", 46: "[h]:mm:ss", 47: "mmss.0",
    48: "##0.0E+0", 49: "@",
}


# ---------------------------------------------------------------- dates

def serial_to_datetime(x):
    x = float(x)
    if x < 0:
        raise ValueError("negative date")
    if x < 60:
        return dt.datetime(1899, 12, 31) + dt.timedelta(days=x)
    if x < 61:  # Excel's fictional 1900-02-29
        return dt.datetime(1900, 2, 28) + dt.timedelta(days=x - 60)
    return EPOCH + dt.timedelta(days=x)


def datetime_to_serial(d):
    if isinstance(d, dt.datetime):
        delta = d - EPOCH
    elif isinstance(d, dt.date):
        delta = dt.datetime(d.year, d.month, d.day) - EPOCH
    elif isinstance(d, dt.time):
        return (d.hour * 3600 + d.minute * 60 + d.second + d.microsecond / 1e6) / 86400
    elif isinstance(d, dt.timedelta):
        return d.total_seconds() / 86400
    else:
        raise TypeError(d)
    serial = delta.days + (delta.seconds + delta.microseconds / 1e6) / 86400
    if serial < 61:
        serial -= 1
    return float(serial) if serial % 1 else float(int(serial))


# ---------------------------------------------------------------- general

def format_general(v):
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if v == 0:
        return "0"
    a = abs(v)
    if float(v).is_integer() and a < 1e11:
        return str(int(v))
    if 1e-9 <= a < 1e11 and not (a < 1e-4):
        for p in range(10, 0, -1):
            s = f"{v:.{p}g}"
            if "e" in s:
                break
            if len(s.lstrip("-")) <= 11:
                return s
    return _sci_general(v)


def _sci_general(v, maxlen=11):
    for p in range(5, -1, -1):
        s = f"{v:.{p}E}"
        mant, exp = s.split("E")
        if "." in mant:
            mant = mant.rstrip("0").rstrip(".")
        e = int(exp)
        s = f"{mant}E{'+' if e >= 0 else '-'}{abs(e):02d}"
        if len(s.lstrip("-")) <= maxlen:
            return s
    return s


def format_general_fit(v, max_chars):
    """General format squeezed to at most max_chars (as Excel does in narrow columns)."""
    s = format_general(v)
    if len(s) <= max_chars:
        return s
    if float(v).is_integer() and abs(v) < 1e11:
        return _sci_general(v, max_chars)
    if "E" not in s:
        intlen = len(str(int(abs(v)))) + (1 if v < 0 else 0)
        dec = max_chars - intlen - 1
        if dec >= 1:
            s2 = f"{v:.{dec}f}".rstrip("0").rstrip(".")
            if len(s2) <= max_chars:
                return s2
        if intlen <= max_chars:
            return str(round(v))
    return _sci_general(v, max_chars)


# ---------------------------------------------------------------- format parsing

def _split_sections(fmt):
    sections, cur, i, n = [], [], 0, len(fmt)
    while i < n:
        ch = fmt[i]
        if ch == '"':
            j = fmt.find('"', i + 1)
            j = n - 1 if j < 0 else j
            cur.append(fmt[i:j + 1])
            i = j + 1
            continue
        if ch == "\\" and i + 1 < n:
            cur.append(fmt[i:i + 2])
            i += 2
            continue
        if ch == "[":
            j = fmt.find("]", i)
            j = n - 1 if j < 0 else j
            cur.append(fmt[i:j + 1])
            i = j + 1
            continue
        if ch == ";":
            sections.append("".join(cur))
            cur = []
        else:
            cur.append(ch)
        i += 1
    sections.append("".join(cur))
    return sections


_DATE_LETTERS = set("ymdhsYMDHS")
_COND_RE = re.compile(r"^(<=|>=|<>|<|>|=)(-?\d+(?:\.\d*)?)$")


class Section:
    __slots__ = ("tokens", "color", "is_date", "is_text", "is_general", "cond", "has_ampm")

    def __init__(self):
        self.tokens = []
        self.color = None
        self.is_date = False
        self.is_text = False
        self.is_general = False
        self.cond = None
        self.has_ampm = False


def _tokenize(sec):
    s = Section()
    toks = s.tokens
    i, n = 0, len(sec)
    while i < n:
        ch = sec[i]
        if ch == '"':
            j = sec.find('"', i + 1)
            j = n if j < 0 else j
            toks.append(("lit", sec[i + 1:j]))
            i = j + 1
        elif ch == "\\":
            toks.append(("lit", sec[i + 1:i + 2]))
            i += 2
        elif ch == "_":
            toks.append(("lit", " "))
            i += 2
        elif ch == "*":
            i += 2
        elif ch == "[":
            j = sec.find("]", i)
            j = n if j < 0 else j
            content = sec[i + 1:j]
            low = content.lower()
            if low in COLORS:
                s.color = COLORS[low]
            elif low.startswith("color"):
                pass
            elif content.startswith("$"):
                sym = content[1:].split("-", 1)[0]
                if sym:
                    toks.append(("lit", sym))
            elif low and set(low) <= {"h"} or set(low) <= {"m"} or set(low) <= {"s"}:
                if low:
                    toks.append(("elapsed", low[0], len(low)))
                    s.is_date = True
            else:
                m = _COND_RE.match(content)
                if m:
                    s.cond = (m.group(1), float(m.group(2)))
            i = j + 1
        elif sec[i:i + 7].lower() == "general":
            toks.append(("general",))
            s.is_general = True
            i += 7
        elif sec[i:i + 5].upper() == "AM/PM":
            toks.append(("ampm", "AM/PM"))
            s.has_ampm = True
            i += 5
        elif sec[i:i + 3].upper() == "A/P":
            toks.append(("ampm", "A/P"))
            s.has_ampm = True
            i += 3
        elif ch == "@":
            toks.append(("text",))
            s.is_text = True
            i += 1
        elif ch in "0#?.,%":
            toks.append(("n", ch))
            i += 1
        elif ch in "Ee" and i + 1 < n and sec[i + 1] in "+-":
            toks.append(("exp", sec[i + 1]))
            i += 2
        elif ch in _DATE_LETTERS:
            j = i
            while j < n and sec[j].lower() == ch.lower():
                j += 1
            toks.append(("d", ch.lower() * (j - i)))
            s.is_date = True
            i = j
        elif ch == "/" and not s.is_date and any(t[0] == "n" for t in toks):
            toks.append(("lit", "/"))  # fraction formats degrade to decimals
            i += 1
        else:
            toks.append(("lit", ch))
            i += 1
    if s.is_date:
        _resolve_minutes(toks)
    return s


def _resolve_minutes(toks):
    """'m' after an hour or before a second means minutes."""
    idx = [i for i, t in enumerate(toks) if t[0] in ("d", "elapsed")]
    for pos, i in enumerate(idx):
        t = toks[i]
        if t[0] == "d" and t[1][0] == "m" and len(t[1]) <= 2:
            prev = toks[idx[pos - 1]] if pos > 0 else None
            nxt = toks[idx[pos + 1]] if pos + 1 < len(idx) else None
            if (prev and prev[0] in ("d", "elapsed") and prev[1][0] == "h") or \
               (nxt and nxt[0] in ("d", "elapsed") and nxt[1][0] == "s"):
                toks[i] = ("min", t[1])


@lru_cache(maxsize=512)
def compile_format(fmt):
    if not fmt or fmt.lower() == "general":
        return None
    return [_tokenize(s) for s in _split_sections(fmt)]


def is_date_format(fmt):
    secs = compile_format(fmt)
    return bool(secs) and secs[0].is_date


def is_text_format(fmt):
    secs = compile_format(fmt)
    return bool(secs) and len(secs) == 1 and secs[0].is_text


# ---------------------------------------------------------------- rendering

def format_value(v, fmt="General"):
    """Return (text, color_hex_or_None)."""
    if v is None:
        return "", None
    if isinstance(v, XLError):
        return v.code, None
    if isinstance(v, bool):
        return ("TRUE" if v else "FALSE"), None
    secs = compile_format(fmt)
    if isinstance(v, str):
        if secs:
            sec = secs[3] if len(secs) >= 4 else (secs[0] if secs[0].is_text else None)
            if sec is not None and sec.is_text:
                out = []
                for t in sec.tokens:
                    if t[0] == "text":
                        out.append(v)
                    elif t[0] == "lit":
                        out.append(t[1])
                return "".join(out), sec.color
        return v, None
    if not secs:
        return format_general(v), None
    sec, neg_section = _pick_section(secs, v)
    if sec.is_text and not any(t[0] in ("n", "general", "d") for t in sec.tokens):
        return format_general(v), sec.color
    try:
        if sec.is_date:
            return _render_date(sec, v), sec.color
        return _render_number(sec, abs(v) if neg_section else v), sec.color
    except (ValueError, OverflowError):
        return "#" * 8, None


def _pick_section(secs, v):
    numeric = [s for s in secs[:3] if not (s.is_text and not s.is_date and not any(t[0] == "n" for t in s.tokens))] or secs
    if any(s.cond for s in numeric[:2]):
        for s in numeric[:2]:
            if s.cond and _cond_ok(s.cond, v):
                return s, s is not numeric[0] and v < 0
        return numeric[min(2, len(numeric) - 1)], False
    if len(numeric) == 1:
        return numeric[0], False
    if v > 0 or (v == 0 and len(numeric) == 2):
        return numeric[0], False
    if v < 0:
        return numeric[1], True
    return numeric[2], False


def _cond_ok(cond, v):
    op, x = cond
    return {"<": v < x, ">": v > x, "=": v == x, "<=": v <= x, ">=": v >= x, "<>": v != x}[op]


def _render_number(sec, v):
    toks = sec.tokens
    if sec.is_general and not any(t[0] == "n" for t in toks):
        out = []
        for t in toks:
            if t[0] == "general":
                out.append(format_general(v))
            elif t[0] == "lit":
                out.append(t[1])
        return "".join(out)

    slash = next((i for i, t in enumerate(toks) if t == ("lit", "/")), None)
    if slash is not None and slash > 0 and toks[slash - 1][0] == "n" and toks[slash - 1][1] in "0#?":
        return _render_fraction(toks, slash, v)

    ntok = [(i, t[1]) for i, t in enumerate(toks) if t[0] == "n"]
    exp_i = next((i for i, t in enumerate(toks) if t[0] == "exp"), None)
    dot_i = next((i for i, ch in ntok if ch == "." and (exp_i is None or i < exp_i)), None)

    digit = "0#?"
    int_pos = [i for i, ch in ntok if ch in digit and (dot_i is None or i < dot_i) and (exp_i is None or i < exp_i)]
    frac_pos = [i for i, ch in ntok if ch in digit and dot_i is not None and i > dot_i and (exp_i is None or i < exp_i)]
    exp_pos = [i for i, ch in ntok if ch in digit and exp_i is not None and i > exp_i]

    pct = sum(1 for _, ch in ntok if ch == "%")
    v = v * (100 ** pct)

    # thousands separators and scaling commas (commas after the last digit placeholder,
    # or right before the decimal point, divide by 1000 each)
    grouping = False
    if int_pos:
        last_int = int_pos[-1]
        last_digit = max(int_pos + frac_pos)
        for i, ch in ntok:
            if ch == ",":
                if int_pos[0] < i < last_int:
                    grouping = True
                elif (i > last_digit or (i > last_int and dot_i is not None and i < dot_i)) and \
                        (exp_i is None or i < exp_i):
                    v /= 1000

    frac_spec = "".join(toks[i][1] for i in frac_pos)
    ndec = len(frac_spec)
    neg = v < 0
    a = abs(v)

    exp_str = None
    if exp_i is not None:
        n_int = max(1, len(int_pos))
        if a == 0:
            e = 0
        else:
            e = math.floor(math.log10(a))
            if n_int > 1 and "#" in "".join(toks[i][1] for i in int_pos):
                e = e - (e % n_int)
            else:
                e = e - (n_int - 1)
        mant = a / (10 ** e) if a else 0.0
        if round(mant, ndec) >= 10 ** n_int:
            mant /= 10
            e += 1
        a = mant
        sign = "-" if e < 0 else ("+" if toks[exp_i][1] == "+" else "")
        exp_digits = str(abs(e)).rjust(sum(1 for i in exp_pos if toks[i][1] == "0"), "0")
        exp_str = "E" + sign + exp_digits

    s = f"{a:.{ndec}f}"
    int_str, _, frac_str = s.partition(".")
    if ndec:
        req = len(frac_spec.rstrip("#?"))
        keep = len(frac_str.rstrip("0"))
        keep = max(keep, req)
        tail = frac_str[keep:]
        frac_str = frac_str[:keep] + "".join(" " if frac_spec[k] == "?" else "" for k in range(keep, ndec))
        _ = tail
    min_int = sum(1 for i in int_pos if toks[i][1] == "0")
    if int_str == "0" and min_int == 0:
        int_str = ""
    int_str = int_str.rjust(min_int, "0")
    if grouping and int_str:
        int_str = f"{int(int_str):,}".rjust(len(int_str), "0") if int_str.isdigit() else int_str

    # are the integer placeholders interrupted by literals? (e.g. phone numbers)
    contiguous = not int_pos or all(
        toks[k][0] == "n" for k in range(int_pos[0], int_pos[-1] + 1))

    out = []
    placed_int = False
    if not contiguous:
        digits = list(int_str.replace(",", ""))
        slots = {}
        for p in reversed(int_pos):
            slots[p] = digits.pop() if digits else ("0" if toks[p][1] == "0" else "")
        if digits:
            slots[int_pos[0]] = "".join(digits) + slots[int_pos[0]]
    frac_iter = iter(frac_str)
    for i, t in enumerate(toks):
        kind = t[0]
        if kind == "lit":
            out.append(t[1])
        elif kind == "general":
            out.append(format_general(a))
        elif kind == "n":
            ch = t[1]
            if i in int_pos or (ch == "," and int_pos and int_pos[0] < i < int_pos[-1]):
                if contiguous:
                    if not placed_int:
                        out.append(int_str)
                        placed_int = True
                elif ch != ",":
                    out.append(slots.get(i, ""))
            elif i == dot_i:
                out.append(".")
            elif i in frac_pos:
                out.append(next(frac_iter, ""))
            elif ch == "%":
                out.append("%")
        elif kind == "exp":
            if not placed_int and contiguous:
                out.append(int_str)
                placed_int = True
            out.append(exp_str)
        elif kind == "text":
            pass
    res = "".join(out)
    if neg and any(ch.isdigit() and ch != "0" for ch in res):
        res = "-" + res
    return res


def _render_fraction(toks, slash, v):
    """Formats like '# ?/?', '?/??', '# ??/16'."""
    from fractions import Fraction
    digit = lambda t: t[0] == "n" and t[1] in "0#?"
    j = slash - 1
    while j >= 0 and digit(toks[j]):
        j -= 1
    num_start = j + 1
    num_width = slash - num_start
    int_pos = [i for i in range(num_start) if digit(toks[i])]
    k = slash + 1
    den_digits = 0
    fixed = ""
    while k < len(toks) and (digit(toks[k]) or (toks[k][0] == "lit" and toks[k][1].isdigit())):
        if digit(toks[k]):
            den_digits += 1
        else:
            fixed += toks[k][1]
        k += 1
    a = abs(v)
    ip = math.floor(a) if int_pos else 0
    frac = a - ip
    if fixed:
        den = int(fixed)
        num = round(frac * den)
    else:
        f = Fraction(frac).limit_denominator(10 ** max(1, den_digits) - 1)
        num, den = f.numerator, f.denominator
    if num == den and int_pos:
        ip, num = ip + 1, 0
    prefix = "".join(t[1] for t in toks[:int_pos[0] if int_pos else num_start] if t[0] == "lit")
    suffix = "".join(t[1] for t in toks[k:] if t[0] == "lit")
    if int_pos:
        sep = "".join(t[1] for t in toks[int_pos[-1] + 1:num_start] if t[0] == "lit")
        if num == 0:
            core = str(ip) + " " * (len(sep) + num_width + 1 + max(den_digits, len(fixed)))
        else:
            core = (str(ip) if ip else "") + (sep if ip else " " * len(sep)) + \
                f"{str(num).rjust(num_width)}/{str(den).ljust(max(den_digits, len(fixed)))}"
    else:
        core = f"{str(num).rjust(num_width)}/{str(den).ljust(max(den_digits, len(fixed)))}"
    return ("-" if v < 0 and (ip or num) else "") + prefix + core + suffix


def _render_date(sec, v):
    v = float(v)
    has_frac_sec = False
    toks = sec.tokens
    # fractional seconds: ".0", ".00" after a seconds token
    frac_digits = 0
    for i, t in enumerate(toks):
        if t[0] == "n" and t[1] == "." and i > 0:
            j = i + 1
            while j < len(toks) and toks[j] == ("n", "0"):
                j += 1
            frac_digits = j - i - 1
            has_frac_sec = frac_digits > 0
    if not has_frac_sec:
        v = round(v * 86400) / 86400
    d = serial_to_datetime(v) if v >= 0 else None
    if d is None:
        return "#" * 8
    out = []
    skip_n = False
    for i, t in enumerate(toks):
        kind = t[0]
        if kind == "lit":
            out.append(t[1])
        elif kind == "d":
            code = t[1]
            ch = code[0]
            n = len(code)
            if ch == "y":
                out.append(f"{d.year % 100:02d}" if n <= 2 else f"{d.year:04d}")
            elif ch == "m":
                if n == 1:
                    out.append(str(d.month))
                elif n == 2:
                    out.append(f"{d.month:02d}")
                elif n == 3:
                    out.append(MONTHS[d.month - 1][:3])
                elif n == 5:
                    out.append(MONTHS[d.month - 1][0])
                else:
                    out.append(MONTHS[d.month - 1])
            elif ch == "d":
                if n == 1:
                    out.append(str(d.day))
                elif n == 2:
                    out.append(f"{d.day:02d}")
                elif n == 3:
                    out.append(DAYS[d.weekday()][:3])
                else:
                    out.append(DAYS[d.weekday()])
            elif ch == "h":
                h = d.hour
                if sec.has_ampm:
                    h = h % 12 or 12
                out.append(f"{h:02d}" if n >= 2 else str(h))
            elif ch == "s":
                out.append(f"{d.second:02d}" if n >= 2 else str(d.second))
        elif kind == "min":
            out.append(f"{d.minute:02d}" if len(t[1]) >= 2 else str(d.minute))
        elif kind == "elapsed":
            unit, n = t[1], t[2]
            total = v * {"h": 24, "m": 1440, "s": 86400}[unit]
            out.append(str(int(math.floor(total + 1e-9))).rjust(n, "0"))
        elif kind == "ampm":
            pm = d.hour >= 12
            out.append(("PM" if pm else "AM") if t[1] == "AM/PM" else ("P" if pm else "A"))
        elif kind == "n":
            if t[1] == "." and has_frac_sec and not skip_n:
                frac = d.microsecond / 1e6
                out.append("." + f"{frac:.{frac_digits}f}"[2:])
                skip_n = True
            elif t[1] == "0" and skip_n:
                continue
            else:
                out.append(t[1])
        elif kind == "general":
            out.append(format_general(v))
    return "".join(out)


# ---------------------------------------------------------------- input parsing

_NUM_RE = re.compile(r"^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$")
_RICH_NUM_RE = re.compile(
    r"^(?P<neg1>-)?(?P<paren>\()?(?P<neg2>-)?(?P<cur>[$€£¥])?(?P<neg3>-)?"
    r"(?P<num>\d{1,3}(,\d{3})+(\.\d*)?|\d+\.?\d*|\.\d+)(?P<exp>[eE][+-]?\d+)?"
    r"(?P<pct>%)?(?P<close>\))?$")
_ISO_DATE_RE = re.compile(r"^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2})(\.\d+)?)?)?$")
_US_DATE_RE = re.compile(r"^(\d{1,2})/(\d{1,2})(?:/(\d{2}|\d{4}))?(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AaPp][Mm])?)?$")
_TIME_RE = re.compile(r"^(\d{1,2}):(\d{2})(?::(\d{2})(\.\d+)?)?\s*([AaPp][Mm])?$")
_MON_DATE_RE = re.compile(r"^(\d{1,2})[- ]([A-Za-z]{3,9})[- ,]*(\d{2}|\d{4})?$")
_MON_FIRST_RE = re.compile(r"^([A-Za-z]{3,9})\.? (\d{1,2}),? (\d{4})$")
_MONTH_LOOKUP = {m[:3].lower(): i + 1 for i, m in enumerate(MONTHS)}


def _month(name):
    name = name.lower()
    if name[:3] in _MONTH_LOOKUP and (len(name) == 3 or MONTHS[_MONTH_LOOKUP[name[:3]] - 1].lower().startswith(name)):
        return _MONTH_LOOKUP[name[:3]]
    return None


def _decimals_fmt(numtext):
    if "." in numtext:
        dec = len(numtext.split(".", 1)[1])
        return dec
    return 0


def parse_input(text, typed=True):
    """Convert typed/imported text to (value, suggested_format_or_None).

    Formulas are not handled here.  Leading-zero digit strings ("00123") and
    numbers longer than 15 digits are kept as text so IDs survive.
    """
    if text is None:
        return None, None
    s = text.strip()
    if s == "":
        return (None, None) if typed else ("" if text else None, None)
    if text.startswith("'") and typed:
        return text[1:], None
    if _NUM_RE.match(s):
        digits = s.lstrip("+-")
        if len(digits) > 1 and digits[0] == "0" and digits[1].isdigit():
            return text, None
        if len(digits.replace(".", "").lstrip("0")) > 15 and "e" not in digits.lower():
            return text, None
        v = float(s)
        if v.is_integer() and abs(v) < 1e15 and "e" not in s.lower() and "." not in s:
            v = float(int(s))
        fmt = None
        if "." in s and s.endswith("0") and "e" not in s.lower():
            fmt = "0." + "0" * _decimals_fmt(s)
        return v, fmt
    up = s.upper()
    if up == "TRUE":
        return True, None
    if up == "FALSE":
        return False, None
    m = _RICH_NUM_RE.match(s)
    if m and (m.group("paren") is None) == (m.group("close") is None):
        num = m.group("num")
        try:
            v = float(num.replace(",", "") + (m.group("exp") or ""))
        except ValueError:
            return text, None
        neg = bool(m.group("neg1") or m.group("neg2") or m.group("neg3") or m.group("paren"))
        if neg:
            v = -v
        dec = _decimals_fmt(num)
        dpart = ("." + "0" * dec) if dec else ""
        if m.group("pct"):
            v /= 100
            return v, "0" + dpart + "%"
        if m.group("cur"):
            sym = m.group("cur")
            q = f'"{sym}"' if sym != "$" else "$"
            return v, f"{q}#,##0{dpart}"
        if "," in num:
            return v, "#,##0" + dpart
        return v, ("0" + dpart) if dec else None
    d = _parse_datetime(s)
    if d is not None:
        return d
    return text, None


def _ymd(m):
    """ISO date format matching the digit counts that were typed (2024-1-5 vs 2024-01-05)."""
    return "yyyy-" + ("mm" if len(m.group(2)) == 2 else "m") + "-" + ("dd" if len(m.group(3)) == 2 else "d")


def _parse_datetime(s):
    m = _ISO_DATE_RE.match(s)
    if m:
        try:
            y, mo, da = int(m.group(1)), int(m.group(2)), int(m.group(3))
            if m.group(4) is not None:
                sec = int(m.group(6) or 0)
                frac = float(m.group(7) or 0)
                d = dt.datetime(y, mo, da, int(m.group(4)), int(m.group(5)), sec, int(frac * 1e6))
                fmt = _ymd(m) + (" hh" if len(m.group(4)) == 2 else " h") + ":mm" + \
                    (":ss" if m.group(6) is not None else "")
            else:
                d = dt.datetime(y, mo, da)
                fmt = _ymd(m)
            return datetime_to_serial(d), fmt
        except ValueError:
            return None
    m = _US_DATE_RE.match(s)
    if m:
        try:
            mo, da = int(m.group(1)), int(m.group(2))
            ytxt = m.group(3)
            if ytxt is None:
                y = dt.date.today().year
                fmt = "d-mmm"
            else:
                y = int(ytxt)
                if len(ytxt) == 2:
                    y += 2000 if y < 30 else 1900
                fmt = ("mm/" if len(m.group(1)) == 2 else "m/") + ("dd/" if len(m.group(2)) == 2 else "d/") + \
                    ("yy" if len(ytxt) == 2 else "yyyy")
            d = dt.datetime(y, mo, da)
            if m.group(4) is not None:
                h = int(m.group(4))
                ap = m.group(7)
                if ap:
                    h = h % 12 + (12 if ap.lower() == "pm" else 0)
                d = d.replace(hour=h, minute=int(m.group(5)), second=int(m.group(6) or 0))
                fmt += " h:mm" + (":ss" if m.group(6) else "") + (" AM/PM" if ap else "")
            return datetime_to_serial(d), fmt
        except ValueError:
            return None
    m = _TIME_RE.match(s)
    if m:
        h, mi = int(m.group(1)), int(m.group(2))
        sec = int(m.group(3) or 0)
        ap = m.group(5)
        if ap:
            if h > 12:
                return None
            h = h % 12 + (12 if ap.lower() == "pm" else 0)
        if mi > 59 or sec > 59:
            return None
        frac = float(m.group(4) or 0)
        v = (h * 3600 + mi * 60 + sec + frac) / 86400
        fmt = ("h:mm" if not ap else "h:mm") + (":ss" if m.group(3) else "") + (" AM/PM" if ap else "")
        if h >= 24 and not ap:
            fmt = "[h]:mm" + (":ss" if m.group(3) else "")
        return v, fmt
    m = _MON_DATE_RE.match(s)
    if m:
        mo = _month(m.group(2))
        if mo:
            try:
                y = int(m.group(3)) if m.group(3) else dt.date.today().year
                if m.group(3) and len(m.group(3)) == 2:
                    y += 2000 if y < 30 else 1900
                d = dt.datetime(y, mo, int(m.group(1)))
                fmt = "d-mmm-yy" if m.group(3) and len(m.group(3)) == 2 else ("d-mmm-yyyy" if m.group(3) else "d-mmm")
                return datetime_to_serial(d), fmt
            except ValueError:
                return None
    m = _MON_FIRST_RE.match(s)
    if m:
        mo = _month(m.group(1))
        if mo:
            try:
                d = dt.datetime(int(m.group(3)), mo, int(m.group(2)))
                return datetime_to_serial(d), "mmmm d, yyyy" if len(m.group(1)) > 3 else "mmm d, yyyy"
            except ValueError:
                return None
    return None


def edit_text(v, fmt="General"):
    """Text shown when a constant cell is opened for editing."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, XLError):
        return v.code
    if isinstance(v, str):
        if v.startswith("=") or (parse_input(v)[0] != v and v.strip()):
            return "'" + v
        return v
    secs = compile_format(fmt)
    if secs and secs[0].is_date:
        try:
            d = serial_to_datetime(v)
        except (ValueError, OverflowError):
            return repr_number(v)
        has_time = abs(v - math.floor(v)) > 1e-9
        if v < 1 and has_time:
            return f"{d.hour}:{d.minute:02d}:{d.second:02d}" if d.second else f"{d.hour}:{d.minute:02d}"
        base = f"{d.month}/{d.day}/{d.year}"
        if has_time:
            base += f" {d.hour}:{d.minute:02d}" + (f":{d.second:02d}" if d.second else "")
        return base
    if secs and "%" in fmt and not secs[0].is_text:
        return repr_number(round(v * 100, 10)) + "%"
    return repr_number(v)


def repr_number(v):
    if float(v).is_integer() and abs(v) < 1e16:
        return str(int(v))
    s = repr(float(v))
    if "e" in s:
        m, e = s.split("e")
        return f"{m}E{int(e):+d}"
    # trim floating noise like 0.30000000000000004
    t = f"{v:.15g}"
    return t if "e" not in t else s
