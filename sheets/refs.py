"""Cell addressing helpers: A1 notation <-> (row, col), all 0-based internally."""
import re

MAX_ROWS = 100_000_000  # Excel stops at 1,048,576; big CSVs need more
XLSX_MAX_ROWS = 1_048_576
MAX_COLS = 16_384
_COL_BITS = 14  # 2**14 == MAX_COLS


def key(r, c):
    return (r << _COL_BITS) | c


def unkey(k):
    return k >> _COL_BITS, k & (MAX_COLS - 1)


def col_name(c):
    s = ""
    c += 1
    while c:
        c, rem = divmod(c - 1, 26)
        s = chr(65 + rem) + s
    return s


def col_index(name):
    n = 0
    for ch in name.upper():
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def addr(r, c, abs_row=False, abs_col=False):
    return f"{'$' if abs_col else ''}{col_name(c)}{'$' if abs_row else ''}{r + 1}"


def range_addr(r1, c1, r2, c2):
    if r1 == r2 and c1 == c2:
        return addr(r1, c1)
    if r1 == 0 and r2 >= MAX_ROWS - 1:
        return f"{col_name(c1)}:{col_name(c2)}"
    if c1 == 0 and c2 >= MAX_COLS - 1:
        return f"{r1 + 1}:{r2 + 1}"
    return f"{addr(r1, c1)}:{addr(r2, c2)}"


_CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")
_COLS_RE = re.compile(r"^\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})$")
_ROWS_RE = re.compile(r"^\$?(\d+):\$?(\d+)$")


def parse_addr(text):
    m = _CELL_RE.match(text.strip())
    if not m:
        return None
    c = col_index(m.group(1))
    r = int(m.group(2)) - 1
    if not (0 <= r < MAX_ROWS and 0 <= c < MAX_COLS):
        return None
    return r, c


def parse_range(text):
    """Parse 'A1', 'A1:B5', 'A:C', '3:7' -> (r1, c1, r2, c2) normalized, or None."""
    t = text.strip()
    if "!" in t:
        t = t.split("!", 1)[1]
    m = _COLS_RE.match(t)
    if m:
        a, b = sorted((col_index(m.group(1)), col_index(m.group(2))))
        if b >= MAX_COLS:
            return None
        return 0, a, MAX_ROWS - 1, b
    m = _ROWS_RE.match(t)
    if m:
        a, b = sorted((int(m.group(1)) - 1, int(m.group(2)) - 1))
        if a < 0 or b >= MAX_ROWS:
            return None
        return a, 0, b, MAX_COLS - 1
    if ":" in t:
        p, q = t.split(":", 1)
        a, b = parse_addr(p), parse_addr(q)
        if not a or not b:
            return None
        return min(a[0], b[0]), min(a[1], b[1]), max(a[0], b[0]), max(a[1], b[1])
    a = parse_addr(t)
    if not a:
        return None
    return a[0], a[1], a[0], a[1]
