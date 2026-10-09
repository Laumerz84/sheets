"""Editing operations. Each returns a dict {key: (content, style)} of new
states so the UI can apply them as one undoable step."""
import math
import re

from .errors import XLError
from .formula import ParseError, normalize, parse, shift_formula
from .numfmt import (DAYS, MONTHS, compile_format, edit_text, format_value,
                     is_date_format, is_text_format, parse_input)
from .refs import MAX_COLS, MAX_ROWS, key
from .values import is_num, sort_key
from .workbook import DEFAULT_STYLE, intern_style


# ---------------------------------------------------------------- typed input

def input_state(sheet, r, c, text):
    """New state for typing `text` into a cell."""
    k = key(r, c)
    style = sheet.styles.get(k, DEFAULT_STYLE)
    if text is None or text == "":
        return (None, style)
    if text.startswith("=") and len(text) > 1 and not is_text_format(style.numfmt):
        try:
            parse(text)
            text = normalize(text)
        except ParseError:
            pass
        return (("f", text), style)
    if text.startswith("+") and len(text) > 1 and not _looks_numeric(text):
        return input_state(sheet, r, c, "=" + text[1:])
    if is_text_format(style.numfmt):
        return (("v", text), style)
    v, fmt = parse_input(text)
    if fmt and (style.numfmt == "General" or (is_date_format(fmt) != is_date_format(style.numfmt))):
        if not (fmt.startswith("0.") and style.numfmt == "General" and not text.strip().endswith("0")):
            style = style.with_(numfmt=fmt)
    return (("v", v) if v is not None else None, style)


def _looks_numeric(t):
    v, _ = parse_input(t)
    return is_num(v)


def edit_text_for(sheet, r, c):
    k = key(r, c)
    f = sheet.formulas.get(k)
    if f is not None:
        return f.text
    v = sheet.values.get(k)
    return edit_text(v, sheet.styles.get(k, DEFAULT_STYLE).numfmt)


def display_text(sheet, r, c):
    k = key(r, c)
    v = sheet.value_at_key(k)
    if v is None:
        return ""
    st = sheet.styles.get(k, DEFAULT_STYLE)
    return format_value(v, st.numfmt)[0]


# ---------------------------------------------------------------- clearing / formatting

def iter_rect(rect):
    r1, c1, r2, c2 = rect
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            yield r, c


def rect_keys(sheet, rect, include_styles=True):
    """Keys inside rect that hold anything (sparse-friendly)."""
    r1, c1, r2, c2 = rect
    area = (r2 - r1 + 1) * (c2 - c1 + 1)
    sources = [sheet.values, sheet.formulas] + ([sheet.styles] if include_styles else [])
    if area <= 200000:
        out = []
        for r in range(r1, r2 + 1):
            base = r << 14
            for c in range(c1, c2 + 1):
                k = base | c
                if any(k in d for d in sources):
                    out.append(k)
        return out
    out = set()
    for d in sources:
        for k in d:
            r, c = k >> 14, k & 0x3FFF
            if r1 <= r <= r2 and c1 <= c <= c2:
                out.add(k)
    return sorted(out)


def clear_contents(sheet, rects):
    states = {}
    for rect in rects:
        for k in rect_keys(sheet, rect, include_styles=False):
            states[k] = (None, sheet.styles.get(k, DEFAULT_STYLE))
    return states


def clear_formats(sheet, rects):
    states = {}
    for rect in rects:
        for k in rect_keys(sheet, rect):
            content, _ = sheet.get_state(k)
            states[k] = (content, DEFAULT_STYLE)
    return states


def clear_all(sheet, rects):
    states = {}
    for rect in rects:
        for k in rect_keys(sheet, rect):
            states[k] = (None, DEFAULT_STYLE)
    return states


def restyle(sheet, rects, fn):
    """Apply fn(Style)->Style to every cell in rects (all cells, not just used ones)."""
    states = {}
    for rect in rects:
        r1, c1, r2, c2 = rect
        area = (r2 - r1 + 1) * (c2 - c1 + 1)
        if area > 500000:
            # whole columns/rows: only touch used cells
            r2 = min(r2, max(sheet.used_extent()[0], r1))
            c2 = min(c2, max(sheet.used_extent()[1], c1))
        for r in range(r1, r2 + 1):
            base = r << 14
            for c in range(c1, c2 + 1):
                k = base | c
                content, st = sheet.get_state(k)
                new = intern_style(fn(st))
                if new != st:
                    states[k] = (content, new)
    return states


def borders_states(sheet, rect, mode, side=("thin", "#000000")):
    """mode: all | outline | inside | bottom | top | left | right | none | thick_outline."""
    r1, c1, r2, c2 = rect
    states = {}
    if mode == "thick_outline":
        side, mode = ("medium", "#000000"), "outline"
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            k = key(r, c)
            content, st = sheet.get_state(k)
            L, T, R, B = st.border
            if mode == "none":
                L = T = R = B = None
            elif mode == "all":
                L = T = R = B = side
            elif mode == "outline":
                if c == c1:
                    L = side
                if c == c2:
                    R = side
                if r == r1:
                    T = side
                if r == r2:
                    B = side
            elif mode == "inside":
                if c > c1:
                    L = side
                if c < c2:
                    R = side
                if r > r1:
                    T = side
                if r < r2:
                    B = side
            elif mode == "bottom" and r == r2:
                B = side
            elif mode == "top" and r == r1:
                T = side
            elif mode == "left" and c == c1:
                L = side
            elif mode == "right" and c == c2:
                R = side
            new = st.with_(border=(L, T, R, B))
            if new != st:
                states[k] = (content, new)
    return states


# ---------------------------------------------------------------- clipboard

class Clip:
    """Copied block: rows x cols of (content, style) plus origin."""

    def __init__(self, sheet, rect, cells, cut=False, rows=None):
        self.sheet = sheet
        self.rect = rect
        self.cells = cells  # list of rows of (content, style)
        self.cut = cut
        self.text = None
        # sheet row of each clip row (rows hidden by a filter are skipped when copying)
        self.rows = list(rows) if rows is not None else list(range(rect[0], rect[0] + len(cells)))

    @property
    def nrows(self):
        return len(self.cells)

    @property
    def ncols(self):
        return len(self.cells[0]) if self.cells else 0


def copy_rect(sheet, rect, rows=None):
    """rows: explicit list of sheet rows to include (skips filtered rows)."""
    r1, c1, r2, c2 = rect
    rows = list(rows) if rows is not None else list(range(r1, r2 + 1))
    cells = []
    for r in rows:
        cells.append([sheet.get_state(key(r, c)) for c in range(c1, c2 + 1)])
    return Clip(sheet, rect, cells, rows=rows)


def clip_to_text(sheet, clip, rows=None):
    """Tab-separated text of displayed values (what Excel puts on the clipboard)."""
    r1, c1, r2, c2 = clip.rect
    rows = rows if rows is not None else range(r1, r2 + 1)
    lines = []
    for r in rows:
        parts = []
        for c in range(c1, c2 + 1):
            t = display_text(sheet, r, c)
            if any(ch in t for ch in '\t\n"'):
                t = '"' + t.replace('"', '""') + '"'
            parts.append(t)
        lines.append("\t".join(parts))
    return "\r\n".join(lines) + "\r\n"


def parse_clipboard_text(text):
    """TSV (Excel style, with quoted multi-line cells) -> list of rows of strings."""
    if text.endswith("\r\n"):
        text = text[:-2]
    elif text.endswith("\n"):
        text = text[:-1]
    rows, row, cell = [], [], []
    i, n = 0, len(text)
    at_start = True
    while i < n:
        ch = text[i]
        if at_start and ch == '"':
            # quoted cell
            j = i + 1
            buf = []
            closed = False
            while j < n:
                if text[j] == '"':
                    if j + 1 < n and text[j + 1] == '"':
                        buf.append('"')
                        j += 2
                        continue
                    closed = True
                    j += 1
                    break
                buf.append(text[j])
                j += 1
            if closed and (j >= n or text[j] in "\t\r\n"):
                cell = buf
                i = j
                at_start = False
                continue
            cell.append(ch)
            i += 1
            at_start = False
            continue
        if ch == "\t":
            row.append("".join(cell))
            cell = []
            at_start = True
        elif ch == "\n" or ch == "\r":
            row.append("".join(cell))
            rows.append(row)
            row, cell = [], []
            at_start = True
            if ch == "\r" and i + 1 < n and text[i + 1] == "\n":
                i += 1
        else:
            cell.append(ch)
            at_start = False
        i += 1
    row.append("".join(cell))
    rows.append(row)
    return rows


def paste_states(sheet, clip, dest_rect, values_only=False, formats_only=False, transpose=False):
    """States for pasting clip at dest (tiles if dest is a multiple of the clip size)."""
    cells = clip.cells
    if transpose:
        cells = [list(r) for r in zip(*cells)]
    h, w = len(cells), len(cells[0]) if cells else 0
    if not h or not w:
        return {}
    dr1, dc1, dr2, dc2 = dest_rect
    th, tw = dr2 - dr1 + 1, dc2 - dc1 + 1
    reps_r = th // h if th % h == 0 and th >= h else 1
    reps_c = tw // w if tw % w == 0 and tw >= w else 1
    sr1, sc1 = clip.rect[0], clip.rect[1]
    states = {}
    src_sheet = clip.sheet
    for rr in range(reps_r):
        for cc in range(reps_c):
            for i in range(h):
                for j in range(w):
                    r, c = dr1 + rr * h + i, dc1 + cc * w + j
                    if r >= MAX_ROWS or c >= MAX_COLS:
                        continue
                    content, style = cells[i][j]
                    k = key(r, c)
                    cur_content, cur_style = sheet.get_state(k)
                    if formats_only:
                        states[k] = (cur_content, style)
                        continue
                    if content is not None and content[0] == "f":
                        if values_only:
                            si, sj = (j, i) if transpose else (i, j)
                            v = src_sheet.value(clip.rows[si], sc1 + sj) if src_sheet is not None else None
                            content = None if v is None else ("v", v)
                        elif not clip.cut:
                            si, sj = (j, i) if transpose else (i, j)
                            content = ("f", shift_formula(content[1], r - clip.rows[si], c - (sc1 + sj)))
                    if values_only:
                        style = cur_style.with_(numfmt=style.numfmt) if style.numfmt != "General" and \
                            cur_style.numfmt == "General" else cur_style
                    states[k] = (content, style)
    return states


def text_paste_states(sheet, rows, r0, c0):
    states = {}
    for i, row in enumerate(rows):
        for j, t in enumerate(row):
            r, c = r0 + i, c0 + j
            if r >= MAX_ROWS or c >= MAX_COLS:
                continue
            states[key(r, c)] = input_state(sheet, r, c, t)
    return states


# ---------------------------------------------------------------- fill

_TEXT_NUM_RE = re.compile(r"^(.*?)(\d+)(\D*)$")
_MONTH_FULL = [m.lower() for m in MONTHS]
_MONTH_ABBR = [m[:3].lower() for m in MONTHS]
_DAY_FULL = [d.lower() for d in DAYS]
_DAY_ABBR = [d[:3].lower() for d in DAYS]


def _name_series(s):
    low = s.lower()
    for lst in (_MONTH_FULL, _MONTH_ABBR, _DAY_FULL, _DAY_ABBR):
        if low in lst:
            return lst, lst.index(low)
    return None


def _match_case(template, word):
    if template.isupper():
        return word.upper()
    if template[:1].isupper():
        return word.capitalize()
    return word


def _series_values(src, count, sheet_styles):
    """Given source states (content, style), produce `count` continuation states.
    Returns list or None (meaning: plain repeat)."""
    contents = [s[0] for s in src]
    if any(c is None or c[0] == "f" for c in contents):
        return None
    vals = [c[1] for c in contents]
    n = len(vals)
    if all(is_num(v) for v in vals):
        if n == 1:
            # single numbers repeat, single dates count up by a day
            if is_date_format(src[0][1].numfmt):
                return [vals[0] + i + 1 for i in range(count)]
            return None
        xs = list(range(n))
        mx = sum(xs) / n
        my = sum(vals) / n
        sxx = sum((x - mx) ** 2 for x in xs)
        slope = sum((x - mx) * (y - my) for x, y in zip(xs, vals)) / sxx
        icpt = my - slope * mx
        out = []
        for i in range(count):
            v = icpt + slope * (n + i)
            if all(float(x).is_integer() for x in vals) and float(slope).is_integer():
                v = float(round(v))
            else:
                v = round(v, 12)
            out.append(v)
        return out
    if all(isinstance(v, str) for v in vals):
        names = [_name_series(v) for v in vals]
        if all(names) and len({id(x[0]) for x in names}) == 1:
            lst = names[0][0]
            idxs = [x[1] for x in names]
            step = (idxs[-1] - idxs[0]) // (n - 1) if n > 1 else 1
            if n > 1 and any((idxs[i + 1] - idxs[i]) % len(lst) != step % len(lst) for i in range(n - 1)):
                return None
            return [_match_case(vals[-1], lst[(idxs[-1] + step * (i + 1)) % len(lst)]) for i in range(count)]
        ms = [_TEXT_NUM_RE.match(v) for v in vals]
        if all(ms) and len({(m.group(1), m.group(3)) for m in ms}) == 1:
            nums = [int(m.group(2)) for m in ms]
            step = (nums[-1] - nums[0]) // (n - 1) if n > 1 else 1
            if n > 1 and any(nums[i + 1] - nums[i] != step for i in range(n - 1)):
                return None
            width = len(ms[-1].group(2)) if ms[-1].group(2).startswith("0") else 0
            pre, post = ms[0].group(1), ms[0].group(3)
            out = []
            for i in range(count):
                x = nums[-1] + step * (i + 1)
                out.append(f"{pre}{str(abs(x)).zfill(width) if x >= 0 else x}{post}")
            return out
    return None


def fill_states(sheet, src_rect, dst_rect, series=True):
    """Fill dst_rect (which extends src_rect in one direction) from src_rect."""
    sr1, sc1, sr2, sc2 = src_rect
    dr1, dc1, dr2, dc2 = dst_rect
    states = {}
    if dr2 > sr2 or dr1 < sr1:
        vertical = True
        forward = dr2 > sr2
    else:
        vertical = False
        forward = dc2 > sc2
    lines = range(sc1, sc2 + 1) if vertical else range(sr1, sr2 + 1)
    for line in lines:
        if vertical:
            src_pos = list(range(sr1, sr2 + 1))
            dst_pos = list(range(sr2 + 1, dr2 + 1)) if forward else list(range(sr1 - 1, dr1 - 1, -1))
            src = [sheet.get_state(key(p, line)) for p in src_pos]
        else:
            src_pos = list(range(sc1, sc2 + 1))
            dst_pos = list(range(sc2 + 1, dc2 + 1)) if forward else list(range(sc1 - 1, dc1 - 1, -1))
            src = [sheet.get_state(key(line, p)) for p in src_pos]
        if not dst_pos:
            continue
        ordered = src if forward else list(reversed(src))
        seq = _series_values(ordered, len(dst_pos), None) if series else None
        n = len(src)
        for i, p in enumerate(dst_pos):
            si = i % n
            s_content, s_style = ordered[si]
            spos = src_pos[si] if forward else list(reversed(src_pos))[si]
            r, c = (p, line) if vertical else (line, p)
            if seq is not None:
                content = ("v", seq[i])
            elif s_content is not None and s_content[0] == "f":
                dr, dc = (p - spos, 0) if vertical else (0, p - spos)
                content = ("f", shift_formula(s_content[1], dr, dc))
            else:
                content = s_content
            states[key(r, c)] = (content, s_style)
    return states


def fill_down_states(sheet, rect, direction="down"):
    """Ctrl+D / Ctrl+R: copy the first row/column of rect into the rest."""
    r1, c1, r2, c2 = rect
    states = {}
    if direction == "down":
        if r1 == r2:
            if r1 == 0:
                return states
            r1 -= 1
        for c in range(c1, c2 + 1):
            content, style = sheet.get_state(key(r1, c))
            for r in range(r1 + 1, r2 + 1):
                ct = ("f", shift_formula(content[1], r - r1, 0)) if content and content[0] == "f" else content
                states[key(r, c)] = (ct, style)
    else:
        if c1 == c2:
            if c1 == 0:
                return states
            c1 -= 1
        for r in range(r1, r2 + 1):
            content, style = sheet.get_state(key(r, c1))
            for c in range(c1 + 1, c2 + 1):
                ct = ("f", shift_formula(content[1], 0, c - c1)) if content and content[0] == "f" else content
                states[key(r, c)] = (ct, style)
    return states


# ---------------------------------------------------------------- sort

def guess_header(sheet, rect):
    r1, c1, r2, c2 = rect
    if r2 <= r1:
        return False
    first = [sheet.value(r1, c) for c in range(c1, c2 + 1)]
    second = [sheet.value(r1 + 1, c) for c in range(c1, c2 + 1)]
    if not all(isinstance(v, str) for v in first if v is not None) or all(v is None for v in first):
        return False
    if any(is_num(v) for v in second):
        return True
    st1 = [sheet.style(r1, c) for c in range(c1, c2 + 1)]
    if any(s.bold for s in st1) and not any(sheet.style(r1 + 1, c).bold for c in range(c1, c2 + 1)):
        return True
    # all-text table starting on the first row: almost always has a header row
    labels = [str(v).lower() for v in first if v is not None]
    return r1 == 0 and len(set(labels)) == len(labels)


def sort_states(sheet, rect, keys, header=False, case_sensitive=False, skip_rows=None):
    """keys: [(col_index_absolute, ascending)].  Rows move with formulas shifted.
    Blanks always sort last; ties keep their original order."""
    r1, c1, r2, c2 = rect
    if header:
        r1 += 1
    rows = [r for r in range(r1, r2 + 1) if not skip_rows or r not in skip_rows]
    if len(rows) < 2:
        return {}
    order = list(rows)
    value = sheet.value
    for col, asc in reversed(keys):
        def k(r, col=col, asc=asc):
            v = value(r, col)
            # group flag keeps blanks last in both directions (reverse flips groups too)
            if v is None or v == "":
                return (asc, (0, 0, ""))
            if isinstance(v, XLError):
                sk = (3, 0, "")
            elif case_sensitive and isinstance(v, str):
                sk = (1, 0, v)
            else:
                sk = sort_key(v)
            return (not asc, sk)
        order.sort(key=k, reverse=not asc)
    if order == rows:
        return {}
    values, formulas, styles = sheet.values, sheet.formulas, sheet.styles
    snapshot = {}
    for r in rows:
        base = r << 14
        line = []
        for c in range(c1, c2 + 1):
            kk = base | c
            f = formulas.get(kk)
            content = ("f", f.text) if f is not None else (("v", values[kk]) if kk in values else None)
            line.append((content, styles.get(kk, DEFAULT_STYLE)))
        snapshot[r] = line
    states = {}
    for dst, src in zip(rows, order):
        if dst == src:
            continue
        base = dst << 14
        for j, (content, style) in enumerate(snapshot[src]):
            if content is not None and content[0] == "f":
                content = ("f", shift_formula(content[1], dst - src, 0))
            states[base | (c1 + j)] = (content, style)
    return states


# ---------------------------------------------------------------- navigation helpers

def current_region(sheet, r, c):
    """The contiguous block of non-empty cells around (r, c), like Ctrl+A / Ctrl+*."""
    has = sheet.has_content
    r1 = r2 = r
    c1 = c2 = c
    if not has(r, c):
        # look at neighbours so a click just below/right of a table still finds it
        found = False
        for dr, dc in ((-1, 0), (0, -1), (1, 0), (0, 1), (-1, -1)):
            rr, cc = r + dr, c + dc
            if rr >= 0 and cc >= 0 and has(rr, cc):
                r1 = r2 = rr
                c1 = c2 = cc
                found = True
                break
        if not found:
            return (r, c, r, c)
    def row_has(r):
        return any(has(r, x) for x in range(max(0, c1 - 1), min(MAX_COLS - 1, c2 + 1) + 1))

    def col_has(c):
        return any(has(y, c) for y in range(max(0, r1 - 1), min(MAX_ROWS - 1, r2 + 1) + 1))

    changed = True
    while changed:
        changed = False
        while r1 > 0 and row_has(r1 - 1):
            r1 -= 1
            changed = True
        while r2 < MAX_ROWS - 1 and row_has(r2 + 1):
            r2 += 1
            changed = True
        while c1 > 0 and col_has(c1 - 1):
            c1 -= 1
            changed = True
        while c2 < MAX_COLS - 1 and col_has(c2 + 1):
            c2 += 1
            changed = True
    return (r1, c1, r2, c2)


def data_edge(sheet, r, c, dr, dc, limit_r, limit_c):
    """Ctrl+Arrow target from (r, c) moving by (dr, dc)."""
    has = sheet.has_content

    def ok(rr, cc):
        return 0 <= rr <= limit_r and 0 <= cc <= limit_c

    nr, nc = r + dr, c + dc
    if not ok(nr, nc):
        return r, c
    if has(r, c) and has(nr, nc):
        while ok(nr + dr, nc + dc) and has(nr + dr, nc + dc):
            nr, nc = nr + dr, nc + dc
        return nr, nc
    # jump to next non-empty cell (fast path along sparse data)
    if dr:
        rows = sorted({k >> 14 for d in (sheet.values, sheet.formulas) for k in d if (k & 0x3FFF) == c})
        if dr > 0:
            nxt = [x for x in rows if x > r]
            return (nxt[0], c) if nxt else (limit_r, c)
        prv = [x for x in rows if x < r]
        return (prv[-1], c) if prv else (0, c)
    cols = sorted({k & 0x3FFF for d in (sheet.values, sheet.formulas) for k in d if (k >> 14) == r})
    if dc > 0:
        nxt = [x for x in cols if x > c]
        return (r, nxt[0]) if nxt else (r, limit_c)
    prv = [x for x in cols if x < c]
    return (r, prv[-1]) if prv else (r, 0)


def autosum_range(sheet, r, c):
    """Guess the range for AutoSum at (r, c): numbers above, else to the left."""
    def numeric(rr, cc):
        v = sheet.value(rr, cc)
        return is_num(v)
    rr = r - 1
    while rr >= 0 and not sheet.has_content(rr, c):
        rr -= 1
    if rr >= 0 and numeric(rr, c):
        top = rr
        while top - 1 >= 0 and (numeric(top - 1, c) or (sheet.has_content(top - 1, c) and sheet.formulas.get(key(top - 1, c)) is not None)):
            top -= 1
        return (top, c, rr, c)
    cc = c - 1
    while cc >= 0 and not sheet.has_content(r, cc):
        cc -= 1
    if cc >= 0 and numeric(r, cc):
        left = cc
        while left - 1 >= 0 and numeric(r, left - 1):
            left -= 1
        return (r, left, r, cc)
    return None


# ---------------------------------------------------------------- find / replace

def find_matches(sheet, needle, match_case=False, whole=False, in_formulas=True, rect=None):
    """Sorted list of (r, c) whose text matches."""
    if needle == "":
        return []
    if any(ch in needle for ch in "*?"):
        from .functions import wild_match
        test = lambda s: wild_match(needle, s, whole, match_case)
    else:
        n = needle if match_case else needle.lower()
        if whole:
            test = lambda s: (s if match_case else s.lower()) == n
        else:
            test = lambda s: n in (s if match_case else s.lower())
    out = []
    keys = set(sheet.values) | set(sheet.formulas)
    for k in keys:
        r, c = k >> 14, k & 0x3FFF
        if rect and not (rect[0] <= r <= rect[2] and rect[1] <= c <= rect[3]):
            continue
        f = sheet.formulas.get(k)
        if f is not None and in_formulas:
            s = f.text
        else:
            s = display_text(sheet, r, c)
        if test(s):
            out.append((r, c))
    out.sort()
    return out


def replace_in_text(s, needle, repl, match_case=False, whole=False):
    if whole:
        same = s == needle if match_case else s.lower() == needle.lower()
        return repl if same else s
    if any(ch in needle for ch in "*?"):
        from .functions import _wild_compile
        rx = _wild_compile(needle)
        flags = 0 if match_case else re.IGNORECASE
        return re.sub(re.compile(rx.pattern.replace(".*", ".*?"), flags | re.DOTALL), lambda m: repl, s)
    if match_case:
        return s.replace(needle, repl)
    return re.sub(re.escape(needle), lambda m: repl, s, flags=re.IGNORECASE)


def replace_states(sheet, cells, needle, repl, match_case=False, whole=False):
    states = {}
    for r, c in cells:
        k = key(r, c)
        f = sheet.formulas.get(k)
        old = f.text if f is not None else edit_text_for(sheet, r, c)
        new = replace_in_text(old, needle, repl, match_case, whole)
        if new != old:
            states[k] = input_state(sheet, r, c, new)
    return states


# ---------------------------------------------------------------- filter

def filter_values_for(sheet, col, r1, r2):
    """Display strings present in a filter column (sorted, blanks flagged)."""
    seen = {}
    blanks = False
    for r in range(r1, r2 + 1):
        t = display_text(sheet, r, col)
        if t == "":
            blanks = True
            continue
        if t not in seen:
            v = sheet.value(r, col)
            seen[t] = v
    items = sorted(seen.items(), key=lambda kv: (sort_key(kv[1]) if kv[1] is not None else (9,)))
    return [t for t, _ in items], blanks


def row_passes(sheet, r, spec, col):
    t = display_text(sheet, r, col)
    allowed = spec.get("values")
    if allowed is not None:
        if t == "":
            if not spec.get("blanks", True):
                return False
        elif t not in allowed:
            return False
    cond = spec.get("cond")
    if cond:
        op, arg = cond
        v = sheet.value(r, col)
        return _cond(op, arg, v, t)
    return True


def _cond(op, arg, v, t):
    low = t.lower()
    a = str(arg).lower()
    if op == "contains":
        return a in low
    if op == "not_contains":
        return a not in low
    if op == "begins":
        return low.startswith(a)
    if op == "ends":
        return low.endswith(a)
    if op == "equals":
        x, _ = parse_input(str(arg))
        if is_num(x) and is_num(v):
            return v == x
        return low == a
    if op == "not_equals":
        return not _cond("equals", arg, v, t)
    x, _ = parse_input(str(arg))
    if not is_num(x) or not is_num(v):
        return False
    return {">": v > x, ">=": v >= x, "<": v < x, "<=": v <= x}.get(op, True)


def filter_extent(sheet):
    """Extend the autofilter range down to the last used row in its columns."""
    r1, c1, r2, c2 = sheet.autofilter
    last = r2
    has = sheet.has_content
    # grow over rows added directly below (contiguous data only, like Excel)
    while last + 1 < MAX_ROWS and any(has(last + 1, c) for c in range(c1, c2 + 1)):
        last += 1
    return (r1, c1, last, c2)


def compute_filter_hidden(sheet):
    if not sheet.autofilter or not sheet.filters:
        return set()
    r1, c1, r2, c2 = filter_extent(sheet)
    sheet.autofilter = (r1, c1, r2, c2)
    hidden = set()
    for r in range(r1 + 1, r2 + 1):
        for col, spec in sheet.filters.items():
            if not row_passes(sheet, r, spec, col):
                hidden.add(r)
                break
    return hidden
