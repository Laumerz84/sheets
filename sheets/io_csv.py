"""CSV / TSV / TXT reading and writing."""
import csv
import re
import io
import os

from .numfmt import compile_format, format_value, parse_input, repr_number
from .refs import MAX_COLS, MAX_ROWS, key
from .workbook import DEFAULT_STYLE, Formula, Workbook, intern_style, Sheet
from .errors import XLError


def detect_encoding(raw):
    if raw.startswith(b"\xef\xbb\xbf"):
        return "utf-8-sig", True
    if raw.startswith(b"\xff\xfe") or raw.startswith(b"\xfe\xff"):
        return "utf-16", True
    try:
        raw.decode("utf-8")
        return "utf-8", False
    except UnicodeDecodeError as e:
        # a truncated multi-byte char at the very end of the sample is fine
        if e.start >= len(raw) - 4:
            return "utf-8", False
        return "cp1252", False


def detect_delimiter(sample, ext):
    if ext in (".tsv", ".tab"):
        return "\t"
    lines = [ln for ln in sample.splitlines()[:50] if ln.strip()]
    if not lines:
        return ","
    best, best_score = ",", -1
    for d in (",", "\t", ";", "|"):
        counts = []
        for ln in lines:
            try:
                counts.append(len(next(csv.reader([ln], delimiter=d))) - 1)
            except (csv.Error, StopIteration):
                counts.append(0)
        if not counts or max(counts) == 0:
            continue
        common = max(set(counts), key=counts.count)
        consistency = counts.count(common) / len(counts)
        score = consistency * 10 + min(common, 50) / 50
        if common > 0 and score > best_score:
            best, best_score = d, score
    return best


def load_csv(path, progress=None):
    with open(path, "rb") as fh:
        raw = fh.read()
    enc, bom = detect_encoding(raw[:65536])
    try:
        text = raw.decode(enc)
    except UnicodeDecodeError:
        enc = "cp1252"
        text = raw.decode(enc, errors="replace")
    newline = "\r\n" if "\r\n" in text[:65536] else "\n"
    ext = os.path.splitext(path)[1].lower()
    delim = detect_delimiter(text[:65536], ext)

    wb = Workbook()
    name = os.path.splitext(os.path.basename(path))[0][:31] or "Sheet1"
    for ch in "[]:*?/\\":
        name = name.replace(ch, "_")
    sh = Sheet(wb, name)
    wb.sheets.append(sh)
    wb.path = path
    wb.file_format = "csv"
    wb.csv_options = {"encoding": enc, "delimiter": delim, "bom": bom, "newline": newline}

    values = sh.values
    styles = sh.styles
    style_cache = {}
    formulas = []
    rt_cache = {}
    reader = csv.reader(io.StringIO(text, newline=""), delimiter=delim)
    max_c = -1
    r = -1
    total = max(1, len(text))
    for r, row in enumerate(reader):
        if r >= MAX_ROWS:
            break
        if progress and r % 5000 == 0:
            progress(min(0.99, (r * 40) / total if r else 0))
        if len(row) > max_c + 1:
            max_c = min(len(row), MAX_COLS) - 1
        base = r << 14
        for c, field in enumerate(row[:MAX_COLS]):
            if field == "":
                continue
            if field[0] == "=" and len(field) > 1:
                formulas.append((base | c, field))
                continue
            v, fmt = parse_input(field, typed=False)
            if v is None:
                continue
            if v.__class__ is not str and not (
                    (fmt is None and v.__class__ is float and field.isdigit() and (field[0] != "0" or field == "0"))
                    or rt_cache.get((v, fmt, field)) or _remember(rt_cache, v, fmt, field)):
                fmt = _sci_format(field) if fmt is None else None
                if fmt is None or format_value(v, fmt)[0] != field:
                    v, fmt = field, None  # keep exactly what the file says
            values[base | c] = v
            if fmt:
                st = style_cache.get(fmt)
                if st is None:
                    st = style_cache[fmt] = intern_style(DEFAULT_STYLE.with_(numfmt=fmt))
                styles[base | c] = st
    sh.max_row = r
    sh.max_col = max_c
    for k, text_f in formulas:
        sh.formulas[k] = Formula(text_f)
    sh.recompute_extent()
    _auto_widths(sh)
    wb.rebuild_dependencies()
    wb.recalc(full=True)
    return wb


def _round_trips(v, fmt, field):
    """Convert a CSV field only if saving it again gives back the same text
    (so '1/2', 'true', '(5)' or ' 42 ' aren't silently rewritten)."""
    if isinstance(v, bool):
        return field in ("TRUE", "FALSE")
    if fmt:
        return format_value(v, fmt)[0] == field
    return field == repr_number(v)


def _remember(cache, v, fmt, field):
    ok = _round_trips(v, fmt, field)
    if len(cache) < 100000:
        cache[(v, fmt, field)] = ok
    return ok


def _sci_format(field):
    """'1.23457E+11' -> '0.00000E+00' so scientific values keep their look."""
    m = re.match(r"^-?\d(?:\.(\d+))?E([+-])(\d+)$", field)
    if not m:
        return None
    dec = len(m.group(1) or "")
    return "0" + ("." + "0" * dec if dec else "") + "E+" + "0" * len(m.group(3))


def _auto_widths(sh, sample_rows=2000):
    """Rough column widths from content so imported data is readable."""
    widths = {}
    for k, v in sh.values.items():
        r = k >> 14
        if r > sample_rows:
            continue
        c = k & 0x3FFF
        n = len(v) if isinstance(v, str) else 10
        if n > widths.get(c, 0):
            widths[c] = n
    for c, n in widths.items():
        px = min(400, max(64, int(n * 7.2 + 12)))
        if px > 64:
            sh.col_widths[c] = px


def cell_text_for_csv(sh, k):
    v = sh.value_at_key(k)
    if v is None:
        return ""
    if isinstance(v, str):
        return v
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, XLError):
        return v.code
    st = sh.styles.get(k)
    if st is not None and st.numfmt != "General":
        return format_value(v, st.numfmt)[0].strip()
    return repr_number(v)


def save_csv(wb, sheet, path, options=None):
    opts = dict(wb.csv_options or {})
    if options:
        opts.update(options)
    ext = os.path.splitext(path)[1].lower()
    delim = opts.get("delimiter") or ("\t" if ext in (".tsv", ".tab") else ",")
    if ext in (".tsv", ".tab"):
        delim = "\t"
    enc = opts.get("encoding", "utf-8")
    if enc == "utf-16":
        enc = "utf-16"
    elif opts.get("bom") and enc.startswith("utf-8"):
        enc = "utf-8-sig"
    elif enc == "utf-8-sig" and not opts.get("bom"):
        enc = "utf-8"
    newline = opts.get("newline", "\r\n")
    mr, mc = sheet.used_extent(include_styles=False)
    tmp = path + ".tmp~"
    with open(tmp, "w", encoding=enc, newline="", errors="replace") as fh:
        w = csv.writer(fh, delimiter=delim, lineterminator=newline)
        has = sheet.values.keys() | sheet.formulas.keys()
        for r in range(mr + 1):
            w.writerow([cell_text_for_csv(sheet, k) if k in has else ""
                        for k in (key(r, c) for c in range(mc + 1))])
    os.replace(tmp, path)
