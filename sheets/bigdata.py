"""Big-file mode: CSV files with millions of rows (100 million and more) kept in a pyarrow
table of text columns instead of one Python object per cell.

A BigSheet looks like an ordinary Sheet to the rest of Ekxel: `values` and `styles` are
mappings that read through to the file's data, with the user's edits layered on top. The
file itself is never modified in memory; edits live in `BigData.edits`, keyed by *data* row
(the row's position in the file), so they stay with their row through sorting and filtering.

Sorting and filtering don't move any data: they set `order` (data rows in sorted order) and
`view` (the rows shown, top to bottom). Sheet row r shows data row view[r]; the row header
shows the file's row number, like Excel's blue numbers on a filtered list.

Iterating `values` yields only edited cells. Anything that would have to touch every cell
one by one (formatting a whole column of 100M cells, copying it, ...) raises TooBig, which
the window shows as a message; whole-column work that matters (sum/average/count, formulas
like SUM(E:E), sort, filter, find, replace, save) has vectorised versions here."""
import csv
import io
import os
import re
import sys
import threading
import time
from collections import OrderedDict

from .refs import MAX_ROWS, col_name
from .workbook import DEFAULT_STYLE, Sheet, Workbook, intern_style

BIG_FILE_BYTES = 100 * 1024 * 1024  # CSV files bigger than this open in big-file mode
CELL_BUDGET = 2_000_000             # cell-by-cell operations on a big sheet stop above this
ROW_CACHE = 20_000                  # converted rows kept around for painting
BLOCK = 256                         # rows fetched from the table at once

_MISSING = object()


class _Cleared:
    def __repr__(self):
        return "CLEARED"


CLEARED = _Cleared()  # an edit that empties a cell the file has a value in


class TooBig(Exception):
    """An operation that would visit too many cells of a big sheet one by one."""


def arrow():
    import pyarrow as pa
    import pyarrow.compute as pc
    return pa, pc


_POOL = None


def pmap(fn, arr):
    """Apply an Arrow kernel function to each chunk of a ChunkedArray on all CPU cores
    (Arrow kernels release the GIL) and return a ChunkedArray with the same chunking."""
    global _POOL
    pa, pc = arrow()
    if not isinstance(arr, pa.ChunkedArray):
        return fn(arr)
    chunks = arr.chunks
    if len(chunks) <= 1:
        return pa.chunked_array([fn(ch) for ch in chunks]) if chunks else fn(arr)
    if _POOL is None:
        from concurrent.futures import ThreadPoolExecutor
        _POOL = ThreadPoolExecutor(max(2, min(16, os.cpu_count() or 4)))
    return pa.chunked_array(list(_POOL.map(fn, chunks)))


def pmap2(fn, a, b):
    """pmap over two ChunkedArrays chunked the same way (a column and its numeric version)."""
    global _POOL
    pa, pc = arrow()
    if not isinstance(a, pa.ChunkedArray) or a.num_chunks != b.num_chunks or a.num_chunks <= 1:
        return fn(a, b)
    if _POOL is None:
        pmap(lambda x: x, pa.chunked_array([pa.array([1]), pa.array([2])]))
    return pa.chunked_array(list(_POOL.map(fn, a.chunks, b.chunks)))


def arange(n):
    """Int64Array 0..n-1 built in C (a Python range is slow at 100M)."""
    pa, pc = arrow()
    if n <= 0:
        return pa.array([], pa.int64())
    return pc.subtract(pc.cumulative_sum(pa.repeat(pa.scalar(1, pa.int64()), n)), 1)


def _position_mask(n, positions):
    """Boolean array of length n, true at the given sorted positions."""
    pa, pc = arrow()
    return pc.is_in(arange(n), value_set=pa.array(positions, pa.int64()))


def is_big(sheet):
    return getattr(sheet, "big", None) is not None


def check_area(sheet, rect, what="That"):
    """Raise TooBig when rect (clamped to the used part) is too many cells for cell-by-cell work."""
    if not is_big(sheet):
        return
    r1, c1, r2, c2 = rect
    r2 = min(r2, max(sheet.max_row, r1))
    c2 = min(c2, max(sheet.max_col, c1))
    n = (r2 - r1 + 1) * (c2 - c1 + 1)
    if n > CELL_BUDGET:
        raise TooBig(f"{what} covers {n:,} cells. In big-file mode Ekxel changes at most "
                     f"{CELL_BUDGET:,} cells one by one; select fewer rows. (Sum/Average/Count, "
                     "sorting, filtering, Find/Replace and saving work on the whole file.)")


# ================================================================ memory
def free_memory():
    """Bytes of physical memory available now (an estimate off Windows), or None."""
    try:
        if sys.platform == "win32":
            import ctypes

            class MS(ctypes.Structure):
                _fields_ = [("len", ctypes.c_ulong), ("load", ctypes.c_ulong)] + \
                           [(n, ctypes.c_ulonglong) for n in ("tp", "ap", "tf", "af", "tv", "av", "ae")]
            ms = MS()
            ms.len = ctypes.sizeof(MS)
            ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(ms))
            return ms.ap
        if sys.platform == "darwin":
            import subprocess
            total = int(subprocess.run(["sysctl", "-n", "hw.memsize"], capture_output=True,
                                       text=True, timeout=5).stdout.strip())
            return int(total * 0.6)  # macOS frees/compresses memory on demand
        with open("/proc/meminfo") as fh:
            for line in fh:
                if line.startswith("MemAvailable:"):
                    return int(line.split()[1]) * 1024
    except Exception:
        pass
    return None


# ================================================================ loading
def wants_big(path):
    try:
        return os.path.getsize(path) > BIG_FILE_BYTES
    except OSError:
        return False


def load_big_csv(path, progress=None):
    from .io_csv import detect_delimiter, detect_encoding
    pa, pc = arrow()
    import pyarrow.csv as pcsv

    size = os.path.getsize(path)
    with open(path, "rb") as fh:
        head = fh.read(1 << 20)
    enc, bom = detect_encoding(head[:65536])
    text = head.decode("utf-8" if enc == "utf-8-sig" else enc, errors="replace").lstrip("﻿")
    newline = "\r\n" if "\r\n" in text[:65536] else "\n"
    ext = os.path.splitext(path)[1].lower()
    delim = detect_delimiter(text[:65536], ext)
    sample = text[:text.rfind("\n") + 1] if size > len(head) else text
    try:
        ncols = max((len(row) for row in csv.reader(io.StringIO(sample, newline=""), delimiter=delim)), default=1)
    except csv.Error:
        ncols = max(1, sample.split("\n", 1)[0].count(delim) + 1)
    ncols = max(1, ncols)

    need = int(size * 1.8) + (512 << 20)
    free = free_memory()
    if free is not None and need > free:
        raise MemoryError(f"This file is {size / 1e9:.1f} GB. Opening it needs about {need / 1e9:.0f} GB "
                          f"of free memory and only {free / 1e9:.1f} GB is free. Close other programs "
                          "and try again.")

    names = [f"c{i}" for i in range(ncols)]
    arrow_enc = {"utf-8": "utf8", "utf-8-sig": "utf8"}.get(enc, enc)

    def read(newlines):
        skipped = []

        def bad_row(row):
            if len(skipped) < 1_000_000:
                skipped.append(row.number)
            return "skip"
        tbl = pcsv.read_csv(
            path,
            read_options=pcsv.ReadOptions(column_names=names, block_size=1 << 24, encoding=arrow_enc),
            parse_options=pcsv.ParseOptions(delimiter=delim, newlines_in_values=newlines, invalid_row_handler=bad_row),
            convert_options=pcsv.ConvertOptions(column_types={n: pa.string() for n in names},
                                                strings_can_be_null=False, quoted_strings_can_be_null=False))
        return tbl, skipped

    result = {}

    def work():
        try:
            result["v"] = read(False)
            if result["v"][1] and '"' in text:
                # quoted fields with line breaks inside: slower but exact
                again = read(True)
                if len(again[1]) < len(result["v"][1]):
                    result["v"] = again
        except BaseException as e:  # noqa: BLE001 - re-raised below in the caller's thread
            result["e"] = e
    t0 = time.time()
    th = threading.Thread(target=work, daemon=True)
    th.start()
    expected = size / 0.8e9 + 1.0
    while th.is_alive():
        th.join(0.1)
        if progress:
            progress(min(0.95, (time.time() - t0) / expected * 0.95))
    if "e" in result:
        raise result["e"]
    table, skipped = result["v"]
    if table.num_rows > MAX_ROWS:
        raise MemoryError(f"This file has {table.num_rows:,} rows; Ekxel sheets hold at most {MAX_ROWS:,}.")

    wb = Workbook()
    name = os.path.splitext(os.path.basename(path))[0][:31] or "Sheet1"
    for ch in "[]:*?/\\":
        name = name.replace(ch, "_")
    big = BigData(table)
    sh = BigSheet(wb, name, big)
    wb.sheets.append(sh)
    wb.path = path
    wb.file_format = "csv"
    wb.csv_options = {"encoding": enc, "delimiter": delim, "bom": bom, "newline": newline}
    wb.big_skipped = len(skipped)
    sh.recompute_extent()
    # column widths from the first rows, like the normal loader
    sample_rows = table.slice(0, 2000)
    for c in range(ncols):
        lens = pc.utf8_length(sample_rows.column(c))
        n = pc.max(lens).as_py() or 0
        px = min(400, max(64, int(n * 7.2 + 12)))
        if px > 64:
            sh.col_widths[c] = px
    return wb


# ================================================================ the data
class BigData:
    def __init__(self, table):
        self.table = table
        self.cols = [table.column(i) for i in range(table.num_columns)]
        self.ncols = table.num_columns
        self.N = table.num_rows
        self.order = None   # data rows in sorted order (Int64Array) or None = file order
        self.view = None    # data rows shown (sorted + filtered) or None = file order
        self._vm = None     # memoryview of view for fast indexing
        self.edits = {}     # (data_row << 14) | col -> value or CLEARED
        self._rows = OrderedDict()   # data row -> tuple of (value, fmt) or None per column
        self._num = {}      # col -> float64 ChunkedArray (null where not a number)
        self._inv = None    # data row -> sheet row, for edited rows while a view is set
        self._rt = {}       # round-trip cache for field conversion
        self._fmt_styles = {}
        self.lock = threading.RLock()
        self.sheet = None       # the BigSheet showing this data
        self.calc_vals = {}     # col -> (kind, num, err) arrays of a calculated column (bigcalc.py)
        self.calc_defs = {}     # col -> {"text", "anchor", "inputs"}
        self.calc_auto = set()  # edit keys bigcalc.recompute_row filled in (not typed by the user)

    # ------------------------------------------------------------ rows
    @property
    def nview(self):
        return self.N if self._vm is None else len(self._vm)

    def drow(self, r):
        """Data row behind sheet row r (rows past the data are new rows the user typed in)."""
        vm = self._vm
        if vm is None:
            return r
        nv = len(vm)
        return vm[r] if r < nv else self.N + (r - nv)

    def srow(self, d):
        """Sheet row showing data row d, or None when it's filtered out."""
        if self._vm is None:
            return d
        if d >= self.N:
            return len(self._vm) + (d - self.N)
        if self._inv is None:
            self._build_inv()
        return self._inv.get(d)

    def _build_inv(self):
        rows = sorted({dk >> 14 for dk in self.edits if (dk >> 14) < self.N})
        inv = {}
        if rows:
            pa, pc = arrow()
            mask = pc.is_in(self.view, value_set=pa.array(rows, pa.int64()))
            pos = pc.indices_nonzero(mask)
            inv = dict(zip(pc.take(self.view, pos).to_pylist(), pos.to_pylist()))
        self._inv = inv

    def set_state(self, state):
        """state = (order, view) Int64Arrays or Nones."""
        order, view = state
        self.order = order
        self.view = view
        if view is None:
            self._vm = None
        else:
            pa, pc = arrow()
            arr = view.combine_chunks() if isinstance(view, pa.ChunkedArray) else view
            if arr.type != pa.int64():
                arr = arr.cast(pa.int64())
            self.view = arr
            buf = arr.buffers()[1]
            self._vm = memoryview(buf)[arr.offset * 8:(arr.offset + len(arr)) * 8].cast("q")
        self._inv = None

    def state(self):
        return (self.order, self.view)

    # ------------------------------------------------------------ cells
    def backing(self, d, c, r=None):
        """(value, numfmt) the file has at data row d, column c, or None when empty."""
        if d >= self.N or c >= self.ncols:
            return None
        if self.calc_vals and c in self.calc_vals:
            from .bigcalc import NOT_CALC, cell
            v = cell(self, d, c)
            if v is not NOT_CALC:
                return None if v is None else (v, None)
        row = self._rows.get(d)
        if row is None:
            self._fetch(d, r)
            row = self._rows.get(d)
            if row is None:
                return None
        return row[c]

    def _fetch(self, d, r):
        """Convert a block of rows around d (around sheet row r when a view is set)."""
        from .io_csv import convert_field
        pa, pc = arrow()
        with self.lock:
            if self._vm is None or r is None:
                start = (d // BLOCK) * BLOCK
                n = min(BLOCK, self.N - start)
                ds = list(range(start, start + n))
                cols = [col.slice(start, n).to_pylist() for col in self.cols]
            else:
                start = (r // BLOCK) * BLOCK
                n = max(0, min(BLOCK, len(self._vm) - start))
                ds = list(self._vm[start:start + n])
                if d not in ds:
                    ds = [d]
                idx = pa.array(ds, pa.int64())
                cols = [pc.take(col, idx).to_pylist() for col in self.cols]
            rt = self._rt
            rows = self._rows
            for i, dd in enumerate(ds):
                line = []
                for col in cols:
                    f = col[i]
                    line.append(convert_field(f, rt) if f else None)
                rows[dd] = tuple(line)
                rows.move_to_end(dd)
            while len(rows) > ROW_CACHE:
                rows.popitem(last=False)
            if len(rt) > 200_000:
                rt.clear()

    def fmt_style(self, fmt):
        st = self._fmt_styles.get(fmt)
        if st is None:
            st = self._fmt_styles[fmt] = intern_style(DEFAULT_STYLE.with_(numfmt=fmt))
        return st

    def set_edit(self, r, c, v):
        d = self.drow(r)
        self.edits[(d << 14) | c] = v
        if self._inv is not None and d < self.N:
            self._inv[d] = r
        self.after_edit(d, c)

    def after_edit(self, d, c):
        """A cell the user changed: it's no longer an auto-computed one, and calculated columns
        reading it get that row recomputed."""
        self.calc_auto.discard((d << 14) | c)
        if self.calc_defs and self.sheet is not None:
            from .bigcalc import recompute_row
            recompute_row(self.sheet, d, c)

    # ------------------------------------------------------------ numbers
    def numeric(self, c):
        """Column c of the file as float64 (null where the cell isn't a number), cached.
        Plain numbers, 1,234.5 / $12.50 / 45% and ISO dates (as date serials)."""
        got = self._num.get(c)
        if got is not None:
            return got
        self._num[c] = got = pmap(_to_numbers, self.cols[c])
        return got

    def rows_array(self, r1, r2):
        """Data rows for sheet rows r1..r2 (within the data part), as an Int64Array or a slice spec."""
        pa, pc = arrow()
        n = r2 - r1 + 1
        if self.view is None:
            return None, r1, n
        return self.view.slice(r1, n), r1, n

    def column_part(self, arr, r1, r2):
        """arr (aligned to data rows) restricted to sheet rows r1..r2 of the current view."""
        pa, pc = arrow()
        n = r2 - r1 + 1
        if n <= 0:
            return arr.slice(0, 0)
        if self.view is None:
            return arr.slice(r1, n)
        return pc.take(arr, self.view.slice(r1, n))


def _to_numbers(col):
    pa, pc = arrow()
    none = pa.scalar(None, pa.string())
    plain = pc.match_substring_regex(col, r"^-?(0|[1-9]\d*)(\.\d+)?([eE][-+]?\d+)?$")
    x = pc.cast(pc.if_else(plain, col, none), pa.float64())
    if pc.any(pc.match_substring_regex(col, r"[$%,]")).as_py():
        fancy = pc.match_substring_regex(col, r"^-?\$?(\d{1,3}(,\d{3})+|\d+)(\.\d+)?%?$")
        fancy = pc.and_(fancy, pc.invert(plain))
        stripped = pc.replace_substring_regex(pc.if_else(fancy, col, none), r"[$,%]", "")
        y = pc.cast(stripped, pa.float64())
        y = pc.if_else(pc.ends_with(col, "%"), pc.divide(y, 100.0), y)
        x = pc.if_else(fancy, y, x)
    dates = pc.match_substring_regex(col, r"^\d{4}-\d{2}-\d{2}$")
    if pc.any(dates).as_py():
        ts = pc.strptime(pc.if_else(dates, col, none), format="%Y-%m-%d", unit="s", error_is_null=True)
        days = pc.add(pc.divide(pc.cast(pc.cast(ts, pa.int64()), pa.float64()), 86400.0), 25569.0)
        x = pc.if_else(pc.is_valid(days), days, x)
    return x


# ================================================================ mappings
class BigValues:
    """Sheet-keyed cell values: the file's data with the user's edits on top.
    Iteration yields edited cells only (never the 100M file cells)."""

    def __init__(self, big):
        self.big = big

    def get(self, k, default=None):
        big = self.big
        r = k >> 14
        c = k & 0x3FFF
        d = big.drow(r)
        e = big.edits.get((d << 14) | c, _MISSING)
        if e is not _MISSING:
            return default if e is CLEARED else e
        b = big.backing(d, c, r)
        return default if b is None else b[0]

    def __contains__(self, k):
        return self.get(k, _MISSING) is not _MISSING

    def __getitem__(self, k):
        v = self.get(k, _MISSING)
        if v is _MISSING:
            raise KeyError(k)
        return v

    def __setitem__(self, k, v):
        self.big.set_edit(k >> 14, k & 0x3FFF, v)

    def pop(self, k, default=_MISSING):
        big = self.big
        v = self.get(k, _MISSING)
        r, c = k >> 14, k & 0x3FFF
        d = big.drow(r)
        if big.backing(d, c, r) is not None:
            big.set_edit(r, c, CLEARED)
        else:
            big.edits.pop((d << 14) | c, None)
            big.after_edit(d, c)
        if v is _MISSING:
            if default is _MISSING:
                raise KeyError(k)
            return default
        return v

    def __delitem__(self, k):
        self.pop(k)

    def setdefault(self, k, default=None):
        v = self.get(k, _MISSING)
        if v is _MISSING:
            self[k] = default
            return default
        return v

    def __iter__(self):
        big = self.big
        for dk, v in list(big.edits.items()):
            if v is CLEARED:
                continue
            r = big.srow(dk >> 14)
            if r is not None:
                yield (r << 14) | (dk & 0x3FFF)

    def keys(self):
        return set(self)

    def items(self):
        return [(k, self.get(k)) for k in self]

    def values(self):
        return [v for _, v in self.items()]

    def update(self, other=(), **kw):
        for k, v in dict(other).items():
            self[k] = v

    def copy(self):
        return dict(self.items())

    def __len__(self):
        # a size, not a count of what iteration yields: callers use it to pick strategies
        return self.big.nview * self.big.ncols + len(self.big.edits)

    def __bool__(self):
        return bool(self.big.N and self.big.ncols) or bool(self.big.edits)

    def snapshot(self):
        return BigSnapshot(dict(self.big.edits), self.big.state())


class BigSnapshot:
    __slots__ = ("edits", "state")

    def __init__(self, edits, state):
        self.edits = edits
        self.state = state


class BigStyles:
    """Sheet-keyed styles: explicit ones (by position) plus the number formats the file's
    values imply (dates, currency...) for cells the user hasn't edited."""

    def __init__(self, big):
        self.big = big
        self.d = {}

    def get(self, k, default=None):
        s = self.d.get(k)
        if s is not None:
            return s
        big = self.big
        r = k >> 14
        c = k & 0x3FFF
        d = big.drow(r)
        if ((d << 14) | c) in big.edits:
            return default
        b = big.backing(d, c, r)
        if b is None or b[1] is None:
            return default
        return big.fmt_style(b[1])

    def __contains__(self, k):
        return self.get(k) is not None

    def __getitem__(self, k):
        s = self.get(k)
        if s is None:
            raise KeyError(k)
        return s

    def __setitem__(self, k, v):
        self.d[k] = v

    def pop(self, k, default=None):
        old = self.d.pop(k, None)
        if self.get(k) is not None:
            self.d[k] = DEFAULT_STYLE  # hide the format the file's value implies
        return old if old is not None else default

    def __delitem__(self, k):
        self.pop(k)

    def __iter__(self):
        return iter(list(self.d))

    def keys(self):
        return self.d.keys()

    def items(self):
        return self.d.items()

    def values(self):
        return self.d.values()

    def copy(self):
        return dict(self.d)

    def __len__(self):
        return len(self.d) + 1

    def __bool__(self):
        return True


# ================================================================ the sheet
class BigSheet(Sheet):
    def __init__(self, wb, name, big):
        self.big = big
        big.sheet = self
        self._values = BigValues(big)
        self._styles = BigStyles(big)
        super().__init__(wb, name)

    @property
    def values(self):
        return self._values

    @values.setter
    def values(self, v):
        if isinstance(v, BigSnapshot):
            self.big.edits = dict(v.edits)
            self.big.set_state(v.state)
        elif isinstance(v, BigValues):
            self._values = v
        elif not v:
            pass  # Sheet.__init__
        else:
            raise TooBig("Inserting or deleting rows or columns isn't available in big-file mode.")

    @property
    def styles(self):
        return self._styles

    @styles.setter
    def styles(self, v):
        if isinstance(v, BigStyles):
            self._styles = v
        else:
            self._styles.d = dict(v)

    @property
    def big_state(self):
        return self.big.state()

    @big_state.setter
    def big_state(self, state):
        self.big.set_state(state)
        self.recompute_extent()
        if any(s.formulas for s in self.wb.sheets):
            self.wb.recalc(full=True)  # formulas see the rows as now shown

    @property
    def big_calc(self):
        from .bigcalc import state
        return state(self)

    @big_calc.setter
    def big_calc(self, st):
        from .bigcalc import set_state
        set_state(self, st)

    def formula_text(self, r, c):
        f = self.formulas.get((r << 14) | c)
        if f is not None:
            return f.text
        info = self.big.calc_defs.get(c)
        if info and info["text"]:
            d = self.big.drow(r)
            if ((d << 14) | c) in self.big.edits:
                return None
            from .bigcalc import NOT_CALC, cell
            if cell(self.big, d, c) is not NOT_CALC:
                from .formula import shift_formula
                return shift_formula(info["text"], r - info["anchor"], 0)
        return None

    @property
    def big_cols(self):
        return list(self.big.cols)

    @big_cols.setter
    def big_cols(self, cols):
        set_columns(self, cols)

    def row_label(self, r):
        return self.big.drow(r) + 1

    def _extent(self, include_styles):
        big = self.big
        mr = big.nview - 1 if (big.N and big.ncols) else -1
        mc = big.ncols - 1 if big.N else -1
        keys = [iter(self._values), iter(self.formulas)]
        if include_styles:
            keys.append(iter(self._styles.d))
        for it in keys:
            for k in it:
                r, c = k >> 14, k & 0x3FFF
                if r > mr:
                    mr = r
                if c > mc:
                    mc = c
        return mr, mc

    def recompute_extent(self):
        self.max_row, self.max_col = self._extent(False)

    def used_extent(self, include_styles=True):
        return self._extent(include_styles)

    def keys_in(self, r1, c1, r2, c2):
        r2 = min(r2, self.max_row)
        c2 = min(c2, self.max_col)
        check_area(self, (r1, c1, r2, c2))
        vals, forms = self._values, self.formulas
        for r in range(r1, r2 + 1):
            base = r << 14
            for c in range(c1, c2 + 1):
                k = base | c
                if k in forms or k in vals:
                    yield k

    def content_keys(self):
        return set(self._values) | set(self.formulas)

    def _structural(self, axis, at, count):
        raise TooBig("Inserting or deleting rows or columns isn't available in big-file mode.")


# ================================================================ statistics
def _segments(n, excluded):
    """[start, stop) pieces of range(n) that skip the sorted positions in excluded."""
    out = []
    a = 0
    for p in excluded:
        if p > a:
            out.append((a, p))
        a = p + 1
    if a < n:
        out.append((a, n))
    return out


def overrides(sheet, r1, c1, r2, c2):
    return _overrides(sheet, r1, c1, r2, c2)


def _overrides(sheet, r1, c1, r2, c2):
    """{col: {sheet_row: value}} for edited or formula cells inside the rect (sheet coords)."""
    big = sheet.big
    out = {}
    for dk in list(big.edits):
        r = big.srow(dk >> 14)
        c = dk & 0x3FFF
        if r is not None and r1 <= r <= r2 and c1 <= c <= c2:
            out.setdefault(c, {})[r] = sheet.value(r, c)
    for k in list(sheet.formulas):
        r, c = k >> 14, k & 0x3FFF
        if r1 <= r <= r2 and c1 <= c <= c2:
            out.setdefault(c, {})[r] = sheet.value_at_key(k)
    return out


def range_stats(sheet, r1, c1, r2, c2, over=None):
    """Whole-range numbers for a big sheet without visiting cells one by one.
    Returns dict(count=non-empty cells, nums=numbers, sum, min, max). `over`: overrides()
    computed beforehand (on the main thread, since it evaluates formulas)."""
    from .values import is_num
    pa, pc = arrow()
    big = sheet.big
    r2 = min(r2, sheet.max_row)
    c2 = min(c2, sheet.max_col)
    out = {"count": 0, "nums": 0, "sum": 0.0, "min": None, "max": None}
    if r2 < r1 or c2 < c1:
        return out

    def add_num(x):
        out["nums"] += 1
        out["sum"] += x
        out["min"] = x if out["min"] is None else min(out["min"], x)
        out["max"] = x if out["max"] is None else max(out["max"], x)

    if over is None:
        over = _overrides(sheet, r1, c1, r2, c2)
    dr2 = min(r2, big.nview - 1)  # sheet rows backed by the file
    for c in range(c1, c2 + 1):
        oc = over.get(c, {})
        if c < big.ncols and dr2 >= r1:
            num = big.column_part(big.numeric(c), r1, dr2)
            txt = big.column_part(big.cols[c], r1, dr2)
            n = dr2 - r1 + 1
            excl = sorted(r - r1 for r in oc if r1 <= r <= dr2)
            for a, b in _segments(n, excl):
                part = num.slice(a, b - a)
                k = pc.count(part).as_py()
                if k:
                    out["nums"] += k
                    out["sum"] += pc.sum(part).as_py() or 0.0
                    mm = pc.min_max(part).as_py()
                    out["min"] = mm["min"] if out["min"] is None else min(out["min"], mm["min"])
                    out["max"] = mm["max"] if out["max"] is None else max(out["max"], mm["max"])
                out["count"] += pc.sum(pmap(lambda ch: pc.not_equal(ch, ""), txt.slice(a, b - a))).as_py() or 0
        for r, v in oc.items():
            if v is None or v == "":
                continue
            out["count"] += 1
            if is_num(v):
                add_num(v)
    return out


AGG_FUNCS = {"SUM", "COUNT", "COUNTA", "AVERAGE", "MIN", "MAX"}


def aggregate(name, args):
    """Fast path for SUM/COUNT/COUNTA/AVERAGE/MIN/MAX when an argument is a big range.
    Returns the result, or _MISSING to let the normal code run."""
    from . import errors
    from .values import RangeRef, flat, is_num
    if not any(isinstance(a, RangeRef) and is_big(a.sheet) and
               (min(a.r2, a.sheet.max_row) - a.r1 + 1) * (a.c2 - a.c1 + 1) > 20_000 for a in args):
        return _MISSING
    tot = {"count": 0, "nums": 0, "sum": 0.0, "min": None, "max": None}
    for a in args:
        if isinstance(a, RangeRef) and is_big(a.sheet):
            s = range_stats(a.sheet, a.r1, a.c1, a.r2, a.c2)
            tot["count"] += s["count"]
            tot["nums"] += s["nums"]
            tot["sum"] += s["sum"]
            for m, f in (("min", min), ("max", max)):
                if s[m] is not None:
                    tot[m] = s[m] if tot[m] is None else f(tot[m], s[m])
            continue
        direct = not isinstance(a, (RangeRef, list))
        for v in flat(a):
            if isinstance(v, errors.XLError) and name != "COUNT" and name != "COUNTA":
                raise v
            if v is None or v == "":
                continue
            tot["count"] += 1
            if is_num(v) and not isinstance(v, bool):
                tot["nums"] += 1
                tot["sum"] += v
                tot["min"] = v if tot["min"] is None else min(tot["min"], v)
                tot["max"] = v if tot["max"] is None else max(tot["max"], v)
            elif direct and isinstance(v, (bool, str)) and name in ("SUM", "AVERAGE", "MIN", "MAX"):
                from .values import to_num
                x = to_num(v)
                tot["nums"] += 1
                tot["sum"] += x
                tot["min"] = x if tot["min"] is None else min(tot["min"], x)
                tot["max"] = x if tot["max"] is None else max(tot["max"], x)
    if name == "SUM":
        return float(tot["sum"])
    if name == "COUNT":
        return float(tot["nums"])
    if name == "COUNTA":
        return float(tot["count"])
    if name == "AVERAGE":
        if not tot["nums"]:
            raise errors.DIV0
        return tot["sum"] / tot["nums"]
    if name == "MIN":
        return float(tot["min"] or 0.0)
    if name == "MAX":
        return float(tot["max"] or 0.0)
    return _MISSING


IFS_FUNCS = {"SUMIF", "SUMIFS", "COUNTIF", "COUNTIFS", "AVERAGEIF", "AVERAGEIFS", "MAXIFS", "MINIFS"}


def _crit_mask(col, num, crit):
    """Arrow version of functions.make_criteria for one chunk: text column + its numbers."""
    from .functions import _has_wild, _wild_compile
    from .numfmt import parse_input
    from .values import MISSING, is_num, scalar
    from .errors import XLError
    pa, pc = arrow()
    crit = scalar(crit)
    if isinstance(crit, XLError):
        raise crit
    isnum = pc.is_valid(num)
    blank = pc.equal(col, "")
    is_text = pc.and_(pc.invert(isnum), pc.invert(blank))
    if crit is None or crit is MISSING:
        return blank
    if isinstance(crit, bool):
        return pc.match_substring_regex(col, "^" + ("TRUE" if crit else "FALSE") + "$", ignore_case=True)
    if is_num(crit):
        return pc.fill_null(pc.equal(num, float(crit)), False)
    st = str(crit)
    op = "="
    for p_ in (">=", "<=", "<>", ">", "<", "="):
        if st.startswith(p_):
            op, st = p_, st[len(p_):]
            break
    x, _ = parse_input(st) if st.strip() else (None, None)
    if op in ("=", "<>"):
        if st == "":
            m = blank
        elif is_num(x):
            m = pc.or_(pc.fill_null(pc.equal(num, float(x)), False),
                       pc.match_substring_regex(col, r"^\s*" + re.escape(st.strip()) + r"\s*$", ignore_case=True))
        elif st.upper() in ("TRUE", "FALSE"):
            m = pc.match_substring_regex(col, "^" + st.upper() + "$", ignore_case=True)
        elif _has_wild(st):
            m = pc.and_(is_text, pc.match_substring_regex(col, "^(?s:" + _wild_compile(st).pattern + ")$",
                                                          ignore_case=True))
        else:
            m = pc.and_(is_text, pc.match_substring_regex(col, "^" + re.escape(st) + "$", ignore_case=True))
        return pc.invert(m) if op == "<>" else m
    fn = {">": pc.greater, "<": pc.less, ">=": pc.greater_equal, "<=": pc.less_equal}[op]
    if is_num(x):
        return pc.fill_null(fn(num, float(x)), False)
    return pc.and_(is_text, pc.fill_null(fn(pc.utf8_lower(col), st.lower()), False))


def ifs(name, target, pairs):
    """SUMIF(S)/COUNTIF(S)/AVERAGEIF(S)/MAXIFS/MINIFS over single-column ranges of one big sheet,
    vectorised. target: the range summed/averaged (None for COUNTIF(S)). Returns the result, or
    _MISSING when the arguments don't fit this path (the normal code then runs)."""
    from . import errors
    from .functions import make_criteria
    from .values import RangeRef, is_num
    pa, pc = arrow()
    rngs = [r for r, _ in pairs] + ([target] if target is not None else [])
    if not all(isinstance(r, RangeRef) for r in rngs):
        return _MISSING
    sheet = rngs[0].sheet
    if not is_big(sheet) or any(r.sheet is not sheet or r.c1 != r.c2 for r in rngs):
        return _MISSING
    big = sheet.big
    r1 = pairs[0][0].r1
    height = pairs[0][0].r2 - r1 + 1
    if any(r.r2 - r.r1 + 1 != height for r, _ in pairs):
        raise errors.VALUE
    if target is not None:
        t_r1 = target.r1
    if any(r.r1 != r1 for r, _ in pairs) or (target is not None and t_r1 != r1):
        return _MISSING  # offset ranges: rare, leave to the normal code (guarded by the cell budget)
    r2 = min(r1 + height - 1, sheet.max_row)
    cols = sorted({r.c1 for r in rngs})
    over = _overrides(sheet, r1, min(cols), r2, max(cols))
    special = sorted({r for c in cols for r in over.get(c, {})})
    preds = [(r.c1, make_criteria(cr)) for r, cr in pairs]
    tc = target.c1 if target is not None else None

    total, count, lo, hi = 0.0, 0, None, None
    dr2 = min(r2, big.nview - 1)
    if dr2 >= r1:
        n = dr2 - r1 + 1
        mask = None
        for (rng, cr) in pairs:
            c = rng.c1
            if c < big.ncols:
                m = pmap2(lambda ch, nm, cr=cr: _crit_mask(ch, nm, cr),
                          big.column_part(big.cols[c], r1, dr2), big.column_part(big.numeric(c), r1, dr2))
            else:  # a column past the file's data: every row is blank there
                m = pa.repeat(pa.scalar(bool(make_criteria(cr)(None))), n)
            mask = m if mask is None else pc.and_(mask, m)
        if isinstance(mask, pa.ChunkedArray):
            mask = mask.combine_chunks()
        inside = [r - r1 for r in special if r <= dr2]
        if inside:
            mask = pc.replace_with_mask(mask, _position_mask(n, inside), pa.repeat(False, len(inside)))
        if target is None:
            count += pc.sum(mask).as_py() or 0
        elif tc < big.ncols:
            vals = pc.filter(big.column_part(big.numeric(tc), r1, dr2), mask)
            k = pc.count(vals).as_py()
            if k:
                count += k
                total += pc.sum(vals).as_py() or 0.0
                mm = pc.min_max(vals).as_py()
                lo, hi = mm["min"], mm["max"]
    # edited/formula cells and rows past the data, the normal (cell by cell) way
    rest = [r for r in special if r <= dr2] + [r for r in range(max(r1, dr2 + 1), r2 + 1)]
    for r in rest:
        if all(pred(sheet.value(r, c)) for c, pred in preds):
            if target is None:
                count += 1
                continue
            v = sheet.value(r, tc)
            if is_num(v) and not isinstance(v, bool):
                count += 1
                total += v
                lo = v if lo is None else min(lo, v)
                hi = v if hi is None else max(hi, v)
    if target is None and r1 + height - 1 > r2 and all(pred(None) for _, pred in preds):
        count += (r1 + height - 1) - r2  # blank rows below the used area match "" criteria too
    if name in ("SUMIF", "SUMIFS"):
        return float(total)
    if name in ("COUNTIF", "COUNTIFS"):
        return float(count)
    if name in ("AVERAGEIF", "AVERAGEIFS"):
        if not count:
            raise errors.DIV0
        return total / count
    if name == "MAXIFS":
        return float(hi if hi is not None else 0.0)
    if name == "MINIFS":
        return float(lo if lo is not None else 0.0)
    return _MISSING


class LazyColumn:
    """A tall single-column range of a big sheet as a sequence for the lookup functions
    (VLOOKUP/MATCH/XLOOKUP...): binary searches read only the few cells they visit, and exact
    matches run vectorised through find_exact. Iterating it whole raises #CALC! instead of
    reading millions of cells one by one."""

    def __init__(self, sheet, r1, r2, c):
        self.sheet, self.r1, self.c = sheet, r1, c
        self.n = max(0, min(r2, sheet.max_row) - r1 + 1)

    def __len__(self):
        return self.n

    def __getitem__(self, i):
        from . import errors
        if isinstance(i, slice):
            raise errors.CALC
        if i < 0:
            i += self.n
        if not 0 <= i < self.n:
            raise IndexError(i)
        return self.sheet.value(self.r1 + i, self.c)

    def __iter__(self):
        from . import errors
        raise errors.CALC

    def _search(self, arrow_mask, pred, pick):
        """Positions where the cell matches: arrow_mask(text_chunk, number_chunk) for the file's
        cells, pred(value) for edited/formula cells and rows past the data; pick chooses one."""
        pa, pc = arrow()
        big = self.sheet.big
        r1, r2 = self.r1, self.r1 + self.n - 1
        c = self.c
        special = sorted(over_rows for over_rows in _overrides(self.sheet, r1, c, r2, c).get(c, {}))
        dr2 = min(r2, big.nview - 1)
        found = []
        if dr2 >= r1 and c < big.ncols:
            n = dr2 - r1 + 1
            mask = pmap2(arrow_mask, big.column_part(big.cols[c], r1, dr2), big.column_part(big.numeric(c), r1, dr2))
            if isinstance(mask, pa.ChunkedArray):
                mask = mask.combine_chunks()
            inside = [r - r1 for r in special if r <= dr2]
            if inside:
                mask = pc.replace_with_mask(mask, _position_mask(n, inside), pa.repeat(False, len(inside)))
            found = pick(mask)
        rest = [r for r in special if r <= dr2] + list(range(max(r1, dr2 + 1), r2 + 1))
        for r in rest:
            if pred(self.sheet.value(r, c)):
                found.append(r - r1)
        return found

    def find_exact(self, x, wildcard=True, reverse=False):
        from .functions import _has_wild, _lookup_eq, _wild_compile, wild_match
        from .values import is_num
        pa, pc = arrow()

        def text_mask(rx):
            def m(ch, nm):
                is_text = pc.and_(pc.invert(pc.is_valid(nm)), pc.not_equal(ch, ""))
                return pc.and_(is_text, pc.match_substring_regex(ch, rx, ignore_case=True))
            return m
        if isinstance(x, str) and wildcard and _has_wild(x):
            amask = text_mask("^(?s:" + _wild_compile(x).pattern + ")$")
            pred = lambda v: isinstance(v, str) and wild_match(x, v)
        elif isinstance(x, str):
            amask = text_mask("^" + re.escape(x) + "$")
            xl = x.lower()
            pred = lambda v: isinstance(v, str) and v.lower() == xl
        elif isinstance(x, bool):
            amask = lambda ch, nm: pc.match_substring_regex(ch, "^" + ("TRUE" if x else "FALSE") + "$", ignore_case=True)
            pred = lambda v: _lookup_eq(v, x)
        elif is_num(x):
            amask = lambda ch, nm: pc.fill_null(pc.equal(nm, float(x)), False)
            pred = lambda v: _lookup_eq(v, x)
        else:
            return -1

        def pick(mask):
            if reverse:
                idx = pc.indices_nonzero(mask)
                return [idx[len(idx) - 1].as_py()] if len(idx) else []
            i = pc.index(mask, True).as_py()
            return [i] if i >= 0 else []
        found = self._search(amask, pred, pick)
        if not found:
            return -1
        return max(found) if reverse else min(found)

    def find_nearest(self, x, mode):
        """XLOOKUP/XMATCH match_mode -1 (next smaller) / 1 (next larger) for numbers."""
        from . import errors
        from .values import is_num
        pa, pc = arrow()
        if not is_num(x) or isinstance(x, bool):
            raise errors.CALC
        x = float(x)
        cmp = pc.less if mode == -1 else pc.greater

        def amask(ch, nm):
            return pc.fill_null(cmp(nm, x), False)

        cand = {}

        def pick(mask):
            big = self.sheet.big
            nums = big.column_part(big.numeric(self.c), self.r1, self.r1 + len(mask) - 1)
            vals = pc.filter(nums, mask)
            if not len(vals):
                return []
            best = (pc.max if mode == -1 else pc.min)(vals).as_py()
            hit = pc.fill_null(pc.equal(nums, best), False)
            if isinstance(hit, pa.ChunkedArray):
                hit = hit.combine_chunks()
            hit = pc.and_(hit, mask)
            i = pc.index(hit, True).as_py()
            cand[i] = best
            return [i]

        def pred(v):
            return is_num(v) and not isinstance(v, bool) and (v < x if mode == -1 else v > x)
        found = self._search(amask, pred, pick)
        best_i, best_v = -1, None
        for i in sorted(found):
            v = cand.get(i)
            if v is None:
                v = self.sheet.value(self.r1 + i, self.c)
            if best_v is None or (v > best_v if mode == -1 else v < best_v):
                best_i, best_v = i, v
        return best_i


def lazy_column(rng):
    """LazyColumn for a tall single-column (or first-column) range of a big sheet, else None."""
    if getattr(rng.sheet, "big", None) is None:
        return None
    if min(rng.r2, rng.sheet.max_row) - rng.r1 + 1 <= 50_000:
        return None
    return LazyColumn(rng.sheet, rng.r1, rng.r2, rng.c1)


# ================================================================ navigation
def current_region(sheet, r, c):
    """The data block around (r, c): the whole file's table when the cell is inside it."""
    big = sheet.big
    if r < big.nview and c < big.ncols and big.N:
        return (0, 0, big.nview - 1, big.ncols - 1)
    return None


def data_edge(sheet, r, c, dr, limit_r):
    """Ctrl+Up/Down inside a big column: next edge of non-blank cells (edits not considered)."""
    pa, pc = arrow()
    big = sheet.big
    if c >= big.ncols or not big.N:
        return None
    nv = big.nview
    if r >= nv:
        return None
    col = big.cols[c]
    if dr > 0:
        if r >= nv - 1:
            return (min(limit_r, nv - 1 if r < nv - 1 else limit_r), c)
        part = big.column_part(col, r, nv - 1)
        blank = pc.equal(part, "")
        here_full = not blank[0].as_py()
        next_full = len(blank) > 1 and not blank[1].as_py()
        if here_full and next_full:
            i = pc.index(blank, True).as_py()
            return (r + (i - 1 if i >= 0 else len(blank) - 1), c)
        i = pc.index(blank.slice(1), False).as_py()
        return (r + 1 + i, c) if i >= 0 else (limit_r, c)
    if r == 0:
        return (0, c)
    part = big.column_part(col, 0, min(r, nv - 1))
    blank = pc.equal(part, "")
    here_full = not blank[len(blank) - 1].as_py()
    prev_full = len(blank) > 1 and not blank[len(blank) - 2].as_py()
    if here_full and prev_full:
        idx = pc.indices_nonzero(blank)
        return ((idx[len(idx) - 1].as_py() + 1) if len(idx) else 0, c)
    idx = pc.indices_nonzero(pc.invert(blank.slice(0, len(blank) - 1)))
    return (idx[len(idx) - 1].as_py(), c) if len(idx) else (0, c)


# ================================================================ sort
def _key_arrays(big, c, base, asc, case):
    """Sort key columns for column c over data rows `base`: (kind, number, text)."""
    pa, pc = arrow()
    num = pc.take(big.numeric(c), base)
    txt = pc.take(big.cols[c], base)
    patch = {dk >> 14: v for dk, v in big.edits.items() if (dk & 0x3FFF) == c and (dk >> 14) < big.N}
    if patch:
        rows = pa.array(sorted(patch), pa.int64())
        mask = pc.is_in(base, value_set=rows)
        pos = pc.indices_nonzero(mask).to_pylist()
        if pos:
            ds = pc.take(base, pa.array(pos, pa.int64())).to_pylist()
            from .values import is_num
            nums, txts = [], []
            for d in ds:
                v = patch[d]
                if v is CLEARED or v is None:
                    nums.append(None)
                    txts.append("")
                elif is_num(v) and not isinstance(v, bool):
                    nums.append(float(v))
                    txts.append(str(v))
                else:
                    nums.append(None)
                    txts.append(str(v))
            full = _position_mask(len(base), pos)
            num = pc.replace_with_mask(num, full, pa.array(nums, pa.float64()))
            txt = pc.replace_with_mask(txt, full, pa.array(txts, pa.string()))
    blank = pmap(lambda ch: pc.equal(ch, ""), txt)
    isnum = pc.is_valid(num)
    first, second = (0, 1) if asc else (1, 0)
    kind = pc.if_else(isnum, pa.scalar(first, pa.int8()),
                      pc.if_else(blank, pa.scalar(2, pa.int8()), pa.scalar(second, pa.int8())))
    n_num = pc.sum(isnum).as_py() or 0
    n_blank = pc.sum(blank).as_py() or 0
    has_text = n_num + n_blank < len(txt)
    if has_text and not case:
        txt = pmap(pc.utf8_lower, txt)
    return kind, (num if n_num else None), (txt if has_text else None), bool(n_blank or (n_num and has_text))


def sorted_state(sheet, keys, header=True, case=False):
    """New (order, view) with the data sorted by keys [(col, ascending)], header row kept on top.
    Ties keep their current order; blanks always last."""
    pa, pc = arrow()
    big = sheet.big
    cur = big.order if big.order is not None else arange(big.N)
    top = 1 if header else 0
    head, base = cur.slice(0, top), cur.slice(top)
    cols = {}
    sort_keys = []
    for i, (c, asc) in enumerate(keys):
        if c >= big.ncols:
            continue
        kind, num, txt, mixed = _key_arrays(big, c, base, asc, case)
        order = "ascending" if asc else "descending"
        # only the keys that matter: a pure number column sorts on the numbers alone (3x faster)
        if mixed:
            cols[f"k{i}"] = kind
            sort_keys.append((f"k{i}", "ascending"))
        if num is not None:
            cols[f"n{i}"] = num
            sort_keys.append((f"n{i}", order))
        if txt is not None:
            cols[f"t{i}"] = txt
            sort_keys.append((f"t{i}", order))
    if not sort_keys:
        return big.state()
    idx = pc.sort_indices(pa.table(cols), sort_keys=sort_keys)
    new_order = pa.concat_arrays([head, pc.take(base, idx)])
    return new_order, filtered_view(sheet, new_order, sheet.filters)


# ================================================================ filter
def _cond_mask(big, c, op, arg):
    from .numfmt import parse_input
    from .values import is_num
    pa, pc = arrow()
    col = big.cols[c]
    al = str(arg).lower()
    text_ops = {
        "contains": lambda ch: pc.match_substring(ch, al, ignore_case=True),
        "not_contains": lambda ch: pc.invert(pc.match_substring(ch, al, ignore_case=True)),
        "begins": lambda ch: pc.starts_with(ch, al, ignore_case=True),
        "ends": lambda ch: pc.ends_with(ch, al, ignore_case=True),
    }
    if op in text_ops:
        return pmap(text_ops[op], col)
    x, _ = parse_input(str(arg))
    same_text = lambda ch: pc.match_substring_regex(ch, "^" + re.escape(al) + "$", ignore_case=True)
    if op in ("equals", "not_equals"):
        if is_num(x):
            m = pmap2(lambda ch, nm: pc.or_(pc.fill_null(pc.equal(nm, float(x)), False), same_text(ch)),
                      col, big.numeric(c))
        else:
            m = pmap(same_text, col)
        return m if op == "equals" else pmap(pc.invert, m)
    fn = {">": pc.greater, ">=": pc.greater_equal, "<": pc.less, "<=": pc.less_equal}.get(op)
    if fn is None or not is_num(x):
        return pmap(lambda ch: pc.equal(ch, "\x00never") if fn is not None else pc.not_equal(ch, "\x00never"), col)
    return pmap(lambda nm: pc.fill_null(fn(nm, float(x)), False), big.numeric(c))


def _spec_mask(big, c, spec):
    """Boolean ChunkedArray over all data rows: does the row pass this column's filter?"""
    pa, pc = arrow()
    col = big.cols[c] if c < big.ncols else None
    if col is None:
        return None
    mask = None
    allowed = spec.get("values")
    if allowed is not None:
        vs = pa.array(sorted(allowed), pa.string())
        keep_blank = spec.get("blanks", True)

        def one(ch):
            blank = pc.equal(ch, "")
            m = pc.is_in(ch, value_set=vs)
            return pc.or_(m, blank) if keep_blank else pc.and_(m, pc.invert(blank))
        mask = pmap(one, col)
    cond = spec.get("cond")
    if cond and cond[0]:
        m = _cond_mask(big, c, cond[0], cond[1])
        mask = m if mask is None else pc.and_(mask, m)
    return mask


def filtered_view(sheet, order, filters, header=True):
    """Rows of `order` passing all filters (header row always shown), or order itself."""
    pa, pc = arrow()
    big = sheet.big
    if not filters:
        return order
    from .ops import row_passes
    if order is None:
        order = arange(big.N)
    total = None
    for c, spec in filters.items():
        m = _spec_mask(big, c, spec)
        if m is not None:
            total = m if total is None else pc.and_(total, m)
    if total is None:
        return order
    if isinstance(total, pa.ChunkedArray):
        total = total.combine_chunks()
    top = 1 if header else 0
    rest = order.slice(top)
    keep = pc.fill_null(pc.take(total, rest), False)
    # rows with edited cells in filtered columns: decide those the normal, cell-by-cell way
    fix = sorted({dk >> 14 for dk in big.edits if (dk >> 14) < big.N and (dk & 0x3FFF) in filters})
    if fix:
        pos = pc.indices_nonzero(pc.is_in(rest, value_set=pa.array(fix, pa.int64()))).to_pylist()
        if pos:
            saved = big.state()
            big.set_state((order, order))
            try:
                ok = [all(row_passes(sheet, p + top, spec, c) for c, spec in filters.items()) for p in pos]
            finally:
                big.set_state(saved)
            keep = pc.replace_with_mask(keep, _position_mask(len(keep), pos), pa.array(ok, pa.bool_()))
    return pa.concat_arrays([order.slice(0, top), pc.filter(rest, keep)])


def filter_values(sheet, col, filters, limit=10_000):
    """(display values, has_blanks) offered in the filter list for col; [] when there are more
    than `limit` different values (then only the condition filter is offered)."""
    pa, pc = arrow()
    big = sheet.big
    if col >= big.ncols:
        return [], False
    other = {k: v for k, v in filters.items() if k != col}
    data = big.cols[col]
    mask = None
    for c, spec in other.items():
        m = _spec_mask(big, c, spec)
        if m is not None:
            mask = m if mask is None else pc.and_(mask, m)
    data = data.slice(1)  # header row
    if mask is not None:
        data = pc.filter(data, pc.fill_null(mask.slice(1), False))
    probe = pc.unique(data.slice(0, min(len(data), 2_000_000)))
    if len(probe) > limit:
        return [], True
    uniq = pc.unique(data).to_pylist()
    if len(uniq) > limit + 1:
        return [], True
    blanks = "" in uniq
    from .io_csv import convert_field
    from .values import sort_key
    rt = {}
    vals = [u for u in uniq if u != ""]

    def key(t):
        conv = convert_field(t, rt)
        return sort_key(conv[0]) if conv else (9,)
    vals.sort(key=key)
    return vals, blanks


# ================================================================ find / replace
def _matcher(needle, match_case, whole):
    pa, pc = arrow()
    if any(ch in needle for ch in "*?"):
        rx = "".join(".*" if ch == "*" else "." if ch == "?" else re.escape(ch) for ch in needle)
        rx = "^" + rx + "$" if whole else rx
        return lambda ch: pc.match_substring_regex(ch, rx, ignore_case=not match_case)
    if whole:
        if match_case:
            return lambda ch: pc.equal(ch, needle)
        rx = "^" + re.escape(needle) + "$"
        return lambda ch: pc.match_substring_regex(ch, rx, ignore_case=True)
    return lambda ch: pc.match_substring(ch, needle, ignore_case=not match_case)


def find_rows(sheet, needle, match_case=False, whole=False, start=0, limit=1_000_000):
    """Sorted [(r, c)] of file cells matching needle in sheet coordinates, scanning sheet rows from
    `start` (then wrapping round) in blocks and stopping once `limit` hits are found. Edited
    cells are left out; callers add those themselves."""
    pa, pc = arrow()
    big = sheet.big
    nv = big.nview
    if not nv or not big.ncols:
        return []
    test = _matcher(needle, match_case, whole)
    edited = set()
    for dk in big.edits:
        r = big.srow(dk >> 14)
        if r is not None:
            edited.add((r, dk & 0x3FFF))
    block = 4_000_000
    start = max(0, min(start, nv - 1))
    spans = [(a, min(a + block, nv)) for a in range(start, nv, block)] +             [(a, min(a + block, start)) for a in range(0, start, block)]
    hits = []
    for a, stop in spans:
        n = stop - a
        found = []
        for c in range(big.ncols):
            part = big.column_part(big.cols[c], a, a + n - 1)
            m = pmap(test, part) if isinstance(part, pa.ChunkedArray) else test(part)
            idx = pc.indices_nonzero(m)
            if len(idx):
                found.extend((a + i, c) for i in idx.slice(0, limit).to_pylist())
        found = [h for h in found if h not in edited]
        found.sort()
        hits.extend(found)
        if len(hits) >= limit:
            break
    hits.sort()
    return hits[:limit] if len(hits) > limit else hits


def replace_in_columns(sheet, needle, repl, match_case=False, whole=False):
    """New column list with needle replaced in the file's text (edited cells untouched)
    and the number of cells changed."""
    pa, pc = arrow()
    big = sheet.big
    new_cols = list(big.cols)
    total = 0
    for c, col in enumerate(big.cols):
        if whole:
            m = pc.equal(col, needle) if match_case else pc.equal(pc.utf8_lower(col), needle.lower())
            n = pc.sum(m).as_py() or 0
            if n:
                new_cols[c] = pc.if_else(m, pa.scalar(repl, pa.string()), col)
        else:
            m = pc.match_substring(col, needle, ignore_case=not match_case)
            n = pc.sum(m).as_py() or 0
            if n:
                pat = re.escape(needle) if match_case else "(?i)" + re.escape(needle)
                new_cols[c] = pc.replace_substring_regex(col, pat, repl.replace("\\", "\\\\"))
        total += n
    return new_cols, total


def set_columns(sheet, cols):
    big = sheet.big
    big.cols = list(cols)
    big._rows.clear()
    big._num = {}
    sheet.wb.recalc(full=True)


# ================================================================ save
def _csv_bytes(arrays, delim, eol):
    """CSV text for equally long string arrays, quoting like Python's csv module (only fields
    holding the delimiter, a quote or a line break), built with Arrow string kernels."""
    pa, pc = arrow()
    special = "[" + re.escape(delim) + '"\r\n]'
    fields = []
    for a in arrays:
        if isinstance(a, pa.ChunkedArray):
            a = a.combine_chunks()
        need = pc.match_substring_regex(a, special)
        if pc.any(need).as_py():
            quoted = pc.binary_join_element_wise('"', pc.replace_substring(a, '"', '""'), '"', "")
            a = pc.if_else(need, quoted, a)
        fields.append(a)
    line = pc.binary_join_element_wise(*fields, delim) if len(fields) > 1 else fields[0]
    line = pc.binary_join_element_wise(line, eol, "")
    if isinstance(line, pa.ChunkedArray):
        line = line.combine_chunks()
    if line.type != pa.large_string():
        line = line.cast(pa.large_string())
    offs = line.buffers()[1]
    data = line.buffers()[2]
    om = memoryview(offs).cast("q")
    start, stop = om[line.offset], om[line.offset + len(line)]
    return memoryview(data)[start:stop].tobytes() if data is not None else b""


def save_big_csv(sheet, path, options, progress=None):
    """Write the file's data in the current sort order (filtered-out rows included, like Excel),
    with the user's edits and formula results, plus rows/columns typed past the data."""
    pa, pc = arrow()
    from .io_csv import cell_text_for_csv
    big = sheet.big
    delim = options.get("delimiter") or ","
    enc = options.get("encoding", "utf-8")
    eol = options.get("newline", "\r\n")
    bom = options.get("bom") or enc == "utf-8-sig"
    py_enc = "utf-8" if enc in ("utf-8", "utf-8-sig") else enc
    order = big.order
    # edits and formula results, by data row, as file text
    patch = {}
    saved = big.state()
    big.set_state((order, order))
    try:
        sheet.recompute_extent()
        sheet.wb.recalc(full=True)  # formula results as with no filter
        mr, mc = sheet.used_extent(include_styles=False)
        for k in set(sheet.values) | set(sheet.formulas):
            r, c = k >> 14, k & 0x3FFF
            patch.setdefault(big.drow(r), {})[c] = cell_text_for_csv(sheet, k)
        for dk, v in big.edits.items():
            if v is CLEARED:
                patch.setdefault(dk >> 14, {}).setdefault(dk & 0x3FFF, "")
    finally:
        big.set_state(saved)
        sheet.recompute_extent()
        sheet.wb.recalc(full=True)
    ncols = max(mc + 1, big.ncols, 1)
    extra_rows = sorted(d for d in patch if d >= big.N)
    total = big.N + (extra_rows[-1] - big.N + 1 if extra_rows else 0)
    chunk = 1_000_000
    tmp = path + ".tmp~"

    def write(raw, data):
        if py_enc != "utf-8":
            data = data.decode("utf-8").encode(py_enc, errors="replace")
        raw.write(data)

    with open(tmp, "wb") as raw:
        if bom and py_enc == "utf-8":
            raw.write(b"\xef\xbb\xbf")
        for start in range(0, big.N, chunk):
            n = min(chunk, big.N - start)
            if order is None:
                drows = None
                arrays = [col.slice(start, n) for col in big.cols]
            else:
                drows = order.slice(start, n)
                arrays = [pc.take(col, drows) for col in big.cols]
            arrays += [pa.repeat("", n) for _ in range(ncols - big.ncols)]
            if patch:
                if drows is None:
                    hit = [(d - start, d) for d in patch if start <= d < start + n]
                else:
                    want = pa.array(sorted(d for d in patch if d < big.N), pa.int64())
                    pos = pc.indices_nonzero(pc.is_in(drows, value_set=want)).to_pylist()
                    hit = list(zip(pos, pc.take(drows, pa.array(pos, pa.int64())).to_pylist())) if pos else []
                hit.sort()
                for c in range(ncols):
                    rows = [(p, patch[d][c]) for p, d in hit if c in patch[d]]
                    if rows:
                        arr = arrays[c]
                        if isinstance(arr, pa.ChunkedArray):
                            arr = arr.combine_chunks()
                        arrays[c] = pc.replace_with_mask(arr, _position_mask(n, [p for p, _ in rows]),
                                                         pa.array([t for _, t in rows], pa.string()))
            write(raw, _csv_bytes(arrays, delim, eol))
            if progress:
                progress(min(0.99, (start + n) / max(1, total)))
        if extra_rows:
            text = io.StringIO()
            w = csv.writer(text, delimiter=delim, lineterminator=eol)
            for d in range(big.N, extra_rows[-1] + 1):
                cells = patch.get(d, {})
                width = max(cells) + 1 if cells else 0
                w.writerow([cells.get(c, "") for c in range(width)])
            write(raw, text.getvalue().encode("utf-8"))
    os.replace(tmp, path)


def too_many_rows_for_xlsx(sheet):
    return sheet.used_extent(include_styles=False)[0] + 1 > 1_048_576


__all__ = ["BigSheet", "BigData", "TooBig", "load_big_csv", "wants_big", "is_big", "check_area",
           "range_stats", "aggregate", "sorted_state", "filtered_view", "filter_values", "find_rows",
           "replace_in_columns", "set_columns", "save_big_csv", "current_region", "data_edge",
           "col_name", "MAX_ROWS"]
