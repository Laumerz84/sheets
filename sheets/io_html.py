"""Read HTML tables (web/bank exports that are saved as .xls)."""
import os
import re
from html.parser import HTMLParser

from .numfmt import parse_input
from .workbook import DEFAULT_STYLE, Sheet, Workbook, intern_style


class _TableParser(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.tables = []
        self.depth = 0
        self.row = None
        self.cell = None
        self.cell_bold = False
        self.span = 1

    def handle_starttag(self, tag, attrs):
        if tag == "table":
            self.depth += 1
            if self.depth == 1:
                self.tables.append([])
        elif self.depth == 1 and tag == "tr":
            self.row = []
        elif self.depth == 1 and tag in ("td", "th") and self.row is not None:
            self.cell = []
            self.cell_bold = tag == "th"
            try:
                self.span = max(1, int(dict(attrs).get("colspan", 1) or 1))
            except ValueError:
                self.span = 1
        elif tag == "br" and self.cell is not None:
            self.cell.append("\n")
        elif tag in ("b", "strong") and self.cell is not None:
            self.cell_bold = True

    def handle_endtag(self, tag):
        if tag == "table":
            self.depth = max(0, self.depth - 1)
        elif self.depth == 1 and tag in ("td", "th") and self.cell is not None and self.row is not None:
            text = re.sub(r"[ \t\r\n\xa0]+", " ", "".join(self.cell).replace("\n", "\x00")).replace("\x00", "\n").strip()
            self.row.append((text, self.cell_bold))
            for _ in range(self.span - 1):
                self.row.append(("", False))
            self.cell = None
        elif self.depth == 1 and tag == "tr" and self.row is not None:
            self.tables[-1].append(self.row)
            self.row = None

    def handle_data(self, data):
        if self.cell is not None:
            self.cell.append(data)


def load_html(path):
    with open(path, "rb") as fh:
        raw = fh.read()
    m = re.search(rb'charset=["\']?([\w-]+)', raw[:4096])
    enc = m.group(1).decode("ascii", "ignore") if m else "utf-8"
    try:
        text = raw.decode(enc)
    except (LookupError, UnicodeDecodeError):
        text = raw.decode("cp1252", errors="replace")
    p = _TableParser()
    p.feed(text)
    tables = [t for t in p.tables if t]
    if not tables:
        raise ValueError("No table found in this file.")
    wb = Workbook()
    wb.path = path
    wb.file_format = "xls"
    bold = intern_style(DEFAULT_STYLE.with_(bold=True))
    for i, rows in enumerate(tables):
        sh = Sheet(wb, f"Table{i + 1}" if len(tables) > 1 else
                   (os.path.splitext(os.path.basename(path))[0][:31] or "Sheet1"))
        wb.sheets.append(sh)
        for r, row in enumerate(rows):
            for c, (t, is_bold) in enumerate(row):
                if not t:
                    continue
                v, fmt = parse_input(t, typed=False)
                k = (r << 14) | c
                sh.values[k] = v
                st = DEFAULT_STYLE
                if fmt:
                    st = st.with_(numfmt=fmt)
                if is_bold:
                    st = st.with_(bold=True) if st is not DEFAULT_STYLE else bold
                if st is not DEFAULT_STYLE:
                    sh.styles[k] = st
        sh.recompute_extent()
        from .io_csv import _auto_widths
        _auto_widths(sh)
    wb.rebuild_dependencies()
    wb.recalc(full=True)
    return wb
