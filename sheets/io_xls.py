"""Legacy Excel .xls reading through xlrd (values and formatting; formulas
come in as their last calculated values because xlrd cannot read them)."""
import os

from . import errors
from .workbook import Sheet, Style, Workbook, intern_style, DEFAULT_STYLE

_HALIGN = {1: "left", 2: "center", 3: "right", 6: "center"}
_VALIGN = {0: "top", 1: "center"}
_BORDER = {1: "thin", 2: "medium", 3: "dashed", 4: "dotted", 5: "thick", 6: "double",
           7: "hair", 8: "mediumDashed", 9: "dashDot", 10: "mediumDashDot",
           11: "dashDotDot", 12: "mediumDashDotDot", 13: "slantDashDot"}


def _rgb(book, idx, default=None):
    if idx is None:
        return default
    rgb = book.colour_map.get(idx)
    if not rgb:
        return default
    return "#%02X%02X%02X" % rgb


def load_xls(path, progress=None, formatting=True):
    import xlrd
    book = xlrd.open_workbook(path, formatting_info=formatting, on_demand=True)
    wb = Workbook()
    wb.path = path
    wb.file_format = "xls"
    style_cache = {}
    date_offset = 1462.0 if book.datemode == 1 else 0.0

    def style_for(xf_index):
        st = style_cache.get(xf_index)
        if st is not None:
            return st
        xf = book.xf_list[xf_index]
        font = book.font_list[xf.font_index]
        kw = {
            "bold": bool(font.bold), "italic": bool(font.italic),
            "underline": bool(font.underlined), "strike": bool(font.struck_out),
        }
        if font.name and font.name not in ("Arial", "Calibri"):
            kw["font"] = font.name
        size = font.height / 20
        if size and size not in (10, 11):
            kw["size"] = size
        col = _rgb(book, font.colour_index)
        if col and col != "#000000":
            kw["color"] = col
        bg = xf.background
        if bg.fill_pattern:
            fill = _rgb(book, bg.pattern_colour_index)
            if fill:
                kw["fill"] = fill
        al = xf.alignment
        if al.hor_align in _HALIGN:
            kw["halign"] = _HALIGN[al.hor_align]
        if al.vert_align in _VALIGN:
            kw["valign"] = _VALIGN[al.vert_align]
        if al.text_wrapped:
            kw["wrap"] = True
        b = xf.border
        sides = []
        for ls, ci in ((b.left_line_style, b.left_colour_index), (b.top_line_style, b.top_colour_index),
                       (b.right_line_style, b.right_colour_index), (b.bottom_line_style, b.bottom_colour_index)):
            sides.append((_BORDER.get(ls, "thin"), _rgb(book, ci, "#000000")) if ls else None)
        kw["border"] = tuple(sides)
        fmt = book.format_map.get(xf.format_key)
        if fmt is not None and fmt.format_str and fmt.format_str != "General":
            kw["numfmt"] = fmt.format_str
        st = intern_style(Style(**kw))
        style_cache[xf_index] = st
        return st

    nsheets = book.nsheets
    for si in range(nsheets):
        xs = book.sheet_by_index(si)
        sh = Sheet(wb, xs.name)
        wb.sheets.append(sh)
        values, styles = sh.values, sh.styles
        for r in range(xs.nrows):
            if progress and r % 5000 == 0:
                progress((si + r / max(1, xs.nrows)) / nsheets)
            base = r << 14
            for c in range(min(xs.ncols, 16384)):
                ctype = xs.cell_type(r, c)
                try:
                    xf = xs.cell_xf_index(r, c) if formatting else None
                except (IndexError, AttributeError):
                    xf = None
                if xf is not None:
                    st = style_for(xf)
                    if st is not DEFAULT_STYLE:
                        styles[base | c] = st
                if ctype in (0, 6):
                    continue
                v = xs.cell_value(r, c)
                if ctype == 1:
                    if v != "":
                        values[base | c] = v
                elif ctype == 2:
                    values[base | c] = float(v)
                elif ctype == 3:
                    values[base | c] = float(v) + date_offset
                elif ctype == 4:
                    values[base | c] = bool(v)
                elif ctype == 5:
                    values[base | c] = errors.from_code(xlrd.error_text_from_code.get(v, "#VALUE!")) or errors.VALUE
        for c, info in (xs.colinfo_map.items() if formatting else ()):
            sh.col_widths[c] = max(2, int(round(info.width / 256 * 7 + 5)))
            if info.hidden:
                sh.hidden_cols.add(c)
        for r, info in (xs.rowinfo_map.items() if formatting else ()):
            if info.height:
                px = int(round(info.height / 20 * 96 / 72))
                if px != 20:
                    sh.row_heights[r] = px
            if info.hidden:
                sh.hidden_rows.add(r)
        for rlo, rhi, clo, chi in (xs.merged_cells if formatting else ()):
            if rhi - rlo > 1 or chi - clo > 1:
                sh.merges.append((rlo, clo, rhi - 1, chi - 1))
        if getattr(xs, "panes_are_frozen", 0):
            sh.freeze = (xs.horz_split_pos or 0, xs.vert_split_pos or 0)
        sh.recompute_extent()
        book.unload_sheet(si)
    if not wb.sheets:
        wb.sheets.append(Sheet(wb, "Sheet1"))
    wb.rebuild_dependencies()
    wb.recalc(full=True)
    return wb
