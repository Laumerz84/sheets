"""Open / save dispatch by file extension."""
import os

from .workbook import new_workbook

OPEN_FILTER = ("All spreadsheets (*.xlsx *.xlsm *.xls *.csv *.tsv *.txt);;"
               "Excel workbook (*.xlsx *.xlsm);;Excel 97-2003 (*.xls);;"
               "CSV (*.csv);;Tab-separated (*.tsv *.txt);;All files (*)")
SAVE_FILTERS = {
    ".xlsx": "Excel workbook (*.xlsx)",
    ".xlsm": "Excel macro-enabled workbook (*.xlsm)",
    ".csv": "CSV - comma separated (*.csv)",
    ".tsv": "Tab-separated text (*.tsv)",
    ".txt": "Tab-separated text (*.txt)",
}
READABLE = {".xlsx", ".xlsm", ".xls", ".csv", ".tsv", ".txt", ".tab"}
WRITABLE = {".xlsx", ".xlsm", ".csv", ".tsv", ".txt"}


def open_file(path, progress=None):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        from .io_xlsx import load_xlsx
        return load_xlsx(path, progress)
    if ext == ".xls":
        return _open_xls(path, progress)
    from .bigdata import load_big_csv, wants_big
    if wants_big(path):
        return load_big_csv(path, progress)
    from .io_csv import load_csv
    return load_csv(path, progress)


def _open_xls(path, progress):
    """Real .xls, or one of the common impostors (xlsx, HTML table, CSV) renamed to .xls."""
    with open(path, "rb") as fh:
        head = fh.read(512)
    if head.startswith(b"PK"):
        from .io_xlsx import load_xlsx
        return load_xlsx(path, progress)
    if head.lstrip()[:1] == b"<":
        from .io_html import load_html
        return load_html(path)
    from .io_xls import load_xls
    try:
        return load_xls(path, progress)
    except Exception:
        try:
            return load_xls(path, progress, formatting=False)
        except Exception:
            if head.startswith(b"\xd0\xcf\x11\xe0"):
                raise
    from .io_csv import load_csv
    wb = load_csv(path, progress)
    wb.file_format = "xls"
    return wb


def save_file(wb, path, sheet=None, progress=None):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        from .refs import XLSX_MAX_ROWS
        for sh in wb.sheets:
            if sh.used_extent(include_styles=False)[0] + 1 > XLSX_MAX_ROWS:
                raise ValueError(f"'{sh.name}' has {sh.used_extent(include_styles=False)[0] + 1:,} rows. "
                                 f"Excel files hold at most {XLSX_MAX_ROWS:,} rows per sheet; "
                                 "save it as CSV instead.")
        from .io_xlsx import save_xlsx
        save_xlsx(wb, path)
        wb.file_format = "xlsx"
    else:
        from .io_csv import save_csv
        sh = sheet or wb.sheets[wb.active]
        opts = None
        if wb.file_format != "csv":
            opts = {"encoding": "utf-8", "bom": True, "newline": "\r\n",
                    "delimiter": "\t" if ext in (".tsv", ".txt", ".tab") else ","}
        if getattr(sh, "big", None) is not None:
            from .bigdata import save_big_csv
            o = dict(wb.csv_options or {})
            o.update(opts or {})
            if ext in (".tsv", ".tab"):
                o["delimiter"] = "\t"
            save_big_csv(sh, path, o, progress)
        else:
            save_csv(wb, sh, path, opts)
        if wb.file_format != "csv" or wb.csv_options is None:
            wb.csv_options = opts
        wb.file_format = "csv"
    wb.path = path


__all__ = ["open_file", "save_file", "new_workbook", "OPEN_FILTER", "SAVE_FILTERS"]
