"""MCP server that Claude Code starts (stdio). It forwards every tool call to the
running Sheets window identified by SHEETS_TARGET, over Sheets' local socket.

Run as:  python -m sheets.ai.mcp_server   (env: SHEETS_TARGET, optional SHEETS_SERVER_NAME)"""
import json
import os
import sys
from typing import Any

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from mcp.server.mcpserver import MCPServer  # noqa: E402
from PySide6.QtCore import QCoreApplication  # noqa: E402
from PySide6.QtNetwork import QLocalSocket  # noqa: E402

SERVER_NAME = os.environ.get("SHEETS_SERVER_NAME") or ("SheetsSpreadsheetApp-" + (os.environ.get("USERNAME") or os.environ.get("USER") or "user"))
TARGET = os.environ.get("SHEETS_TARGET", "")

_qapp = QCoreApplication.instance() or QCoreApplication([])
_sock = None


def call(tool, **args):
    """Send one request to Sheets and wait for its one-line JSON reply."""
    global _sock
    args = {k: v for k, v in args.items() if v is not None}
    if _sock is None or _sock.state() != QLocalSocket.ConnectedState:
        _sock = QLocalSocket()
        _sock.connectToServer(SERVER_NAME)
        if not _sock.waitForConnected(3000):
            _sock = None
            return {"error": "Ekxel isn't running (couldn't connect)."}
    _sock.write((json.dumps({"rpc": tool, "args": args, "target": TARGET}) + "\n").encode("utf-8"))
    _sock.flush()
    buf = b""
    while b"\n" not in buf:
        if not _sock.waitForReadyRead(120000):
            break
        buf += bytes(_sock.readAll())
    if b"\n" in buf:
        return json.loads(buf.split(b"\n", 1)[0].decode("utf-8"))
    # Never resend: the request may already have been carried out (e.g. inserting rows).
    _sock = None
    return {"error": "Ekxel didn't answer in time; check the workbook before retrying."}


server = MCPServer(
    "sheets",
    instructions=(
        "Tools for the spreadsheet open in the Ekxel app. Ranges use A1 notation ('B2:D10', "
        "'Sheet2!A1', 'C:C'). Formulas start with '=' and use Excel functions. Read before you write, "
        "check formula_errors in write results, and keep changes to what was asked."),
)


@server.tool()
def workbook_info() -> dict[str, Any]:
    """Sheets in the workbook, their used ranges, the active sheet, the user's current selection
    and active cell. Call this first."""
    return call("workbook_info")


@server.tool()
def read_range(range: str, sheet: str | None = None, include_formats: bool = False) -> dict[str, Any]:
    """Read cells: raw values, displayed text and any formulas. Whole columns like 'A:D' are
    trimmed to the used rows. Up to 5000 cells per call."""
    return call("read_range", range=range, sheet=sheet, include_formats=include_formats)


@server.tool()
def write_range(start_cell: str, rows: list[list[Any]], sheet: str | None = None) -> dict[str, Any]:
    """Write a block of cells starting at start_cell (top-left). rows is a list of rows; each item is
    a number, text, true/false, null (to clear) or a formula string starting with '='. Text like
    '$1,200', '12%' or '2024-03-05' is converted the way typing it would be. The result lists any
    cells whose formulas evaluate to an error."""
    return call("write_range", start_cell=start_cell, rows=rows, sheet=sheet)


@server.tool()
def format_range(range: str, sheet: str | None = None, bold: bool | None = None, italic: bool | None = None,
                 underline: bool | None = None, font_color: str | None = None, fill_color: str | None = None,
                 font_size: float | None = None, number_format: str | None = None, align: str | None = None,
                 wrap: bool | None = None, borders: str | None = None) -> dict[str, Any]:
    """Format cells. Colors are '#RRGGBB' or a basic name. number_format is an Excel format such as
    '"$"#,##0.00', '0.0%', 'yyyy-mm-dd' or 'General'. align: left/center/right/general.
    borders: all/outline/thick_outline/inside/top/bottom/left/right/none. Only the options you pass
    change."""
    return call("format_range", range=range, sheet=sheet, bold=bold, italic=italic, underline=underline,
                font_color=font_color, fill_color=fill_color, font_size=font_size,
                number_format=number_format, align=align, wrap=wrap, borders=borders)


@server.tool()
def clear_range(range: str, sheet: str | None = None, what: str = "contents") -> dict[str, Any]:
    """Clear cells. what: 'contents' (keep formatting), 'formats', or 'all'."""
    return call("clear_range", range=range, sheet=sheet, what=what)


@server.tool()
def add_sheet(name: str, position: int | None = None) -> dict[str, Any]:
    """Add a new empty sheet (position 0 = first; default last). Write to it with sheet=name."""
    return call("add_sheet", name=name, position=position)


@server.tool()
def sort_range(range: str, column: str, sheet: str | None = None, ascending: bool = True,
               has_header: bool = True) -> dict[str, Any]:
    """Sort the rows of a range by one column (a letter inside the range). Formulas move with
    their rows."""
    return call("sort_range", range=range, column=column, sheet=sheet, ascending=ascending,
                has_header=has_header)


@server.tool()
def insert_or_delete(action: str, at: str, count: int = 1, sheet: str | None = None) -> dict[str, Any]:
    """Insert or delete whole rows or columns. action: insert_rows / delete_rows (at = row number,
    e.g. '5') or insert_columns / delete_columns (at = column letter, e.g. 'C'). Formulas that
    refer to moved cells are updated."""
    return call("insert_or_delete", action=action, at=at, count=count, sheet=sheet)


@server.tool()
def find(text: str, sheet: str | None = None, match_case: bool = False,
         whole_cell: bool = False) -> dict[str, Any]:
    """Find cells whose text or formula contains `text` (* and ? wildcards work). Searches every
    sheet unless sheet is given."""
    return call("find", text=text, sheet=sheet, match_case=match_case, whole_cell=whole_cell)


@server.tool()
def set_column_width(columns: str, width: float | None = None, sheet: str | None = None) -> dict[str, Any]:
    """Set column width in characters for columns like 'B' or 'A:D'. Leave width out to autofit."""
    return call("set_column_width", columns=columns, width=width, sheet=sheet)


@server.tool()
def select_range(range: str, sheet: str | None = None) -> dict[str, Any]:
    """Select a range in the app so the user sees it (switches sheet if needed)."""
    return call("select_range", range=range, sheet=sheet)


@server.tool()
def add_control(kind: str, linked_cell: str, place: str | None = None, min: float = 0, max: float = 100,
                step: float = 1, sheet: str | None = None) -> dict[str, Any]:
    """Add a form control linked to a cell, like Excel's Scroll Bar / Spin Button. kind: 'scrollbar'
    (a slider to drag) or 'spinner' (up/down arrows). Moving it sets linked_cell to a number between
    min and max in multiples of step, live, so formulas using that cell update as the user drags.
    place = the cells it covers (e.g. 'C4:F4' for a slider; default: next to the linked cell)."""
    return call("add_control", kind=kind, linked_cell=linked_cell, place=place, min=min, max=max,
                step=step, sheet=sheet)


@server.tool()
def add_chart(data_range: str, chart_type: str = "column", title: str | None = None, place: str | None = None,
              series_in: str | None = None, x_axis_title: str | None = None, y_axis_title: str | None = None,
              legend: str | None = None, data_labels: bool = False, sheet: str | None = None) -> dict[str, Any]:
    """Add an Excel-style chart linked to cells, like Insert > Chart: it follows the cells (edit a number and
    the chart updates). data_range = the cells to plot including the header row / first label column,
    e.g. 'A1:D7'; Excel's rules pick series (a header row gives series names, a first column of text gives
    the category labels; series_in='columns' or 'rows' forces the direction). chart_type: column,
    stacked_column, percent_column, bar, stacked_bar, percent_bar, line, line_markers, pie, doughnut,
    area, stacked_area, percent_area, scatter (first column = X values), scatter_lines, scatter_smooth,
    combo (columns + lines) or combo_secondary (lines on a second axis). place = top-left cell like 'H2'
    or a range like 'H2:N18' for its size (default: right of the data). legend: right, left, top, bottom or
    none. The user can click the chart to move, resize and format it. Large ranges (big files) are sampled."""
    return call("add_chart", data_range=data_range, chart_type=chart_type, title=title, place=place,
                series_in=series_in, x_axis_title=x_axis_title, y_axis_title=y_axis_title, legend=legend,
                data_labels=data_labels, sheet=sheet)


if __name__ == "__main__":
    server.run("stdio")
