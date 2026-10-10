# Sheets

A lightweight Excel-style spreadsheet for Windows and macOS (Apple Silicon or Intel): view and edit
CSV and Excel files without Excel.
On this PC it shows as **Ekxel** (window title, Start menu, About); the package, folder, settings
(`%APPDATA%\Sheets`) and app id stay "Sheets".

## Run

- First time: double-click `setup.bat` (makes `.venv`, installs the libraries, checks the Claude panel's
  setup). Or ask Claude Code in this folder to "set up Ekxel"; `CLAUDE.md` tells it how.
- Start menu: **Ekxel**
- Or: `G:\Sheets\.venv\Scripts\pythonw.exe G:\Sheets\launch.pyw [file ...]`
- Double-click a file after registering file types (File → *Make Ekxel the default for CSV/Excel files...*),
  or run `.venv\Scripts\python.exe -m sheets.register` (`--remove` to undo).

### macOS (M1-M5 or Intel)

1. Python 3.11 or newer (the one built into macOS is too old): `brew install python@3.12`
   (or the installer from python.org).
2. Claude Code, for the Claude panel (optional): `curl -fsSL https://claude.ai/install.sh | bash`, then run
   `claude` once and sign in.
3. `git clone https://github.com/Laumerz84/sheets.git && cd sheets && ./setup.sh`
4. Double-click `Ekxel.command` in Finder (first time: right-click > Open if macOS asks).

On a Mac, Ctrl shortcuts are Cmd (Cmd+S, Cmd+Shift+A for Claude, ...). Settings live in
`~/Library/Application Support/Sheets`; Open/Save start in `~/Documents/Spreadsheets`.
Not on a Mac: Start menu/taskbar shortcut and "Make Ekxel the default" (Windows only).

Only one copy runs at a time: opening another file hands it to the running app (new window per workbook).

## Files

| Format | Open | Save |
| --- | --- | --- |
| `.xlsx` / `.xlsm` | values, formulas, formatting, merges, widths, freeze panes, filters, conditional formatting, data validation | yes; keeps comments, data validation, conditional formatting, defined names of the original file |
| `.xls` (Excel 97-2003) | values + formatting (formulas come in as values) | saves as `.xlsx` |
| `.csv` / `.tsv` / `.txt` | auto-detects delimiter and encoding; keeps leading zeros and long IDs as text | yes, same delimiter/encoding as opened |
| HTML tables named `.xls` | yes (common bank/web export) | saves as `.xlsx` |

Charts made in Ekxel are saved as real Excel charts (and restored when you reopen the file in Ekxel, see
Charts below). Charts and pictures that came from another program can't be shown or edited, and Ekxel may
simplify or drop them when it re-saves the file, so it warns before saving over such a file.

### 100 million rows

Sheets have 1,000,000,000 rows (Excel: 1,048,576). Excel files still hold at most 1,048,576 rows, so a
sheet with more saves as CSV only.

**Big-file mode.** CSV/TSV files over 100 MB open in big-file mode (`sheets/bigdata.py`): the file is kept
column by column in memory with pyarrow instead of one object per cell. A 100-million-row, 5 GB CSV opens
in seconds and needs about 1.8x the file size in free RAM (Ekxel checks first). What works on the
whole file:
- scrolling, Go To, typing/editing/clearing cells (your edits sit on top of the file's data and stay with
  their row through sorting and filtering), new rows below and columns to the right
- status-bar Sum/Average/Count for whole columns, and these functions over big ranges, computed
  column-wise in about 0.5-3 s on 100M rows: `SUM`/`COUNT`/`COUNTA`/`AVERAGE`/`MIN`/`MAX`,
  `SUMIF(S)`/`COUNTIF(S)`/`AVERAGEIF(S)`/`MAXIFS`/`MINIFS`, `VLOOKUP`/`MATCH`/`XLOOKUP`/`XMATCH`/`INDEX`
  (e.g. `=SUMIFS(E:E,C:C,"North")`). Other functions over more than 2,000,000 cells of a big file give
  `#CALC!` rather than freezing Ekxel.
- Sort (stable, numbers before text, blanks last), AutoFilter (value lists for columns with up to 10,000
  different values, condition filters for any column; the row headers show the file's row numbers in blue)
- Find (Ctrl+F) and Replace All
- filling a formula or value down millions of rows (double-click the fill handle, drag it, or Ctrl+D):
  the column becomes a calculated column, stored once and computed for every row at once
  (`sheets/bigcalc.py`), like an Excel Table column. It can use cells of the same row (`=E2*F2`), fixed
  cells (`$J$1`), + - * / ^ & %, comparisons and IF, IFERROR, AND, OR, NOT, ROUND/ROUNDUP/ROUNDDOWN, INT,
  ABS, SQRT, MOD, SUM, AVERAGE, MIN, MAX, COUNT, LEFT, RIGHT, MID, LEN, UPPER, LOWER, TRIM, CONCAT, VALUE,
  YEAR, MONTH, DAY, ISBLANK/ISNUMBER/ISTEXT/ISERROR. Editing an input cell recomputes that row.
- Save / Save As CSV or TSV (written in the current sort order, all rows including filtered-out ones)
- undo for edits, sort, filter and Replace All

Limits in big-file mode: inserting/deleting rows or columns and duplicating the sheet aren't available;
cell-by-cell actions (formatting, copying, pasting, filling) work on up to 2,000,000 cells at a time and say
so if you select more. While a filter is on, formulas see the rows that are shown. Lines with a different
number of fields than the rest are left out (Ekxel says how many and won't save over the original).

## Features

Excel-like grid (frozen panes, merged cells, overflow text, wrap, borders, fills, number formats),
180 functions (SUM/IF/VLOOKUP/XLOOKUP/INDEX/MATCH/SUMIFS/TEXT/dates...), click-to-reference and
arrow-key pointing while typing formulas, F4 for `$`, function autocomplete and hints, fill handle
(series: numbers, dates, months, "Item 1"), copy/paste with Excel (TSV) and relative formula shifting,
paste values/formats/transposed, undo/redo for everything, sort (multi-level), AutoFilter with value
lists and conditions, find/replace across sheets, remove duplicates, format painter, AutoSum, insert/delete/
hide rows and columns, autofit, multiple sheets, zoom, status-bar Sum/Average/Count.

Conditional formatting (Home toolbar / Format > Conditional Formatting, Alt H L): highlight rules
(greater/less than, between, equal, text contains, duplicates), top/bottom and above/below average,
data bars, color scales, formula rules, New Rule, Clear Rules and Manage Rules (order, Stop If True,
Applies to). Data validation (Data > Data Validation, Alt A V V): whole number, decimal, list, date,
time, text length and custom-formula rules, in-cell dropdown lists (click the arrow or Alt+Down),
input messages and Stop/Warning/Information error alerts on typed entries. Both are saved in xlsx files
Excel reads.

PivotTables (Insert > PivotTable, Alt N V): pick the data (the table around the cursor by default) and a
new or existing sheet, then build it in the PivotTable Fields panel - tick fields or drag them into
Filters, Columns, Rows and Values (the panel has a field search box and Defer Layout Update). Like Excel 365:
- Value Field Settings (double-click a value field, or right-click a value cell): Sum, Count, Average,
  Max, Min, Product, Count Numbers, StdDev/StdDevp, Var/Varp, Distinct Count; custom name; number
  format; Show Values As (% of Grand/Column/Row/Parent Total, % Of, Difference From, % Difference From,
  Running Total In, Rank, Index). Two or more value fields add a "Σ Values" field you can put in Columns or Rows.
- Date fields dropped into Rows/Columns group into Years > Quarters > Months (Group... / Ungroup on the
  right-click menu: seconds to years; number fields group by ranges with start, end and step).
- Dropdown buttons on Row Labels / Column Labels / field headers / report filters: sort A-Z, Z-A or by a
  value field, item checklist with search, Label Filters (equals, begins with, contains, greater than,
  between...) and Value Filters (Top 10 items/percent/sum, greater than...). Report filters take several items.
- +/- buttons expand and collapse outer items (also Expand/Collapse Entire Field); double-click a value
  for Show Details (a new sheet listing its source rows).
- Right-click > Design / PivotTable Options: Compact, Outline or Tabular form, repeat item labels,
  subtotals at top/bottom/off, grand totals for rows/columns, blank line after each item, text for empty
  cells, show items with no data, a dozen PivotTable styles with banded rows/columns.
- Calculated fields (e.g. `=Units*Price`, computed on the sums like Excel) and `GETPIVOTDATA`.
- Pivots also work on big-file sheets (grouped with pyarrow; Show Details isn't available there yet).
Refresh in the panel or Data > Refresh All (Ctrl+Alt+F5) after the source changes. The result is ordinary
cells, so Excel shows the numbers; the pivot's setup is kept in Ekxel's hidden sheet so Ekxel can
refresh and change it after reopening (Excel sees a plain table, not an interactive PivotTable).

Charts (Insert > Chart, Alt N C; Alt+F1 = a clustered column chart of the selection right away, F11 = a chart on a
new sheet): the gallery shows every type drawn with your own data - Clustered / Stacked / 100% Stacked Column and
Bar, Line (with or without markers), Pie, Doughnut, Area / Stacked / 100% Stacked Area, Scatter (markers, straight
or smooth lines) and Combo (columns + lines, optionally on a secondary axis, with a type per series). The data is
the selection or the table around the cursor, with Excel's rules: a header row gives the series names, a first
column of text (or dates, or years) gives the category labels, more columns than rows puts the series in rows.
- The chart floats over the cells and is linked to them: edit a number, undo, insert or delete rows, rename a
  sheet, and the chart follows. Click it to select; drag to move, drag a handle to resize (it moves and sizes
  with its cells), Delete removes it, Ctrl+C / X / V / D copy, cut, paste and duplicate (a picture goes to the
  clipboard too), arrow keys nudge it. Right-click: Change Chart Type, Select Data (data range, Switch
  Row/Column, add / edit / remove / reorder series, category labels), Move Chart (another sheet or a new
  one), Add Chart Element, Chart Style, Bring to Front / Send to Back, Format. Hover a bar or point for its value.
- The green "+" and brush buttons next to a selected chart (and the right-click menu) add chart elements: chart
  title, axis titles, data labels, gridlines, legend position; and change the style (five looks) and colors (four
  colorful and four monochromatic palettes).
- Double-click an element (or press Enter on the chart) for the Format Chart pane: chart area (fill, border,
  gap width, doughnut hole), title, plot area, legend, axes (title, minimum / maximum / major unit, log scale,
  number format, categories in reverse order) and each series (name, color, data labels, smoothing,
  column / line / area and secondary axis in a combo). Every change is one undo step.
- Hidden rows and columns are left out of a chart (so filters drive it), like Excel. In big-file mode a range
  with millions of rows is sampled down to about 2,000 points with pyarrow (a note on the chart says so); a pie
  or doughnut with more than 120 slices is refused with a message.
- Saving .xlsx writes each chart as an Excel chart (same type, data links, titles, axes, legend, colors,
  position), and keeps Ekxel's full setup in the hidden sheet so reopening in Ekxel restores it exactly.
  Not carried over to Excel: chart styles (Excel gets the colors), fonts other than Calibri-ish grey text, and
  anything Ekxel can't draw. Changes made to an Ekxel chart inside Excel are not read back by Ekxel (it restores
  its own copy). Not in Ekxel: 3-D, radar, stock, surface, bubble, histogram, waterfall and map charts,
  trendlines, error bars, sparklines, PivotCharts, slicers, individual point formatting, text boxes inside charts.
- Claude can add charts too (the `add_chart` tool: type, range, title, axis titles, legend, where it goes).

Keyboard shortcuts follow Excel; Help → Keyboard Shortcuts (F1) lists them.

## Development

```
.venv\Scripts\python.exe -m pytest            # tests (offscreen Qt)
.venv\Scripts\python.exe tools\snapshot.py out.png [file]   # render the window to a PNG
.venv\Scripts\python.exe -u tools\bench.py 200000           # performance check
```

Code: `sheets/` engine (formula parser/evaluator, number formats, workbook model, file I/O),
`sheets/ui/` PySide6 interface (custom grid widget in `grid.py`).
