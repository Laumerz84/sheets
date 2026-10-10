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

Charts and pictures in xlsx files can't be kept; Sheets warns before saving over such a file.

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
Filters, Columns, Rows and Values; right-click for Sum/Count/Average/Max/Min/Distinct Count, filters,
moving and removing. Subtotals and grand totals, month/weekday names in calendar order, columns autofit.
Refresh in the panel or Data > Refresh All (Ctrl+Alt+F5) after the source changes. The result is ordinary
cells, so Excel shows the numbers; the pivot's setup is kept in Ekxel's hidden sheet so Ekxel can
refresh and change it after reopening (Excel sees a plain table, not an interactive PivotTable).

Keyboard shortcuts follow Excel; Help → Keyboard Shortcuts (F1) lists them.

## Development

```
.venv\Scripts\python.exe -m pytest            # tests (offscreen Qt)
.venv\Scripts\python.exe tools\snapshot.py out.png [file]   # render the window to a PNG
.venv\Scripts\python.exe -u tools\bench.py 200000           # performance check
```

Code: `sheets/` engine (formula parser/evaluator, number formats, workbook model, file I/O),
`sheets/ui/` PySide6 interface (custom grid widget in `grid.py`).
