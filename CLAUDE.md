# Ekxel (repo/package name: Sheets) - notes for Claude

Excel-style spreadsheet (PySide6). Engine in `sheets/`, UI in `sheets/ui/` (custom grid in `grid.py`),
Claude panel in `sheets/ai/` (runs the user's Claude Code headless with an MCP bridge as its only tools).
Rows go to 1,000,000,000 (`refs.MAX_ROWS`): never loop to MAX_ROWS, clamp to `sheet.max_row`. CSVs over
100 MB open as a `bigdata.BigSheet` (pyarrow columns + edits overlay; iterating `values` yields edits only,
see the module docstring). Whole-file work there must be vectorised in `bigdata.py`; cell-by-cell code
calls `ops._guard`/`bigdata.check_area`, which raises `TooBig` (shown as a message by `app.py`).
Charts: model in `charts.py`, QPainter drawing in `ui/chart_paint.py`, the grid overlay / mouse / menus in
`ui/charts_ui.py`, dialogs in `ui/chart_dialogs.py`, Format pane in `ui/chart_pane.py`, xlsx in `chart_xlsx.py`
(stored in `Sheet.charts` like `controls`; JSON in the hidden `_EkxelControls` sheet; `add_chart` Claude tool).
Tests: `.venv\Scripts\python.exe -m pytest` (offscreen Qt). Screenshots: `tools\snapshot.py`.
Live Claude check (uses the user's Claude usage): `tools\e2e_claude.py "request"`.

## SETUP: if the user asks you to set up Ekxel (or its Claude chat), do exactly this

**What you are setting up.** Ekxel has a chat panel docked in its window. The user opens it with
**Ctrl+Shift+A**, or View > Claude. That chat is the user's own Claude Code, using their login.
No API key is needed. In that chat, Claude can only read and change the open workbook.
You do not write any code for this. It is already built. You only install and check.

Run every command from the repo folder (the folder this file is in). The commands are for
PowerShell or cmd on Windows. **On a Mac** use the same steps with these swaps: `./setup.sh` instead of
`setup.bat`; `.venv/bin/python` instead of `.venv\Scripts\python.exe`; in step 4 run
`open "<repo folder>/Ekxel.command"`; Cmd instead of Ctrl in what you tell the user; skip the Windows-only
optional items (Start menu, register). Per-OS differences live in `sheets/osinfo.py`.

**Step 0. Get the code and the two programs it needs.** Do each part only if it's missing, and
ask the user before installing anything system-wide (these are their computer's programs):
- The repo: if you aren't already in it, `git clone https://github.com/Laumerz84/sheets.git` into a folder
  the user picks, then work from there. No git? Windows: `winget install Git.Git`. Mac: `xcode-select --install`.
- Python 3.11 or newer (`python --version`; on a Mac `python3 --version`, the built-in one is too old).
  Windows: `winget install Python.Python.3.12`. Mac: `brew install python@3.12` (no Homebrew? the
  installer from https://www.python.org/downloads/ works too).
- Claude Code, for the chat panel: Windows PowerShell `irm https://claude.ai/install.ps1 | iex`;
  Mac `curl -fsSL https://claude.ai/install.sh | bash`. If you are running inside Claude Code
  right now, it is already installed.
Everything else (PySide6, openpyxl, xlrd, mcp, pywin32 on Windows) is installed by step 1 into
`.venv` inside the repo. Nothing else goes system-wide.

**Step 1. Install.** Run:

    cmd /c "<repo folder>\setup.bat"

It creates `.venv`, installs the libraries, then prints a check list. It is safe to run again.

**Step 2. Read the check list.** Each line starts with PASS, FAIL or NOTE.
- Every FAIL line has a `FIX:` line under it. Do that fix, then run the check again with
  `.venv\Scripts\python.exe tools\doctor.py`.
- Repeat until the last line says `READY`.
- Some fixes are only the user's to do. If the fix is installing Python or Claude Code, tell the
  user the exact command from the FIX line and ask before running it.

**Step 3. Sign-in.** The check list can't see whether the user is signed in to Claude Code.
Ask them: "Have you signed in to Claude Code on this PC? If not, open a terminal, type
`claude`, and log in." The user does the login themselves. Never type a password or a code for them.

**Step 4. Start Ekxel.** Start it so it keeps running after your session ends:

    Invoke-CimMethod Win32_Process -MethodName Create -Arguments @{ CommandLine = '"<repo folder>\.venv\Scripts\pythonw.exe" "<repo folder>\launch.pyw"'; CurrentDirectory = '<repo folder>' }

Put the real folder in place of `<repo folder>`. `ReturnValue` 0 means it started.

**Step 5. Give the user a short tour, once Ekxel is open.** Keep it brief and friendly, about
8 short bullet points, not a manual. On a Mac say **Cmd** wherever this says Ctrl. Cover, in this order:

1. **Ctrl+Shift+A: the headline feature. Make it stand out (bold it, put it first or give it its own
   paragraph).** It opens a Claude chat docked beside the sheet (also View > Claude). That Claude can
   see the workbook, the cells the user has selected and what's in them. It can read, write, format,
   sort, insert rows, add formulas, sheets and slider/spinner controls. Uses their Claude Code login,
   no API key. While it works, the sheet is locked so edits don't collide. **Ctrl+Z undoes everything it
   did in one step.** Suggest a first try: select some numbers and ask "add a total row under this
   and make it bold", or "make me a loan calculator".
2. It opens and saves `.xlsx`, `.xls`, `.csv` and `.tsv` files and works like Excel: formulas (180
   functions), fill handle, copy/paste to and from real Excel, sort and filter, find/replace, freeze
   panes, multiple sheets.
3. **Alt** shows Excel-style KeyTips (Alt, H, O, I autofits columns, as in Excel). Windows only; Option
   on a Mac doesn't do this reliably. **F1** lists all keyboard shortcuts.
4. Insert > Slider / Spin Button links a control to a cell. Drag it and every formula updates live.
   Insert > Chart (Alt N C, or Alt+F1 for an instant one) makes Excel-style charts linked to the cells.
5. View > Skin switches to an Excel 95 look; View > Cursor sets a custom mouse cursor (just for fun).
6. File > Set Default Folder picks where Open/Save start.
7. Only one copy runs. Opening another file adds a window to it.

Then ask whether they want to try the Claude chat now.

**Optional. Ask the user first, then do only what they say yes to.**
- Start menu shortcut:
  `.venv\Scripts\python.exe -c "from sheets import winshell; winshell.write_shortcut(winshell.START_MENU_LNK)"`
- Show Ekxel in "Open with" for csv/xlsx files: `.venv\Scripts\python.exe -m sheets.register`.
  `--remove` undoes it.
- A full live test of the chat: `.venv\Scripts\python.exe tools\e2e_claude.py "put =1+1 in A1"`.
  It uses a little of the user's Claude usage. It should print a transcript and the changed cells.

**If the chat shows an error later:**

| Message in the chat panel | Fix |
| --- | --- |
| "Claude Code isn't installed" | Run `tools\doctor.py` and follow the Claude Code FIX line. |
| "Couldn't connect Claude to this workbook" | Run `setup.bat` again (the `mcp` library is missing). |
| "Claude Code failed ... sign in" | Step 3. |


How it's wired, for changes: `sheets/ai/panel.py` writes a per-window MCP config in `.aiwork\` and starts
`claude -p --output-format stream-json --mcp-config <cfg> --strict-mcp-config --tools "" --allowedTools mcp__sheets
--setting-sources ""`. The `--setting-sources ""` keeps the user's own CLAUDE.md, hooks and settings out of the
panel. A follow-up message resumes the chat with `--resume <session id>`. The bridge (`mcp_server.py`) talks to
the running Ekxel over its QLocalServer socket. `sheets/ai/tools.py` holds the tool implementations, and each
request's edits are one undo step. The model picker in the panel passes `--model` (opus / sonnet / haiku).

## Feature wishlist - ask the user about it

`docs/excel-wishlist.md` lists Excel features Sheets doesn't have yet (generated from the `KEYTIPS`
table in `sheets/ui/keytips.py`; entries with `action=None`). When the user presses one of their
KeyTips, Sheets counts a try in `%APPDATA%\Sheets\wishlist-tried.json` (also shown in
Help > Feature Wishlist).

At the start of a session where the user works on Sheets, and roughly once per session after
that, ask about **one or two** wishlist items: most-tried first (read `wishlist-tried.json`),
otherwise the most generally useful. One line each: what it does in Excel and a rough build cost.
If they say yes, build it (tests + screenshot check), give its KEYTIPS entry an action, run
`tools\make_wishlist.py`, and commit. If they say no, mark it `declined` in the doc's Status column
(keep the row) and don't ask about it again unless they bring it up. Don't ask more than once per
session, and skip it entirely when they're in the middle of something urgent.

## Conventions
- Never use the C: drive for project files, caches or output (user rule); scratch output goes in `.shots/`.
- Every change to cells goes through the undo stack (`_push_states` / `_push_meta` / `SnapshotCommand`).
- PowerShell 5.1 mangles UTF-8 on `Get-Content`/`Set-Content` round-trips: edit files with Python or the Edit tool.
