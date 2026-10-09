# Exkel (repo/package name: Sheets) - notes for Claude

Excel-style spreadsheet (PySide6). Engine in `sheets/`, UI in `sheets/ui/` (custom grid in `grid.py`),
Claude panel in `sheets/ai/` (runs the user's Claude Code headless with an MCP bridge as its only tools).
Tests: `.venv\Scripts\python.exe -m pytest` (offscreen Qt). Screenshots: `tools\snapshot.py`.
Live Claude check (uses the user's Claude usage): `tools\e2e_claude.py "request"`.

## Fresh clone: set Exkel up and hook yourself in as its chat

Exkel has a **Claude panel** (View > Claude, **Ctrl+Shift+A**): a chat docked in the Exkel window that
works on the open workbook. It isn't a separate AI. It runs *the user's own Claude Code*
(`claude -p`, their login and plan, no API key) with one MCP server, `sheets/ai/mcp_server.py`,
as its only tools. That Claude can read and write cells, formulas and formatting in that window, and
has no file, shell or web access. Ctrl+Z undoes everything it did for one request. To hook it up:

1. **Install Exkel** (Windows, Python 3.11+): from the repo folder run `python -m venv .venv`, then
   `.venv\Scripts\python.exe -m pip install -r requirements.txt`. This includes `mcp`, which the
   bridge needs. Start it with `.venv\Scripts\pythonw.exe launch.pyw [file]`. Launch it detached
   from your session (e.g. `Invoke-CimMethod Win32_Process -MethodName Create -Arguments
   @{ CommandLine = '"<repo>\.venv\Scripts\pythonw.exe" "<repo>\launch.pyw"'; CurrentDirectory = '<repo>' }`),
   because processes started from a Claude Code session can die with it. Optional, ask first:
   a Start menu shortcut with the icon,
   `.venv\Scripts\python.exe -c "from sheets import winshell; winshell.write_shortcut(winshell.START_MENU_LNK)"`,
   and "Open with" entries via `.venv\Scripts\python.exe -m sheets.register` (`--remove` undoes; HKCU only).
2. **Make sure Claude Code is installed for the user.** Exkel looks for `claude` on PATH, then
   `%USERPROFILE%\.local\bin\claude.exe` (where the native installer puts it:
   `irm https://claude.ai/install.ps1 | iex`). If it lives somewhere else, set the user env var
   `SHEETS_CLAUDE_CMD` to a JSON list, e.g. `["C:\\path\\to\\claude.exe"]`.
3. **Signed in:** the user must have run `claude` once in a terminal and logged in. They do the login
   themselves; never type their credentials.
4. **Check the hookup.** Run `.venv\Scripts\python.exe tools\e2e_claude.py "put =1+1 in A1"`. It
   drives a real panel request offscreen and uses a little of the user's Claude usage, so ask before
   running it. Or have the user open Exkel, press Ctrl+Shift+A and ask something small.
   Panel messages and what they mean:
   - "Claude Code isn't installed": step 2.
   - "Couldn't connect Claude to this workbook (failed)": the bridge didn't start. Usually `mcp` is
     missing from `.venv`; re-run the pip install.
   - "Claude Code failed ... sign in": step 3.

How it's wired, for changes: `sheets/ai/panel.py` writes a per-window MCP config in `.aiwork\` and starts
`claude -p --output-format stream-json --mcp-config <cfg> --strict-mcp-config --tools "" --allowedTools mcp__sheets
--setting-sources ""`. The `--setting-sources ""` keeps the user's own CLAUDE.md, hooks and settings out of the
panel. A follow-up message resumes the chat with `--resume <session id>`. The bridge (`mcp_server.py`) talks to
the running Exkel over its QLocalServer socket. `sheets/ai/tools.py` holds the tool implementations, and each
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
