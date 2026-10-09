# Sheets - notes for Claude

Excel-style spreadsheet (PySide6). Engine in `sheets/`, UI in `sheets/ui/` (custom grid in `grid.py`),
Claude panel in `sheets/ai/` (runs the user's Claude Code headless with an MCP bridge as its only tools).
Tests: `.venv\Scripts\python.exe -m pytest` (offscreen Qt). Screenshots: `tools\snapshot.py`.
Live Claude check (uses the user's Claude usage): `tools\e2e_claude.py "request"`.

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
