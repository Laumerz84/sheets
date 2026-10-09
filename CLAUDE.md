# Exkel (repo/package name: Sheets) - notes for Claude

Excel-style spreadsheet (PySide6). Engine in `sheets/`, UI in `sheets/ui/` (custom grid in `grid.py`),
Claude panel in `sheets/ai/` (runs the user's Claude Code headless with an MCP bridge as its only tools).
Tests: `.venv\Scripts\python.exe -m pytest` (offscreen Qt). Screenshots: `tools\snapshot.py`.
Live Claude check (uses the user's Claude usage): `tools\e2e_claude.py "request"`.

## SETUP: if the user asks you to set up Exkel (or its Claude chat), do exactly this

**What you are setting up.** Exkel has a chat panel docked in its window. The user opens it with
**Ctrl+Shift+A**, or View > Claude. That chat is the user's own Claude Code, using their login.
No API key is needed. In that chat, Claude can only read and change the open workbook.
You do not write any code for this. It is already built. You only install and check.

Run every command from the repo folder (the folder this file is in). The commands are for
PowerShell or cmd on Windows.

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

**Step 4. Start Exkel.** Start it so it keeps running after your session ends:

    Invoke-CimMethod Win32_Process -MethodName Create -Arguments @{ CommandLine = '"<repo folder>\.venv\Scripts\pythonw.exe" "<repo folder>\launch.pyw"'; CurrentDirectory = '<repo folder>' }

Put the real folder in place of `<repo folder>`. `ReturnValue` 0 means it started.

**Step 5. Tell the user how to use it.** Say: "In Exkel, press **Ctrl+Shift+A** to open the Claude
chat. Ask it something small, like 'put 1 to 10 in column A'. **Ctrl+Z** undoes what it did."

**Optional. Ask the user first, then do only what they say yes to.**
- Start menu shortcut:
  `.venv\Scripts\python.exe -c "from sheets import winshell; winshell.write_shortcut(winshell.START_MENU_LNK)"`
- Show Exkel in "Open with" for csv/xlsx files: `.venv\Scripts\python.exe -m sheets.register`.
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
