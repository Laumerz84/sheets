"""Claude side panel: runs Claude Code headless (`claude -p`) with the Sheets
bridge as its only tools, and streams the reply into a chat transcript."""
import json
import os
import shutil
import subprocess
import sys
import threading

from PySide6.QtCore import QObject, Qt, QTimer, Signal
from PySide6.QtGui import QDesktopServices, QFont, QTextCursor
from PySide6.QtWidgets import (QComboBox, QDockWidget, QHBoxLayout, QLabel,
                               QPlainTextEdit, QPushButton, QTextBrowser,
                               QToolButton, QVBoxLayout, QWidget)

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
WORK_DIR = os.path.join(ROOT, ".aiwork")
BRIDGE = os.path.join(ROOT, "sheets", "ai", "mcp_server.py")

MODELS = [("Default model", None), ("Opus", "opus"), ("Sonnet (faster)", "sonnet"), ("Haiku (fastest)", "haiku")]

SYSTEM = (
    "You are Claude, working inside Exkel, a desktop spreadsheet app. The user is looking at the "
    "workbook; act on it only through the sheets tools (you have no file, shell or web access). "
    "Each message starts with a [Exkel: ...] line giving the active sheet and selection; 'this', "
    "'here' or 'the selection' mean that range. Read cells before changing them, use formulas "
    "(not hard-coded results) when values depend on other cells, check formula_errors after "
    "writing, and don't delete or overwrite data the user didn't ask about. Keep replies short "
    "and plain: say what you did with cell addresses, or answer the question. All your changes "
    "from one request are a single undo step (Ctrl+Z) for the user.")


def find_claude():
    """Command (list) that runs Claude Code. SHEETS_CLAUDE_CMD (JSON list) overrides, for tests."""
    override = os.environ.get("SHEETS_CLAUDE_CMD")
    if override:
        return json.loads(override)
    exe = shutil.which("claude")
    if not exe:
        guess = os.path.join(os.path.expanduser("~"), ".local", "bin", "claude.exe")
        exe = guess if os.path.exists(guess) else None
    return [exe] if exe else None


def bridge_python():
    """python.exe (not pythonw) so the bridge has working stdio."""
    exe = sys.executable
    cand = os.path.join(os.path.dirname(exe), "python.exe")
    return cand if os.path.exists(cand) else exe


class _Reader(QObject):
    line = Signal(int, str)
    finished = Signal(int, object, str)


class ClaudePanel(QDockWidget):
    def __init__(self, win):
        super().__init__("Claude", win)
        self.win = win
        self.setObjectName("ClaudePanel")
        self.setAllowedAreas(Qt.RightDockWidgetArea | Qt.LeftDockWidgetArea)
        self.session_id = None
        self.proc = None
        self.generation = 0      # ignores events from a run that was stopped
        self.cfg_path = None
        self.md = ""             # transcript (markdown)
        self.live = ""           # assistant text of the running turn
        self.streamed = False    # got text deltas for the current assistant message
        self.reader = _Reader()
        self.reader.line.connect(self._on_line)
        self.reader.finished.connect(self._on_finished)
        self._render_timer = QTimer(self)
        self._render_timer.setSingleShot(True)
        self._render_timer.setInterval(40)
        self._render_timer.timeout.connect(self._render)

        body = QWidget()
        v = QVBoxLayout(body)
        v.setContentsMargins(8, 6, 8, 8)
        v.setSpacing(6)
        top = QHBoxLayout()
        self.model_box = QComboBox()
        for label, model in MODELS:
            self.model_box.addItem(label, model)
        saved = self._settings().get("claude_model")
        idx = next((i for i, (_, m) in enumerate(MODELS) if m == saved), 0)
        self.model_box.setCurrentIndex(idx)
        self.model_box.currentIndexChanged.connect(self._model_changed)
        new_btn = QToolButton()
        new_btn.setText("New chat")
        new_btn.setToolTip("Start a fresh conversation")
        new_btn.clicked.connect(self.new_chat)
        top.addWidget(self.model_box, 1)
        top.addWidget(new_btn)
        v.addLayout(top)
        self.view = QTextBrowser()
        # Claude's text can contain links (possibly prompted by untrusted cell data): never open
        # files or other schemes, only ordinary web pages, and only when clicked.
        self.view.setOpenLinks(False)
        self.view.setOpenExternalLinks(False)
        self.view.anchorClicked.connect(_open_web_link)
        self.view.setStyleSheet("QTextBrowser { background: #FFFFFF; border: 1px solid #D4D4D4; padding: 4px; }")
        self.view.document().setDefaultFont(QFont("Segoe UI", 10))
        v.addWidget(self.view, 1)
        self.status = QLabel("")
        self.status.setStyleSheet("color: #6B6B6B;")
        self.status.setWordWrap(True)
        v.addWidget(self.status)
        self.input = _Input()
        self.input.setPlaceholderText("Ask Claude to do something with this workbook...  (Enter to send)")
        self.input.setFixedHeight(76)
        self.input.submit.connect(self.send)
        v.addWidget(self.input)
        row = QHBoxLayout()
        row.addStretch(1)
        self.stop_btn = QPushButton("Stop")
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self.stop)
        self.send_btn = QPushButton("Send")
        self.send_btn.setDefault(True)
        self.send_btn.clicked.connect(self.send)
        row.addWidget(self.stop_btn)
        row.addWidget(self.send_btn)
        v.addLayout(row)
        self.setWidget(body)
        self.setMinimumWidth(320)
        self._intro()

    # ------------------------------------------------------------ settings
    def _settings(self):
        from ..ui.mainwindow import load_settings
        return load_settings()

    def _model_changed(self, _):
        from ..ui.mainwindow import load_settings, save_settings
        s = load_settings()
        s["claude_model"] = self.model_box.currentData()
        save_settings(s)

    # ------------------------------------------------------------ transcript
    def _intro(self):
        self.md = ("*Ask in plain English, e.g. \"total each column\", \"make row 1 a bold header with a "
                   "blue fill\", \"why is E14 #N/A?\". Claude works on this workbook only, and "
                   "**Ctrl+Z** undoes everything it did in one go. Uses your Claude Code login.*\n\n")
        self._render()

    def _append(self, text):
        self.md += text
        self._render_timer.start()

    def _render(self):
        self.view.setMarkdown(self.md + self.live)
        self.view.moveCursor(QTextCursor.End)
        self.view.ensureCursorVisible()

    def new_chat(self):
        if self.proc is not None:
            return
        self.session_id = None
        self.live = ""
        self._intro()
        self.status.setText("")

    # ------------------------------------------------------------ running
    def busy(self):
        return self.proc is not None

    def send(self):
        text = self.input.toPlainText().strip()
        if not text or self.busy():
            return
        cmd = find_claude()
        if not cmd or not cmd[0]:
            self._append("\n\n**Claude Code isn't installed** (couldn't find `claude`). Install it, sign in "
                         "once in a terminal with `claude`, then try again.\n\n")
            return
        win = self.win
        win._prep()
        from ..app import SERVER_NAME
        cfg = os.path.join(WORK_DIR, f"mcp-{win.ai_token}.json")
        try:
            os.makedirs(WORK_DIR, exist_ok=True)
            with open(cfg, "w", encoding="utf-8") as fh:
                json.dump({"mcpServers": {"sheets": {
                    "type": "stdio", "command": bridge_python(), "args": [BRIDGE],
                    "env": {"SHEETS_TARGET": win.ai_token, "SHEETS_SERVER_NAME": SERVER_NAME}}}}, fh)
        except OSError as e:
            self._append(f"\n\n**Couldn't prepare the Claude connection:** {_md_escape(str(e))}\n\n")
            return
        self.cfg_path = cfg
        self.input.clear()
        self._append(f"\n\n---\n\n**You:** {_md_escape(text)}\n\n")
        self.live = ""
        self.streamed = False
        g = win.grid
        sel = ", ".join(_rng(g.sheet, rc) for rc in g.sel.rects)
        context = (f"[Exkel: file {os.path.basename(win.wb.path) if win.wb.path else 'unsaved new workbook'}, "
                   f"active sheet '{g.sheet.name}', selection {sel}]")
        args = cmd + ["-p", "--output-format", "stream-json", "--verbose", "--include-partial-messages",
                      "--mcp-config", cfg, "--strict-mcp-config", "--tools", "",
                      "--allowedTools", "mcp__sheets", "--append-system-prompt", SYSTEM,
                      "--setting-sources", "", "--effort", "medium"]
        model = self.model_box.currentData()
        if model:
            args += ["--model", model]
        if self.session_id:
            args += ["--resume", self.session_id]
        self.status.setText("Starting Claude...")
        self.send_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        flags = 0x08000000 if sys.platform == "win32" else 0  # CREATE_NO_WINDOW
        try:
            self.proc = subprocess.Popen(args, cwd=WORK_DIR, stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                                         stderr=subprocess.PIPE, creationflags=flags)
        except OSError as e:
            self._done_ui()
            self._append(f"**Couldn't start Claude Code:** {_md_escape(str(e))}\n\n")
            return
        win.ai_tools.begin_turn(text[:50])
        win.grid.read_only = True
        self.generation += 1
        gen = self.generation
        proc = self.proc
        prompt = (context + "\n" + text).encode("utf-8")

        def pump():
            err = b""
            try:
                proc.stdin.write(prompt)
                proc.stdin.close()
            except OSError:
                pass

            def drain_err():
                nonlocal err
                err = proc.stderr.read()
            t = threading.Thread(target=drain_err, daemon=True)
            t.start()
            for raw in proc.stdout:
                self.reader.line.emit(gen, raw.decode("utf-8", "replace"))
            rc = proc.wait()
            t.join(2)
            self.reader.finished.emit(gen, rc, err.decode("utf-8", "replace")[-2000:])
        threading.Thread(target=pump, daemon=True).start()

    def stop(self, finish_now=False):
        if self.proc is None:
            return
        self.win.ai_tools.stopped = True   # refuse any tool call still in the pipe
        pid = self.proc.pid
        if sys.platform == "win32":
            subprocess.run(["taskkill", "/T", "/F", "/PID", str(pid)], capture_output=True,
                           creationflags=0x08000000)
        else:
            self.proc.kill()
        self.status.setText("Stopped.")
        if finish_now:  # window closing: don't wait for the process to report back
            self._on_finished(self.generation, None, "")

    def _done_ui(self):
        self.send_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)

    # ------------------------------------------------------------ stream-json events
    def _on_line(self, gen, line):
        if gen != self.generation or self.proc is None:
            return
        line = line.strip()
        if not line:
            return
        try:
            ev = json.loads(line)
        except ValueError:
            return
        t = ev.get("type")
        if t == "system" and ev.get("subtype") == "init":
            self.session_id = ev.get("session_id") or self.session_id
            servers = {s.get("name"): s.get("status") for s in ev.get("mcp_servers", [])}
            if servers.get("sheets") not in (None, "connected"):
                self._append(f"*Couldn't connect Claude to this workbook ({servers.get('sheets')}).*\n\n")
            self.status.setText("Thinking...")
        elif t == "stream_event":
            e = ev.get("event", {})
            if e.get("type") == "content_block_delta" and e.get("delta", {}).get("type") == "text_delta":
                self.live += e["delta"].get("text", "")
                self.streamed = True
                self._render_timer.start()
            elif e.get("type") == "message_start":
                self.streamed = False
        elif t == "assistant":
            for block in ev.get("message", {}).get("content", []):
                if block.get("type") == "text" and not self.streamed:
                    self.live += block.get("text", "")
                elif block.get("type") == "tool_use":
                    self._flush_live()
                    name = block.get("name", "").replace("mcp__sheets__", "")
                    self.status.setText(f"Working: {name.replace('_', ' ')}...")
                    if name not in ("workbook_info", "read_range", "find"):
                        self._append(f"*▸ {_describe(name, block.get('input', {}))}*\n\n")
            self._flush_live()
            self.streamed = False
        elif t == "user":
            for block in ev.get("message", {}).get("content", []) if isinstance(ev.get("message", {}).get("content"), list) else []:
                if block.get("type") == "tool_result" and block.get("is_error"):
                    self._append(f"*⚠ {_tool_error_text(block)}*\n\n")
        elif t == "result":
            self.session_id = ev.get("session_id") or self.session_id
            self._flush_live()
            if ev.get("is_error") or ev.get("subtype") != "success":
                msg = ev.get("result") or ev.get("subtype") or "error"
                self._append(f"\n\n**Claude stopped:** {_md_escape(str(msg))[:600]}\n\n")
            secs = (ev.get("duration_ms") or 0) / 1000
            self.status.setText(f"Done in {secs:.1f}s." if secs else "Done.")

    def _flush_live(self):
        if self.live.strip():
            self.md += self.live.rstrip() + "\n\n"
        self.live = ""
        self._render_timer.start()

    def _on_finished(self, gen, rc, err):
        if gen != self.generation or self.proc is None:
            return
        self.proc = None
        self._flush_live()
        self.win.ai_tools.end_turn()
        self.win.grid.read_only = False
        self._done_ui()
        if self.cfg_path:
            try:
                os.remove(self.cfg_path)
            except OSError:
                pass
            self.cfg_path = None
        if rc not in (0, None) and not self.status.text().startswith(("Done", "Stopped")):
            hint = err.strip().splitlines()[-1] if err.strip() else f"exit code {rc}"
            if "login" in err.lower() or "auth" in err.lower():
                hint += "  (Open a terminal, run `claude`, and sign in once.)"
            self._append(f"\n\n**Claude Code failed:** {_md_escape(hint)[:500]}\n\n")
            self.status.setText("")
        self.win.grid.invalidate()
        self.input.setFocus()


class _Input(QPlainTextEdit):
    submit = Signal()

    def keyPressEvent(self, e):
        if e.key() in (Qt.Key_Return, Qt.Key_Enter) and not (e.modifiers() & Qt.ShiftModifier):
            self.submit.emit()
            return
        super().keyPressEvent(e)


def _rng(sheet, rect):
    from ..refs import MAX_COLS, MAX_ROWS, range_addr
    return range_addr(*rect)


def _describe(name, inp):
    where = inp.get("range") or inp.get("start_cell") or inp.get("columns") or inp.get("at") or ""
    sheet = inp.get("sheet")
    loc = f"{sheet}!{where}" if sheet and where else (where or sheet or "")
    label = {"write_range": "Wrote", "format_range": "Formatted", "clear_range": "Cleared",
             "add_sheet": "Added sheet", "sort_range": "Sorted", "insert_or_delete": inp.get("action", "").replace("_", " ").capitalize(),
             "set_column_width": "Column width", "select_range": "Selected"}.get(name, name)
    if name == "add_sheet":
        loc = inp.get("name", "")
    return f"{label} {loc}".strip()


def _tool_error_text(block):
    c = block.get("content")
    if isinstance(c, list):
        c = " ".join(x.get("text", "") for x in c if isinstance(x, dict))
    return _md_escape(str(c))[:300]


def _open_web_link(url):
    if url.scheme() in ("http", "https"):
        QDesktopServices.openUrl(url)


def _md_escape(s):
    out = s
    for ch in "\\`*_[]#<>":
        out = out.replace(ch, "\\" + ch)
    return out
