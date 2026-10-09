"""End-to-end check of the Claude panel with the real Claude Code (uses your Claude usage).

usage: python tools/e2e_claude.py "request text" [model]
Runs an invisible Sheets window with a small sales table, sends the request through the
panel exactly as a user would, and prints the transcript plus the resulting cells."""
import os
import sys
import time
import uuid

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QT_QPA_FONTDIR", r"C:\Windows\Fonts")
os.environ["APPDATA"] = os.path.join(ROOT, ".shots")  # keep test settings out of the real ones

from PySide6.QtWidgets import QApplication  # noqa: E402

from sheets import app as sheets_app  # noqa: E402
from sheets import ops  # noqa: E402
from sheets.refs import key  # noqa: E402
from sheets.ui.mainwindow import MainWindow  # noqa: E402

request = sys.argv[1] if len(sys.argv) > 1 else "Add a Total row under the data and make the header row bold."
model = sys.argv[2] if len(sys.argv) > 2 else None

app = QApplication([])
name = "SheetsE2E-" + uuid.uuid4().hex[:8]
sheets_app.SERVER_NAME = name
server = sheets_app._start_server(lambda p: None, name=name)
w = MainWindow()
w.resize(1300, 800)
w.show()
sh = w.grid.sheet
rows = [["Region", "Q1", "Q2", "Q3"], ["North", "120", "135", "150"], ["South", "90", "80", "110"],
        ["East", "200", "190", "240"], ["West", "60", "75", "70"]]
for r, row in enumerate(rows):
    for c, t in enumerate(row):
        sh.set_state(key(r, c), ops.input_state(sh, r, c, t))
w.wb.recalc()
panel = w.claude_panel
panel.show()
if model:
    i = panel.model_box.findData(model)
    panel.model_box.setCurrentIndex(i if i >= 0 else 0)
before = w.undo.index()
t0 = time.time()
panel.input.setPlainText(request)
panel.send()
while panel.busy() and time.time() - t0 < 300:
    app.processEvents()
    time.sleep(0.02)
app.processEvents()
print(f"--- took {time.time() - t0:.1f}s, undo steps added: {w.undo.index() - before}")
print("--- transcript:\n" + (panel.md + panel.live)[-3000:])
mr, mc = sh.used_extent()
print("--- sheet:")
for r in range(mr + 1):
    print([ops.display_text(sh, r, c) + (f" [{sh.formula_text(r, c)}]" if sh.formula_text(r, c) else "")
           for c in range(mc + 1)], "bold" if sh.style(r, 0).bold else "")
w.grab().save(os.path.join(ROOT, ".shots", "e2e.png"))
w.undo.setClean()
w.close()
server.close()
