"""Rough performance check on a large CSV.  usage: python -u tools/bench.py [rows]"""
import csv
import os
import random
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
os.environ["QT_QPA_PLATFORM"] = "offscreen"
os.environ["QT_QPA_FONTDIR"] = r"C:\Windows\Fonts"
os.environ["APPDATA"] = os.path.join(ROOT, ".shots")  # keep benchmark files out of the real recent list

N = int(sys.argv[1]) if len(sys.argv) > 1 else 200000
p = os.path.join(ROOT, ".shots", f"big{N}.csv")


class T:
    def __init__(self, label):
        self.label = label

    def __enter__(self):
        self.t = time.perf_counter()
        print(f"{self.label} ...", end=" ", flush=True)

    def __exit__(self, *a):
        print(f"{time.perf_counter() - self.t:.2f}s", flush=True)


if not os.path.exists(p):
    random.seed(1)
    with open(p, "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(["id", "name", "city", "amount", "qty", "date", "flag", "code", "score", "note"])
        cities = ["Boston", "Austin", "Denver", "Seattle", "Miami"]
        for i in range(N):
            w.writerow([i, f"Person {i}", random.choice(cities), f"{random.uniform(1, 1000):.2f}",
                        random.randint(1, 50), f"2024-{random.randint(1, 12):02d}-{random.randint(1, 28):02d}",
                        random.choice(["TRUE", "FALSE"]), f"{random.randint(0, 99999):05d}",
                        round(random.random(), 4), "lorem ipsum" if i % 7 == 0 else ""])

from sheets import ops  # noqa: E402
from sheets.fileio import open_file  # noqa: E402
from sheets.refs import key  # noqa: E402

with T("load csv"):
    wb = open_file(p)
sh = wb.sheets[0]
print("cells", len(sh.values), flush=True)

from PySide6.QtWidgets import QApplication  # noqa: E402

app = QApplication([])
from sheets.ui.mainwindow import MainWindow  # noqa: E402
from sheets.ui.style import apply_palette  # noqa: E402

apply_palette(app)
with T("window"):
    w = MainWindow(wb)
    w.resize(2560, 1400)
    w.show()
    app.processEvents()
with T("10 paints"):
    for _ in range(10):
        w.grid.repaint()
with T("scroll + paint"):
    w.grid.scroll_rows(N // 2)
    w.grid.repaint()
with T("=SUM(D:D)"):
    w._push_states({key(0, 11): ops.input_state(sh, 0, 11, "=SUM(D:D)")}, "x")
print(" ->", sh.value(0, 11))
with T("edit + recalc"):
    w._push_states({key(5, 3): ops.input_state(sh, 5, 3, "1000000")}, "x")
print(" ->", sh.value(0, 11))
with T("sort desc by amount"):
    w.grid.set_selection([(0, 0, 0, 0)])
    w.quick_sort(False, col=3)
with T("filter on"):
    w.grid.set_selection([(0, 0, 0, 0)])
    w.toggle_filter()
with T("filter city=Boston"):
    w.apply_filter(2, {"values": {"Boston"}, "blanks": False})
    app.processEvents()
print(" ->", w.filter_lbl.text())
with T("paint filtered"):
    w.grid.repaint()
with T("select all stats"):
    w.grid.select_all()
    w._update_stats()
with T("undo x4"):
    for _ in range(4):
        w.undo.undo()
with T("save xlsx"):
    w.do_save(os.path.join(ROOT, ".shots", "big_out.xlsx"))
with T("save csv"):
    w.do_save(os.path.join(ROOT, ".shots", "big_out.csv"))
w.undo.setClean()
w.close()
