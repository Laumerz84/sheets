"""Render the main window offscreen to a PNG (for checking the UI without a display).

usage: python tools/snapshot.py out.png [file-to-open] [--script script.py]
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QT_QPA_FONTDIR", r"C:\Windows\Fonts")
os.environ["APPDATA"] = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), ".shots")

from PySide6.QtWidgets import QApplication  # noqa: E402

from sheets import ops  # noqa: E402
from sheets.refs import key, parse_addr  # noqa: E402
from sheets.ui.mainwindow import MainWindow, app_icon  # noqa: E402
from sheets.ui.style import apply_palette  # noqa: E402
from sheets.workbook import new_workbook  # noqa: E402


def demo_workbook():
    wb = new_workbook()
    sh = wb.sheets[0]
    rows = [
        ["Region", "Product", "Units", "Price", "Revenue", "Date", "Notes"],
        ["North", "Widget", "120", "$4.50", "=C2*D2", "2024-01-15", "First order of the year, shipped early"],
        ["South", "Gadget", "75", "$12.00", "=C3*D3", "2024-02-03", ""],
        ["East", "Widget", "210", "$4.50", "=C4*D4", "2024-02-20", "Bulk"],
        ["West", "Doohickey", "33", "$27.25", "=C5*D5", "2024-03-11", ""],
        ["North", "Gadget", "98", "$12.00", "=C6*D6", "2024-03-28", "Repeat customer"],
        ["", "", "", "", "", "", ""],
        ["Total", "", "=SUM(C2:C6)", "", "=SUM(E2:E6)", "", ""],
        ["Average", "", "=AVERAGE(C2:C6)", "", "=AVERAGE(E2:E6)", "", ""],
        ["Check", "", '=IF(C8>500,"big","small")', "", "=E8/0", "", ""],
    ]
    for r, row in enumerate(rows):
        for c, text in enumerate(row):
            if text:
                sh.set_state(key(r, c), ops.input_state(sh, r, c, text))
    wb.recalc()
    for c in range(7):
        sh.set_style(0, c, sh.style(0, c).with_(bold=True, fill="#E2EFDA"))
    sh.set_style(7, 0, sh.style(7, 0).with_(bold=True))
    sh.col_widths[6] = 120
    return wb


def main():
    out = sys.argv[1]
    app = QApplication([])
    apply_palette(app)
    app.setWindowIcon(app_icon())
    args = sys.argv[2:]
    script = None
    if "--script" in args:
        i = args.index("--script")
        script = open(args[i + 1], encoding="utf-8-sig").read()
        args = args[:i] + args[i + 2:]
    if args:
        from sheets.fileio import open_file
        wb = open_file(args[0])
    else:
        wb = demo_workbook()
    w = MainWindow(wb)
    w.resize(1500, 800)
    w.show()
    app.processEvents()
    if script:
        exec(script, {"w": w, "app": app, "g": w.grid, "sh": w.grid.sheet, "ops": ops})
        app.processEvents()
    w.grab().save(out)
    print("saved", out)


if __name__ == "__main__":
    main()
