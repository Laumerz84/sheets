"""Form controls linked to a cell, like Excel's Scroll Bar and Spin Button (Developer > Insert).

A control floats over the cells in `place` (r1, c1, r2, c2) and drives the number in `link` (r, c):
dragging a slider or clicking a spinner writes the linked cell live, so every formula that uses
it recalculates as you go, and the whole gesture becomes one undo step.

Controls are stored per sheet in `Sheet.controls` as plain dicts (never mutated: edits replace
the list, which MetaCommand undoes) and saved in xlsx files in a very hidden sheet
(see io_xlsx: CONTROLS_SHEET)."""
import math

from PySide6.QtCore import QRect, Qt, QTimer, Signal
from PySide6.QtWidgets import (QDialog, QDialogButtonBox, QDoubleSpinBox, QFormLayout, QHBoxLayout,
                               QLabel, QLineEdit, QMenu, QMessageBox, QScrollBar, QSizePolicy, QToolButton,
                               QVBoxLayout, QWidget)

from ..refs import MAX_COLS, MAX_ROWS, addr, parse_addr, parse_range, range_addr
from ..values import is_num

KINDS = {"scrollbar": "Slider (scroll bar)", "spinner": "Spin button"}


def make_control(kind, place, link, vmin=0.0, vmax=100.0, step=1.0):
    if kind not in KINDS:
        raise ValueError("kind must be 'scrollbar' or 'spinner'")
    vmin, vmax, step = float(vmin), float(vmax), float(step)
    for v in (vmin, vmax, step):
        if not math.isfinite(v):
            raise ValueError("min, max and step must be numbers")
    if vmax <= vmin:
        raise ValueError("max must be greater than min")
    if step <= 0:
        raise ValueError("step must be greater than 0")
    if (vmax - vmin) / step > 1_000_000:
        raise ValueError("too many steps between min and max")
    r1, c1, r2, c2 = place
    if not (0 <= r1 <= r2 < MAX_ROWS and 0 <= c1 <= c2 < MAX_COLS):
        raise ValueError("bad placement")
    return {"kind": kind, "place": tuple(place), "link": tuple(link), "min": vmin, "max": vmax, "step": step}


def describe(ctl):
    return (f"{KINDS[ctl['kind']]} at {range_addr(*ctl['place'])} -> {addr(*ctl['link'])} "
            f"({_fmt(ctl['min'])} to {_fmt(ctl['max'])}, step {_fmt(ctl['step'])})")


def _fmt(x):
    return str(int(x)) if float(x).is_integer() else f"{x:g}"


def snap(ctl, v):
    """Clamp v to [min, max] on the step grid."""
    n = round((v - ctl["min"]) / ctl["step"])
    n = max(0, min(n, int(round((ctl["max"] - ctl["min"]) / ctl["step"]))))
    x = ctl["min"] + n * ctl["step"]
    return round(x, 10)


def remap(ctl, move_row, move_col, remap_rect):
    """Shift a control after rows/cols are inserted or deleted; None if its cells were deleted."""
    place = remap_rect(ctl["place"])
    lr, lc = ctl["link"]
    lr, lc = move_row(lr), move_col(lc)
    if place is None or lr is None or lc is None:
        return None
    return dict(ctl, place=tuple(place), link=(lr, lc))


# ---------------------------------------------------------------- widgets

class _Spinner(QWidget):
    stepped = Signal(int)  # +1 / -1

    def __init__(self, parent, horizontal):
        super().__init__(parent)
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setAutoFillBackground(True)   # solid, like Excel's spin button (cell text mustn't show through)
        self.setStyleSheet("_Spinner { background: palette(button); }")
        lay = (QHBoxLayout if horizontal else QVBoxLayout)(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)
        for sign, arrow in ((-1, Qt.LeftArrow if horizontal else Qt.DownArrow),
                            (1, Qt.RightArrow if horizontal else Qt.UpArrow)):
            b = QToolButton(self)
            b.setArrowType(arrow)
            b.setAutoRepeat(True)
            b.setAutoRaise(False)
            b.setAutoFillBackground(True)
            b.setAutoRepeatDelay(350)
            b.setAutoRepeatInterval(60)
            b.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
            b.clicked.connect(lambda _=False, s=sign: self.stepped.emit(s))
            if horizontal:
                lay.addWidget(b)
            else:
                lay.insertWidget(0, b)  # up arrow on top


class ControlLayer:
    """Keeps a widget over each control of the grid's current sheet and wires it to its cell."""

    def __init__(self, win):
        self.win = win
        self.grid = win.grid
        self.widgets = []       # (ctl dict, widget)
        self._for = None        # (sheet, list identity) the widgets were built for
        self.gesture = None     # {"ctl":..., "key":..., "old": state} while the user drags/clicks
        self._commit = QTimer(self.grid)
        self._commit.setSingleShot(True)
        self._commit.setInterval(450)
        self._commit.timeout.connect(self._finish_gesture)

    # ------------------------------------------------------------ building / placing
    def sync(self):
        sh = self.grid.sheet
        if sh is None:
            return
        controls = getattr(sh, "controls", [])
        if self._for != (sh, id(controls), len(controls)):
            if self.gesture:
                self._finish_gesture()
            for _, w in self.widgets:
                w.hide()
                w.deleteLater()
            self.widgets = [(c, self._make(c)) for c in controls]
            self._for = (sh, id(controls), len(controls))
        self.place()

    def _make(self, ctl):
        r1, c1, r2, c2 = ctl["place"]
        horizontal = (c2 - c1 + 1) * 64 >= (r2 - r1 + 1) * 20
        if ctl["kind"] == "scrollbar":
            w = QScrollBar(Qt.Horizontal if horizontal else Qt.Vertical, self.grid)
            n = int(round((ctl["max"] - ctl["min"]) / ctl["step"]))
            w.setRange(0, n)
            w.setSingleStep(1)
            w.setPageStep(max(1, n // 10))
            w.valueChanged.connect(lambda i, c=ctl: self._set(c, c["min"] + i * c["step"]))
            w.sliderReleased.connect(self._finish_gesture)
        else:
            w = _Spinner(self.grid, horizontal)
            w.stepped.connect(lambda s, c=ctl: self._step(c, s))
        w.setToolTip(describe(ctl) + "\nRight-click to change or delete.")
        w.setContextMenuPolicy(Qt.CustomContextMenu)
        w.customContextMenuRequested.connect(lambda pos, c=ctl, ww=w: self._menu(c, ww.mapToGlobal(pos)))
        w.show()
        return w

    def place(self):
        g = self.grid
        locked = g.read_only or self.win.ai_tools.user_locked()
        for ctl, w in self.widgets:
            r1, c1, r2, c2 = ctl["place"]
            a, b = g.cell_rect(r1, c1), g.cell_rect(r2, c2)
            rect = QRect(a.topLeft(), b.bottomRight()).adjusted(1, 1, -1, -1)
            visible = rect.width() > 4 and rect.height() > 4 and rect.top() >= g.hh and rect.left() >= g.rw \
                and rect.top() < g.height() and rect.left() < g.width()
            w.setVisible(visible)
            if not visible:
                continue
            w.setGeometry(rect)
            w.setEnabled(not locked)
            if isinstance(w, QScrollBar) and not (self.gesture and self.gesture["ctl"] is ctl) and not w.isSliderDown():
                v = g.sheet.value(*ctl["link"])
                i = int(round((snap(ctl, v) - ctl["min"]) / ctl["step"])) if is_num(v) else 0
                if w.value() != i:
                    w.blockSignals(True)
                    w.setValue(i)
                    w.blockSignals(False)

    # ------------------------------------------------------------ changing the linked cell
    def _current(self, ctl):
        v = self.grid.sheet.value(*ctl["link"])
        return snap(ctl, v) if is_num(v) else ctl["min"]

    def _step(self, ctl, sign):
        self._set(ctl, self._current(ctl) + sign * ctl["step"])

    def _set(self, ctl, value):
        from .. import ops
        from ..refs import key
        win = self.win
        if self.grid.read_only or win.ai_tools.user_locked():
            win._claude_busy_message()
            return
        sh = self.grid.sheet
        value = snap(ctl, value)
        k = key(*ctl["link"])
        if self.gesture is None or self.gesture["ctl"] is not ctl:
            if self.gesture:
                self._finish_gesture()
            self.gesture = {"ctl": ctl, "key": k, "old": sh.get_state(k), "sheet": sh}
        text = repr(value) if not float(value).is_integer() else str(int(value))
        sh.set_state(k, ops.input_state(sh, *ctl["link"], text))
        win.wb.recalc()  # live: everything that depends on the cell updates while dragging
        self.grid.invalidate()
        win._selection_changed()
        self._commit.start()

    def _finish_gesture(self):
        """Turn the live changes into one undo step."""
        self._commit.stop()
        g = self.gesture
        self.gesture = None
        if not g:
            return
        sh, k = g["sheet"], g["key"]
        new = sh.get_state(k)
        if new == g["old"]:
            return
        sh.set_state(k, g["old"])       # put the old value back quietly...
        self.win._push_states({k: new}, "Change control value", sheet=sh)  # ...and redo it as one step

    # ------------------------------------------------------------ editing controls
    def _menu(self, ctl, pos):
        m = QMenu(self.grid)
        m.addAction("Format Control...", lambda: self.win.edit_control(ctl))
        m.addAction("Delete Control", lambda: self.win.delete_control(ctl))
        m.exec(pos)


class ControlDialog(QDialog):
    """Create or change a control: linked cell, where it sits, min / max / step."""

    def __init__(self, parent, kind, link, place, vmin=0.0, vmax=100.0, step=1.0, title=None):
        super().__init__(parent)
        self.kind = kind
        self.setWindowTitle(title or f"Insert {KINDS[kind]}")
        lay = QVBoxLayout(self)
        lay.addWidget(QLabel("The control changes the number in the linked cell; formulas that use "
                             "that cell update as you drag or click."))
        form = QFormLayout()
        self.link_edit = QLineEdit(addr(*link))
        self.place_edit = QLineEdit(range_addr(*place))
        self.min_box, self.max_box, self.step_box = (QDoubleSpinBox() for _ in range(3))
        for box, v in ((self.min_box, vmin), (self.max_box, vmax), (self.step_box, step)):
            box.setDecimals(4)
            box.setRange(-1e9, 1e9)
            box.setValue(v)
        form.addRow("Linked cell:", self.link_edit)
        form.addRow("Place over cells:", self.place_edit)
        form.addRow("Minimum value:", self.min_box)
        form.addRow("Maximum value:", self.max_box)
        form.addRow("Step (per click):", self.step_box)
        lay.addLayout(form)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self._ok)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)
        self.result_ctl = None

    def _ok(self):
        link = parse_addr(self.link_edit.text().replace("$", ""))
        place = parse_range(self.place_edit.text().replace("$", ""))
        try:
            if link is None:
                raise ValueError("Linked cell must be a single cell like B4.")
            if place is None or place[2] >= MAX_ROWS - 1 or place[3] >= MAX_COLS - 1:
                raise ValueError("Place over cells must be a cell or range like C4:F4.")
            self.result_ctl = make_control(self.kind, place, link, self.min_box.value(),
                                           self.max_box.value(), self.step_box.value())
        except ValueError as e:
            QMessageBox.warning(self, self.windowTitle(), str(e))
            return
        self.accept()


def default_place(kind, link):
    """Next to the linked cell: three cells wide for a slider, one cell for a spinner."""
    r, c = link
    c1 = min(c + 1, MAX_COLS - 4)
    return (r, c1, r, c1 + (2 if kind == "scrollbar" else 0))
