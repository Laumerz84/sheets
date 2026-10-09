"""Undoable commands."""
from PySide6.QtGui import QUndoCommand, QUndoStack


class GuardedUndoStack(QUndoStack):
    """Undo stack that refuses the user's own changes while `guard()` is true
    (Claude is working: they would otherwise end up inside Claude's undo step)."""

    def __init__(self, parent, guard, on_blocked):
        super().__init__(parent)
        self.guard = guard
        self.on_blocked = on_blocked
        self._blocked_depth = 0

    def _blocked(self):
        return self._blocked_depth > 0 or self.guard()

    def push(self, cmd):
        if self._blocked():
            self.on_blocked()
            return
        super().push(cmd)

    def beginMacro(self, text):
        if self._blocked():
            self._blocked_depth += 1
            self.on_blocked()
            return
        super().beginMacro(text)

    def endMacro(self):
        if self._blocked_depth:
            self._blocked_depth -= 1
            return
        super().endMacro()

    def undo(self):
        if self.guard():
            self.on_blocked()
            return
        super().undo()

    def redo(self):
        if self.guard():
            self.on_blocked()
            return
        super().redo()


class _Base(QUndoCommand):
    def __init__(self, win, sheet, text):
        super().__init__(text)
        self.win = win
        self.sheet = sheet
        self.sel_before = win.grid.sel.copy() if win.grid.sheet is sheet else None
        self.sel_after = None
        self._first = True

    def _show(self, sel):
        self.win.show_sheet(self.sheet)
        if sel is not None:
            self.win.grid.set_selection(sel.rects, active=sel.active)


class StatesCommand(_Base):
    """Cell contents/styles change (typing, paste, clear, format, sort, fill...)."""

    def __init__(self, win, sheet, states, text, select=None):
        super().__init__(win, sheet, text)
        self.new = states
        self.old = None
        self.select = select

    def redo(self):
        wb = self.win.wb
        old = wb.apply_states(self.sheet, self.new)
        if self.old is None:
            self.old = old
        if self._first:
            self._first = False
            if self.select is not None:
                self.win.grid.set_selection(*self.select)
            self.sel_after = self.win.grid.sel.copy()
        else:
            self._show(self.sel_after)
        self.win.after_change(self.sheet)

    def undo(self):
        self.win.wb.apply_states(self.sheet, self.old)
        self._show(self.sel_before)
        self.win.after_change(self.sheet)


class MetaCommand(_Base):
    """Sheet attribute change: widths, heights, hidden rows/cols, merges, freeze, filters."""

    def __init__(self, win, sheet, changes, text, relayout=True):
        super().__init__(win, sheet, text)
        self.changes = changes  # {attr: (old, new)}
        self.relayout = relayout

    def _apply(self, idx):
        for attr, pair in self.changes.items():
            val = pair[idx]
            if isinstance(val, dict):
                val = dict(val)
            elif isinstance(val, set):
                val = set(val)
            elif isinstance(val, list):
                val = list(val)
            setattr(self.sheet, attr, val)

    def redo(self):
        self._apply(1)
        if self._first:
            self._first = False
            self.sel_after = self.win.grid.sel.copy()
        else:
            self._show(self.sel_after)
        self.win.after_change(self.sheet, relayout=self.relayout)

    def undo(self):
        self._apply(0)
        self._show(self.sel_before)
        self.win.after_change(self.sheet, relayout=self.relayout)


class SnapshotCommand(_Base):
    """Structural change (insert/delete rows or columns, sheet add/remove/rename/move).
    The change is performed by `action()` on first redo; undo/redo restore snapshots."""

    def __init__(self, win, sheet, text, action, sheets=None):
        super().__init__(win, sheet, text)
        self.action = action
        self.sheets = sheets
        self.before = None
        self.after = None

    def redo(self):
        wb = self.win.wb
        if self.before is None:
            self.before = wb.snapshot(self.sheets)
            self.action()
            self.after = wb.snapshot(self.sheets)
            self.sel_after = self.win.grid.sel.copy()
            self.win.after_structure()
        else:
            wb.restore(self.after)
            self.win.after_structure()
            if self.sheet in wb.sheets:
                self._show(self.sel_after)

    def undo(self):
        self.win.wb.restore(self.before)
        self.win.after_structure()
        if self.sheet in self.win.wb.sheets:
            self._show(self.sel_before)
