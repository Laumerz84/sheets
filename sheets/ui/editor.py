"""In-cell editor / formula bar text widget."""
import re

from PySide6.QtCore import QEvent, QPoint, Qt, Signal
from PySide6.QtGui import (QColor, QFont, QFontMetrics, QSyntaxHighlighter,
                           QTextCharFormat, QTextCursor, QTextOption)
from PySide6.QtWidgets import (QFrame, QLabel, QListWidget, QListWidgetItem,
                               QPlainTextEdit)

from ..formula import ref_spans, toggle_absolute
from ..functions import ALL_NAMES, SIGNATURES
from .style import REF_COLORS

POINT_AFTER = set("=(,;+-*/^&<>:")


class RefHighlighter(QSyntaxHighlighter):
    """Colors cell references in a formula the same way the grid outlines them."""

    def highlightBlock(self, text):
        doc_text = self.document().toPlainText()
        if not doc_text.startswith("="):
            return
        block_start = self.currentBlock().position()
        try:
            spans = ref_spans(doc_text)
        except Exception:
            return
        colors = {}
        for start, end, info in spans:
            name = info.text().replace("$", "").upper()
            idx = colors.setdefault(name, len(colors))
            fmt = QTextCharFormat()
            fmt.setForeground(QColor(REF_COLORS[idx % len(REF_COLORS)]))
            s = start - block_start
            e = end - block_start
            if e <= 0 or s >= len(text):
                continue
            self.setFormat(max(0, s), min(len(text), e) - max(0, s), fmt)


class FunctionPopup(QListWidget):
    """Autocomplete list of function names."""
    chosen = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.ToolTip | Qt.FramelessWindowHint)
        self.setFocusPolicy(Qt.NoFocus)
        self.setStyleSheet("QListWidget { border: 1px solid #A0A0A0; background: white; font: 9pt 'Segoe UI'; }"
                           "QListWidget::item { padding: 2px 6px; }"
                           "QListWidget::item:selected { background: #CFE5D7; color: black; }")
        self.itemClicked.connect(lambda it: self.chosen.emit(it.text()))

    def show_for(self, names, global_pos):
        self.clear()
        for n in names[:12]:
            QListWidgetItem(n, self)
        self.setCurrentRow(0)
        h = min(12, len(names)) * 21 + 4
        self.setFixedSize(220, h)
        self.move(global_pos)
        self.show()

    def move_sel(self, d):
        r = max(0, min(self.count() - 1, self.currentRow() + d))
        self.setCurrentRow(r)

    def current(self):
        it = self.currentItem()
        return it.text() if it else None


class HintLabel(QLabel):
    def __init__(self):
        super().__init__(None)
        self.setWindowFlags(Qt.ToolTip | Qt.FramelessWindowHint)
        self.setStyleSheet("QLabel { background: #FFFFE8; border: 1px solid #B0B0B0; padding: 2px 6px;"
                           " font: 9pt 'Segoe UI'; color: #333; }")


_TOKEN_BEFORE = re.compile(r"([A-Za-z_][A-Za-z0-9_.]*)$")


class CellEditor(QPlainTextEdit):
    """Plain-text editor with Excel key handling.

    Emits `commit(dr, dc, fill_selection)` / `cancel()` and lets the owner
    insert references while pointing (`point_key` for arrow keys)."""
    commit = Signal(int, int, bool)
    cancel = Signal()
    point_key = Signal(int, int, bool)  # dr, dc, extend
    edited = Signal(str)

    def __init__(self, parent=None, in_bar=False):
        super().__init__(parent)
        self.in_bar = in_bar
        self.mode = "edit"          # 'enter' (arrows commit) or 'edit' (arrows move caret)
        self.point_span = None      # (start, end) of a reference inserted by pointing
        self.setFrameShape(QFrame.NoFrame)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setTabChangesFocus(False)
        self.setWordWrapMode(QTextOption.NoWrap if not in_bar else QTextOption.WrapAnywhere)
        self.document().setDocumentMargin(1 if not in_bar else 3)
        self.highlighter = RefHighlighter(self.document())
        self.popup = FunctionPopup()
        self.popup.chosen.connect(self._accept_completion)
        self.hint = HintLabel()
        self._suppress = False
        self.textChanged.connect(self._on_text_changed)
        if not in_bar:
            self.setStyleSheet("QPlainTextEdit { background: white; border: 2px solid #217346; }")

    # ------------------------------------------------------------ text helpers
    def text(self):
        return self.toPlainText()

    def set_text(self, t, cursor_end=True):
        self._suppress = True
        self.setPlainText(t)
        self._suppress = False
        if cursor_end:
            c = self.textCursor()
            c.movePosition(QTextCursor.End)
            self.setTextCursor(c)

    def cursor_pos(self):
        return self.textCursor().position()

    def can_point(self):
        t = self.text()
        if not t.startswith("="):
            return False
        pos = self.cursor_pos()
        if self.point_span and self.point_span[0] <= pos <= self.point_span[1]:
            return True
        before = t[:pos].rstrip()
        return bool(before) and before[-1] in POINT_AFTER

    def insert_ref(self, ref):
        """Insert or replace the pointed reference."""
        t = self.text()
        if self.point_span:
            s, e = self.point_span
        else:
            s = e = self.cursor_pos()
        new = t[:s] + ref + t[e:]
        self.set_text(new, cursor_end=False)
        c = self.textCursor()
        c.setPosition(s + len(ref))
        self.setTextCursor(c)
        self.point_span = (s, s + len(ref))
        self.edited.emit(new)

    def _on_text_changed(self):
        if self._suppress:
            return
        self.edited.emit(self.text())
        self._update_completion()

    # ------------------------------------------------------------ completion & hints
    def _current_func_token(self):
        t = self.text()
        if not t.startswith("="):
            return None
        pos = self.cursor_pos()
        before = t[:pos]
        if before.count('"') % 2:
            return None
        m = _TOKEN_BEFORE.search(before)
        if not m:
            return None
        start = m.start(1)
        if start > 0 and before[start - 1] not in POINT_AFTER and before[start - 1] != " ":
            return None
        return m.group(1), start

    def _enclosing_function(self):
        t = self.text()[:self.cursor_pos()]
        depth = 0
        in_str = False
        for i in range(len(t) - 1, -1, -1):
            ch = t[i]
            if ch == '"':
                in_str = not in_str
            if in_str:
                continue
            if ch == ")":
                depth += 1
            elif ch == "(":
                if depth == 0:
                    m = _TOKEN_BEFORE.search(t[:i])
                    return m.group(1).upper() if m else None
                depth -= 1
        return None

    def _update_completion(self):
        if not self.hasFocus():
            self.popup.hide()
            self.hint.hide()
            return
        tok = self._current_func_token()
        if tok and len(tok[0]) >= 1 and not re.fullmatch(r"[A-Za-z]{1,3}\d+", tok[0]):
            prefix = tok[0].upper()
            names = [n for n in ALL_NAMES if n.startswith(prefix)]
            if names and not (len(names) == 1 and names[0] == prefix):
                rect = self.cursorRect()
                self.popup.show_for(names, self.viewport().mapToGlobal(QPoint(rect.left(), rect.bottom() + 4)))
                self.hint.hide()
                return
        self.popup.hide()
        fn = self._enclosing_function()
        if fn and fn in SIGNATURES:
            self.hint.setText(SIGNATURES[fn])
            self.hint.adjustSize()
            rect = self.cursorRect()
            self.hint.move(self.viewport().mapToGlobal(QPoint(rect.left(), rect.bottom() + 4)))
            self.hint.show()
        else:
            self.hint.hide()

    def _accept_completion(self, name):
        tok = self._current_func_token()
        if not tok:
            return
        word, start = tok
        t = self.text()
        pos = self.cursor_pos()
        new = t[:start] + name + "(" + t[pos:]
        self.set_text(new, cursor_end=False)
        c = self.textCursor()
        c.setPosition(start + len(name) + 1)
        self.setTextCursor(c)
        self.popup.hide()
        self.edited.emit(new)
        self._update_completion()

    def hide_popups(self):
        try:
            self.popup.hide()
            self.hint.hide()
        except RuntimeError:  # already destroyed during shutdown
            pass

    def focusOutEvent(self, e):
        self.hide_popups()
        super().focusOutEvent(e)

    def hideEvent(self, e):
        self.hide_popups()
        super().hideEvent(e)

    # ------------------------------------------------------------ keys
    def keyPressEvent(self, e):
        k = e.key()
        mods = e.modifiers()
        ctrl = bool(mods & Qt.ControlModifier)
        shift = bool(mods & Qt.ShiftModifier)
        alt = bool(mods & Qt.AltModifier)

        if self.popup.isVisible():
            if k in (Qt.Key_Down, Qt.Key_Up):
                self.popup.move_sel(1 if k == Qt.Key_Down else -1)
                return
            if k == Qt.Key_Tab:
                self._accept_completion(self.popup.current())
                return
            if k == Qt.Key_Escape:
                self.popup.hide()
                return

        if k in (Qt.Key_Return, Qt.Key_Enter):
            if alt:
                self.insertPlainText("\n")
                return
            self.hide_popups()
            self.commit.emit(-1 if shift else 1, 0, ctrl)
            return
        if k == Qt.Key_Tab or k == Qt.Key_Backtab:
            self.hide_popups()
            self.commit.emit(0, -1 if (shift or k == Qt.Key_Backtab) else 1, False)
            return
        if k == Qt.Key_Escape:
            self.hide_popups()
            self.cancel.emit()
            return
        if k == Qt.Key_F2:
            self.mode = "edit" if self.mode == "enter" else "enter"
            self.point_span = None
            return
        if k == Qt.Key_F4:
            t, pos = toggle_absolute(self.text(), self.cursor_pos())
            self.set_text(t, cursor_end=False)
            c = self.textCursor()
            c.setPosition(pos)
            self.setTextCursor(c)
            self.edited.emit(t)
            return
        arrows = {Qt.Key_Left: (0, -1), Qt.Key_Right: (0, 1), Qt.Key_Up: (-1, 0), Qt.Key_Down: (1, 0)}
        if k in arrows and not ctrl and not alt and not self.in_bar:
            if self.mode == "enter" or self.point_span:
                if self.can_point():
                    dr, dc = arrows[k]
                    self.point_key.emit(dr, dc, shift)
                    return
                if self.mode == "enter":
                    dr, dc = arrows[k]
                    self.hide_popups()
                    self.commit.emit(dr, dc, False)
                    return
        if e.text() and not ctrl:
            self.point_span = None
        if k in (Qt.Key_Left, Qt.Key_Right, Qt.Key_Home, Qt.Key_End):
            self.point_span = None
        super().keyPressEvent(e)

    def event(self, e):
        # make Tab reach keyPressEvent instead of moving focus
        if e.type() == QEvent.KeyPress and e.key() in (Qt.Key_Tab, Qt.Key_Backtab):
            self.keyPressEvent(e)
            return True
        return super().event(e)

    def mousePressEvent(self, e):
        self.point_span = None
        if self.mode == "enter":
            self.mode = "edit"
        super().mousePressEvent(e)
