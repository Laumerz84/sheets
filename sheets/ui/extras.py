"""Smaller features kept out of mainwindow.py: format painter, remove duplicates."""
from PySide6.QtWidgets import (QCheckBox, QDialog, QDialogButtonBox, QLabel,
                               QListWidget, QListWidgetItem, QVBoxLayout)
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QHeaderView, QTableWidget, QTableWidgetItem

from .. import ops
from ..refs import col_name, key
from ..workbook import DEFAULT_STYLE


def copy_styles(sheet, rect):
    r1, c1, r2, c2 = rect
    return [[sheet.style(r, c) for c in range(c1, c2 + 1)] for r in range(r1, r2 + 1)]


def paint_states(sheet, styles, dest):
    """Tile copied styles over dest (contents untouched)."""
    h, w = len(styles), len(styles[0])
    r1, c1, r2, c2 = dest
    if r1 == r2 and c1 == c2:
        r2, c2 = r1 + h - 1, c1 + w - 1
    out = {}
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            k = key(r, c)
            content, _ = sheet.get_state(k)
            out[k] = (content, styles[(r - r1) % h][(c - c1) % w])
    return out


class RemoveDuplicatesDialog(QDialog):
    def __init__(self, parent, sheet, rect, header):
        super().__init__(parent)
        self.setWindowTitle("Remove Duplicates")
        self.sheet = sheet
        self.rect = rect
        lay = QVBoxLayout(self)
        lay.addWidget(QLabel("Rows are duplicates when all of the checked columns match:"))
        self.header_chk = QCheckBox("My data has headers")
        self.header_chk.setChecked(header)
        lay.addWidget(self.header_chk)
        self.list = QListWidget()
        lay.addWidget(self.list)
        self.header_chk.toggled.connect(self._fill)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)
        self._fill()

    def _fill(self):
        r1, c1, r2, c2 = self.rect
        checked = {self.list.item(i).data(Qt.UserRole) for i in range(self.list.count())
                   if self.list.item(i).checkState() == Qt.Checked} if self.list.count() else None
        self.list.clear()
        for c in range(c1, c2 + 1):
            label = f"Column {col_name(c)}"
            if self.header_chk.isChecked():
                v = self.sheet.value(r1, c)
                if v not in (None, ""):
                    label = str(v)
            it = QListWidgetItem(label)
            it.setData(Qt.UserRole, c)
            it.setFlags(it.flags() | Qt.ItemIsUserCheckable)
            it.setCheckState(Qt.Checked if checked is None or c in checked else Qt.Unchecked)
            self.list.addItem(it)

    def columns(self):
        return [self.list.item(i).data(Qt.UserRole) for i in range(self.list.count())
                if self.list.item(i).checkState() == Qt.Checked]


def dedupe_states(sheet, rect, cols, header):
    """States that keep the first occurrence of each row (compared by displayed
    text, case-insensitive like Excel) and pack the rest up. -> (states, removed, kept)"""
    r1, c1, r2, c2 = rect
    if header:
        r1 += 1
    seen = set()
    keep = []
    for r in range(r1, r2 + 1):
        sig = tuple(ops.display_text(sheet, r, c).lower() for c in cols)
        if sig in seen:
            continue
        seen.add(sig)
        keep.append(r)
    removed = (r2 - r1 + 1) - len(keep)
    if not removed:
        return {}, 0, len(keep)
    snap = {r: [sheet.get_state(key(r, c)) for c in range(c1, c2 + 1)] for r in keep}
    states = {}
    from ..formula import shift_formula
    for i, src in enumerate(keep):
        dst = r1 + i
        for j, (content, style) in enumerate(snap[src]):
            if content is not None and content[0] == "f" and dst != src:
                content = ("f", shift_formula(content[1], dst - src, 0))
            states[key(dst, c1 + j)] = (content, style)
    for dst in range(r1 + len(keep), r2 + 1):
        for c in range(c1, c2 + 1):
            states[key(dst, c)] = (None, DEFAULT_STYLE)
    return states, removed, len(keep)


class WishlistDialog(QDialog):
    """Excel features Sheets doesn't have yet, most-wanted (most tried via KeyTips) first."""

    def __init__(self, parent):
        super().__init__(parent)
        from .keytips import KEYTIPS, load_tried
        self.setWindowTitle("Feature Wishlist")
        self.resize(900, 560)
        lay = QVBoxLayout(self)
        lay.addWidget(QLabel("Excel features Exkel doesn't have yet. Pressing one of their KeyTips counts as a "
                             "vote; Claude asks about the most-wanted ones when you work on Exkel."))
        tried = load_tried()
        rows = [(seq, label, excel, tried.get(seq, {}).get("count", 0))
                for seq, label, action, excel in KEYTIPS if action is None]
        rows.sort(key=lambda r: (-r[3], r[1]))
        table = QTableWidget(len(rows), 4)
        table.setHorizontalHeaderLabels(["Tried", "KeyTip", "Feature", "What it does in Excel"])
        table.verticalHeader().setVisible(False)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setWordWrap(True)
        for i, (seq, label, excel, n) in enumerate(rows):
            for j, text in enumerate((str(n) if n else "", "Alt " + " ".join(seq), label, excel)):
                table.setItem(i, j, QTableWidgetItem(text))
        h = table.horizontalHeader()
        for j in range(3):
            h.setSectionResizeMode(j, QHeaderView.ResizeToContents)
        h.setSectionResizeMode(3, QHeaderView.Stretch)
        table.resizeRowsToContents()
        lay.addWidget(table)
        bb = QDialogButtonBox(QDialogButtonBox.Close)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)
