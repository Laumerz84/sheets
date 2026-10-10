"""Charts on the grid: the layer that keeps a widget over each chart of the current sheet, mouse
handling (select, move, resize with handles), the "+" and style buttons, the right-click menu and
the commands that change charts (every change is one undo step through MetaCommand).

Mirrors ui/controls.py: charts live in `Sheet.charts`; `ChartLayer.sync()` runs after every grid
paint, so a chart follows cell edits, scrolling, zoom, undo/redo and structure changes."""
import copy

from PySide6.QtCore import QEvent, QObject, QPoint, QPointF, QRect, QRectF, QSize, Qt, QTimer, Signal
from PySide6.QtGui import QColor, QIcon, QKeySequence, QPainter, QPen, QPixmap
from PySide6.QtWidgets import QApplication, QMenu, QMessageBox, QToolButton, QToolTip, QWidget

from .. import charts as C
from ..refs import MAX_COLS, MAX_ROWS
from . import chart_paint as CP
from . import style as S
from .commands import MetaCommand, _Base

HANDLE = 5          # px the widget extends past the chart when selected (room for the handles)
CLIP_MARK = "[Ekxel chart]"


def chart_icon(size=18, color="#217346"):
    """A small bar-chart icon for menus and the toolbar."""
    pm = QPixmap(size * 2, size * 2)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    s = size * 2
    p.setPen(QPen(QColor("#555555"), 2))
    p.drawLine(int(s * 0.12), int(s * 0.1), int(s * 0.12), int(s * 0.88))
    p.drawLine(int(s * 0.12), int(s * 0.88), int(s * 0.92), int(s * 0.88))
    p.setPen(Qt.NoPen)
    for i, (h, col) in enumerate(((0.38, "#4472C4"), (0.62, "#ED7D31"), (0.5, "#A5A5A5"))):
        x = s * (0.22 + i * 0.25)
        p.setBrush(QColor(col))
        p.drawRect(QRectF(x, s * 0.88 - s * h, s * 0.18, s * h))
    p.end()
    pm.setDevicePixelRatio(2)
    return QIcon(pm)


def type_icon(ctype, size=28):
    """Tiny schematic of a chart type for menus."""
    from .chart_dialogs import sample_chart_image
    return QIcon(QPixmap.fromImage(sample_chart_image(ctype, size * 2, int(size * 1.5))))


# ================================================================ the chart widget
class ChartWidget(QWidget):
    """One chart floating over the grid."""

    def __init__(self, layer, ch):
        super().__init__(layer.grid)
        self.layer = layer
        self.cid = ch["id"]
        self.ch = ch
        self.data = None
        self.virtual = QRect()        # where the chart sits in grid coordinates (may be partly off screen)
        self.regions = []
        self._pix = None
        self._key = None
        self.gesture = None
        self.setMouseTracking(True)
        self.setFocusPolicy(Qt.StrongFocus)
        self.setContextMenuPolicy(Qt.NoContextMenu)
        self.setAttribute(Qt.WA_OpaquePaintEvent, False)

    # ---- state
    @property
    def selected(self):
        return self.layer.selected == self.cid

    def margin(self):
        return HANDLE if self.selected else 0

    def local_rect(self):
        """The chart's rect in this widget's coordinates."""
        g = self.geometry()
        return QRectF(self.virtual.translated(-g.topLeft()))

    def _scale(self):
        return self.layer.grid.zoom

    # ---- painting
    def _render(self):
        v = self.virtual
        w, h = max(v.width(), 1), max(v.height(), 1)
        dpr = self.devicePixelRatioF()
        pm = QPixmap(int(w * dpr), int(h * dpr))
        pm.setDevicePixelRatio(dpr)
        pm.fill(Qt.transparent)
        p = QPainter(pm)
        sel = self.layer.elem if self.selected else None
        self.regions = CP.paint_chart(p, QRectF(0, 0, w, h), self.ch, self.data, self._scale(), sel)
        p.end()
        self._pix = pm

    def paintEvent(self, ev):
        if self.data is None or self.virtual.isEmpty():
            return
        sel = self.layer.elem if self.selected else None
        key = (id(self.ch), id(self.data), self.virtual.size().toTuple(), self._scale(), repr(sel), self.devicePixelRatioF())
        if self._pix is None or key != self._key:
            self._render()
            self._key = key
        p = QPainter(self)
        off = self.virtual.topLeft() - self.geometry().topLeft()
        p.drawPixmap(off, self._pix)
        if self.selected:
            self._paint_handles(p)
        p.end()

    def handle_points(self):
        r = self.local_rect()
        cx, cy = r.center().x(), r.center().y()
        return {"tl": r.topLeft(), "t": QPointF(cx, r.top()), "tr": r.topRight(), "r": QPointF(r.right(), cy),
                "br": r.bottomRight(), "b": QPointF(cx, r.bottom()), "bl": r.bottomLeft(), "l": QPointF(r.left(), cy)}

    def _paint_handles(self, p):
        p.setRenderHint(QPainter.Antialiasing, True)
        p.setPen(QPen(QColor("#8C8C8C"), 1))
        p.setBrush(QColor("#FFFFFF"))
        for name, pt in self.handle_points().items():
            p.drawEllipse(pt, 3.5, 3.5)

    def handle_at(self, pos):
        if not self.selected:
            return None
        for name, pt in self.handle_points().items():
            if abs(pos.x() - pt.x()) <= 6 and abs(pos.y() - pt.y()) <= 6:
                return name
        return None

    # ---- mouse
    _CURSORS = {"tl": Qt.SizeFDiagCursor, "br": Qt.SizeFDiagCursor, "tr": Qt.SizeBDiagCursor,
                "bl": Qt.SizeBDiagCursor, "t": Qt.SizeVerCursor, "b": Qt.SizeVerCursor,
                "l": Qt.SizeHorCursor, "r": Qt.SizeHorCursor}

    def _hit(self, pos):
        off = self.virtual.topLeft() - self.geometry().topLeft()
        return CP.hit_test(self.regions, QPointF(pos) - QPointF(off))

    def mousePressEvent(self, e):
        if e.button() not in (Qt.LeftButton, Qt.RightButton):
            return
        lay = self.layer
        was_selected = self.selected
        hit = self._hit(e.position().toPoint())
        if e.button() == Qt.RightButton:
            lay.select(self.cid, (hit.kind, hit.key) if hit and hit.kind not in ("chart",) else None)
            lay.context_menu(self, e.globalPosition().toPoint(), hit)
            return
        if lay.win.grid.read_only or lay.win.ai_tools.user_locked():
            lay.win._claude_busy_message()
            return
        h = self.handle_at(e.position().toPoint()) if was_selected else None
        lay.select(self.cid, None if not was_selected else lay.elem)
        self.setFocus()
        self.gesture = {"mode": "resize" if h else "pending", "handle": h, "start": e.globalPosition().toPoint(),
                        "rect": QRect(self.virtual), "hit": hit, "was": was_selected}
        QToolTip.hideText()

    def mouseMoveEvent(self, e):
        g = self.gesture
        pos = e.position().toPoint()
        if g is None:
            h = self.handle_at(pos)
            if h:
                self.setCursor(self._CURSORS[h])
                QToolTip.hideText()
            else:
                self.setCursor(Qt.ArrowCursor)
                hit = self._hit(pos)
                if hit is not None and hit.tip:
                    QToolTip.showText(e.globalPosition().toPoint(), hit.tip, self)
                else:
                    QToolTip.hideText()
            return
        d = e.globalPosition().toPoint() - g["start"]
        if g["mode"] == "pending":
            if abs(d.x()) + abs(d.y()) < 4:
                return
            g["mode"] = "move"
            self.setCursor(Qt.SizeAllCursor)
        r = QRect(g["rect"])
        if g["mode"] == "move":
            r.translate(d)
        else:
            h = g["handle"]
            minw, minh = int(C.MIN_SIZE[0] * self._scale()), int(C.MIN_SIZE[1] * self._scale())
            if "l" in h:
                r.setLeft(min(r.left() + d.x(), r.right() - minw))
            if "r" in h:
                r.setRight(max(r.right() + d.x(), r.left() + minw))
            if "t" in h:
                r.setTop(min(r.top() + d.y(), r.bottom() - minh))
            if "b" in h:
                r.setBottom(max(r.bottom() + d.y(), r.top() + minh))
        self.virtual = r
        self.layer.apply_geometry(self)
        self.update()

    def mouseReleaseEvent(self, e):
        g, self.gesture = self.gesture, None
        if g is None:
            return
        self.unsetCursor()
        if g["mode"] == "pending":
            hit = g["hit"]
            if g["was"]:
                self.layer.select(self.cid, (hit.kind, hit.key) if hit and hit.kind != "chart" else None)
            return
        if self.virtual != g["rect"]:
            self.layer.commit_rect(self.cid, self.virtual, "Move Chart" if g["mode"] == "move" else "Resize Chart")

    def mouseDoubleClickEvent(self, e):
        hit = self._hit(e.position().toPoint())
        self.layer.select(self.cid, (hit.kind, hit.key) if hit and hit.kind != "chart" else None)
        self.layer.show_pane()

    def wheelEvent(self, e):
        self.layer.grid.wheelEvent(e)

    # ---- keys
    def event(self, e):
        if e.type() == QEvent.ShortcutOverride:
            k, m = e.key(), e.modifiers()
            ctrl = bool(m & Qt.ControlModifier)
            if (ctrl and k in (Qt.Key_C, Qt.Key_X, Qt.Key_V, Qt.Key_D)) or k in (Qt.Key_Delete, Qt.Key_Escape):
                e.accept()
                return True
        return super().event(e)

    def keyPressEvent(self, e):
        k, m = e.key(), e.modifiers()
        ctrl = bool(m & Qt.ControlModifier)
        lay = self.layer
        if k == Qt.Key_Delete:
            lay.delete(self.cid)
        elif k == Qt.Key_Escape:
            lay.deselect(focus_grid=True)
        elif ctrl and k == Qt.Key_C:
            lay.copy(self.cid)
        elif ctrl and k == Qt.Key_X:
            lay.copy(self.cid, cut=True)
        elif ctrl and k == Qt.Key_V:
            lay.paste()
        elif ctrl and k == Qt.Key_D:
            lay.duplicate(self.cid)
        elif k in (Qt.Key_Left, Qt.Key_Right, Qt.Key_Up, Qt.Key_Down):
            step = 1 if ctrl else 8
            dx = {Qt.Key_Left: -step, Qt.Key_Right: step}.get(k, 0)
            dy = {Qt.Key_Up: -step, Qt.Key_Down: step}.get(k, 0)
            lay.commit_rect(self.cid, self.virtual.translated(int(dx * self._scale()), int(dy * self._scale())),
                            "Move Chart")
        elif k == Qt.Key_F2 or k == Qt.Key_Return:
            lay.show_pane()
        else:
            super().keyPressEvent(e)


class _Buttons:
    """The '+' (Chart Elements) and brush (Chart Styles) buttons next to the selected chart."""

    def __init__(self, layer):
        self.layer = layer
        grid = layer.grid
        self.plus = self._button(grid, "+", "Chart Elements", layer.elements_menu_for_button)
        self.brush = self._button(grid, "", "Chart Styles", layer.styles_menu_for_button)
        self.brush.setIcon(self._brush_icon())
        self.hide()

    @staticmethod
    def _brush_icon():
        pm = QPixmap(36, 36)
        pm.fill(Qt.transparent)
        p = QPainter(pm)
        p.setRenderHint(QPainter.Antialiasing)
        p.setPen(Qt.NoPen)
        p.setBrush(QColor("#217346"))
        p.drawRoundedRect(QRectF(5, 5, 22, 10), 2, 2)
        p.setBrush(QColor("#7F7F7F"))
        p.drawRect(QRectF(15, 15, 4, 9))
        p.setBrush(QColor("#217346"))
        p.drawRoundedRect(QRectF(12, 24, 10, 8), 2, 2)
        p.end()
        pm.setDevicePixelRatio(2)
        return QIcon(pm)

    def _button(self, parent, text, tip, handler):
        b = QToolButton(parent)
        b.setText(text)
        b.setToolTip(tip)
        b.setFixedSize(26, 26)
        b.setIconSize(QSize(18, 18))
        b.setStyleSheet("QToolButton { background: #FFFFFF; border: 1px solid #BFBFBF; border-radius: 2px; "
                        "color: #217346; font: bold 16px 'Segoe UI'; } QToolButton:hover { background: #E1EFE6; }")
        b.setFocusPolicy(Qt.NoFocus)
        b.clicked.connect(lambda _=False, bb=b, h=handler: h(bb))
        return b

    def hide(self):
        self.plus.hide()
        self.brush.hide()

    def place(self, rect, grid_w):
        """Next to the chart's top-right corner (inside it when there's no room)."""
        x = rect.right() + 8
        if x + 28 > grid_w:
            x = rect.right() - 30
        y = rect.top() + (0 if x > rect.right() else 6)
        self.plus.move(x, y)
        self.brush.move(x, y + 30)
        for b in (self.plus, self.brush):
            b.setVisible(True)
            b.raise_()


# ================================================================ the layer
class ChartLayer(QObject):
    changed = Signal()      # the selected chart, its element or its content changed (for the Format pane)

    def __init__(self, win):
        super().__init__(win.grid)
        self.win = win
        self.grid = win.grid
        self.widgets = {}        # chart id -> ChartWidget
        self.selected = None
        self.elem = None         # (kind, key) of the selected element, None = the whole chart
        self._sheet = None
        self._z = None
        self.pane = None
        self.buttons = _Buttons(self)
        self.clip = None         # copied chart dict
        self.grid.installEventFilter(self)

    # ------------------------------------------------------------ helpers
    @property
    def sheet(self):
        return self.grid.sheet

    def charts(self, sheet=None):
        return list((sheet or self.sheet).charts) if (sheet or self.sheet) is not None else []

    def find(self, cid, sheet=None):
        for c in self.charts(sheet):
            if c["id"] == cid:
                return c
        return None

    def selected_chart(self):
        return self.find(self.selected) if self.selected else None

    def eventFilter(self, obj, ev):
        if obj is self.grid and ev.type() == QEvent.MouseButtonPress and self.selected:
            self.deselect()
        return False

    # ------------------------------------------------------------ building / placing (after every grid paint)
    def sync(self):
        sh = self.grid.sheet
        if sh is None:
            return
        if self._sheet is not sh:
            for w in self.widgets.values():
                w.hide()
                w.deleteLater()
            self.widgets = {}
            self._sheet = sh
            if self.selected:
                self.selected, self.elem = None, None
                self.changed.emit()
        charts = getattr(sh, "charts", [])
        ids = [c["id"] for c in charts]
        for cid in list(self.widgets):
            if cid not in ids:
                w = self.widgets.pop(cid)
                w.hide()
                w.deleteLater()
                if self.selected == cid:
                    self.selected, self.elem = None, None
                    self.changed.emit()
        for c in charts:
            w = self.widgets.get(c["id"])
            if w is None:
                w = self.widgets[c["id"]] = ChartWidget(self, c)
                w.show()
            if w.ch is not c:
                w.ch = c
                w._pix = None
                if c["id"] == self.selected:
                    self.changed.emit()
        if self._z != ids:
            for cid in ids:
                self.widgets[cid].raise_()
            self._z = ids
        self.place()

    def virtual_rect(self, ch):
        g = self.grid
        z = g.zoom
        r1, c1, dx1, dy1 = ch["from"]
        r2, c2, dx2, dy2 = ch["to"]
        x1 = g.col_x(c1) + min(int(dx1 * z), g.cols.size(c1))
        y1 = g.row_y(r1) + min(int(dy1 * z), g.rows.size(r1))
        x2 = g.col_x(c2) + min(int(dx2 * z), g.cols.size(c2))
        y2 = g.row_y(r2) + min(int(dy2 * z), g.rows.size(r2))
        w = max(int(C.MIN_SIZE[0] * z), x2 - x1)
        h = max(int(C.MIN_SIZE[1] * z), y2 - y1)
        return QRect(x1, y1, w, h)

    def viewport(self):
        g = self.grid
        return QRect(g.rw, g.hh, max(0, g.width() - g.rw), max(0, g.height() - g.hh))

    def apply_geometry(self, w):
        """Put widget w where its virtual rect says, clipped to the cell area."""
        vp = self.viewport()
        m = w.margin()
        vis = w.virtual.adjusted(-m, -m, m, m).intersected(vp)
        if vis.isEmpty():
            w.hide()
            return False
        if w.geometry() != vis:
            w.setGeometry(vis)
        if not w.isVisible():
            w.show()
        return True

    def place(self):
        g = self.grid
        sh = g.sheet
        for ch in sh.charts:
            w = self.widgets.get(ch["id"])
            if w is None:
                continue
            if w.gesture is None:
                v = self.virtual_rect(ch)
                if v != w.virtual:
                    w.virtual = v
                    w._pix = None
            if not self.apply_geometry(w):
                continue
            if w.data is None or w.gesture is None:
                data = self._data(ch, sh)
                if w.data is None or not _same_data(w.data, data):
                    w.data = data
                    w._pix = None
            w.update()
        self._place_buttons()

    def _data(self, ch, sh):
        try:
            return C.resolve(self.win.wb, ch, sh)
        except Exception as e:      # a chart must never take the grid down
            return {"series": [], "cats": [], "sampled": 1, "xfmt": "General", "error": f"Can't draw this chart ({e})"}

    def _place_buttons(self):
        w = self.widgets.get(self.selected) if self.selected else None
        if w is None or not w.isVisible() or self.grid.read_only:
            self.buttons.hide()
            return
        self.buttons.place(w.virtual, self.grid.width())

    # ------------------------------------------------------------ selection
    def select(self, cid, elem=None):
        changed = cid != self.selected or elem != self.elem
        prev = self.widgets.get(self.selected) if self.selected else None
        self.selected, self.elem = cid, elem
        w = self.widgets.get(cid)
        if w is not None:
            w.setFocus()
        for ww in (prev, w):
            if ww is not None:
                ww._pix = None
                self.apply_geometry(ww)
                ww.update()
        ch = self.find(cid)
        if ch is not None:
            self.win.name_box.setText(ch.get("name") or "Chart")
        self._place_buttons()
        if changed:
            self.changed.emit()

    def deselect(self, focus_grid=False):
        if self.selected is None:
            return
        self.select(None, None)
        self.buttons.hide()
        if focus_grid:
            self.grid.setFocus()
            self.win._selection_changed()

    # ------------------------------------------------------------ changing charts (each = one undo step)
    def set_charts(self, new, text, sheet=None):
        sh = sheet or self.sheet
        self.win.undo.push(MetaCommand(self.win, sh, {"charts": (list(sh.charts), list(new))}, text, relayout=False))
        self.grid.update()

    def add(self, ch, text="Insert Chart", sheet=None):
        sh = sheet or self.sheet
        self.set_charts(sh.charts + [ch], text, sh)
        if sh is self.grid.sheet:
            QTimer.singleShot(0, lambda: self._select_when_ready(ch["id"]))

    def _select_when_ready(self, cid):
        self.sync()
        if cid in self.widgets:
            self.select(cid)

    def edit(self, cid, fn, text="Format Chart"):
        """Replace chart `cid` by fn(chart) (a new dict) as one undo step."""
        sh = self.sheet
        new = []
        done = False
        for c in sh.charts:
            if c["id"] == cid:
                n = fn(c)
                if n is None or n == c:
                    return False
                new.append(n)
                done = True
            else:
                new.append(c)
        if done:
            self.set_charts(new, text, sh)
        return done

    def update(self, cid, text="Format Chart", **changes):
        return self.edit(cid, lambda c: dict(c, **changes), text)

    def delete(self, cid):
        if self.selected == cid:
            self.selected, self.elem = None, None
            self.buttons.hide()
            self.changed.emit()
        self.set_charts([c for c in self.sheet.charts if c["id"] != cid], "Delete Chart")
        self.grid.setFocus()

    def commit_rect(self, cid, vrect, text):
        """Store the virtual rect (grid pixels) the user dragged the chart to."""
        g = self.grid
        z = g.zoom

        def anchor_at(x, y):
            x = max(x, g.rw)
            y = max(y, g.hh)
            c, r = g.col_at(x), g.row_at(y)
            r, c = min(r, MAX_ROWS - 1), min(c, MAX_COLS - 1)
            dx = int(max(0, x - g.col_x(c)) / z)
            dy = int(max(0, y - g.row_y(r)) / z)
            return [r, c, dx, dy]
        frm = anchor_at(vrect.left(), vrect.top())
        to = anchor_at(vrect.right() + 1, vrect.bottom() + 1)
        self.edit(cid, lambda c: dict(c, **{"from": frm, "to": to}), text)
        w = self.widgets.get(cid)
        if w is not None:
            w.gesture = None
        self.grid.update()

    def reorder(self, cid, front):
        charts = list(self.sheet.charts)
        item = self.find(cid)
        if item is None:
            return
        charts.remove(item)
        charts = charts + [item] if front else [item] + charts
        self.set_charts(charts, "Bring Chart to Front" if front else "Send Chart to Back")

    # ------------------------------------------------------------ clipboard
    def copy(self, cid, cut=False):
        ch = self.find(cid)
        if ch is None:
            return
        self.clip = copy.deepcopy(ch)
        self.clip["_home"] = self.sheet.name
        w = self.widgets.get(cid)
        self.clip_img = None
        if w is not None and w.data is not None:
            # the clipboard gets a picture of the chart (pastes into Word etc.); Ekxel recognises it by its pixels
            self.clip_img = CP.render_image(ch, w.data, w.virtual.width(), w.virtual.height(), self.grid.zoom)
            QApplication.clipboard().setImage(self.clip_img)
        else:
            QApplication.clipboard().setText(CLIP_MARK + " " + (ch.get("name") or ""))
        self.clip_text = QApplication.clipboard().text()
        self.win.statusBar().showMessage("Chart copied: press Ctrl+V to paste it", 3000)
        if cut:
            self.delete(cid)

    def try_copy(self, cut=False):
        """Called by the window's Copy / Cut: copies the selected chart when one has the focus."""
        w = self.widgets.get(self.selected) if self.selected else None
        if w is not None and w.hasFocus():
            self.copy(self.selected, cut)
            return True
        return False

    def try_paste(self):
        """True (and a chart pasted) when the clipboard still holds the chart Ekxel copied last."""
        if self.clip is None:
            return False
        cb = QApplication.clipboard()
        img = getattr(self, "clip_img", None)
        if img is not None:
            md = cb.mimeData()
            if md is None or not md.hasImage() or md.hasText():
                return False
            from PySide6.QtGui import QImage
            same = cb.image().convertToFormat(QImage.Format_ARGB32) == img.convertToFormat(QImage.Format_ARGB32)
        else:
            same = cb.text() == getattr(self, "clip_text", None)
        if same:
            self.paste()
        return same

    def try_delete(self):
        w = self.widgets.get(self.selected) if self.selected else None
        if w is not None and w.hasFocus():
            self.delete(self.selected)
            return True
        return False

    def paste(self):
        if self.clip is None:
            return
        sh = self.sheet
        ch = {k: v for k, v in copy.deepcopy(self.clip).items() if k != "_home"}
        ch["id"] = C.new_id()
        ch["name"] = C.chart_name(sh.charts)
        x, y, w, h = C.chart_px_rect(sh, ch)
        if self.clip.get("_home") == sh.name:
            x, y = x + 24, y + 24
        ch["from"], ch["to"] = C.anchors_for(sh, x, y, w, h)
        self.add(ch, "Paste Chart")

    def duplicate(self, cid):
        ch = self.find(cid)
        if ch is None:
            return
        n = copy.deepcopy(ch)
        n["id"] = C.new_id()
        n["name"] = C.chart_name(self.sheet.charts)
        x, y, w, h = C.chart_px_rect(self.sheet, n)
        n["from"], n["to"] = C.anchors_for(self.sheet, x + 24, y + 24, w, h)
        self.add(n, "Duplicate Chart")

    # ------------------------------------------------------------ creating charts
    def selection_rect(self):
        """The data for a new chart: the selected block, or the table around the active cell."""
        from .. import ops
        g = self.grid
        sh = g.sheet
        rect = g.sel.rects[-1]
        single = rect[0] == rect[2] and rect[1] == rect[3]
        if single:
            r, c = g.sel.active
            rect = ops.current_region(sh, r, c)
        return C.clamp_rect(sh, rect)

    def default_anchors(self, rect, w=C.DEFAULT_SIZE[0], h=C.DEFAULT_SIZE[1]):
        """Beside the data when there's room on screen, else below it, else near the top left of the window."""
        g = self.grid
        sh = g.sheet
        z = g.zoom
        vp = self.viewport()
        r1, c1, r2, c2 = rect
        cands = [(r1, c2 + 2), (r2 + 2, c1), (g.top + 1, g.left + 1)]
        for r, c in cands:
            x, y = g.col_x(c), g.row_y(r)
            if x >= vp.left() and y >= vp.top() and x + w * z <= vp.right() and y + h * z <= vp.bottom()                     and C.overlapping(sh, C.axis_pos(sh, "col", c), C.axis_pos(sh, "row", r), w, h) is None:
                return C.place_at_cell(sh, r, c, w, h)
        return C.free_place(sh, g.top + 1, g.left + 1, w, h)

    def insert_dialog(self, ctype=None):
        from .chart_dialogs import InsertChartDialog
        win = self.win
        win._prep()
        sh = self.sheet
        rect = self.selection_rect()
        dlg = InsertChartDialog(win, sh, rect, ctype)
        if dlg.exec() != dlg.Accepted:
            return
        self.create(sh, dlg.rect, dlg.chart_type, dlg.by, dlg.combo, dlg.data_sheet)

    def create(self, sh, rect, ctype, by=None, combo=None, data_sheet=None):
        """New chart on sheet `sh` of the cells `rect` (of `data_sheet`, default the same sheet)."""
        ds = data_sheet or sh
        try:
            ch = C.make_chart(ds, rect, ctype, by, home=sh,
                              anchors=self.default_anchors(rect if ds is sh else self.selection_rect()))
        except ValueError as e:
            QMessageBox.information(self.win, "Insert Chart", str(e))
            return None
        if combo:
            ch = C.apply_type(ch, "combo")
            ch["series"] = [dict(s, type=t, secondary=sec) for s, (t, sec) in zip(ch["series"], combo)]
        self.add(ch)
        return ch

    def create_selected(self, ctype):
        """Insert a chart of the given type from the selected data (no dialog)."""
        self.win._prep()
        self.create(self.sheet, self.selection_rect(), ctype)

    def quick_chart(self):
        """Alt+F1: a clustered column chart of the selected data right away."""
        self.win._prep()
        self.create(self.sheet, self.selection_rect(), "col")

    def chart_on_new_sheet(self, ctype="col"):
        """F11: a chart of the selected data on a new sheet of its own (filling the window)."""
        win = self.win
        win._prep()
        src = self.sheet
        rect = self.selection_rect()
        try:
            probe = C.make_chart(src, rect, ctype)
        except ValueError as e:
            QMessageBox.information(win, "Insert Chart", str(e))
            return
        win.undo.beginMacro("Chart on New Sheet")
        win.add_sheet()
        sh = win.sheet
        n = 1
        while win.wb.get_sheet(f"Chart{n}"):
            n += 1
        win.undo.push(_RenameCommand(win, sh, f"Chart{n}"))
        g = self.grid
        w = max(480, int((g.width() - g.rw - 40) / g.zoom))
        h = max(288, int((g.height() - g.hh - 40) / g.zoom))
        probe["from"], probe["to"] = C.place_at_cell(sh, 1, 1, w, h)
        probe["name"] = "Chart 1"
        self.set_charts([probe], "Insert Chart", sh)
        win.undo.endMacro()
        QTimer.singleShot(0, lambda: self._select_when_ready(probe["id"]))

    # ------------------------------------------------------------ menus
    def context_menu(self, w, pos, hit):
        ch = self.find(self.selected)
        if ch is None:
            return
        m = QMenu(self.grid)
        cid = ch["id"]
        a = m.addAction("Cu&t", lambda: self.copy(cid, cut=True))
        a.setShortcut(QKeySequence("Ctrl+X"))
        a = m.addAction("&Copy", lambda: self.copy(cid))
        a.setShortcut(QKeySequence("Ctrl+C"))
        m.addAction("&Paste", self.paste).setEnabled(self.clip is not None)
        m.addAction("Duplicate", lambda: self.duplicate(cid)).setShortcut(QKeySequence("Ctrl+D"))
        m.addSeparator()
        m.addAction("Change Chart &Type...", lambda: self.change_type(cid))
        m.addAction("Select &Data...", lambda: self.select_data(cid))
        m.addAction("&Move Chart...", lambda: self.move_chart(cid))
        m.addSeparator()
        em = m.addMenu("Add Chart &Element")
        self.fill_elements_menu(em, ch)
        sm = m.addMenu("Chart &Style")
        self.fill_styles_menu(sm, ch)
        m.addSeparator()
        m.addAction("Bring to Front", lambda: self.reorder(cid, True))
        m.addAction("Send to Back", lambda: self.reorder(cid, False))
        m.addSeparator()
        kind = hit.kind if hit else "chart"
        label = {"title": "Format Chart Title...", "legend": "Format Legend...", "plot": "Format Plot Area...",
                 "xaxis": "Format Axis...", "yaxis": "Format Axis...", "y2axis": "Format Axis...",
                 "xtitle": "Format Axis Title...", "ytitle": "Format Axis Title...", "y2title": "Format Axis Title...",
                 "series": "Format Data Series..."}.get(kind, "Format Chart Area...")
        m.addAction(label, self.show_pane)
        m.addSeparator()
        m.addAction("&Delete", lambda: self.delete(cid)).setShortcut(QKeySequence("Delete"))
        m.exec(pos)

    def elements_menu_for_button(self, btn):
        ch = self.selected_chart()
        if ch is None:
            return
        m = QMenu(self.grid)
        self.fill_elements_menu(m, ch)
        m.exec(btn.mapToGlobal(QPoint(btn.width() + 2, 0)))

    def styles_menu_for_button(self, btn):
        ch = self.selected_chart()
        if ch is None:
            return
        m = QMenu(self.grid)
        self.fill_styles_menu(m, ch)
        m.exec(btn.mapToGlobal(QPoint(btn.width() + 2, 0)))

    def fill_elements_menu(self, m, ch):
        cid = ch["id"]
        fam = C.family(ch["type"])
        pie = fam in ("pie", "doughnut")

        def chk(menu, text, on, fn):
            a = menu.addAction(text)
            a.setCheckable(True)
            a.setChecked(bool(on))
            a.triggered.connect(lambda _=False: fn(not on))
            return a
        chk(m, "Chart Title", ch["title"] is not None,
            lambda on: self.update(cid, "Chart Title", title="" if on else None))
        if not pie:
            at = m.addMenu("Axis Titles")
            chk(at, "Primary Horizontal", ch["x_title"],
                lambda on: self.update(cid, "Axis Title", x_title="Axis Title" if on else ""))
            chk(at, "Primary Vertical", ch["y_title"],
                lambda on: self.update(cid, "Axis Title", y_title="Axis Title" if on else ""))
            if any(s.get("secondary") for s in ch["series"]):
                chk(at, "Secondary Vertical", ch["y2_title"],
                    lambda on: self.update(cid, "Axis Title", y2_title="Axis Title" if on else ""))
        chk(m, "Data Labels", ch["labels"], lambda on: self.update(cid, "Data Labels", labels=on))
        if not pie:
            gm = m.addMenu("Gridlines")
            chk(gm, "Primary Major Horizontal", ch["grid_y"], lambda on: self.update(cid, "Gridlines", grid_y=on))
            chk(gm, "Primary Major Vertical", ch["grid_x"], lambda on: self.update(cid, "Gridlines", grid_x=on))
        lm = m.addMenu("Legend")
        for key, label in (("right", "Right"), ("top", "Top"), ("left", "Left"), ("bottom", "Bottom"), ("none", "None")):
            a = lm.addAction(label)
            a.setCheckable(True)
            a.setChecked(ch["legend"] == key)
            a.triggered.connect(lambda _=False, k=key: self.update(cid, "Legend", legend=k))

    def fill_styles_menu(self, m, ch):
        cid = ch["id"]
        sm = m.addMenu("Style")
        for k, (label, opts) in C.LOOKS.items():
            a = sm.addAction(label)
            a.setCheckable(True)
            a.setChecked(ch.get("look") == k)
            a.triggered.connect(lambda _=False, kk=k: self.update(cid, "Chart Style", look=kk, **C.LOOKS[kk][1]))
        cm = m.addMenu("Change Colors")
        for k, (label, spec) in C.PALETTES.items():
            a = cm.addAction(S.color_square_icon(C.palette_colors(k, 3)[0], 14), label)
            a.setCheckable(True)
            a.setChecked(ch.get("style") == k)
            a.triggered.connect(lambda _=False, kk=k: self.update(cid, "Change Colors", style=kk,
                                                                   series=[dict(s, color=None) for s in ch["series"]]))

    # ------------------------------------------------------------ dialogs
    def change_type(self, cid):
        from .chart_dialogs import ChangeTypeDialog
        ch = self.find(cid)
        if ch is None:
            return
        dlg = ChangeTypeDialog(self.win, self.win.wb, self.sheet, ch)
        if dlg.exec() == dlg.Accepted:
            self.edit(cid, lambda c: dlg.apply(c), "Change Chart Type")

    def select_data(self, cid):
        from .chart_dialogs import SelectDataDialog
        ch = self.find(cid)
        if ch is None:
            return
        dlg = SelectDataDialog(self.win, self.win.wb, self.sheet, ch)
        if dlg.exec() == dlg.Accepted:
            self.edit(cid, lambda c: dlg.result_chart, "Select Data")

    def move_chart(self, cid):
        from .chart_dialogs import MoveChartDialog
        ch = self.find(cid)
        if ch is None:
            return
        dlg = MoveChartDialog(self.win, self.win.wb, self.sheet)
        if dlg.exec() != dlg.Accepted:
            return
        win = self.win
        src = self.sheet
        if dlg.target_new:
            win.undo.beginMacro("Move Chart")
            self.set_charts([c for c in src.charts if c["id"] != cid], "Move Chart", src)
            win.add_sheet()
            tgt = win.sheet
            n = 1
            while win.wb.get_sheet(f"Chart{n}"):
                n += 1
            win.undo.push(_RenameCommand(win, tgt, f"Chart{n}"))
        else:
            tgt = dlg.target_sheet
            if tgt is src:
                return
            win.undo.beginMacro("Move Chart")
            self.set_charts([c for c in src.charts if c["id"] != cid], "Move Chart", src)
        moved = dict(ch)
        x, y, w, h = C.chart_px_rect(src, ch)
        if dlg.target_new:
            g = self.grid
            w = max(480, int((g.width() - g.rw - 40) / g.zoom))
            h = max(288, int((g.height() - g.hh - 40) / g.zoom))
            x, y = C.axis_pos(tgt, "col", 1), C.axis_pos(tgt, "row", 1)
        moved["from"], moved["to"] = C.anchors_for(tgt, x, y, w, h)
        # keep the data refs pointing at the source sheet (they carry the sheet name already)
        self.set_charts(tgt.charts + [moved], "Move Chart", tgt)
        win.undo.endMacro()
        win.show_sheet(tgt)

    def show_pane(self):
        from .chart_pane import ChartPane
        if self.pane is None:
            self.pane = ChartPane(self.win, self)
            self.win.addDockWidget(Qt.RightDockWidgetArea, self.pane)
        self.pane.show()
        self.pane.raise_()
        self.pane.refresh()


class _RenameCommand(_Base):
    """Part of a macro: give the sheet we just added a name (undoes with the rest)."""

    def __init__(self, win, sheet, name):
        super().__init__(win, sheet, "Rename Sheet")
        self.name = name
        self.old = sheet.name

    def redo(self):
        self.win.wb.rename_sheet(self.sheet, self.name)
        self.win.rebuild_tabs()
        self.win.after_change(self.sheet)

    def undo(self):
        self.win.wb.rename_sheet(self.sheet, self.old)
        self.win.rebuild_tabs()
        self.win.after_change(self.sheet)


def _same_data(a, b):
    if a is b:
        return True
    if a["error"] != b["error"] or a["sampled"] != b["sampled"] or len(a["series"]) != len(b["series"]) \
            or a["cats"] != b["cats"]:
        return False
    for x, y in zip(a["series"], b["series"]):
        if x["name"] != y["name"] or x["y"] != y["y"] or x["x"] != y["x"] or x["fmt"] != y["fmt"] or \
                x["src"] is not y["src"] and x["src"] != y["src"]:
            return False
    return True


def build_chart_menu(win, menu):
    """The toolbar's Insert Chart drop-down: the gallery, Chart Now, and one entry per chart type."""
    menu.addAction(win.a_ins_chart)
    menu.addAction(win.a_chart_quick)
    menu.addAction(win.a_chart_sheet)
    menu.addSeparator()
    group = None
    for t, (label, grp) in C.TYPES.items():
        if grp != group:
            if group is not None:
                menu.addSeparator()
            group = grp
        if t == "combo_sec":
            continue
        a = menu.addAction(type_icon(t), label)
        a.triggered.connect(lambda _=False, tt=t: win.chart_layer.create_selected(tt))
