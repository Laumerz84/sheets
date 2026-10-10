"""The spreadsheet grid: painting, selection, scrolling, frozen panes, editing."""
from bisect import bisect_left

from PySide6.QtCore import QEvent, QPoint, QRect, QRectF, Qt, QTimer, Signal
from PySide6.QtGui import (QBrush, QColor, QCursor, QFont, QFontMetrics,
                           QPainter, QPainterPath, QPen, QPolygon, QRegion)
from PySide6.QtWidgets import QScrollBar, QWidget

from .. import ops
from ..condfmt import CFEngine
from ..errors import XLError
from ..formula import ref_spans
from ..numfmt import format_general_fit, format_value
from ..refs import MAX_COLS, MAX_ROWS, col_name, range_addr, addr
from ..workbook import DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT, DEFAULT_STYLE
from .editor import CellEditor
from . import style as S  # read at paint time: the skin (View > Skin) can change them

PAD = 3


class Axis:
    """Positions of rows or columns with sparse custom sizes and hidden entries."""

    def __init__(self, default, count):
        self.base_default = default
        self.count = count
        self.D = default
        self.spec = []
        self.acc = [0]
        self.sizes = {}

    def configure(self, sizes, hidden, zoom):
        self.D = max(1, int(round(self.base_default * zoom)))
        sp = {}
        for i, s in sizes.items():
            sp[i] = max(0, int(round(s * zoom)))
        for i in hidden:
            sp[i] = 0
        self.sizes = sp
        self.spec = sorted(sp)
        acc = [0]
        D = self.D
        for i in self.spec:
            acc.append(acc[-1] + sp[i] - D)
        self.acc = acc

    def size(self, i):
        return self.sizes.get(i, self.D)

    def pos(self, i):
        return i * self.D + self.acc[bisect_left(self.spec, i)]

    def index_at(self, p):
        """First index whose extent ends after p (skips zero-size entries)."""
        if p < 0:
            p = 0
        lo, hi = 0, self.count - 1
        while lo < hi:
            mid = (lo + hi) // 2
            if self.pos(mid + 1) > p:
                hi = mid
            else:
                lo = mid + 1
        return lo

    def next_visible(self, i, step=1):
        while 0 <= i < self.count and self.size(i) == 0:
            if step > 0:
                j = self.index_at(self.pos(i))
                if j <= i:
                    i += 1
                else:
                    i = j
            else:
                i -= 1
        return max(0, min(self.count - 1, i))


class Band:
    __slots__ = ("items", "p0", "p1", "mapf", "first", "last", "frozen")

    def __init__(self, items, p0, p1, mapf, frozen):
        self.items = items
        self.p0 = p0
        self.p1 = p1
        self.mapf = mapf
        self.frozen = frozen
        self.first = items[0][0] if items else 0
        self.last = items[-1][0] if items else -1


class Selection:
    def __init__(self):
        self.rects = [(0, 0, 0, 0)]
        self.active = (0, 0)
        self.anchor = (0, 0)
        self.end = (0, 0)

    def copy(self):
        s = Selection()
        s.rects = list(self.rects)
        s.active, s.anchor, s.end = self.active, self.anchor, self.end
        return s

    @property
    def last(self):
        return self.rects[-1]

    def contains(self, r, c):
        return any(a <= r <= b and x <= c <= y for a, x, b, y in self.rects)

    def is_single_cell(self, sheet=None):
        if len(self.rects) != 1:
            return False
        r1, c1, r2, c2 = self.rects[0]
        if r1 == r2 and c1 == c2:
            return True
        return sheet is not None and sheet.merge_at(r1, c1) == self.rects[0]


def expand_merges(sheet, rect):
    r1, c1, r2, c2 = rect
    if not sheet.merges:
        return rect
    changed = True
    while changed:
        changed = False
        for m in sheet.merges:
            if m[0] <= r2 and m[2] >= r1 and m[1] <= c2 and m[3] >= c1:
                nr1, nc1, nr2, nc2 = min(r1, m[0]), min(c1, m[1]), max(r2, m[2]), max(c2, m[3])
                if (nr1, nc1, nr2, nc2) != (r1, c1, r2, c2):
                    r1, c1, r2, c2 = nr1, nc1, nr2, nc2
                    changed = True
    return (r1, c1, r2, c2)


class Grid(QWidget):
    selection_changed = Signal()
    selection_done = Signal()                   # mouse selection finished
    escape_pressed = Signal()
    read_only_hit = Signal()
    commit_requested = Signal(int, int, str, bool)
    edit_state_changed = Signal(str)            # Ready / Enter / Edit / Point
    edit_text_changed = Signal(str)
    context_menu_requested = Signal(str, QPoint)  # area: cell/row/col/corner
    fill_requested = Signal(object, object)
    filter_popup_requested = Signal(int, QPoint)
    clear_requested = Signal()
    col_widths_changed = Signal(object, object)   # old, new
    row_heights_changed = Signal(object, object)
    autofit_cols_requested = Signal(object)
    autofit_rows_requested = Signal(object)
    zoom_changed = Signal(float)
    scrolled = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFocusPolicy(Qt.StrongFocus)
        self.setMouseTracking(True)
        self.setAttribute(Qt.WA_OpaquePaintEvent)
        self.setAttribute(Qt.WA_InputMethodEnabled)
        self.sheet = None
        self.rows = Axis(DEFAULT_ROW_HEIGHT, MAX_ROWS)
        self.cols = Axis(DEFAULT_COL_WIDTH, MAX_COLS)
        self.zoom = 1.0
        self.top = 0
        self.left = 0
        self.hh = 20
        self.rw = 36
        self.sel = Selection()
        self._view_states = {}
        self._dcache = {}
        self._fonts = {}
        self.show_formulas = False
        self.marquee = None          # rect being copied/cut
        self._marquee_phase = 0
        self.fill_target = None
        self.ref_highlights = []
        self.drag_mode = None
        self._drag_info = None
        self.editing = False
        self.edit_cell = None
        self.edit_widget = None
        self.point_anchor = None
        self.point_cell = None
        self.bar = None
        self.read_only = False       # set while Claude works on the workbook

        self.vbar = QScrollBar(Qt.Vertical)
        self.hbar = QScrollBar(Qt.Horizontal)
        self.vbar.valueChanged.connect(self._vbar_moved)
        self.hbar.valueChanged.connect(self._hbar_moved)

        self.editor = CellEditor(self)
        self.editor.hide()
        self.editor.commit.connect(self.finish_edit)
        self.editor.cancel.connect(self.cancel_edit)
        self.editor.point_key.connect(self._point_key)
        self.editor.edited.connect(self._editor_edited)

        self._auto_timer = QTimer(self)
        self._auto_timer.setInterval(40)
        self._auto_timer.timeout.connect(self._autoscroll_tick)
        self._last_mouse = QPoint()
        self._marquee_timer = QTimer(self)
        self._marquee_timer.setInterval(120)
        self._marquee_timer.timeout.connect(self._marquee_tick)
        self.header_font = QFont(S.HEADER_FONT_FAMILY)

    # ================================================================ setup
    def attach_formula_bar(self, bar):
        self.bar = bar
        bar.commit.connect(self.finish_edit)
        bar.cancel.connect(self.cancel_edit)
        bar.edited.connect(self._bar_edited)
        bar.installEventFilter(self)
        self.editor.installEventFilter(self)

    def eventFilter(self, obj, ev):
        if ev.type() == QEvent.FocusIn and self.sheet is not None:
            if obj is self.bar:
                if not self.editing:
                    self.begin_edit(mode="edit", widget=self.bar, keep_bar_cursor=True)
                else:
                    self.edit_widget = self.bar
            elif obj is self.editor and self.editing:
                self.edit_widget = self.editor
        return False

    def set_sheet(self, sheet):
        if self.editing:
            self.finish_edit(0, 0, False)
        if self.sheet is not None:
            self._view_states[self.sheet] = (self.sel.copy(), self.top, self.left)
        self.sheet = sheet
        self.zoom = sheet.zoom
        st = self._view_states.get(sheet)
        if st:
            self.sel, self.top, self.left = st[0].copy(), st[1], st[2]
        else:
            fr, fc = sheet.freeze
            self.sel = Selection()
            self.sel.rects = [(fr, fc, fr, fc)]
            self.sel.active = self.sel.anchor = self.sel.end = (fr, fc)
            self.top, self.left = fr, fc
        self.marquee = None if (self.marquee and self.marquee[0] is not sheet) else self.marquee
        self.relayout()
        self.selection_changed.emit()

    def forget_sheet(self, sheet):
        self._view_states.pop(sheet, None)

    def relayout(self):
        sh = self.sheet
        if sh is None:
            return
        self._fonts = {}
        self._dcache = {}
        self._cf = None
        self.rows.configure(sh.row_heights, sh.hidden_rows | sh.filter_hidden, self.zoom)
        self.cols.configure(sh.col_widths, sh.hidden_cols, self.zoom)
        self.hh = max(14, int(round(20 * min(max(self.zoom, 0.6), 2.0))))
        self.header_font = QFont(S.HEADER_FONT_FAMILY)
        self.header_font.setPixelSize(max(8, int(round(12 * min(max(self.zoom, 0.6), 2.0)))))
        fr, fc = sh.freeze
        self.top = self.rows.next_visible(max(self.top, fr))
        self.left = self.cols.next_visible(max(self.left, fc))
        self._update_row_header_width()
        self._update_scrollbars()
        self._place_editor()
        self.update()

    def invalidate(self, relayout=False):
        self._dcache = {}
        self._cf = None
        if relayout:
            self.relayout()
        else:
            self._update_scrollbars()
            self.update()

    # ================================================================ geometry
    # Frozen panes show rows fo_r..fr-1 (and cols fo_c..fc-1) pinned at the top/left;
    # fo_* is the first frozen row/col on screen (Excel's topLeftCell when freezing while scrolled).
    @property
    def fr(self):
        return self.sheet.freeze[0]

    @property
    def fc(self):
        return self.sheet.freeze[1]

    @property
    def fo_r(self):
        return min(self.sheet.freeze_origin[0], self.fr) if self.fr else 0

    @property
    def fo_c(self):
        return min(self.sheet.freeze_origin[1], self.fc) if self.fc else 0

    def frozen_h(self):
        return self.rows.pos(self.fr) - self.rows.pos(self.fo_r) if self.fr else 0

    def frozen_w(self):
        return self.cols.pos(self.fc) - self.cols.pos(self.fo_c) if self.fc else 0

    def row_y(self, r):
        if r < self.fr:
            return self.hh + self.rows.pos(r) - self.rows.pos(self.fo_r)
        return self.hh + self.frozen_h() + self.rows.pos(r) - self.rows.pos(self.top)

    def col_x(self, c):
        if c < self.fc:
            return self.rw + self.cols.pos(c) - self.cols.pos(self.fo_c)
        return self.rw + self.frozen_w() + self.cols.pos(c) - self.cols.pos(self.left)

    def row_at(self, y):
        yy = y - self.hh
        fh = self.frozen_h()
        if self.fr and yy < fh:
            return self.rows.index_at(max(0, yy) + self.rows.pos(self.fo_r))
        return self.rows.index_at(yy - fh + self.rows.pos(self.top))

    def col_at(self, x):
        xx = x - self.rw
        fw = self.frozen_w()
        if self.fc and xx < fw:
            return self.cols.index_at(max(0, xx) + self.cols.pos(self.fo_c))
        return self.cols.index_at(xx - fw + self.cols.pos(self.left))

    def cell_rect(self, r, c):
        """Pixel rect of a cell (or the merge containing it)."""
        m = self.sheet.merge_at(r, c)
        if m:
            r1, c1, r2, c2 = m
        else:
            r1, c1, r2, c2 = r, c, r, c
        x = self.col_x(c1)
        y = self.row_y(r1)
        x2 = self.col_x(c2) + self.cols.size(c2)
        y2 = self.row_y(r2) + self.rows.size(r2)
        return QRect(x, y, x2 - x, y2 - y)

    def _bands(self, axis, n_frozen, first, p_start, p_end, f_origin=0):
        bands = []
        if n_frozen:
            items = []
            p = p_start
            for i in range(f_origin, n_frozen):
                s = axis.size(i)
                if s:
                    items.append((i, p, s))
                p += s
            end = p
            base = p_start - axis.pos(f_origin)
            bands.append(Band(items, p_start, min(end, p_end), lambda i, b=base: b + axis.pos(i), True))
            p_start = end
        items = []
        p = p_start
        i = first
        origin = axis.pos(first)
        count = axis.count
        while p < p_end and i < count:
            s = axis.size(i)
            if s == 0:
                j = axis.index_at(axis.pos(i))
                i = j if j > i else i + 1
                continue
            items.append((i, p, s))
            p += s
            i += 1
        base = p_start
        bands.append(Band(items, p_start, p_end, lambda i, b=base, o=origin: b + axis.pos(i) - o, False))
        return bands

    def row_bands(self):
        return self._bands(self.rows, self.fr, self.top, self.hh, self.height(), self.fo_r)

    def col_bands(self):
        return self._bands(self.cols, self.fc, self.left, self.rw, self.width(), self.fo_c)

    def visible_row_count(self):
        h = self.height() - self.hh - self.frozen_h()
        return max(1, h // max(1, self.rows.D))

    def visible_col_count(self):
        w = self.width() - self.rw - self.frozen_w()
        return max(1, w // max(1, self.cols.D))

    def _update_row_header_width(self):
        last = self.top + self.visible_row_count() + 1
        digits = len(str(min(MAX_ROWS, last + 1)))
        fm = QFontMetrics(self.header_font)
        rw = max(int(26 * min(self.zoom, 1.5)), fm.horizontalAdvance("9" * max(digits, 3)) + 12)
        if rw != self.rw:
            self.rw = rw
            self._place_editor()

    # ================================================================ scrolling
    def _update_scrollbars(self):
        if self.sheet is None:
            return
        ur, uc = self.sheet.max_row, self.sheet.max_col
        mr, mc = self.sheet.used_extent() if (ur < 0 and self.sheet.styles) else (ur, uc)
        ar, ac = self.sel.active
        vis_r, vis_c = self.visible_row_count(), self.visible_col_count()
        vmax = max(self.fr, mr, ar, self.top, vis_r)
        hmax = max(self.fc, mc, ac, self.left, vis_c // 2)
        self.vbar.blockSignals(True)
        self.vbar.setRange(self.fr, min(MAX_ROWS - 1, vmax))
        self.vbar.setPageStep(vis_r)
        self.vbar.setSingleStep(1)
        self.vbar.setValue(self.top)
        self.vbar.blockSignals(False)
        self.hbar.blockSignals(True)
        self.hbar.setRange(self.fc, min(MAX_COLS - 1, hmax))
        self.hbar.setPageStep(vis_c)
        self.hbar.setValue(self.left)
        self.hbar.blockSignals(False)

    def _vbar_moved(self, v):
        self.set_top(v, from_bar=True)

    def _hbar_moved(self, v):
        self.set_left(v, from_bar=True)

    def set_top(self, r, from_bar=False):
        r = max(self.fr, min(MAX_ROWS - 1, r))
        r = self.rows.next_visible(r)
        if r == self.top:
            return
        self.top = r
        self._update_row_header_width()
        if not from_bar:
            self._update_scrollbars()
        elif r >= self.vbar.maximum():
            self._update_scrollbars()
        self._place_editor()
        self.update()
        self.scrolled.emit()

    def set_left(self, c, from_bar=False):
        c = max(self.fc, min(MAX_COLS - 1, c))
        c = self.cols.next_visible(c)
        if c == self.left:
            return
        self.left = c
        if not from_bar or c >= self.hbar.maximum():
            self._update_scrollbars()
        self._place_editor()
        self.update()
        self.scrolled.emit()

    def scroll_rows(self, n):
        r = self.top
        step = 1 if n > 0 else -1
        for _ in range(abs(n)):
            nr = r + step
            if nr < self.fr or nr >= MAX_ROWS:
                break
            nr = self.rows.next_visible(nr, step)
            if self.rows.size(nr) == 0:
                break
            r = nr
        self.set_top(r)

    def scroll_cols(self, n):
        c = self.left
        step = 1 if n > 0 else -1
        for _ in range(abs(n)):
            nc = c + step
            if nc < self.fc or nc >= MAX_COLS:
                break
            nc = self.cols.next_visible(nc, step)
            if self.cols.size(nc) == 0:
                break
            c = nc
        self.set_left(c)

    def ensure_visible(self, r, c):
        if r >= self.fr:
            if r < self.top:
                self.set_top(r)
            else:
                avail = self.height() - self.hh - self.frozen_h()
                bottom = self.rows.pos(r + 1) - self.rows.pos(self.top)
                if bottom > avail:
                    # choose top so that row r is the last fully visible row
                    target = self.rows.pos(r + 1) - avail
                    t = self.rows.index_at(target)
                    if self.rows.pos(t) < target:
                        t += 1
                    self.set_top(max(self.fr, min(t, r)))
        if c >= self.fc:
            if c < self.left:
                self.set_left(c)
            else:
                avail = self.width() - self.rw - self.frozen_w()
                right = self.cols.pos(c + 1) - self.cols.pos(self.left)
                if right > avail:
                    target = self.cols.pos(c + 1) - avail
                    t = self.cols.index_at(target)
                    if self.cols.pos(t) < target:
                        t += 1
                    self.set_left(max(self.fc, min(t, c)))

    def wheelEvent(self, e):
        dy = e.angleDelta().y()
        dx = e.angleDelta().x()
        if e.modifiers() & Qt.ControlModifier:
            z = self.zoom + (0.1 if dy > 0 else -0.1)
            self.set_zoom(z)
            return
        if e.modifiers() & Qt.ShiftModifier and dy:
            self.scroll_cols(-3 if dy > 0 else 3)
            return
        if dy:
            steps = max(1, abs(dy) // 40)
            self.scroll_rows(-steps if dy > 0 else steps)
        if dx:
            self.scroll_cols(-2 if dx > 0 else 2)

    def set_zoom(self, z):
        z = max(0.25, min(4.0, round(z, 2)))
        if self.sheet is None or z == self.zoom:
            return
        self.zoom = z
        self.sheet.zoom = z
        self.relayout()
        self.zoom_changed.emit(z)

    def resizeEvent(self, e):
        self._update_row_header_width()
        self._update_scrollbars()
        self._place_editor()
        super().resizeEvent(e)

    # ================================================================ painting helpers
    def font_for(self, st):
        k = (st.font, st.size, st.bold, st.italic, st.underline, st.strike)
        f = self._fonts.get(k)
        if f is None:
            font = QFont(st.font or S.DISPLAY_FONT_FAMILY)
            px = (st.size or S.DISPLAY_FONT_SIZE) * 96 / 72 * self.zoom
            font.setPixelSize(max(1, int(round(px))))
            font.setBold(st.bold)
            font.setItalic(st.italic)
            font.setUnderline(st.underline)
            font.setStrikeOut(st.strike)
            f = (font, QFontMetrics(font))
            self._fonts[k] = f
        return f

    def display(self, k, st):
        """(text, color, value) for a cell key."""
        sh = self.sheet
        if self.show_formulas:
            f = sh.formulas.get(k)
            if f is not None:
                return f.text, None, f.text
        v = sh.value_at_key(k)
        c = self._dcache.get(k)
        fmt = st.numfmt
        if c is not None and c[1] == fmt and type(c[0]) is type(v) and c[0] == v:
            return c[2], c[3], v
        text, color = format_value(v, fmt)
        self._dcache[k] = (v, fmt, text, color)
        return text, color, v

    # ================================================================ paint
    def paintEvent(self, ev):
        if self.sheet is None:
            return
        p = QPainter(self)
        W, H = self.width(), self.height()
        p.fillRect(0, 0, W, H, S.CELL_BG)
        rbands = self.row_bands()
        cbands = self.col_bands()
        for rb in rbands:
            for cb in cbands:
                clip = QRect(cb.p0, rb.p0, cb.p1 - cb.p0, rb.p1 - rb.p0)
                if clip.isEmpty():
                    continue
                p.save()
                p.setClipRect(clip)
                self._paint_cells(p, rb, cb)
                self._paint_overlays(p, rb, cb, clip)
                p.restore()
        self._paint_headers(p, rbands, cbands)
        p.setPen(QPen(S.FROZEN_LINE, 1))
        if self.fr:
            y = self.hh + self.frozen_h() - 1
            p.drawLine(0, y, W, y)
        if self.fc:
            x = self.rw + self.frozen_w() - 1
            p.drawLine(x, 0, x, H)
        p.end()

    def _visible_merges(self, rb, cb):
        out = []
        for m in self.sheet.merges:
            if m[2] >= rb.first and m[0] <= rb.last and m[3] >= cb.first and m[1] <= cb.last:
                out.append(m)
        return out

    def _paint_cells(self, p, rb, cb):
        sh = self.sheet
        values, formulas, styles = sh.values, sh.formulas, sh.styles
        rows, cols = rb.items, cb.items
        if not rows or not cols:
            return
        merges = self._visible_merges(rb, cb)
        covered = set()
        for r1, c1, r2, c2 in merges:
            for r in range(max(r1, rb.first), min(r2, rb.last) + 1):
                for c in range(max(c1, cb.first), min(c2, cb.last) + 1):
                    covered.add((r << 14) | c)

        # gridlines
        if sh.show_grid:
            p.setPen(QPen(S.GRID_LINE, 1))
            y0, y1 = rb.p0, rb.items[-1][1] + rb.items[-1][2]
            x0, x1 = cb.p0, cb.items[-1][1] + cb.items[-1][2]
            for c, x, w in cols:
                p.drawLine(x + w - 1, y0, x + w - 1, y1)
            for r, y, h in rows:
                p.drawLine(x0, y + h - 1, x1, y + h - 1)

        # conditional formatting overrides for visible cells
        cfmap = {}
        rules = sh.cond_formats
        if rules:
            if self._cf is None or self._cf.sheet is not sh:
                self._cf = CFEngine(sh)
            eng = self._cf
            for r, y, h in rows:
                for c, x, w in cols:
                    for rule in rules:
                        if rule.contains(r, c):
                            ov = eng.style_for(rules, r, c)
                            if ov:
                                cfmap[(r << 14) | c] = ov
                            break

        # fills
        has_styles = bool(styles)
        if has_styles or cfmap:
            for r, y, h in rows:
                base = r << 14
                for c, x, w in cols:
                    k = base | c
                    if k in covered:
                        continue
                    ov = cfmap.get(k)
                    fill = ov.get("fill") if ov else None
                    if fill is None:
                        st = styles.get(k)
                        fill = st.fill if st is not None else None
                    if fill:
                        p.fillRect(x - 1, y - 1, w + 1, h + 1, QColor(fill))
                    if ov and "bar" in ov:
                        frac, color = ov["bar"]
                        bw = int((w - 5) * frac)
                        if bw > 0:
                            col = QColor(color)
                            col.setAlpha(200)
                            p.fillRect(x + 1, y + 2, bw, h - 5, col)

        # text (with overflow into empty neighbours)
        last_col = cb.last
        for r, y, h in rows:
            base = r << 14
            # text overflowing from the left edge of a scrolled band
            if not cb.frozen and cb.first > 0:
                self._paint_left_overflow(p, r, y, h, cb)
            for idx, (c, x, w) in enumerate(cols):
                k = base | c
                if (k not in values and k not in formulas) or k in covered:
                    continue
                st = styles.get(k, DEFAULT_STYLE) if has_styles else DEFAULT_STYLE
                ov = cfmap.get(k) if cfmap else None
                if ov:
                    kw = {f: ov[f] for f in ("color", "bold", "italic") if f in ov}
                    if kw:
                        st = st.with_(**kw)
                self._paint_text(p, k, r, c, x, y, w, h, st, cb)

        # merged cells
        for m in merges:
            self._paint_merge(p, m, rb, cb)

        # borders
        if has_styles:
            for r, y, h in rows:
                base = r << 14
                for c, x, w in cols:
                    st = styles.get(base | c)
                    if st is not None and st.border != (None, None, None, None):
                        self._paint_borders(p, st.border, x, y, w, h)

        # autofilter buttons
        af = sh.autofilter
        if af and rb.first <= af[0] <= rb.last:
            ry = rb.mapf(af[0])
            rh = self.rows.size(af[0])
            if rh:
                for c, x, w in cols:
                    if af[1] <= c <= af[3]:
                        self._paint_filter_button(p, self._filter_button_rect(x, ry, w, rh), c in sh.filters)

    def _filter_button_rect(self, x, y, w, h):
        s = max(12, min(17, h - 3))
        return QRect(x + w - s - 2, y + h - s - 2, s, s)

    def _paint_filter_button(self, p, rect, active):
        p.save()
        p.setPen(QPen(QColor("#A0A0A0"), 1))
        p.setBrush(QColor("#F2F2F2") if not active else QColor("#DDEBE2"))
        p.drawRect(rect.adjusted(0, 0, -1, -1))
        cx = rect.center().x() + 1
        cy = rect.center().y()
        p.setPen(Qt.NoPen)
        p.setBrush(QColor("#505050"))
        if active:
            path = QPainterPath()
            path.moveTo(cx - 4, cy - 3)
            path.lineTo(cx + 3, cy - 3)
            path.lineTo(cx, cy + 1)
            path.lineTo(cx, cy + 4)
            path.lineTo(cx - 1, cy + 4)
            path.lineTo(cx - 1, cy + 1)
            path.closeSubpath()
            p.setBrush(S.ACCENT_DARK)
            p.drawPath(path)
        else:
            p.drawPolygon(QPolygon([QPoint(cx - 4, cy - 2), QPoint(cx + 3, cy - 2), QPoint(cx - 1, cy + 2)]))
        p.restore()

    def _cell_empty(self, r, c):
        k = (r << 14) | c
        sh = self.sheet
        return k not in sh.values and k not in sh.formulas and not (sh.merges and sh.merge_at(r, c))

    def _text_layout(self, k, r, c, st):
        text, color, v = self.display(k, st)
        if isinstance(v, str) and "\n" in text and not st.wrap:
            text = text.replace("\r", "").replace("\n", " ")
        if st.halign:
            ha = st.halign
        elif self.show_formulas and k in self.sheet.formulas:
            ha = "left"
        elif isinstance(v, bool) or isinstance(v, XLError):
            ha = "center"
        elif isinstance(v, (int, float)):
            ha = "right"
        else:
            ha = "left"
        return text, color, v, ha

    def _paint_text(self, p, k, r, c, x, y, w, h, st, cb, clip_extra=None):
        text, color, v, ha = self._text_layout(k, r, c, st)
        if text == "":
            return
        font, fm = self.font_for(st)
        p.setFont(font)
        p.setPen(QColor(color or st.color or "#000000"))
        indent = int(st.indent * 9 * self.zoom) if st.indent else 0
        avail = w - 2 * PAD - indent
        va = {"top": Qt.AlignTop, "center": Qt.AlignVCenter}.get(st.valign, Qt.AlignBottom)
        if st.wrap:
            flags = va | {"left": Qt.AlignLeft, "center": Qt.AlignHCenter, "right": Qt.AlignRight}[ha] | Qt.TextWordWrap
            p.save()
            p.setClipRect(QRect(x, y, w - 1, h - 1), Qt.IntersectClip)
            p.drawText(QRect(x + PAD + indent, y + 1, avail, h - 3), flags, text)
            p.restore()
            return
        tw = fm.horizontalAdvance(text)
        is_number = isinstance(v, (int, float)) and not isinstance(v, bool) and not self.show_formulas
        if is_number and tw > avail:
            if st.numfmt == "General":
                chars = max(1, int(avail / max(1, fm.horizontalAdvance("0"))))
                text = format_general_fit(v, chars)
                tw = fm.horizontalAdvance(text)
            if tw > avail:
                text = "#" * max(1, int(avail / max(1, fm.horizontalAdvance("#"))))
                tw = fm.horizontalAdvance(text)
        rect = QRect(x + PAD + (indent if ha == "left" else 0), y, avail, h - 2)
        if tw > avail and not is_number:
            # overflow into empty neighbouring cells (Excel style)
            if ha == "left":
                right = x + w
                cc = c + 1
                n = 0
                while tw > right - x - 2 * PAD - indent and cc < MAX_COLS and n < 200 and self._cell_empty(r, cc):
                    right += self.cols.size(cc)
                    cc += 1
                    n += 1
                if right > x + w:
                    self._clear_overflow(p, r, c + 1, cc - 1, x + w, y, right, h)
                rect = QRect(x + PAD + indent, y, right - x - 2 * PAD - indent, h - 2)
            elif ha == "right":
                left = x
                cc = c - 1
                n = 0
                while tw > x + w - left - 2 * PAD and cc >= 0 and n < 200 and self._cell_empty(r, cc):
                    left -= self.cols.size(cc)
                    cc -= 1
                    n += 1
                if left < x:
                    self._clear_overflow(p, r, cc + 1, c - 1, left, y, x, h)
                rect = QRect(left + PAD, y, x + w - left - 2 * PAD, h - 2)
            elif ha == "center":
                extra = (tw - avail) // 2 + PAD
                left, right = x, x + w
                cl, cr = c - 1, c + 1
                while x - left < extra and cl >= 0 and self._cell_empty(r, cl):
                    left -= self.cols.size(cl)
                    cl -= 1
                while right - (x + w) < extra and cr < MAX_COLS and self._cell_empty(r, cr):
                    right += self.cols.size(cr)
                    cr += 1
                if left < x:
                    self._clear_overflow(p, r, cl + 1, c - 1, left, y, x, h)
                if right > x + w:
                    self._clear_overflow(p, r, c + 1, cr - 1, x + w, y, right, h)
                span = min(x - left, right - x - w)
                rect = QRect(x - span + PAD, y, w + 2 * span - 2 * PAD, h - 2)
        flags = va | {"left": Qt.AlignLeft, "center": Qt.AlignHCenter, "right": Qt.AlignRight}[ha] | Qt.TextSingleLine
        if va == Qt.AlignBottom:
            rect.adjust(0, 0, 0, -1)
        p.drawText(rect, flags, text)

    def _clear_overflow(self, p, r, c1, c2, x0, y, x1, h):
        """Hide gridlines under overflowing text, keeping neighbours' fills."""
        if not self.sheet.show_grid:
            return
        p.fillRect(QRect(x0 - 1, y, x1 - x0, h - 1), S.CELL_BG)
        styles = self.sheet.styles
        if styles:
            x = x0
            for c in range(c1, c2 + 1):
                w = self.cols.size(c)
                st = styles.get((r << 14) | c)
                if st is not None and st.fill:
                    p.fillRect(x - 1, y - 1, w + 1, h + 1, QColor(st.fill))
                x += w

    def _paint_left_overflow(self, p, r, y, h, cb):
        """Text from a cell left of the visible band that spills into it."""
        sh = self.sheet
        first = cb.first
        c = first - 1
        n = 0
        while c >= 0 and n < 40:
            k = (r << 14) | c
            if k in sh.values or k in sh.formulas:
                st = sh.styles.get(k, DEFAULT_STYLE)
                if st.wrap or (sh.merges and sh.merge_at(r, c)):
                    return
                text, color, v, ha = self._text_layout(k, r, c, st)
                if ha != "left" or not isinstance(v, str):
                    return
                x = cb.mapf(c)
                self._paint_text(p, k, r, c, x, y, self.cols.size(c), h, st, cb)
                return
            if sh.merges and sh.merge_at(r, c):
                return
            c -= 1
            n += 1

    def _paint_merge(self, p, m, rb, cb):
        r1, c1, r2, c2 = m
        sh = self.sheet
        x = cb.mapf(c1)
        x2 = cb.mapf(c2 + 1) if c2 + 1 < MAX_COLS else cb.mapf(c2) + self.cols.size(c2)
        y = rb.mapf(r1)
        y2 = rb.mapf(r2 + 1) if r2 + 1 < MAX_ROWS else rb.mapf(r2) + self.rows.size(r2)
        k = (r1 << 14) | c1
        st = sh.styles.get(k, DEFAULT_STYLE)
        p.fillRect(QRect(x, y, x2 - x - 1, y2 - y - 1), QColor(st.fill) if st.fill else S.CELL_BG)
        if sh.show_grid and not st.fill:
            p.setPen(QPen(S.GRID_LINE, 1))
            p.drawLine(x2 - 1, y, x2 - 1, y2 - 1)
            p.drawLine(x, y2 - 1, x2 - 1, y2 - 1)
        if k in sh.values or k in sh.formulas:
            text, color, v, ha = self._text_layout(k, r1, c1, st)
            font, fm = self.font_for(st)
            p.setFont(font)
            p.setPen(QColor(color or st.color or "#000000"))
            va = {"top": Qt.AlignTop, "center": Qt.AlignVCenter}.get(st.valign, Qt.AlignBottom)
            flags = va | {"left": Qt.AlignLeft, "center": Qt.AlignHCenter, "right": Qt.AlignRight}[ha]
            if st.wrap:
                flags |= Qt.TextWordWrap
            p.save()
            p.setClipRect(QRect(x, y, x2 - x - 1, y2 - y - 1), Qt.IntersectClip)
            p.drawText(QRect(x + PAD, y + 1, x2 - x - 2 * PAD, y2 - y - 3), flags, text)
            p.restore()
        # outer borders from the corner cells
        L = sh.styles.get(k, DEFAULT_STYLE).border[0]
        T = sh.styles.get(k, DEFAULT_STYLE).border[1]
        R = sh.styles.get((r1 << 14) | c2, DEFAULT_STYLE).border[2]
        B = sh.styles.get((r2 << 14) | c1, DEFAULT_STYLE).border[3]
        if L or T or R or B:
            self._paint_borders(p, (L, T, R, B), x, y, x2 - x, y2 - y)

    def _border_pen(self, spec):
        style, color = spec
        width = {"medium": 2, "thick": 3, "mediumDashed": 2, "mediumDashDot": 2,
                 "mediumDashDotDot": 2, "slantDashDot": 2}.get(style, 1)
        pen = QPen(QColor(color or "#000000"), width)
        pen.setCapStyle(Qt.SquareCap)
        if style in ("dashed", "mediumDashed"):
            pen.setStyle(Qt.DashLine)
        elif style in ("dotted", "hair"):
            pen.setStyle(Qt.DotLine)
        elif style in ("dashDot", "mediumDashDot", "slantDashDot"):
            pen.setStyle(Qt.DashDotLine)
        elif style in ("dashDotDot", "mediumDashDotDot"):
            pen.setStyle(Qt.DashDotDotLine)
        return pen

    def _paint_borders(self, p, border, x, y, w, h):
        L, T, R, B = border
        x1, y1, x2, y2 = x - 1, y - 1, x + w - 1, y + h - 1
        for spec, a, b in ((L, (x1, y1), (x1, y2)), (T, (x1, y1), (x2, y1)),
                           (R, (x2, y1), (x2, y2)), (B, (x1, y2), (x2, y2))):
            if not spec:
                continue
            pen = self._border_pen(spec)
            p.setPen(pen)
            p.drawLine(a[0], a[1], b[0], b[1])
            if spec[0] == "double":
                dx = 2 if a[0] == b[0] else 0
                dy = 2 if a[1] == b[1] else 0
                if spec is L or spec is T:
                    p.drawLine(a[0] + dx, a[1] + dy, b[0] + dx, b[1] + dy)
                else:
                    p.drawLine(a[0] - dx, a[1] - dy, b[0] - dx, b[1] - dy)

    def _rect_px(self, rect, rb, cb):
        r1, c1, r2, c2 = rect
        x = cb.mapf(c1)
        y = rb.mapf(r1)
        x2 = cb.mapf(c2) + self.cols.size(c2)
        y2 = rb.mapf(r2) + self.rows.size(r2)
        return QRect(x, y, x2 - x, y2 - y)

    def _paint_overlays(self, p, rb, cb, clip):
        sh = self.sheet
        # formula reference highlights while editing
        for rect, color in self.ref_highlights:
            R = self._rect_px(rect, rb, cb)
            if not R.intersects(clip):
                continue
            col = QColor(color)
            fill = QColor(col)
            fill.setAlpha(28)
            p.fillRect(R, fill)
            p.setPen(QPen(col, 2))
            p.setBrush(Qt.NoBrush)
            p.drawRect(R.adjusted(0, 0, -2, -2))
            p.setPen(Qt.NoPen)
            p.setBrush(col)
            for pt in (R.topLeft(), R.topRight(), R.bottomLeft(), R.bottomRight()):
                p.drawRect(pt.x() - 2, pt.y() - 2, 4, 4)

        # selection
        rects = self.sel.rects
        ar, ac = self.sel.active
        active_rect = self._rect_px(expand_merges(sh, (ar, ac, ar, ac)), rb, cb)
        for rect in rects:
            R = self._rect_px(rect, rb, cb)
            if not R.intersects(clip):
                continue
            if rect[0] != rect[2] or rect[1] != rect[3] or len(rects) > 1:
                region = QRegion(R).subtracted(QRegion(active_rect))
                p.save()
                p.setClipRegion(region, Qt.IntersectClip)
                p.fillRect(R, S.SEL_FILL)
                p.restore()
        if len(rects) == 1:
            R = self._rect_px(rects[0], rb, cb)
            p.setBrush(Qt.NoBrush)
            p.setPen(QPen(S.ACCENT, 2, Qt.SolidLine, Qt.SquareCap, Qt.MiterJoin))
            p.drawRect(R.adjusted(-1, -1, -1, -1))
            if not self.editing:
                hx, hy = R.right(), R.bottom()
                p.fillRect(QRect(hx - 3, hy - 3, 7, 7), S.CELL_BG)
                p.fillRect(QRect(hx - 2, hy - 2, 5, 5), S.ACCENT)
        else:
            p.setPen(QPen(S.ACCENT, 1))
            p.setBrush(Qt.NoBrush)
            p.drawRect(active_rect.adjusted(0, 0, -2, -2))

        # fill-handle drag preview
        if self.fill_target:
            R = self._rect_px(self.fill_target, rb, cb)
            pen = QPen(QColor("#707070"), 1, Qt.DashLine)
            p.setPen(pen)
            p.setBrush(Qt.NoBrush)
            p.drawRect(R.adjusted(0, 0, -1, -1))

        # copy / cut marquee
        if self.marquee and self.marquee[0] is sh:
            R = self._rect_px(self.marquee[1], rb, cb)
            pen = QPen(S.ACCENT, 2, Qt.CustomDashLine)
            pen.setDashPattern([3, 3])
            pen.setDashOffset(self._marquee_phase)
            p.setPen(pen)
            p.setBrush(Qt.NoBrush)
            p.drawRect(R.adjusted(-1, -1, -1, -1))

    def _paint_headers(self, p, rbands, cbands):
        W, H = self.width(), self.height()
        hh, rw = self.hh, self.rw
        p.setFont(self.header_font)
        sel = self.sel.rects
        col_ranges = [(c1, c2, r1 == 0 and r2 >= MAX_ROWS - 1) for r1, c1, r2, c2 in sel]
        row_ranges = [(r1, r2, c1 == 0 and c2 >= MAX_COLS - 1) for r1, c1, r2, c2 in sel]
        # column header
        p.fillRect(QRect(rw, 0, W - rw, hh), S.HEADER_BG)
        for cb in cbands:
            p.save()
            p.setClipRect(QRect(cb.p0, 0, cb.p1 - cb.p0, hh))
            for c, x, w in cb.items:
                inside = [full for a, b, full in col_ranges if a <= c <= b]
                rect = QRect(x, 0, w, hh)
                if S.HEADER_BEVEL:
                    _bevel(p, rect, sunken=bool(inside))
                    p.setPen(S.HEADER_TEXT)
                elif inside:
                    p.fillRect(rect, S.HEADER_FULL_BG if any(inside) else S.HEADER_SEL_BG)
                    p.fillRect(QRect(x, hh - 2, w, 2), S.ACCENT)
                    p.setPen(S.ACCENT_DARK)
                else:
                    p.setPen(S.HEADER_TEXT)
                p.drawText(rect, Qt.AlignCenter, col_name(c))
                if not S.HEADER_BEVEL:
                    p.setPen(S.HEADER_LINE)
                    p.drawLine(x + w - 1, 2, x + w - 1, hh - 1)
            p.restore()
        p.setPen(S.HEADER_LINE)
        p.drawLine(rw, hh - 1, W, hh - 1)
        # row header
        p.fillRect(QRect(0, hh, rw, H - hh), S.HEADER_BG)
        for rb in rbands:
            p.save()
            p.setClipRect(QRect(0, rb.p0, rw, rb.p1 - rb.p0))
            for r, y, h in rb.items:
                inside = [full for a, b, full in row_ranges if a <= r <= b]
                rect = QRect(0, y, rw, h)
                if S.HEADER_BEVEL:
                    _bevel(p, rect, sunken=bool(inside))
                    p.setPen(S.HEADER_TEXT)
                elif inside:
                    p.fillRect(rect, S.HEADER_FULL_BG if any(inside) else S.HEADER_SEL_BG)
                    p.fillRect(QRect(rw - 2, y, 2, h), S.ACCENT)
                    p.setPen(S.ACCENT_DARK)
                else:
                    p.setPen(S.HEADER_TEXT)
                if h >= 6:
                    p.drawText(rect.adjusted(0, 0, -3, 0), Qt.AlignCenter, str(r + 1))
                if not S.HEADER_BEVEL:
                    p.setPen(S.HEADER_LINE)
                    p.drawLine(2, y + h - 1, rw - 1, y + h - 1)
            p.restore()
        p.setPen(S.HEADER_LINE)
        p.drawLine(rw - 1, hh, rw - 1, H)
        # corner
        p.fillRect(QRect(0, 0, rw, hh), S.HEADER_BG)
        if S.HEADER_BEVEL:  # Excel 95: a plain raised button, no triangle
            _bevel(p, QRect(0, 0, rw, hh), sunken=False)
            return
        p.setPen(Qt.NoPen)
        p.setBrush(QColor("#B4B4B4"))
        s = hh - 8
        p.drawPolygon(QPolygon([QPoint(rw - 4, hh - 4), QPoint(rw - 4, hh - 4 - s), QPoint(rw - 4 - s, hh - 4)]))
        p.setPen(S.HEADER_LINE)
        p.drawLine(rw - 1, 0, rw - 1, hh - 1)
        p.drawLine(0, hh - 1, rw - 1, hh - 1)

    # ================================================================ selection API
    def set_active(self, r, c, extend=False, add=False):
        sh = self.sheet
        r = max(0, min(MAX_ROWS - 1, r))
        c = max(0, min(MAX_COLS - 1, c))
        if extend:
            self.sel.end = (r, c)
            a = self.sel.anchor
            rect = (min(a[0], r), min(a[1], c), max(a[0], r), max(a[1], c))
            self.sel.rects[-1] = expand_merges(sh, rect)
        elif add:
            rect = expand_merges(sh, (r, c, r, c))
            self.sel.rects.append(rect)
            self.sel.active = self.sel.anchor = self.sel.end = (rect[0], rect[1])
        else:
            rect = expand_merges(sh, (r, c, r, c))
            self.sel.rects = [rect]
            self.sel.active = self.sel.anchor = self.sel.end = (rect[0], rect[1])
        self._sel_changed()

    def set_selection(self, rects, active=None):
        self.sel.rects = [expand_merges(self.sheet, r) for r in rects] or [(0, 0, 0, 0)]
        a = active or (self.sel.rects[-1][0], self.sel.rects[-1][1])
        self.sel.active = self.sel.anchor = a
        last = self.sel.rects[-1]
        self.sel.end = (last[2], last[3]) if a == (last[0], last[1]) else (last[0], last[1])
        self._sel_changed()

    def _sel_changed(self, scroll=True):
        if scroll:
            er, ec = self.sel.end
            r1, c1, r2, c2 = self.sel.rects[-1]
            if (r2 - r1) < MAX_ROWS - 1 and (c2 - c1) < MAX_COLS - 1:
                self.ensure_visible(er, ec)
            elif (c2 - c1) < MAX_COLS - 1:
                self.ensure_visible(self.top, ec)
            elif (r2 - r1) < MAX_ROWS - 1:
                self.ensure_visible(er, self.left)
        self._update_scrollbars()
        self.update()
        self.selection_changed.emit()

    def select_all(self):
        self.sel.rects = [(0, 0, MAX_ROWS - 1, MAX_COLS - 1)]
        self.sel.anchor = (0, 0)
        self.sel.end = (MAX_ROWS - 1, MAX_COLS - 1)
        self.update()
        self.selection_changed.emit()

    def select_cols(self, c1, c2, add=False):
        rect = (0, min(c1, c2), MAX_ROWS - 1, max(c1, c2))
        if add:
            self.sel.rects.append(rect)
        else:
            self.sel.rects = [rect]
        if add or not (rect[1] <= self.sel.active[1] <= rect[3]):
            self.sel.active = (self.top, c1)
        self.sel.active = (max(self.top, self.sel.active[0]), self.sel.active[1])
        # anchor/end span whole columns so Shift+Arrow keeps extending columns
        self.sel.anchor = (0, c1)
        self.sel.end = (MAX_ROWS - 1, c2)
        self.update()
        self.selection_changed.emit()

    def select_rows(self, r1, r2, add=False):
        rect = (min(r1, r2), 0, max(r1, r2), MAX_COLS - 1)
        if add:
            self.sel.rects.append(rect)
        else:
            self.sel.rects = [rect]
        if add or not (rect[0] <= self.sel.active[0] <= rect[2]):
            self.sel.active = (r1, self.left)
        self.sel.anchor = (r1, 0)
        self.sel.end = (r2, MAX_COLS - 1)
        self.update()
        self.selection_changed.emit()

    def selected_rects(self):
        return list(self.sel.rects)

    def clamp_rect(self, rect):
        """Shrink full-row/column selections to the used area for processing."""
        r1, c1, r2, c2 = rect
        ur, uc = self.sheet.used_extent()
        if r2 >= MAX_ROWS - 1:
            r2 = max(r1, ur)
        if c2 >= MAX_COLS - 1:
            c2 = max(c1, uc)
        return (r1, c1, r2, c2)

    # ================================================================ movement
    def _step(self, axis, i, d):
        j = i + d
        while 0 <= j < axis.count and axis.size(j) == 0:
            j += d
        if j < 0 or j >= axis.count:
            return i
        return j

    def move(self, dr, dc, extend=False):
        sh = self.sheet
        if extend:
            r, c = self.sel.end
        else:
            r, c = self.sel.active
            m = sh.merge_at(r, c)
            if m:
                if dr > 0:
                    r = m[2]
                if dc > 0:
                    c = m[3]
        if dr:
            r = self._step(self.rows, r, dr)
        if dc:
            c = self._step(self.cols, c, dc)
        self.set_active(r, c, extend=extend)

    def move_within_selection(self, dr, dc):
        """Enter/Tab inside a multi-cell selection cycles through it."""
        if len(self.sel.rects) == 1 and not self.sel.is_single_cell(self.sheet):
            r1, c1, r2, c2 = self.sel.rects[0]
            r, c = self.sel.active
            if dr:
                r += dr
                if r > r2:
                    r, c = r1, c + 1 if c < c2 else c1
                elif r < r1:
                    r, c = r2, c - 1 if c > c1 else c2
            elif dc:
                c += dc
                if c > c2:
                    c, r = c1, r + 1 if r < r2 else r1
                elif c < c1:
                    c, r = c2, r - 1 if r > r1 else r2
            self.sel.active = (r, c)
            self.ensure_visible(r, c)
            self.update()
            self.selection_changed.emit()
            return
        self.move(dr, dc)

    def jump(self, dr, dc, extend=False):
        r, c = self.sel.end if extend else self.sel.active
        limit_r, limit_c = MAX_ROWS - 1, MAX_COLS - 1
        nr, nc = ops.data_edge(self.sheet, r, c, dr, dc, limit_r, limit_c)
        self.set_active(nr, nc, extend=extend)

    # ================================================================ editing
    def begin_edit(self, text=None, mode="edit", widget=None, keep_bar_cursor=False):
        if self.sheet is None:
            return
        if self.read_only:
            self.read_only_hit.emit()
            if widget is self.bar:
                self.setFocus()
            return
        if self.editing:
            return
        r, c = self.sel.active
        if text is None:
            text = ops.edit_text_for(self.sheet, r, c)
        self.editing = True
        self.edit_cell = (r, c)
        self.point_anchor = None
        self.point_cell = None
        if widget is not self.bar:
            self.ensure_visible(r, c)
        ed = self.editor
        ed.mode = mode
        ed.point_span = None
        ed.set_text(text)
        st = self.sheet.style(r, c)
        font, fm = self.font_for(st)
        ed.setFont(font)
        self._place_editor()
        ed.show()
        ed.raise_()
        if self.bar is not None:
            self.bar.mode = "edit"
            self.bar.point_span = None
            if not keep_bar_cursor:
                self.bar.set_text(text)
        target = widget or ed
        self.edit_widget = target
        if target is ed:
            ed.setFocus()
        self._update_ref_highlights(text)
        self.edit_state_changed.emit("Enter" if mode == "enter" else "Edit")
        self.update()

    def _place_editor(self):
        if not self.editing or self.sheet is None:
            return
        r, c = self.edit_cell
        R = self.cell_rect(r, c)
        ed = self.editor
        fm = ed.fontMetrics()
        lines = ed.text().split("\n")
        text_w = max(fm.horizontalAdvance(ln) for ln in lines) + 14
        w = max(R.width() + 3, min(text_w, self.width() - R.x()))
        h = max(R.height() + 3, len(lines) * fm.lineSpacing() + 8)
        ed.setGeometry(R.x() - 2, R.y() - 2, w, min(h, self.height() - R.y() + 2))

    def _editor_edited(self, text):
        if self.bar is not None and self.edit_widget is self.editor:
            self.bar.set_text(text)
        self._place_editor()
        self._update_ref_highlights(text)
        self.edit_text_changed.emit(text)

    def _bar_edited(self, text):
        if not self.editing:
            return
        if self.edit_widget is self.bar:
            self.editor.set_text(text)
            self._place_editor()
        self._update_ref_highlights(text)
        self.edit_text_changed.emit(text)

    def _update_ref_highlights(self, text):
        old = self.ref_highlights
        self.ref_highlights = []
        if text.startswith("="):
            try:
                spans = ref_spans(text)
            except Exception:
                spans = []
            seen = {}
            for _, _, info in spans:
                if info.sheet is not None and info.sheet.upper() != self.sheet.name.upper():
                    continue
                name = info.text().replace("$", "").upper()
                idx = seen.setdefault(name, len(seen))
                try:
                    rect = info.bounds()
                except Exception:
                    continue
                self.ref_highlights.append((rect, S.REF_COLORS[idx % len(S.REF_COLORS)]))
        if old or self.ref_highlights:
            self.update()

    def _end_edit(self):
        self.editing = False
        self.editor.hide_popups()
        if self.bar is not None:
            self.bar.hide_popups()
        self.editor.hide()
        self.ref_highlights = []
        self.point_anchor = None
        self.point_cell = None
        self.edit_state_changed.emit("Ready")
        self.update()

    def finish_edit(self, dr=0, dc=0, fill_selection=False):
        if not self.editing:
            return
        text = self.edit_widget.text() if self.edit_widget is not None else self.editor.text()
        r, c = self.edit_cell
        self._end_edit()
        self.commit_requested.emit(r, c, text, fill_selection)
        if dr or dc:
            if self.sel.active != (r, c):
                self.sel.active = (r, c)
            self.move_within_selection(dr, dc)
        self.setFocus()

    def cancel_edit(self):
        if not self.editing:
            return
        self._end_edit()
        self.selection_changed.emit()
        self.setFocus()

    def _point_key(self, dr, dc, extend):
        base = self.point_cell or self.edit_cell
        r = self._step(self.rows, base[0], dr) if dr else base[0]
        c = self._step(self.cols, base[1], dc) if dc else base[1]
        if extend:
            if self.point_anchor is None:
                self.point_anchor = base
        else:
            self.point_anchor = (r, c)
        self.point_cell = (r, c)
        self._insert_point_ref()
        self.ensure_visible(r, c)

    def _insert_point_ref(self):
        a, b = self.point_anchor, self.point_cell
        rect = (min(a[0], b[0]), min(a[1], b[1]), max(a[0], b[0]), max(a[1], b[1]))
        ref = addr(rect[0], rect[1]) if rect[0] == rect[2] and rect[1] == rect[3] else \
            f"{addr(rect[0], rect[1])}:{addr(rect[2], rect[3])}"
        w = self.edit_widget or self.editor
        w.insert_ref(ref)
        if w is self.editor and self.bar is not None:
            self.bar.set_text(w.text())
        elif w is self.bar:
            self.editor.set_text(w.text())
        self.edit_state_changed.emit("Point")

    # ================================================================ mouse
    def _hit(self, pos):
        """-> (area, r, c) area in cell/col/row/corner/colborder/rowborder/none."""
        x, y = pos.x(), pos.y()
        if x < self.rw and y < self.hh:
            return "corner", -1, -1
        if y < self.hh:
            c = self.col_at(x)
            # border hit test (right edge of column c or left edge = right edge of previous)
            right = self.col_x(c) + self.cols.size(c)
            if abs(x - right) <= 3:
                return "colborder", -1, c
            left = self.col_x(c)
            if abs(x - left) <= 3 and c > 0:
                prev = self._step(self.cols, c, -1)
                return "colborder", -1, prev
            return "col", -1, c
        if x < self.rw:
            r = self.row_at(y)
            bottom = self.row_y(r) + self.rows.size(r)
            if abs(y - bottom) <= 2:
                return "rowborder", r, -1
            top = self.row_y(r)
            if abs(y - top) <= 2 and r > 0:
                return "rowborder", self._step(self.rows, r, -1), -1
            return "row", r, -1
        return "cell", self.row_at(y), self.col_at(x)

    def _on_fill_handle(self, pos):
        if self.editing or len(self.sel.rects) != 1:
            return False
        r1, c1, r2, c2 = self.sel.rects[0]
        if r2 >= MAX_ROWS - 1 or c2 >= MAX_COLS - 1:
            return False
        R = self.cell_rect(r2, c2)
        return abs(pos.x() - R.right()) <= 4 and abs(pos.y() - R.bottom()) <= 4

    def _filter_button_hit(self, pos, r, c):
        af = self.sheet.autofilter
        if not af or r != af[0] or not (af[1] <= c <= af[3]):
            return False
        R = QRect(self.col_x(c), self.row_y(r), self.cols.size(c), self.rows.size(r))
        return self._filter_button_rect(R.x(), R.y(), R.width(), R.height()).contains(pos)

    def mousePressEvent(self, e):
        if self.sheet is None:
            return
        pos = e.position().toPoint()
        self._last_mouse = pos
        area, r, c = self._hit(pos)
        mods = e.modifiers()
        shift = bool(mods & Qt.ShiftModifier)
        ctrl = bool(mods & Qt.ControlModifier)
        if e.button() == Qt.RightButton:
            return  # handled by contextMenuEvent
        if e.button() != Qt.LeftButton:
            return

        if self.editing:
            w = self.edit_widget or self.editor
            if area == "cell" and w.can_point():
                self.point_anchor = (r, c)
                self.point_cell = (r, c)
                self._insert_point_ref()
                self.drag_mode = "point"
                w.setFocus()
                return
            self.finish_edit(0, 0, False)

        if area == "corner":
            self.select_all()
            return
        if area in ("colborder", "rowborder") and self.read_only:
            self.read_only_hit.emit()
            return
        if area == "colborder":
            cols = self._resize_targets_cols(c)
            self._drag_info = {"x": pos.x(), "cols": cols, "orig": {k: self.sheet.col_widths.get(k) for k in cols},
                               "w0": self.cols.size(c) / self.zoom}
            self.drag_mode = "colresize"
            return
        if area == "rowborder":
            rows = self._resize_targets_rows(r)
            self._drag_info = {"y": pos.y(), "rows": rows, "orig": {k: self.sheet.row_heights.get(k) for k in rows},
                               "h0": self.rows.size(r) / self.zoom}
            self.drag_mode = "rowresize"
            return
        if area == "col":
            if shift:
                a = self.sel.active[1]
                self.select_cols(a, c)
            else:
                self.select_cols(c, c, add=ctrl)
                self.sel.active = (self.top, c)
            self._drag_info = {"start": self.sel.active[1]}
            self.drag_mode = "colsel"
            self.selection_changed.emit()
            return
        if area == "row":
            if shift:
                a = self.sel.active[0]
                self.select_rows(a, r)
            else:
                self.select_rows(r, r, add=ctrl)
                self.sel.active = (r, self.left)
            self._drag_info = {"start": self.sel.active[0]}
            self.drag_mode = "rowsel"
            self.selection_changed.emit()
            return
        # cell area
        if self._on_fill_handle(pos) and not self.read_only:
            self.drag_mode = "fill"
            self._drag_info = {"src": self.sel.rects[0]}
            return
        if self._filter_button_hit(pos, r, c):
            R = self.cell_rect(r, c)
            self.filter_popup_requested.emit(c, self.mapToGlobal(QPoint(R.x(), R.bottom() + 1)))
            return
        if shift:
            self.set_active(r, c, extend=True)
        else:
            self.set_active(r, c, add=ctrl)
        self.drag_mode = "select"

    def _resize_targets_cols(self, c):
        for r1, c1, r2, c2 in self.sel.rects:
            if r1 == 0 and r2 >= MAX_ROWS - 1 and c1 <= c <= c2:
                return list(range(c1, c2 + 1))
        return [c]

    def _resize_targets_rows(self, r):
        for r1, c1, r2, c2 in self.sel.rects:
            if c1 == 0 and c2 >= MAX_COLS - 1 and r1 <= r <= r2:
                return list(range(r1, r2 + 1))
        return [r]

    def _clamped_cell(self, pos):
        x = max(self.rw, min(self.width() - 1, pos.x()))
        y = max(self.hh, min(self.height() - 1, pos.y()))
        return self.row_at(y), self.col_at(x)

    def mouseMoveEvent(self, e):
        if self.sheet is None:
            return
        pos = e.position().toPoint()
        self._last_mouse = pos
        mode = self.drag_mode
        if not (e.buttons() & Qt.LeftButton) or mode is None:
            self._update_cursor(pos)
            return
        if mode == "colresize":
            d = self._drag_info
            nw = max(0, d["w0"] + (pos.x() - d["x"]) / self.zoom)
            for c in d["cols"]:
                self.sheet.col_widths[c] = int(round(nw))
            self.cols.configure(self.sheet.col_widths, self.sheet.hidden_cols, self.zoom)
            self._place_editor()
            self.update()
            return
        if mode == "rowresize":
            d = self._drag_info
            nh = max(0, d["h0"] + (pos.y() - d["y"]) / self.zoom)
            for r in d["rows"]:
                self.sheet.row_heights[r] = int(round(nh))
            self.rows.configure(self.sheet.row_heights, self.sheet.hidden_rows | self.sheet.filter_hidden, self.zoom)
            self.update()
            return
        self._check_autoscroll(pos)
        r, c = self._clamped_cell(pos)
        if mode == "select":
            if (r, c) != self.sel.end:
                self.set_active(r, c, extend=True)
        elif mode == "colsel":
            self.select_cols(self._drag_info["start"], c)
        elif mode == "rowsel":
            self.select_rows(self._drag_info["start"], r)
        elif mode == "point":
            if (r, c) != self.point_cell:
                self.point_cell = (r, c)
                self._insert_point_ref()
        elif mode == "fill":
            self._update_fill_target(r, c)

    def _update_fill_target(self, r, c):
        r1, c1, r2, c2 = self._drag_info["src"]
        down = r - r2 if r > r2 else 0
        up = r1 - r if r < r1 else 0
        right = c - c2 if c > c2 else 0
        left = c1 - c if c < c1 else 0
        vert = max(down, up)
        horz = max(right, left)
        if vert == 0 and horz == 0:
            target = None
        elif vert >= horz:
            target = (min(r1, r), c1, max(r2, r), c2)
        else:
            target = (r1, min(c1, c), r2, max(c2, c))
        if target != self.fill_target:
            self.fill_target = target
            self.update()

    def mouseReleaseEvent(self, e):
        mode = self.drag_mode
        self.drag_mode = None
        self._auto_timer.stop()
        if mode == "colresize":
            d = self._drag_info
            new = {c: self.sheet.col_widths.get(c) for c in d["cols"]}
            self.col_widths_changed.emit(d["orig"], new)
            self.relayout()
        elif mode == "rowresize":
            d = self._drag_info
            new = {r: self.sheet.row_heights.get(r) for r in d["rows"]}
            self.row_heights_changed.emit(d["orig"], new)
            self.relayout()
        elif mode == "fill":
            target = self.fill_target
            self.fill_target = None
            if target:
                src = self._drag_info["src"]
                self.fill_requested.emit(src, target)
                self.set_selection([target], active=self.sel.active)
            self.update()
        elif mode == "point":
            w = self.edit_widget or self.editor
            w.setFocus()
        elif mode in ("select", "colsel", "rowsel"):
            self.selection_done.emit()

    def mouseDoubleClickEvent(self, e):
        pos = e.position().toPoint()
        area, r, c = self._hit(pos)
        if area == "colborder":
            self.autofit_cols_requested.emit(self._resize_targets_cols(c))
            return
        if area == "rowborder":
            self.autofit_rows_requested.emit(self._resize_targets_rows(r))
            return
        if area == "cell":
            if self._on_fill_handle(pos):
                self._autofill_down()
                return
            if self.editing:
                return
            self.begin_edit(mode="edit")

    def _autofill_down(self):
        r1, c1, r2, c2 = self.sel.rects[0]
        sh = self.sheet
        neighbour = c1 - 1 if c1 > 0 and sh.has_content(r2, c1 - 1) else c2 + 1
        if not sh.has_content(r2, neighbour) and not sh.has_content(r2 + 1, neighbour):
            return
        last = r2
        while last + 1 < MAX_ROWS and sh.has_content(last + 1, neighbour):
            last += 1
        if last > r2:
            target = (r1, c1, last, c2)
            self.fill_requested.emit((r1, c1, r2, c2), target)
            self.set_selection([target], active=self.sel.active)

    def _update_cursor(self, pos):
        from . import cursors
        themed = cursors.pointer()  # View > Cursor: replaces the arrow and the cell plus
        area, r, c = self._hit(pos)
        if area == "colborder":
            self.setCursor(Qt.SplitHCursor)
        elif area == "rowborder":
            self.setCursor(Qt.SplitVCursor)
        elif area in ("col", "row", "corner"):
            self.setCursor(themed or Qt.ArrowCursor)
        elif self._on_fill_handle(pos):
            self.setCursor(Qt.CrossCursor)
        elif area == "cell" and self._filter_button_hit(pos, r, c):
            self.setCursor(themed or Qt.ArrowCursor)
        else:
            self.setCursor(themed or _plus_cursor())

    def _check_autoscroll(self, pos):
        outside = (pos.y() >= self.height() - 2 or pos.x() >= self.width() - 2 or
                   (pos.y() < self.hh + self.frozen_h() and self.top > self.fr) or
                   (pos.x() < self.rw + self.frozen_w() and self.left > self.fc))
        if outside and self.drag_mode in ("select", "point", "fill", "colsel", "rowsel"):
            if not self._auto_timer.isActive():
                self._auto_timer.start()
        else:
            self._auto_timer.stop()

    def _autoscroll_tick(self):
        pos = self._last_mouse
        if pos.y() >= self.height() - 2:
            self.scroll_rows(1)
        elif pos.y() < self.hh + self.frozen_h() and self.top > self.fr:
            self.scroll_rows(-1)
        if pos.x() >= self.width() - 2:
            self.scroll_cols(1)
        elif pos.x() < self.rw + self.frozen_w() and self.left > self.fc:
            self.scroll_cols(-1)
        r, c = self._clamped_cell(pos)
        mode = self.drag_mode
        if mode == "select":
            self.set_active(r, c, extend=True)
        elif mode == "point":
            self.point_cell = (r, c)
            self._insert_point_ref()
        elif mode == "fill":
            self._update_fill_target(r, c)
        elif mode == "colsel":
            self.select_cols(self._drag_info["start"], c)
        elif mode == "rowsel":
            self.select_rows(self._drag_info["start"], r)
        else:
            self._auto_timer.stop()

    def contextMenuEvent(self, e):
        if self.sheet is None:
            return
        if self.editing:
            self.finish_edit(0, 0, False)
        pos = e.pos()
        area, r, c = self._hit(pos)
        if area in ("col", "colborder"):
            if not any(r1 == 0 and r2 >= MAX_ROWS - 1 and c1 <= c <= c2 for r1, c1, r2, c2 in self.sel.rects):
                self.select_cols(c, c)
            self.context_menu_requested.emit("col", e.globalPos())
        elif area in ("row", "rowborder"):
            if not any(c1 == 0 and c2 >= MAX_COLS - 1 and r1 <= r <= r2 for r1, c1, r2, c2 in self.sel.rects):
                self.select_rows(r, r)
            self.context_menu_requested.emit("row", e.globalPos())
        elif area == "cell":
            if not self.sel.contains(r, c):
                self.set_active(r, c)
            self.context_menu_requested.emit("cell", e.globalPos())
        else:
            self.context_menu_requested.emit("corner", e.globalPos())

    # ================================================================ keyboard
    def keyPressEvent(self, e):
        if self.sheet is None:
            return
        k = e.key()
        mods = e.modifiers()
        ctrl = bool(mods & Qt.ControlModifier)
        shift = bool(mods & Qt.ShiftModifier)
        alt = bool(mods & Qt.AltModifier)
        arrows = {Qt.Key_Left: (0, -1), Qt.Key_Right: (0, 1), Qt.Key_Up: (-1, 0), Qt.Key_Down: (1, 0)}
        if k in arrows:
            dr, dc = arrows[k]
            if ctrl:
                self.jump(dr, dc, extend=shift)
            else:
                self.move(dr, dc, extend=shift)
            return
        if k in (Qt.Key_Return, Qt.Key_Enter):
            self.move_within_selection(-1 if shift else 1, 0)
            return
        if k == Qt.Key_Tab:
            self.move_within_selection(0, 1)
            return
        if k == Qt.Key_Backtab:
            self.move_within_selection(0, -1)
            return
        if k == Qt.Key_Home:
            if ctrl:
                self.set_active(self.fr, self.fc, extend=shift)
            else:
                self.set_active(self.sel.end[0] if shift else self.sel.active[0], 0, extend=shift)
            return
        if k == Qt.Key_End and ctrl:
            mr, mc = self.sheet.used_extent()
            self.set_active(max(mr, 0), max(mc, 0), extend=shift)
            return
        if k in (Qt.Key_PageDown, Qt.Key_PageUp) and not ctrl:
            n = self.visible_row_count() if not alt else self.visible_col_count()
            d = 1 if k == Qt.Key_PageDown else -1
            if alt:
                r, c = self.sel.active
                self.scroll_cols(d * n)
                self.set_active(r, max(0, c + d * n), extend=shift)
            else:
                r, c = self.sel.end if shift else self.sel.active
                self.scroll_rows(d * n)
                nr = r
                for _ in range(n):
                    nr = self._step(self.rows, nr, d)
                self.set_active(nr, c, extend=shift)
            return
        if k == Qt.Key_F2:
            self.begin_edit(mode="edit")
            return
        if k == Qt.Key_Delete:
            self.clear_requested.emit()
            return
        if k == Qt.Key_Backspace:
            self.begin_edit(text="", mode="enter")
            return
        if k == Qt.Key_Escape:
            if self.marquee:
                self.set_marquee(None)
            self.escape_pressed.emit()
            return
        if k == Qt.Key_Space and ctrl and not shift:
            c1 = min(r[1] for r in self.sel.rects[-1:])
            c2 = max(r[3] for r in self.sel.rects[-1:])
            self.select_cols(c1, c2)
            return
        if k == Qt.Key_Space and shift and not ctrl:
            r1 = self.sel.rects[-1][0]
            r2 = self.sel.rects[-1][2]
            self.select_rows(r1, r2)
            return
        if k == Qt.Key_A and ctrl:
            r, c = self.sel.active
            region = ops.current_region(self.sheet, r, c)
            if self.sel.rects == [region] or region[0] == region[2] and region[1] == region[3]:
                self.select_all()
            else:
                self.set_selection([region], active=self.sel.active)
            return
        text = e.text()
        if text and text.isprintable() and not ctrl and not (alt and not ctrl):
            self.begin_edit(text=text, mode="enter")
            return
        if text and text.isprintable() and ctrl and alt:  # AltGr characters
            self.begin_edit(text=text, mode="enter")
            return
        super().keyPressEvent(e)

    def inputMethodEvent(self, e):
        if e.commitString() and not self.editing:
            self.begin_edit(text=e.commitString(), mode="enter")
            return
        super().inputMethodEvent(e)

    def focusNextPrevChild(self, nxt):
        return False  # Tab is handled in keyPressEvent

    # ================================================================ marquee
    def set_marquee(self, rect):
        if rect is None:
            self.marquee = None
            self._marquee_timer.stop()
        else:
            self.marquee = (self.sheet, rect)
            self._marquee_timer.start()
        self.update()

    def _marquee_tick(self):
        self._marquee_phase = (self._marquee_phase + 1) % 6
        self.update()


def _bevel(p, rect, sunken):
    """Windows 95 button edge around a header cell: white top/left and grey bottom/right
    (swapped when pressed, i.e. the row/column is selected)."""
    p.fillRect(rect, S.HEADER_BG)
    light, dark = (QColor("#808080"), QColor("#FFFFFF")) if sunken else (QColor("#FFFFFF"), QColor("#808080"))
    x0, y0, x1, y1 = rect.left(), rect.top(), rect.right(), rect.bottom()
    p.setPen(light)
    p.drawLine(x0, y0, x1, y0)
    p.drawLine(x0, y0, x0, y1)
    p.setPen(dark)
    p.drawLine(x0, y1, x1, y1)
    p.drawLine(x1, y0, x1, y1)


_PLUS = None


def _plus_cursor():
    """Excel's fat white plus cursor."""
    global _PLUS
    if _PLUS is None:
        from PySide6.QtGui import QPixmap
        pm = QPixmap(32, 32)
        pm.fill(Qt.transparent)
        p = QPainter(pm)
        p.setRenderHint(QPainter.Antialiasing, False)
        path = QPainterPath()
        a, b = 12, 19
        path.moveTo(a, 5)
        path.lineTo(b, 5)
        path.lineTo(b, a)
        path.lineTo(26, a)
        path.lineTo(26, b)
        path.lineTo(b, b)
        path.lineTo(b, 26)
        path.lineTo(a, 26)
        path.lineTo(a, b)
        path.lineTo(5, b)
        path.lineTo(5, a)
        path.lineTo(a, a)
        path.closeSubpath()
        p.setPen(QPen(QColor("#000000"), 1))
        p.setBrush(QColor("#FFFFFF"))
        p.drawPath(path)
        p.end()
        pm = pm.scaled(20, 20, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        _PLUS = QCursor(pm, 10, 10)
    return _PLUS
