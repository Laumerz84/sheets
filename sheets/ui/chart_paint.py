"""Draws a chart with QPainter, the way Excel 2016+ does (gray text, light gridlines, Office colours).

`paint_chart(painter, rect, chart, data, scale, selected)` draws into `rect` and returns the list of
`Region`s it drew (title, legend, axes, plot area, every bar / point / slice) so the widget can find
what is under the mouse. `data` comes from `charts.resolve`. No Qt widgets here, so the same code
paints the live chart, the gallery previews and the picture put on the clipboard."""
import math

from PySide6.QtCore import QPointF, QRectF, Qt
from PySide6.QtGui import QColor, QFont, QFontMetricsF, QPainter, QPainterPath, QPen, QPolygonF

from .. import charts as C
from ..numfmt import format_value
from . import style as S


class Region:
    __slots__ = ("kind", "key", "shape", "tip")

    def __init__(self, kind, key, shape, tip=""):
        self.kind, self.key, self.shape, self.tip = kind, key, shape, tip

    def contains(self, pt):
        if isinstance(self.shape, QRectF):
            return self.shape.contains(pt)
        return self.shape.contains(pt)

    def bounds(self):
        return self.shape if isinstance(self.shape, QRectF) else self.shape.boundingRect()


def hit_test(regions, pt):
    """The topmost region under pt (later regions are drawn on top)."""
    for reg in reversed(regions):
        if reg.kind != "chart" and reg.contains(pt):
            return reg
    for reg in regions:
        if reg.kind == "chart" and reg.contains(pt):
            return reg
    return None


# ---------------------------------------------------------------- scales
def _fmt_general(v):
    if v == 0:
        return "0"
    return format_value(round(v, 10), "General")[0]


def nice_scale(lo, hi, max_ticks, vmin=None, vmax=None, major=None, zero=True, log=False):
    """-> (min, max, step) like Excel's automatic axis."""
    if log:
        lo_p = max(lo, 1e-300) if lo and lo > 0 else 1.0
        a = math.floor(math.log10(lo_p)) if vmin is None else math.log10(max(vmin, 1e-300))
        b = math.ceil(math.log10(max(hi, lo_p * 1.0001))) if vmax is None else math.log10(max(vmax, 1e-300))
        if b <= a:
            b = a + 1
        return 10.0 ** a, 10.0 ** b, 10.0
    if lo is None or hi is None:
        lo, hi = 0.0, 1.0
    if hi < lo:
        lo, hi = hi, lo
    if vmin is None and zero:
        if lo >= 0 and (hi - lo) > hi / 6.0:
            lo = 0.0
    if vmax is None and zero:
        if hi <= 0 and (hi - lo) > -lo / 6.0:
            hi = 0.0
    if hi == lo:
        if hi == 0:
            hi = 1.0
        else:
            lo, hi = (0.0, hi) if hi > 0 else (hi, 0.0)
    lo_v = vmin if vmin is not None else lo
    hi_v = vmax if vmax is not None else hi
    if hi_v <= lo_v:
        hi_v = lo_v + 1.0
    span = hi_v - lo_v
    if major:
        step = major
    else:
        mag = 10 ** math.floor(math.log10(span / max(max_ticks, 2)))
        step = None
        for k in range(0, 6):
            for m in (1, 2, 5):
                cand = m * mag * (10 ** k)
                lo_c = lo_v if vmin is not None else math.floor(lo_v / cand + 1e-9) * cand
                hi_c = hi_v if vmax is not None else math.ceil((hi_v + 0.05 * (hi_v - lo_v)) / cand - 1e-9) * cand
                if (hi_c - lo_c) / cand <= max_ticks + 1e-9:
                    step = cand
                    break
            if step:
                break
        step = step or span
    lo_o = vmin if vmin is not None else math.floor(lo_v / step + 1e-9) * step
    hi_o = vmax if vmax is not None else math.ceil((hi_v + (0.05 * (hi_v - lo_v) if vmax is None else 0)) / step - 1e-9) * step
    if hi_o <= lo_o:
        hi_o = lo_o + step
    return lo_o, hi_o, step


def ticks_of(lo, hi, step, log=False):
    if log:
        out = []
        v = lo
        while v <= hi * 1.0000001:
            out.append(v)
            v *= 10
        return out
    n = int(round((hi - lo) / step))
    return [round(lo + i * step, 12) for i in range(max(0, min(n, 200)) + 1)]


# ---------------------------------------------------------------- the painter
class _Painter:
    def __init__(self, p, rect, ch, data, scale, selected):
        self.p, self.R, self.ch, self.data, self.s = p, QRectF(rect), ch, data, max(0.3, scale)
        self.sel = selected
        self.regions = []
        self.fill = ch.get("fill") or "#FFFFFF"
        dark = C.is_dark(self.fill)
        self.txt = QColor("#D9D9D9" if dark else "#595959")
        self.grid = QColor("#595959" if dark else "#D9D9D9")
        self.axis = QColor("#7F7F7F" if dark else "#BFBFBF")
        self.fam = C.family(ch["type"])
        self.group = C.grouping(ch["type"])
        self.series = data["series"]
        n_colors = len(self.series)
        self.pie_like = self.fam in ("pie", "doughnut")
        if self.pie_like:
            n_colors = max(len(data["cats"]), 1)
        self.colors = C.palette_colors(ch.get("style", 1), n_colors)

    # ---- fonts / text
    def font(self, px, bold=False):
        f = QFont(S.DISPLAY_FONT_FAMILY)
        f.setPixelSize(max(6, int(round(px * self.s))))
        f.setBold(bold)
        return f

    def fm(self, px, bold=False):
        return QFontMetricsF(self.font(px, bold))

    def text(self, rect, txt, px=12, flags=Qt.AlignCenter, color=None, bold=False):
        self.p.setFont(self.font(px, bold))
        self.p.setPen(color or self.txt)
        self.p.drawText(rect, flags, txt)

    # ---- series info
    def color_of(self, i):
        src = self.series[i]["src"]
        return QColor(src.get("color") or self.colors[i % len(self.colors)])

    def kind_of(self, sd):
        """('col'|'bar'|'line'|'area'|'scatter', markers, lines)"""
        t = sd["src"].get("type") if self.fam == "combo" else self.ch["type"]
        t = t or "col"
        f = C.family(t)
        if f == "line":
            return "line", t == "line_markers" or self.ch["type"] == "line_markers", True
        if f == "scatter":
            return "scatter", True, t != "scatter"
        return f, False, False

    # ================================================================ main
    def paint(self):
        p, ch = self.p, self.ch
        p.save()
        p.setRenderHint(QPainter.Antialiasing, True)
        p.setRenderHint(QPainter.TextAntialiasing, True)
        p.fillRect(self.R, QColor(self.fill))
        if ch.get("border", True):
            p.setPen(QPen(self.grid, max(1.0, self.s)))
            p.setBrush(Qt.NoBrush)
            p.drawRect(self.R.adjusted(0.5, 0.5, -0.5, -0.5))
        self.regions.append(Region("chart", None, QRectF(self.R)))
        pad = 10 * self.s
        inner = self.R.adjusted(pad, pad * 0.8, -pad, -pad)
        inner = self._title(inner)
        if not self.series:
            msg = self.data.get("error") or "No data to display"
            self.text(self.R, msg, 13, Qt.AlignCenter | Qt.TextWordWrap, color=self.txt)
        else:
            inner = self._legend(inner)
            if self.pie_like:
                self.regions.append(Region("plot", None, QRectF(inner)))
                self._pie(inner)
            else:
                self._cartesian(inner)
        if self.data.get("sampled", 1) > 1:
            n = self.data["sampled"]
            self.text(QRectF(self.R.left() + 6 * self.s, self.R.bottom() - 15 * self.s, self.R.width() / 2, 13 * self.s),
                      f"Showing every {n:,}th row", 9, Qt.AlignLeft | Qt.AlignVCenter, QColor("#A6A6A6"))
        self._selection()
        p.restore()
        return self.regions

    # ---- title / legend
    def title_text(self):
        t = self.ch.get("title")
        if t is None:
            return None
        if t == "":
            return self.series[0]["name"] if len(self.series) == 1 or self.fam == "pie" else "Chart Title"
        return t

    def _title(self, inner):
        t = self.title_text()
        if t is None:
            return inner
        fm = self.fm(18.7)
        h = fm.height() + 6 * self.s
        rect = QRectF(inner.left(), inner.top(), inner.width(), h)
        w = min(inner.width(), fm.horizontalAdvance(t) + 12 * self.s)
        trect = QRectF(rect.center().x() - w / 2, rect.top(), w, h)
        self.text(trect, fm.elidedText(t, Qt.ElideRight, trect.width()), 18.7, Qt.AlignCenter)
        self.regions.append(Region("title", None, trect))
        return inner.adjusted(0, h + 2 * self.s, 0, 0)

    def _legend_entries(self):
        if self.pie_like:
            return [(c, self.colors[i % len(self.colors)], "col") for i, c in enumerate(self.data["cats"][:40])]
        out = []
        for i, sd in enumerate(self.series):
            k, mk, ln = self.kind_of(sd)
            out.append((sd["name"], self.color_of(i).name(), k if k != "bar" else "col", mk))
        return out

    def _legend(self, inner):
        pos = self.ch.get("legend", "right")
        if pos == "none":
            return inner
        entries = self._legend_entries()
        if not entries:
            return inner
        fm = self.fm(12)
        fh = fm.height()
        sw = 11 * self.s
        gap = 6 * self.s
        lh = fh + 4 * self.s
        s = self.s
        if pos in ("left", "right"):
            maxw = min(inner.width() * 0.3, max(fm.horizontalAdvance(e[0]) for e in entries) + sw + gap + 4 * s)
            n_fit = max(1, int(inner.height() // lh))
            shown = entries[:n_fit]
            h = lh * len(shown)
            x = inner.right() - maxw if pos == "right" else inner.left()
            rect = QRectF(x, inner.center().y() - h / 2, maxw, h)
            for i, e in enumerate(shown):
                self._legend_item(rect.left(), rect.top() + i * lh, sw, lh, e, rect.width() - sw - gap, fm)
            self.regions.append(Region("legend", None, rect))
            return inner.adjusted(0, 0, -(maxw + 8 * s), 0) if pos == "right" else inner.adjusted(maxw + 8 * s, 0, 0, 0)
        # top / bottom: one row
        widths = [min(fm.horizontalAdvance(e[0]), inner.width() / max(1, len(entries)) - sw - gap - 6 * s) + sw + gap + 10 * s
                  for e in entries]
        total = sum(widths) - 10 * s
        x = inner.center().x() - total / 2
        y = inner.top() if pos == "top" else inner.bottom() - lh
        rect = QRectF(max(inner.left(), x), y, min(total, inner.width()), lh)
        cx = rect.left()
        for e, w in zip(entries, widths):
            self._legend_item(cx, y, sw, lh, e, w - sw - gap - 10 * s, fm)
            cx += w
        self.regions.append(Region("legend", None, rect))
        if pos == "top":
            return inner.adjusted(0, lh + 4 * s, 0, 0)
        return inner.adjusted(0, 0, 0, -(lh + 4 * s))

    def _legend_item(self, x, y, sw, lh, e, text_w, fm):
        p, s = self.p, self.s
        name, color, kind = e[0], QColor(e[1]), e[2]
        marker = len(e) > 3 and e[3]
        cy = y + lh / 2
        if kind in ("line", "scatter"):
            p.setPen(QPen(color, 2.25 * s, Qt.SolidLine, Qt.RoundCap))
            p.drawLine(QPointF(x, cy), QPointF(x + sw, cy))
            if marker or kind == "scatter":
                p.setPen(Qt.NoPen)
                p.setBrush(color)
                p.drawEllipse(QPointF(x + sw / 2, cy), 3 * s, 3 * s)
        else:
            p.setPen(Qt.NoPen)
            p.setBrush(color)
            p.drawRect(QRectF(x, cy - sw / 2, sw, sw))
        self.p.setFont(self.font(12))
        self.p.setPen(self.txt)
        tr = QRectF(x + sw + 6 * s, y, max(text_w, 10), lh)
        self.p.drawText(tr, Qt.AlignLeft | Qt.AlignVCenter, fm.elidedText(name, Qt.ElideRight, tr.width()))

    # ================================================================ pie / doughnut
    def _pie(self, area):
        p, s = self.p, self.s
        ch = self.ch
        labels = ch.get("labels")
        n_ring = len(self.series) if self.fam == "doughnut" else 1
        size = min(area.width(), area.height()) * (0.80 if labels and self.fam == "pie" else 0.92)
        if size < 20:
            return
        cx, cy = area.center().x(), area.center().y()
        R = size / 2
        hole = ch.get("hole", 60) / 100.0 if self.fam == "doughnut" else 0.0
        ring_w = R * (1 - hole) / n_ring
        for ri in range(n_ring):
            sd = self.series[ri]
            vals = [max(0.0, v) if v is not None else 0.0 for v in sd["y"]]
            total = sum(vals)
            if total <= 0:
                continue
            r_out = R - ri * ring_w
            r_in = r_out - ring_w if self.fam == "doughnut" else 0.0
            ang = 90.0
            for j, v in enumerate(vals):
                span = 360.0 * v / total
                if span <= 0:
                    continue
                color = QColor(self.colors[j % len(self.colors)]) if not sd["src"].get("color") or j > 0 \
                    else QColor(sd["src"]["color"])
                path = QPainterPath()
                outer = QRectF(cx - r_out, cy - r_out, 2 * r_out, 2 * r_out)
                if r_in:
                    path.arcMoveTo(outer, ang)
                    path.arcTo(outer, ang, -span)
                    path.arcTo(QRectF(cx - r_in, cy - r_in, 2 * r_in, 2 * r_in), ang - span, span)
                else:
                    path.moveTo(cx, cy)
                    path.arcTo(outer, ang, -span)
                path.closeSubpath()
                p.setPen(QPen(QColor(self.fill), 1.5 * s))
                p.setBrush(color)
                p.drawPath(path)
                cat = self.data["cats"][j] if j < len(self.data["cats"]) else str(j + 1)
                share = v / total
                tip = f'Series "{sd["name"]}" Point "{cat}"\nValue: {self._val_text(v, sd)} ({share:.0%})'
                self.regions.append(Region("series", (ri, j), path, tip))
                if labels and span > 6:
                    mid = math.radians(ang - span / 2)
                    rl = (r_in + r_out) / 2 if self.fam == "doughnut" else r_out * 1.12
                    txt = f"{share:.0%}" if ch.get("label_pct") else self._val_text(v, sd)
                    fm = self.fm(12)
                    w = fm.horizontalAdvance(txt) + 6 * s
                    c_in = self.fam == "doughnut"
                    lx, ly = cx + rl * math.cos(mid), cy - rl * math.sin(mid)
                    box = QRectF(lx - w / 2 if c_in or math.cos(mid) >= -0.15 and abs(math.cos(mid)) < 0.15 else
                                 (lx if math.cos(mid) > 0 else lx - w), ly - fm.height() / 2, w, fm.height())
                    col = QColor("#FFFFFF") if c_in and C.is_dark(color.name()) else self.txt
                    self.text(box, txt, 12, Qt.AlignCenter, col)
                ang -= span

    # ================================================================ cartesian charts
    def _val_text(self, v, sd=None):
        fmt = self.ch.get("y_fmt") or (sd["fmt"] if sd else "General")
        return format_value(v, fmt)[0] if fmt != "General" else _fmt_general(v)

    def _cartesian(self, inner):
        p, s, ch = self.p, self.s, self.ch
        series = self.series
        horizontal = self.fam == "bar"
        scatter = self.fam == "scatter" or any(self.kind_of(sd)[0] == "scatter" for sd in series)
        npts = max(len(sd["y"]) for sd in series)
        cats = self.data["cats"]
        group = self.group if self.fam in ("col", "bar", "area") else "clustered"
        percent = group == "percent"

        # --- value ranges (primary / secondary)
        def collect(sec):
            sub = [(i, sd) for i, sd in enumerate(series) if bool(sd["src"].get("secondary")) == sec]
            if not sub:
                return None
            lo = hi = None
            if group in ("stacked", "percent") and not sec:
                pos = [0.0] * npts
                neg = [0.0] * npts
                for _, sd in sub:
                    for j, v in enumerate(sd["y"]):
                        if v is None:
                            continue
                        if v >= 0:
                            pos[j] += v
                        else:
                            neg[j] += v
                lo, hi = min(neg + [0.0]), max(pos + [0.0])
                if percent:
                    lo, hi = (-1.0 if lo < 0 else 0.0), 1.0
            else:
                for _, sd in sub:
                    for v in sd["y"]:
                        if v is None:
                            continue
                        lo = v if lo is None else min(lo, v)
                        hi = v if hi is None else max(hi, v)
            if lo is None:
                lo, hi = 0.0, 1.0
            # bars and areas always include zero (on a log axis they start at the first decade instead)
            if any(self.kind_of(sd)[0] in ("col", "bar", "area") for _, sd in sub) and not ch.get("log_y"):
                lo, hi = min(lo, 0.0), max(hi, 0.0)
            elif ch.get("log_y"):
                pos = [v for _, sd in sub for v in sd["y"] if v is not None and v > 0]
                lo, hi = (min(pos), max(pos)) if pos else (1.0, 10.0)
            return lo, hi, sub
        prim = collect(False)
        sec = collect(True)
        if prim is None:     # everything on the secondary axis
            prim, sec = sec, None

        def fmt_for(sub):
            if percent:
                return "0%"
            if ch.get("y_fmt"):
                return ch["y_fmt"]
            return sub[0][1]["fmt"] if sub else "General"
        pfmt = fmt_for(prim[2])
        sfmt = fmt_for(sec[2]) if sec else "General"

        def tick_text(v, fmt):
            if fmt == "General":
                return _fmt_general(v)
            return format_value(v, fmt)[0]

        label_px = 12
        fm = self.fm(label_px)
        fh = fm.height()
        log = bool(ch.get("log_y")) and prim[0] > 0
        # --- axis scales need the plot length; estimate it first, refine after labels are measured
        est_len = (inner.width() if horizontal else inner.height()) - 60 * s
        max_ticks = max(3, min(11, int(est_len / (30 * s))))
        manual = (ch.get("y_min"), ch.get("y_max"), ch.get("y_major"))
        if percent:
            manual = (0.0 if manual[0] is None else manual[0], 1.0 if manual[1] is None else manual[1],
                      0.1 if manual[2] is None else manual[2])
        pscale = nice_scale(prim[0], prim[1], max_ticks, *manual, zero=True, log=log)
        sscale = None
        if sec:
            sscale = nice_scale(sec[0], sec[1], max_ticks, ch.get("y2_min"), ch.get("y2_max"), None, zero=True)
        ptick = ticks_of(*pscale, log=log)
        ptext = [tick_text(v, pfmt) for v in ptick]
        stick = ticks_of(*sscale) if sscale else []
        stext = [tick_text(v, sfmt) for v in stick]

        # --- x axis (scatter: numeric)
        xscale = xticks = xtext = None
        if scatter:
            xs = [x for sd in series for x in (sd["x"] or []) if x is not None]
            lo, hi = (min(xs), max(xs)) if xs else (0.0, 1.0)
            xscale = nice_scale(lo, hi, max(3, min(11, int(inner.width() / (60 * s)))), ch.get("x_min"), ch.get("x_max"),
                                None, zero=lo >= 0)
            xticks = ticks_of(*xscale)
            xfmt = self.data.get("xfmt") or "General"
            xtext = [tick_text(v, xfmt) for v in xticks]

        # --- measure axis labels
        left = right = top = bottom = 0.0
        pad_lab = 6 * s
        cat_lab_w = 0.0
        if horizontal:
            cat_lab_w = min(max([fm.horizontalAdvance(c) for c in cats] or [0]), inner.width() * 0.28)
            left += cat_lab_w + pad_lab
            bottom += fh + pad_lab
        else:
            left += max([fm.horizontalAdvance(t) for t in ptext] or [0]) + pad_lab
            if sec:
                right += max([fm.horizontalAdvance(t) for t in stext] or [0]) + pad_lab
        top += fh / 2
        ytitle_w = (self.fm(12).height() + 6 * s)
        if ch.get("y_title") and not horizontal or ch.get("x_title") and horizontal:
            left += ytitle_w
        if ch.get("y2_title") and sec:
            right += ytitle_w
        if not horizontal:
            bottom += fh + pad_lab   # category labels (may grow below)
        if (ch.get("x_title") and not horizontal) or (ch.get("y_title") and horizontal):
            bottom += ytitle_w
        plot = QRectF(inner.left() + left, inner.top() + top, inner.width() - left - right - (0 if sec else 4 * s),
                      inner.height() - top - bottom)
        # x labels: rotation / thinning, which can make the bottom area taller
        rotate = False
        skip = 1
        label_list = [] if scatter or horizontal else cats[:npts]
        if label_list and plot.width() > 20:
            slot = plot.width() / max(1, len(label_list))
            maxw = min(max(fm.horizontalAdvance(c) for c in label_list), 200 * s)
            if maxw + 6 * s > slot:
                if slot >= fh * 0.95 and len(label_list) <= 60:
                    rotate = True
                    extra = maxw * 0.7071 + fh * 0.7071 - fh
                    extra = min(extra, inner.height() * 0.3)
                    plot.setHeight(plot.height() - extra)
                    bottom += extra
                    # the first label leans left of its tick: keep it inside the chart
                    lean = maxw * 0.7071 + fh * 0.35
                    shift = inner.left() - (plot.left() + slot / 2 - lean) + 2 * s
                    if shift > 0:
                        shift = min(shift, inner.width() * 0.3)
                        plot.setLeft(plot.left() + shift)
                else:
                    skip = int(math.ceil((maxw + 6 * s) / slot))
        if plot.width() < 20 or plot.height() < 20:
            return

        # --- value -> pixel
        pmin, pmax, pstep = pscale
        if log:
            lmin, lmax = math.log10(pmin), math.log10(pmax)

        def vy(v, scale=pscale, is_log=log):
            if is_log:
                v = math.log10(max(v, 1e-300))
                lo, hi = lmin, lmax
            else:
                lo, hi = scale[0], scale[1]
            return (v - lo) / (hi - lo)
        if horizontal:
            def pos_v(v, sc=pscale, lg=log):
                return plot.left() + vy(v, sc, lg) * plot.width()

            def pos_c_center(i):
                n = max(npts, 1)
                slot = plot.height() / n
                idx = (n - 1 - i) if not ch.get("x_reverse") else i
                return plot.top() + (idx + 0.5) * slot
        else:
            def pos_v(v, sc=pscale, lg=log):
                return plot.bottom() - vy(v, sc, lg) * plot.height()

            def pos_c_center(i):
                n = max(npts, 1)
                slot = plot.width() / n
                idx = i if not ch.get("x_reverse") else (n - 1 - i)
                return plot.left() + (idx + 0.5) * slot

        def pos_v2(v):
            return pos_v(v, sscale, False)

        # --- plot area background regions
        self.regions.append(Region("plot", None, QRectF(plot)))
        p.setRenderHint(QPainter.Antialiasing, False)
        gpen = QPen(self.grid, max(1.0, s * 0.75))
        # gridlines
        if ch.get("grid_y", True) and not scatter or (scatter and ch.get("grid_y", True)):
            p.setPen(gpen)
            for v in ptick:
                if horizontal:
                    x = round(pos_v(v)) + 0.5
                    p.drawLine(QPointF(x, plot.top()), QPointF(x, plot.bottom()))
                else:
                    y = round(pos_v(v)) + 0.5
                    p.drawLine(QPointF(plot.left(), y), QPointF(plot.right(), y))
        if ch.get("grid_x") or (scatter and ch.get("grid_x")):
            p.setPen(gpen)
            if scatter:
                for v in xticks:
                    x = round(plot.left() + (v - xscale[0]) / (xscale[1] - xscale[0]) * plot.width()) + 0.5
                    p.drawLine(QPointF(x, plot.top()), QPointF(x, plot.bottom()))
            else:
                n = max(npts, 1)
                for i in range(n + 1):
                    if horizontal:
                        y = round(plot.top() + plot.height() * i / n) + 0.5
                        p.drawLine(QPointF(plot.left(), y), QPointF(plot.right(), y))
                    else:
                        x = round(plot.left() + plot.width() * i / n) + 0.5
                        p.drawLine(QPointF(x, plot.top()), QPointF(x, plot.bottom()))
        p.setRenderHint(QPainter.Antialiasing, True)

        # --- draw series, by kind: areas first, then bars, then lines / scatter on top
        order = {"area": 0, "col": 1, "bar": 1, "line": 2, "scatter": 2}
        kinds = [self.kind_of(sd) for sd in series]
        draw_order = sorted(range(len(series)), key=lambda i: (order.get(kinds[i][0], 1), i))
        bar_series = [i for i in draw_order if kinds[i][0] in ("col", "bar")]
        stack_pos = [0.0] * npts
        stack_neg = [0.0] * npts
        tot_pos = [0.0] * npts
        if percent:
            for sd in series:
                for j, v in enumerate(sd["y"]):
                    tot_pos[j] += abs(v or 0.0)
        area_stack = [0.0] * npts
        labels_on = ch.get("labels")
        zero_v = max(pscale[0], min(0.0, pscale[1])) if not log else pscale[0]
        stacked = group in ("stacked", "percent")
        n_bar = len(bar_series)
        self._label_jobs = []
        # bar geometry
        if n_bar:
            slot = (plot.height() if horizontal else plot.width()) / max(npts, 1)
            gap = (ch.get("gap") or (150 if stacked else 219)) / 100.0
            if stacked or n_bar == 1:
                bw = slot / (1 + gap)
                step_in = 0
                n_eff = 1
            else:
                over = 0.27
                bw = slot / (n_bar + over * (n_bar - 1) + gap)
                step_in = bw * (1 + over)
                n_eff = n_bar
        for si in draw_order:
            sd = series[si]
            kind, markers, lines = kinds[si]
            color = self.color_of(si)
            secondary = bool(sd["src"].get("secondary")) and sec is not None
            posv = pos_v2 if secondary else pos_v
            ys = sd["y"]
            if kind == "area":
                self._draw_area(si, sd, color, plot, npts, ys, posv, pos_c_center, stacked, percent, area_stack, tot_pos,
                                zero_v, any(k[0] in ("col", "bar", "line") for k in kinds))
            elif kind in ("col", "bar"):
                bi = bar_series.index(si)
                for j, v in enumerate(ys):
                    if v is None:
                        continue
                    if stacked:
                        vv = v / tot_pos[j] if percent and tot_pos[j] else v
                        base = stack_pos[j] if vv >= 0 else stack_neg[j]
                        a, b = base, base + vv
                        if vv >= 0:
                            stack_pos[j] += vv
                        else:
                            stack_neg[j] += vv
                    else:
                        a, b = zero_v, v
                    c = pos_c_center(j)
                    off = (bi - (n_eff - 1) / 2.0) * step_in if not stacked and n_bar > 1 else 0.0
                    if horizontal:
                        x1, x2 = sorted((posv(a), posv(b)))
                        rect = QRectF(x1, c - bw / 2 + off, max(x2 - x1, 0.5), bw)
                    else:
                        y1, y2 = sorted((posv(a), posv(b)))
                        rect = QRectF(c - bw / 2 + off, y1, bw, max(y2 - y1, 0.5))
                    p.setPen(Qt.NoPen)
                    p.setBrush(color)
                    p.drawRect(rect)
                    cat = cats[j] if j < len(cats) else str(j + 1)
                    self.regions.append(Region("series", (si, j), QRectF(rect),
                                               f'Series "{sd["name"]}" Point "{cat}"\nValue: {self._val_text(v, sd)}'))
                    if labels_on if sd["src"].get("labels") is None else sd["src"]["labels"]:
                        self._label_jobs.append((rect, v, sd, "inside" if stacked else "end", horizontal))
            else:
                self._draw_line(si, sd, color, plot, ys, posv, pos_c_center, kind, markers, lines, scatter,
                                xscale, cats, secondary, sscale if secondary else pscale)
        # labels last so bars don't cover them
        for rect, v, sd, where, hor in self._label_jobs:
            self._bar_label(rect, v, sd, where, hor, pscale)

        # --- axes lines and labels
        p.setRenderHint(QPainter.Antialiasing, False)
        p.setPen(QPen(self.axis, max(1.0, s * 0.75)))
        zero_pos = pos_v(zero_v) if not log else pos_v(pscale[0])
        if horizontal:
            p.drawLine(QPointF(round(zero_pos) + 0.5, plot.top()), QPointF(round(zero_pos) + 0.5, plot.bottom()))
        else:
            p.drawLine(QPointF(plot.left(), round(zero_pos) + 0.5), QPointF(plot.right(), round(zero_pos) + 0.5))
        p.setRenderHint(QPainter.Antialiasing, True)
        self._axis_labels(plot, horizontal, scatter, cats, npts, rotate, skip, ptick, ptext, pos_v, pos_c_center,
                          sec, stick, stext, pos_v2, xticks, xtext, xscale, fh, pad_lab, left, right)
        self._axis_titles(plot, horizontal, sec, fh, pad_lab, left, right, inner)

    # ---- pieces
    def _axis_labels(self, plot, horizontal, scatter, cats, npts, rotate, skip, ptick, ptext, pos_v, pos_c_center,
                     sec, stick, stext, pos_v2, xticks, xtext, xscale, fh, pad_lab, left, right):
        p, s = self.p, self.s
        p.setFont(self.font(12))
        p.setPen(self.txt)
        fm = self.fm(12)
        if horizontal:
            for v, t in zip(ptick, ptext):
                x = pos_v(v)
                r = QRectF(x - 40 * s, plot.bottom() + 3 * s, 80 * s, fh)
                p.drawText(r, Qt.AlignHCenter | Qt.AlignTop, t)
            maxw = min(max([fm.horizontalAdvance(c) for c in cats] or [0]), self.R.width() * 0.28)
            step = max(1, int(math.ceil(fh * 1.1 / max(plot.height() / max(npts, 1), 1))))
            for j in range(0, npts, step):
                c = pos_c_center(j)
                txt = cats[j] if j < len(cats) else str(j + 1)
                r = QRectF(plot.left() - 6 * s - maxw, c - fh / 2, maxw, fh)
                p.drawText(r, Qt.AlignRight | Qt.AlignVCenter, fm.elidedText(txt, Qt.ElideRight, maxw))
            self.regions.append(Region("xaxis", None, QRectF(plot.left() - maxw - 8 * s, plot.top(), maxw + 8 * s, plot.height())))
            self.regions.append(Region("yaxis", None, QRectF(plot.left(), plot.bottom(), plot.width(), fh + 8 * s)))
        else:
            wmax = max([fm.horizontalAdvance(t) for t in ptext] or [0])
            for v, t in zip(ptick, ptext):
                y = pos_v(v)
                r = QRectF(plot.left() - 6 * s - wmax, y - fh / 2, wmax, fh)
                p.drawText(r, Qt.AlignRight | Qt.AlignVCenter, t)
            self.regions.append(Region("yaxis", None, QRectF(plot.left() - wmax - 8 * s, plot.top() - fh / 2,
                                                              wmax + 8 * s, plot.height() + fh)))
            if sec:
                for v, t in zip(stick, stext):
                    y = pos_v2(v)
                    r = QRectF(plot.right() + 6 * s, y - fh / 2, right, fh)
                    p.drawText(r, Qt.AlignLeft | Qt.AlignVCenter, t)
                self.regions.append(Region("y2axis", None, QRectF(plot.right(), plot.top() - fh / 2, right, plot.height() + fh)))
            # category / x labels
            if scatter:
                for v, t in zip(xticks, xtext):
                    x = plot.left() + (v - xscale[0]) / (xscale[1] - xscale[0]) * plot.width()
                    r = QRectF(x - 40 * s, plot.bottom() + 4 * s, 80 * s, fh)
                    p.drawText(r, Qt.AlignHCenter | Qt.AlignTop, t)
                self.regions.append(Region("xaxis", None, QRectF(plot.left(), plot.bottom(), plot.width(), fh + 8 * s)))
            else:
                slot = plot.width() / max(npts, 1)
                maxw = min(max([fm.horizontalAdvance(c) for c in cats[:npts]] or [0]), 200 * s)
                for j in range(0, npts, skip):
                    txt = cats[j] if j < len(cats) else str(j + 1)
                    cx = pos_c_center(j)
                    if rotate:
                        p.save()
                        p.translate(cx, plot.bottom() + 5 * s)
                        p.rotate(-45)
                        p.drawText(QRectF(-maxw, -fh / 2, maxw, fh), Qt.AlignRight | Qt.AlignVCenter,
                                   fm.elidedText(txt, Qt.ElideRight, maxw))
                        p.restore()
                    else:
                        w = slot * skip
                        r = QRectF(cx - w / 2, plot.bottom() + 4 * s, w, fh)
                        p.drawText(r, Qt.AlignHCenter | Qt.AlignTop, fm.elidedText(txt, Qt.ElideRight, w))
                h = (maxw * 0.7071 + fh) if rotate else fh
                self.regions.append(Region("xaxis", None, QRectF(plot.left(), plot.bottom(), plot.width(), h + 8 * s)))

    def _axis_titles(self, plot, horizontal, sec, fh, pad_lab, left, right, inner):
        ch, p, s = self.ch, self.p, self.s
        fm = self.fm(12)
        th = fm.height() + 4 * s
        x_t = ch.get("x_title") or ""
        y_t = ch.get("y_title") or ""
        y2_t = ch.get("y2_title") or ""
        # horizontal-axis title sits at the bottom of the chart (inner), vertical one at the left
        bottom_title = y_t if horizontal else x_t
        left_title = x_t if horizontal else y_t
        key_bottom, key_left = ("ytitle", "xtitle") if horizontal else ("xtitle", "ytitle")
        if bottom_title:
            r = QRectF(plot.left(), inner.bottom() - th, plot.width(), th)
            self.text(r, bottom_title, 12, Qt.AlignCenter)
            tw = min(plot.width(), fm.horizontalAdvance(bottom_title) + 10 * s)
            self.regions.append(Region(key_bottom, None, QRectF(r.center().x() - tw / 2, r.top(), tw, th)))
        if left_title:
            p.save()
            p.translate(inner.left() + th / 2, plot.center().y())
            p.rotate(-90)
            r = QRectF(-plot.height() / 2, -th / 2, plot.height(), th)
            self.text(r, left_title, 12, Qt.AlignCenter)
            p.restore()
            tw = min(plot.height(), fm.horizontalAdvance(left_title) + 10 * s)
            self.regions.append(Region(key_left, None, QRectF(inner.left(), plot.center().y() - tw / 2, th, tw)))
        if y2_t and sec:
            p.save()
            p.translate(inner.right() - th / 2 + 2 * s, plot.center().y())
            p.rotate(90)
            r = QRectF(-plot.height() / 2, -th / 2, plot.height(), th)
            self.text(r, y2_t, 12, Qt.AlignCenter)
            p.restore()
            tw = min(plot.height(), fm.horizontalAdvance(y2_t) + 10 * s)
            self.regions.append(Region("y2title", None, QRectF(inner.right() - th, plot.center().y() - tw / 2, th, tw)))

    def _bar_label(self, rect, v, sd, where, horizontal, pscale):
        s = self.s
        txt = self._val_text(v, sd)
        fm = self.fm(12)
        w, h = fm.horizontalAdvance(txt) + 4 * s, fm.height()
        if where == "inside":
            box = QRectF(rect.center().x() - w / 2, rect.center().y() - h / 2, w, h)
            col = QColor("#FFFFFF") if C.is_dark(self.color_of(sd["i"]).name()) else QColor("#404040")
        else:
            col = self.txt
            if horizontal:
                box = QRectF(rect.right() + 2 * s, rect.center().y() - h / 2, w, h) if v >= 0 else \
                    QRectF(rect.left() - w - 2 * s, rect.center().y() - h / 2, w, h)
            else:
                box = QRectF(rect.center().x() - w / 2, rect.top() - h - 1 * s, w, h) if v >= 0 else \
                    QRectF(rect.center().x() - w / 2, rect.bottom() + 1 * s, w, h)
        self.text(box, txt, 12, Qt.AlignCenter, col)

    def _draw_area(self, si, sd, color, plot, npts, ys, posv, pos_c, stacked, percent, acc, tot, zero_v, between):
        n = npts
        if n < 1:
            return
        # on tick marks (like Excel) unless columns / lines share the axis
        def px(j):
            if between:
                return pos_c(j)
            if n == 1:
                return plot.center().x()
            x = plot.left() + plot.width() * j / (n - 1)
            return plot.right() - (x - plot.left()) if self.ch.get("x_reverse") else x
        top, base = [], []
        for j in range(n):
            v = ys[j] if j < len(ys) and ys[j] is not None else 0.0
            if percent:
                v = v / tot[j] if tot[j] else 0.0
            b = acc[j] if stacked else zero_v
            t = b + v
            if stacked:
                acc[j] = t
            top.append(QPointF(px(j), posv(t)))
            base.append(QPointF(px(j), posv(b)))
        poly = QPolygonF(top + list(reversed(base)))
        fill = QColor(color)
        p = self.p
        p.setPen(Qt.NoPen)
        p.setBrush(fill)
        p.drawPolygon(poly)
        p.setPen(QPen(color.darker(105), 1.5 * self.s))
        p.setBrush(Qt.NoBrush)
        p.drawPolyline(QPolygonF(top))
        path = QPainterPath()
        path.addPolygon(poly)
        path.closeSubpath()
        self.regions.append(Region("series", (si, -1), path, f'Series "{sd["name"]}"'))

    def _draw_line(self, si, sd, color, plot, ys, posv, pos_c, kind, markers, lines, scatter, xscale, cats, secondary,
                   scale):
        p, s = self.p, self.s
        pts = []
        xs = sd["x"]
        for j, v in enumerate(ys):
            if v is None:
                pts.append(None)
                continue
            if kind == "scatter" or scatter:
                xv = xs[j] if xs and j < len(xs) and xs[j] is not None else j + 1.0
                x = plot.left() + (xv - xscale[0]) / (xscale[1] - xscale[0]) * plot.width()
            else:
                x = pos_c(j)
            pts.append(QPointF(x, posv(v)))
        draw_lines = lines or (kind == "line")
        sdef = sd["src"]
        if kind == "scatter":
            t = sdef.get("type") if self.fam == "combo" else self.ch["type"]
            draw_lines = t in ("scatter_lines", "scatter_smooth") or bool(sdef.get("smooth"))
        smooth = bool(sdef.get("smooth"))
        if sdef.get("marker") is not None:
            markers = bool(sdef["marker"])
        if kind == "scatter":
            markers = True if sdef.get("marker") is None else bool(sdef["marker"])
        if draw_lines:
            p.setPen(QPen(color, 2.25 * s * (4 / 3), Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin))
            p.setBrush(Qt.NoBrush)
            seg = []
            for pt in pts + [None]:
                if pt is None:
                    if len(seg) > 1:
                        self._poly(seg, smooth)
                    seg = []
                else:
                    seg.append(pt)
        if markers:
            d = 7 * s
            p.setPen(QPen(color, 1 * s))
            p.setBrush(color)
            for j, pt in enumerate(pts):
                if pt is not None:
                    p.drawEllipse(pt, d / 2, d / 2)
        labels_on = self.ch.get("labels") if sdef.get("labels") is None else sdef["labels"]
        for j, pt in enumerate(pts):
            if pt is None:
                continue
            cat = cats[j] if j < len(cats) else str(j + 1)
            r = QRectF(pt.x() - 5 * s, pt.y() - 5 * s, 10 * s, 10 * s)
            tip = f'Series "{sd["name"]}" Point "{cat}"\nValue: {self._val_text(ys[j], sd)}'
            if kind == "scatter" and sd["x"]:
                tip = f'Series "{sd["name"]}"\nX: {_fmt_general(sd["x"][j])}  Y: {self._val_text(ys[j], sd)}'
            self.regions.append(Region("series", (si, j), r, tip))
            if labels_on:
                txt = self._val_text(ys[j], sd)
                fm = self.fm(12)
                w = fm.horizontalAdvance(txt) + 4 * s
                self.text(QRectF(pt.x() - w / 2, pt.y() - fm.height() - 4 * s, w, fm.height()), txt, 12, Qt.AlignCenter)

    def _poly(self, pts, smooth):
        path = QPainterPath(pts[0])
        if smooth and len(pts) > 2:
            for i in range(len(pts) - 1):
                p0 = pts[i - 1] if i > 0 else pts[i]
                p1, p2 = pts[i], pts[i + 1]
                p3 = pts[i + 2] if i + 2 < len(pts) else p2
                c1 = QPointF(p1.x() + (p2.x() - p0.x()) / 6, p1.y() + (p2.y() - p0.y()) / 6)
                c2 = QPointF(p2.x() - (p3.x() - p1.x()) / 6, p2.y() - (p3.y() - p1.y()) / 6)
                path.cubicTo(c1, c2, p2)
        else:
            for pt in pts[1:]:
                path.lineTo(pt)
        self.p.drawPath(path)

    # ---- selection outline
    def _selection(self):
        sel = self.sel
        if not sel:
            return
        kind, key = sel
        if kind == "chart":
            return
        p = self.p
        pen = QPen(QColor("#7F7F7F"), 1, Qt.DashLine)
        p.setBrush(Qt.NoBrush)
        p.setPen(pen)
        if kind == "series":
            si = key[0] if isinstance(key, tuple) else key
            for reg in self.regions:
                if reg.kind == "series" and reg.key[0] == si:
                    if isinstance(reg.shape, QRectF):
                        p.drawRect(reg.shape)
                    else:
                        p.drawPath(reg.shape)
            return
        for reg in self.regions:
            if reg.kind == kind and (key is None or reg.key == key):
                b = reg.bounds().adjusted(-1, -1, 1, 1)
                p.drawRect(b)
                p.setPen(QPen(QColor("#7F7F7F")))
                p.setBrush(QColor("#FFFFFF"))
                h = 3.5
                for pt in (b.topLeft(), b.topRight(), b.bottomLeft(), b.bottomRight()):
                    p.drawEllipse(pt, h, h)
                return


def paint_chart(painter, rect, ch, data, scale=1.0, selected=None):
    return _Painter(painter, rect, ch, data, scale, selected).paint()


def render_image(ch, data, w, h, scale=1.0):
    from PySide6.QtGui import QImage
    img = QImage(int(w), int(h), QImage.Format_ARGB32)
    img.fill(Qt.white)
    p = QPainter(img)
    paint_chart(p, QRectF(0, 0, w, h), ch, data, scale)
    p.end()
    return img
