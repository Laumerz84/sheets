"""Colors, fonts, palette and toolbar icons."""
from PySide6.QtCore import QPointF, QRectF, Qt
from PySide6.QtGui import (QColor, QFont, QFontDatabase, QIcon, QIconEngine, QPainter,
                           QPainterPath, QPalette, QPen, QPixmap)

ACCENT = QColor("#217346")          # Excel green
ACCENT_DARK = QColor("#185C37")
ACCENT_LIGHT = QColor("#E1EFE6")
GRID_LINE = QColor("#E1E1E1")
HEADER_BG = QColor("#F5F5F5")
HEADER_LINE = QColor("#D4D4D4")
HEADER_TEXT = QColor("#444444")
HEADER_SEL_BG = QColor("#E1E1E1")
HEADER_FULL_BG = QColor("#D3E5DA")
SEL_FILL = QColor(33, 115, 70, 30)
CELL_BG = QColor("#FFFFFF")
FROZEN_LINE = QColor("#A6A6A6")

REF_COLORS = ["#2E64C8", "#C42B1C", "#7B3FB4", "#107C10", "#C46200", "#0099BC", "#B4009E"]

DEFAULT_FONT_FAMILY = "Calibri"   # the workbook's default font (what files are saved with)
DEFAULT_FONT_SIZE = 11.0
DISPLAY_FONT_FAMILY = DEFAULT_FONT_FAMILY  # what unformatted cells are drawn in (the skin may change it)
DISPLAY_FONT_SIZE = DEFAULT_FONT_SIZE
HEADER_FONT_FAMILY = "Segoe UI"
HEADER_BEVEL = False  # Windows 95 raised-button row/column headers

# ---------------------------------------------------------------- skins (View > Skin)
# The skin swaps these module values; the grid and the icons read them when they paint.
SKIN = "modern"
_SKIN_KEYS = ("ACCENT", "ACCENT_DARK", "ACCENT_LIGHT", "GRID_LINE", "HEADER_BG", "HEADER_LINE", "HEADER_TEXT",
              "HEADER_SEL_BG", "HEADER_FULL_BG", "SEL_FILL", "CELL_BG", "FROZEN_LINE", "DISPLAY_FONT_FAMILY",
              "DISPLAY_FONT_SIZE", "HEADER_FONT_FAMILY", "HEADER_BEVEL")
_MODERN = {k: globals()[k] for k in _SKIN_KEYS}
_WIN95 = {
    "ACCENT": QColor("#000000"), "ACCENT_DARK": QColor("#000000"), "ACCENT_LIGHT": QColor("#C0C0C0"),
    "GRID_LINE": QColor("#C0C0C0"), "HEADER_BG": QColor("#C0C0C0"), "HEADER_LINE": QColor("#808080"),
    "HEADER_TEXT": QColor("#000000"), "HEADER_SEL_BG": QColor("#C0C0C0"), "HEADER_FULL_BG": QColor("#C0C0C0"),
    "SEL_FILL": QColor(0, 0, 0, 70), "CELL_BG": QColor("#FFFFFF"), "FROZEN_LINE": QColor("#000000"),
    "DISPLAY_FONT_FAMILY": "Arial", "DISPLAY_FONT_SIZE": 10.0, "HEADER_FONT_FAMILY": "Microsoft Sans Serif",
    "HEADER_BEVEL": True,
}
SKINS = {"modern": "Modern", "excel95": "Excel 95"}


def set_skin(name):
    global SKIN
    SKIN = name if name in SKINS else "modern"
    globals().update(_WIN95 if SKIN == "excel95" else _MODERN)

_icon_font_family = None


def icon_font_family():
    global _icon_font_family
    if _icon_font_family is None:
        fams = set(QFontDatabase.families())
        for f in ("Segoe Fluent Icons", "Segoe MDL2 Assets"):
            if f in fams:
                _icon_font_family = f
                break
        else:
            _icon_font_family = ""
    return _icon_font_family


def apply_palette(app):
    app.setStyle("Fusion")
    pal = QPalette()
    pal.setColor(QPalette.Window, QColor("#F3F3F3"))
    pal.setColor(QPalette.WindowText, QColor("#202020"))
    pal.setColor(QPalette.Base, QColor("#FFFFFF"))
    pal.setColor(QPalette.AlternateBase, QColor("#F7F7F7"))
    pal.setColor(QPalette.Text, QColor("#202020"))
    pal.setColor(QPalette.Button, QColor("#F7F7F7"))
    pal.setColor(QPalette.ButtonText, QColor("#202020"))
    pal.setColor(QPalette.Highlight, ACCENT)
    pal.setColor(QPalette.HighlightedText, QColor("#FFFFFF"))
    pal.setColor(QPalette.ToolTipBase, QColor("#FFFFE1"))
    pal.setColor(QPalette.ToolTipText, QColor("#202020"))
    pal.setColor(QPalette.PlaceholderText, QColor("#8A8A8A"))
    pal.setColor(QPalette.Link, QColor("#0563C1"))
    pal.setColor(QPalette.Disabled, QPalette.Text, QColor("#A0A0A0"))
    pal.setColor(QPalette.Disabled, QPalette.ButtonText, QColor("#A0A0A0"))
    pal.setColor(QPalette.Disabled, QPalette.WindowText, QColor("#A0A0A0"))
    app.setPalette(pal)
    f = QFont("Segoe UI", 9)
    app.setFont(f)
    app.setStyleSheet("""
        QToolBar { background: #FFFFFF; border: none; border-bottom: 1px solid #D4D4D4; spacing: 2px; padding: 3px 6px; }
        QToolBar::separator { background: #D4D4D4; width: 1px; margin: 4px 6px; }
        QToolButton { border: 1px solid transparent; border-radius: 3px; padding: 3px; }
        QToolButton:hover { background: #E8F2EC; border-color: #C5DED0; }
        QToolButton:checked { background: #CFE5D7; border-color: #9CC5AC; }
        QToolButton:pressed { background: #BCDCC8; }
        QToolButton[popupMode="1"] { padding-right: 14px; }
        QMenuBar { background: #217346; color: white; padding: 2px; }
        QMenuBar::item { padding: 4px 10px; background: transparent; border-radius: 3px; }
        QMenuBar::item:selected { background: #185C37; }
        QMenu { background: #FFFFFF; border: 1px solid #C8C8C8; padding: 4px 0; }
        QMenu::item { padding: 5px 28px 5px 28px; }
        QMenu::item:selected { background: #E1EFE6; color: #000; }
        QMenu::item:disabled { color: #A0A0A0; }
        QMenu::separator { height: 1px; background: #E1E1E1; margin: 4px 8px; }
        QStatusBar { background: #F3F3F3; border-top: 1px solid #D4D4D4; }
        QStatusBar QLabel { padding: 0 6px; color: #333; }
        QTabBar::tab { background: transparent; border: none; padding: 5px 16px; margin: 0; color: #333; }
        QTabBar::tab:selected { background: #FFFFFF; color: #217346; font-weight: bold;
                                border-bottom: 3px solid #217346; }
        QTabBar::tab:hover:!selected { background: #E8E8E8; }
        QComboBox { padding: 2px 6px; border: 1px solid #C8C8C8; border-radius: 2px; background: white; }
        QComboBox:hover { border-color: #217346; }
        QLineEdit, QPlainTextEdit, QSpinBox, QDoubleSpinBox { border: 1px solid #C8C8C8; border-radius: 2px; background: white;
                    selection-background-color: #217346; }
        QPushButton { padding: 5px 16px; border: 1px solid #ABABAB; border-radius: 3px; background: #FDFDFD; }
        QPushButton:hover { background: #E8F2EC; border-color: #217346; }
        QPushButton:default { border: 1px solid #217346; }
        QGroupBox { font-weight: bold; border: 1px solid #D4D4D4; border-radius: 3px; margin-top: 10px; padding-top: 6px; }
        QGroupBox::title { subcontrol-origin: margin; left: 8px; padding: 0 3px; }
        QScrollBar:vertical { background: #F3F3F3; width: 14px; margin: 0; }
        QScrollBar:horizontal { background: #F3F3F3; height: 14px; margin: 0; }
        QScrollBar::handle { background: #C2C2C2; border-radius: 3px; min-height: 24px; min-width: 24px; margin: 3px; }
        QScrollBar::handle:hover { background: #A0A0A0; }
        QScrollBar::add-line, QScrollBar::sub-line { width: 0; height: 0; }
        QScrollBar::add-page, QScrollBar::sub-page { background: none; }
    """)


def apply_palette_95(app):
    """Excel 95 / Windows 95: battleship grey, raised and sunken bevels, MS Sans Serif, navy highlights.
    Qt's "Windows" style draws the classic 3D buttons, scrollbars and checkboxes; the stylesheet only
    covers what it doesn't (toolbars, menus, edits) and leaves scrollbars and push buttons to it."""
    app.setStyle("Windows")
    grey, white, dark, black, navy = (QColor(c) for c in ("#C0C0C0", "#FFFFFF", "#808080", "#000000", "#000080"))
    pal = QPalette()
    for role, c in ((QPalette.Window, grey), (QPalette.WindowText, black), (QPalette.Base, white),
                    (QPalette.AlternateBase, grey), (QPalette.Text, black), (QPalette.Button, grey),
                    (QPalette.ButtonText, black), (QPalette.Light, white), (QPalette.Midlight, QColor("#DFDFDF")),
                    (QPalette.Mid, QColor("#A0A0A0")), (QPalette.Dark, dark), (QPalette.Shadow, black),
                    (QPalette.Highlight, navy), (QPalette.HighlightedText, white),
                    (QPalette.ToolTipBase, QColor("#FFFFE1")), (QPalette.ToolTipText, black),
                    (QPalette.PlaceholderText, dark), (QPalette.Link, navy)):
        pal.setColor(role, c)
    for role in (QPalette.Text, QPalette.ButtonText, QPalette.WindowText):
        pal.setColor(QPalette.Disabled, role, dark)
    app.setPalette(pal)
    f = QFont("Microsoft Sans Serif", 8)
    f.setStyleStrategy(QFont.NoAntialias)  # crisp, un-smoothed text like 1995
    app.setFont(f)
    raised = "border: 1px solid; border-color: #FFFFFF #808080 #808080 #FFFFFF;"
    sunken = "border: 1px solid; border-color: #808080 #FFFFFF #FFFFFF #808080;"
    app.setStyleSheet(f"""
        QToolBar {{ background: #C0C0C0; border: none; border-top: 1px solid #FFFFFF;
                    border-bottom: 1px solid #808080; spacing: 1px; padding: 2px 3px; }}
        QToolBar::separator {{ background: #808080; width: 1px; margin: 3px 4px; }}
        QToolButton {{ border: 1px solid transparent; padding: 2px; background: #C0C0C0; }}
        QToolButton:hover {{ {raised} }}
        QToolButton:checked {{ {sunken} background: #DFDFDF; }}
        QToolButton:pressed {{ {sunken} }}
        QToolButton[popupMode="1"] {{ padding-right: 12px; }}
        QMenuBar {{ background: #C0C0C0; color: #000000; padding: 1px; }}
        QMenuBar::item {{ padding: 3px 8px; background: transparent; }}
        QMenuBar::item:selected {{ background: #000080; color: #FFFFFF; }}
        QMenu {{ background: #C0C0C0; border: 2px solid; border-color: #DFDFDF #000000 #000000 #DFDFDF; padding: 2px; }}
        QMenu::item {{ padding: 3px 24px 3px 24px; color: #000000; }}
        QMenu::item:selected {{ background: #000080; color: #FFFFFF; }}
        QMenu::item:disabled {{ color: #808080; }}
        QMenu::separator {{ height: 2px; border-top: 1px solid #808080; border-bottom: 1px solid #FFFFFF; margin: 3px 2px; }}
        QStatusBar {{ background: #C0C0C0; border-top: 1px solid #FFFFFF; }}
        QStatusBar QLabel {{ padding: 0 6px; color: #000000; {sunken} }}
        QTabBar::tab {{ background: transparent; border: none; padding: 3px 18px; margin: 0; color: #000000; }}
        QComboBox {{ padding: 1px 4px; background: #FFFFFF; {sunken} }}
        QLineEdit, QPlainTextEdit, QSpinBox, QDoubleSpinBox {{ background: #FFFFFF; {sunken}
                    selection-background-color: #000080; selection-color: #FFFFFF; }}
        QGroupBox {{ border: 1px solid #808080; margin-top: 10px; padding-top: 6px; }}
        QGroupBox::title {{ subcontrol-origin: margin; left: 8px; padding: 0 3px; }}
    """)


# ---------------------------------------------------------------- icons

ICON_SIZE = 20


class _SkinIcon(QIconEngine):
    """An icon that draws the current skin's version each time it's painted: the modern drawing,
    or the Excel 95 pixel icon for the same key (when pixel95 has one)."""

    def __init__(self, key, modern):
        super().__init__()
        self.key, self.modern = key, modern
        self._base, self._made = {}, {}

    def clone(self):
        return _SkinIcon(self.key, self.modern)

    def _source(self):
        """(pixmap, is_pixel_art) for the current skin."""
        if SKIN not in self._base:
            pm = None
            if SKIN == "excel95":
                from . import pixel95
                pm = pixel95.render(self.key)
            self._base[SKIN] = (pm, True) if pm is not None else (self.modern(), False)
        return self._base[SKIN]

    def _make(self, size, mode, scale):
        key = (SKIN, size.width(), size.height(), getattr(mode, "value", mode), round(scale, 2))
        if key in self._made:
            return self._made[key]
        src, pixel = self._source()
        w, h = max(1, round(size.width() * scale)), max(1, round(size.height() * scale))
        out = QPixmap(w, h)
        out.fill(Qt.transparent)
        p = QPainter(out)
        if pixel:  # whole-pixel scaling only, so pixel art stays crisp
            n = max(1, min(w, h) // src.width())
            art = src.scaled(src.width() * n, src.height() * n, Qt.KeepAspectRatio, Qt.FastTransformation)
            if mode == QIcon.Disabled:  # Win95 "etched" look: white shadow under a grey silhouette
                art_w, art_g = _silhouette(art, QColor("#FFFFFF")), _silhouette(art, QColor("#808080"))
                x, y = (w - art.width()) // 2, (h - art.height()) // 2
                p.drawPixmap(x + n, y + n, art_w)
                p.drawPixmap(x, y, art_g)
            else:
                p.drawPixmap((w - art.width()) // 2, (h - art.height()) // 2, art)
        else:
            p.setRenderHint(QPainter.SmoothPixmapTransform)
            if mode == QIcon.Disabled:
                p.setOpacity(0.35)
            p.drawPixmap(QRectF(0, 0, w, h), src, QRectF(src.rect()))
        p.end()
        out.setDevicePixelRatio(scale)
        self._made[key] = out
        return out

    def pixmap(self, size, mode, state):
        return self._make(size, mode, 1.0)

    def scaledPixmap(self, size, mode, state, scale):
        return self._make(size, mode, scale)

    def paint(self, painter, rect, mode, state):
        dev = painter.device()
        scale = dev.devicePixelRatioF() if dev is not None else 1.0
        painter.drawPixmap(rect, self._make(rect.size(), mode, scale))


def _silhouette(pm, color):
    out = QPixmap(pm.size())
    out.fill(Qt.transparent)
    p = QPainter(out)
    p.drawPixmap(0, 0, pm)
    p.setCompositionMode(QPainter.CompositionMode_SourceIn)
    p.fillRect(out.rect(), color)
    p.end()
    return out


def glyph_icon(code, color="#333333", size=ICON_SIZE, bar=None, fallback=None):
    """Icon from a Segoe Fluent/MDL2 glyph; `bar` adds a colored strip at the bottom."""
    return QIcon(_SkinIcon(("glyph", code, bar, fallback),
                           lambda: _glyph_pm(code, color, size, bar, fallback)))


def text_icon(text, bold=False, italic=False, underline=False, strike=False, color="#333333",
              size=ICON_SIZE, family="Segoe UI", px=None, bar=None):
    return QIcon(_SkinIcon(("text", text, bold, italic, underline, strike, color, bar),
                           lambda: _text_pm(text, bold, italic, underline, strike, color, size, family, px, bar)))


def lines_icon(kind, size=ICON_SIZE, color="#333333"):
    """Alignment / wrap / merge style icons drawn with lines."""
    return QIcon(_SkinIcon(("lines", kind), lambda: _lines_pm(kind, size, color)))


def _canvas(size=ICON_SIZE):
    pm = QPixmap(size * 2, size * 2)
    pm.setDevicePixelRatio(2.0)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    p.setRenderHint(QPainter.TextAntialiasing)
    return pm, p


def _glyph_pm(code, color="#333333", size=ICON_SIZE, bar=None, fallback=None):
    """Icon from a Segoe Fluent/MDL2 glyph; `bar` adds a colored strip at the bottom."""
    pm, p = _canvas(size)
    fam = icon_font_family()
    if fam and code:
        f = QFont(fam)
        f.setPixelSize(int(size * 0.8))
        p.setFont(f)
        p.setPen(QColor(color))
        rect = QRectF(0, 0, size, size - (4 if bar else 0))
        p.drawText(rect, Qt.AlignCenter, code)
    elif fallback:
        f = QFont("Segoe UI", 1)
        f.setPixelSize(int(size * 0.65))
        f.setBold(True)
        p.setFont(f)
        p.setPen(QColor(color))
        p.drawText(QRectF(0, 0, size, size - (4 if bar else 0)), Qt.AlignCenter, fallback)
    if bar:
        p.fillRect(QRectF(2, size - 4, size - 4, 3.5), QColor(bar))
    p.end()
    return pm


def _text_pm(text, bold=False, italic=False, underline=False, strike=False, color="#333333",
              size=ICON_SIZE, family="Segoe UI", px=None, bar=None):
    pm, p = _canvas(size)
    f = QFont(family)
    f.setPixelSize(px or int(size * 0.7))
    f.setBold(bold)
    f.setItalic(italic)
    f.setUnderline(underline)
    f.setStrikeOut(strike)
    p.setFont(f)
    p.setPen(QColor(color))
    p.drawText(QRectF(0, -1, size, size - (3 if bar else 0)), Qt.AlignCenter, text)
    if bar:
        p.fillRect(QRectF(2, size - 4, size - 4, 3.5), QColor(bar))
    p.end()
    return pm


def _lines_pm(kind, size=ICON_SIZE, color="#333333"):
    """Alignment / wrap / merge style icons drawn with lines."""
    pm, p = _canvas(size)
    pen = QPen(QColor(color), 1.6)
    pen.setCapStyle(Qt.RoundCap)
    p.setPen(pen)
    m = 3.5
    w = size - 2 * m
    if kind in ("left", "center", "right"):
        lens = [w, w * 0.6, w, w * 0.6]
        for i, ln in enumerate(lens):
            y = 5 + i * 3.4
            if kind == "left":
                x1 = m
            elif kind == "right":
                x1 = m + w - ln
            else:
                x1 = m + (w - ln) / 2
            p.drawLine(QPointF(x1, y), QPointF(x1 + ln, y))
    elif kind in ("top", "middle", "bottom"):
        p.setPen(QPen(QColor("#9A9A9A"), 1.0))
        p.drawRect(QRectF(m, m, w, w))
        p.setPen(pen)
        ys = {"top": [6.5, 9.5], "middle": [8.5, 11.5], "bottom": [10.5, 13.5]}[kind]
        for y in ys:
            p.drawLine(QPointF(m + 3, y), QPointF(m + w - 3, y))
    elif kind == "wrap":
        p.drawLine(QPointF(m, 5.5), QPointF(m + w, 5.5))
        path = QPainterPath(QPointF(m, 10))
        path.lineTo(m + w - 3, 10)
        path.arcTo(QRectF(m + w - 6, 10, 6, 5), 90, -180)
        path.lineTo(m + 6, 15)
        p.drawPath(path)
        p.drawLine(QPointF(m + 6, 15), QPointF(m + 8.5, 12.8))
        p.drawLine(QPointF(m + 6, 15), QPointF(m + 8.5, 17.2))
    elif kind == "merge":
        p.setPen(QPen(QColor("#9A9A9A"), 1.0))
        p.drawRect(QRectF(m - 1, m + 1, w + 2, w - 2))
        p.setPen(pen)
        y = size / 2
        p.drawLine(QPointF(m + 2, y), QPointF(m + w - 2, y))
        p.drawLine(QPointF(m + 2, y), QPointF(m + 5, y - 2.5))
        p.drawLine(QPointF(m + 2, y), QPointF(m + 5, y + 2.5))
        p.drawLine(QPointF(m + w - 2, y), QPointF(m + w - 5, y - 2.5))
        p.drawLine(QPointF(m + w - 2, y), QPointF(m + w - 5, y + 2.5))
    elif kind == "borders":
        p.setPen(QPen(QColor("#9A9A9A"), 1.0, Qt.DotLine))
        p.drawLine(QPointF(size / 2, m), QPointF(size / 2, m + w))
        p.drawLine(QPointF(m, size / 2), QPointF(m + w, size / 2))
        p.setPen(QPen(QColor(color), 1.4))
        p.drawRect(QRectF(m, m, w, w))
    elif kind == "freeze":
        p.setPen(QPen(QColor("#9A9A9A"), 1.0))
        p.drawRect(QRectF(m, m, w, w))
        p.setPen(QPen(ACCENT, 2.0))
        p.drawLine(QPointF(m, 8), QPointF(m + w, 8))
        p.drawLine(QPointF(8, m), QPointF(8, m + w))
    elif kind == "dec_inc" or kind == "dec_dec":
        f = QFont("Segoe UI")
        f.setPixelSize(8)
        f.setBold(True)
        p.setFont(f)
        p.drawText(QRectF(0, 1, size, 9), Qt.AlignCenter, ".0" if kind == "dec_dec" else ".00")
        p.drawText(QRectF(0, 10, size, 9), Qt.AlignCenter, ".00" if kind == "dec_dec" else ".0")
        arrow_y = 9.5
        p.drawLine(QPointF(2, arrow_y), QPointF(6, arrow_y))
    elif kind == "insert_row" or kind == "delete_row":
        p.setPen(QPen(QColor("#9A9A9A"), 1.0))
        for i in range(3):
            p.drawRect(QRectF(m, m + i * 4.5, w, 4.5))
        col = ACCENT if kind == "insert_row" else QColor("#C42B1C")
        p.fillRect(QRectF(m, m + 4.5, w, 4.5), col)
    elif kind == "sort_az" or kind == "sort_za":
        f = QFont("Segoe UI")
        f.setPixelSize(8)
        f.setBold(True)
        p.setFont(f)
        top, bot = ("A", "Z") if kind == "sort_az" else ("Z", "A")
        p.drawText(QRectF(1, 0, 10, 10), Qt.AlignCenter, top)
        p.drawText(QRectF(1, 10, 10, 10), Qt.AlignCenter, bot)
        p.drawLine(QPointF(15, 3), QPointF(15, 17))
        p.drawLine(QPointF(15, 17), QPointF(12.5, 14))
        p.drawLine(QPointF(15, 17), QPointF(17.5, 14))
    elif kind == "filter":
        path = QPainterPath(QPointF(3.5, 4))
        path.lineTo(16.5, 4)
        path.lineTo(11.5, 10)
        path.lineTo(11.5, 16)
        path.lineTo(8.5, 14.5)
        path.lineTo(8.5, 10)
        path.closeSubpath()
        p.drawPath(path)
    p.end()
    return pm


def color_square_icon(color, size=14):
    pm = QPixmap(size, size)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    if color:
        p.fillRect(1, 1, size - 2, size - 2, QColor(color))
        p.setPen(QColor("#888888"))
        p.drawRect(1, 1, size - 3, size - 3)
    else:
        p.setPen(QPen(QColor("#C42B1C"), 1.5))
        p.drawRect(1, 1, size - 3, size - 3)
        p.drawLine(2, size - 2, size - 2, 2)
    p.end()
    return QIcon(pm)


# Segoe Fluent / MDL2 code points
G_NEW = ""
G_OPEN = ""
G_SAVE = ""
G_UNDO = ""
G_REDO = ""
G_CUT = ""
G_COPY = ""
G_PASTE = ""
G_FIND = ""
G_FONTCOLOR = ""
G_FILL = ""
G_CLEAR = ""
G_ZOOMIN = ""
G_ZOOMOUT = ""
G_ADD = ""
G_DELETE = ""
G_BRUSH = ""
