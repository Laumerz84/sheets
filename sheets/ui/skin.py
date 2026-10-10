"""Optional skins (View > Skin): Modern, or Excel 95.

Excel 95 = Windows 95 grey and bevels, MS Sans Serif, navy highlights, 16-colour pixel toolbar icons
(pixel95.py), raised-button headers and slanted sheet tabs, and Arial 10 for unformatted cells.
It only changes how things look: files are saved exactly the same either way.
Switches live; saved as "skin" in settings.json."""
from PySide6.QtCore import QEvent, QObject, QPoint, QSize, Qt
from PySide6.QtGui import QActionGroup, QColor, QFont, QPainter, QPolygon
from PySide6.QtWidgets import QApplication, QTabBar, QToolBar, QWidget

from . import style as S

_loaded = False


def load_once(settings):
    """Apply the saved skin the first time a window opens."""
    global _loaded
    if not _loaded:
        _loaded = True
        if settings.get("skin", "modern") != "modern":
            apply(settings.get("skin"))


def apply(name):
    """Switch every open window to skin `name`."""
    from .mainwindow import WINDOWS
    S.set_skin(name)
    app = QApplication.instance()
    if app is not None:
        (S.apply_palette_95 if S.SKIN == "excel95" else S.apply_palette)(app)
    for w in WINDOWS:
        refresh(w)


def choose(win, name):
    """Menu action: switch, and remember the choice."""
    from .mainwindow import load_settings, save_settings
    apply(name)
    s = load_settings()
    s["skin"] = S.SKIN
    save_settings(s)
    win.settings = s


def attach(win):
    """Hook a new window up: saved skin, tab painting, icon sizes."""
    load_once(win.settings)
    win.tabs.installEventFilter(_tab_painter(win))
    refresh(win)


def build_menu(menu, win):
    def fill():
        menu.clear()
        group = QActionGroup(menu)
        for name, label in S.SKINS.items():
            a = menu.addAction(label)
            a.setCheckable(True)
            a.setChecked(S.SKIN == name)
            a.triggered.connect(lambda _checked=False, n=name: choose(win, n))
            group.addAction(a)
    menu.aboutToShow.connect(fill)
    fill()


def refresh(win):
    """Redraw one window in the current skin (icons repaint themselves; sizes and fonts need a nudge)."""
    for tb in win.findChildren(QToolBar):
        if tb.property("modern_icon_size") is None:
            tb.setProperty("modern_icon_size", tb.iconSize())
        tb.setIconSize(QSize(16, 16) if S.SKIN == "excel95" else tb.property("modern_icon_size"))
    # strips that carry their own modern stylesheet: the formula bar row and the sheet-tab row
    sunken = "border: 1px solid; border-color: #808080 #FFFFFF #FFFFFF #808080;"
    for widget, css95 in ((getattr(win, "name_box", None) and win.name_box.parentWidget(), "background: #C0C0C0;"),
                          (getattr(win, "tabs", None) and win.tabs.parentWidget(), "background: #C0C0C0;"),
                          (getattr(win, "fbar", None), f"QPlainTextEdit {{ background: white; {sunken} }}")):
        if widget is None:
            continue
        if widget.property("modern_css") is None:
            widget.setProperty("modern_css", widget.styleSheet())
        widget.setStyleSheet(css95 if S.SKIN == "excel95" else widget.property("modern_css"))
    grid = getattr(win, "grid", None)
    if grid is not None and grid.sheet is not None:
        grid.relayout()  # drops cached fonts, so unformatted cells pick up the skin's font
        grid.update()
    if getattr(win, "sheet", None) is not None:
        try:
            win._sync_format_controls()  # font box shows Arial 10 / Calibri 11 for unformatted cells
        except Exception:  # noqa: BLE001 - cosmetic only
            pass
    win.update()
    for child in win.findChildren(QWidget):
        child.update()


class _TabPainter(QObject):
    """Draws the sheet tabs like Excel 95 (slanted grey tabs, the current one white) when that
    skin is on; otherwise lets the tab bar paint itself. Clicks and drags are untouched."""

    def eventFilter(self, obj, ev):
        if ev.type() != QEvent.Paint or S.SKIN != "excel95" or not isinstance(obj, QTabBar):
            return False
        p = QPainter(obj)
        grey, white, black, dark = QColor("#C0C0C0"), QColor("#FFFFFF"), QColor("#000000"), QColor("#808080")
        p.fillRect(obj.rect(), grey)
        p.setPen(black)
        p.drawLine(0, 0, obj.width(), 0)
        cur = obj.currentIndex()
        order = [i for i in range(obj.count()) if i != cur] + ([cur] if cur >= 0 else [])
        for i in order:
            r = obj.tabRect(i)
            sl = max(4, r.height() // 2)
            bottom = r.bottom() - 1
            poly = QPolygon([QPoint(r.left() - sl // 2, 0), QPoint(r.right() + sl // 2, 0),
                             QPoint(r.right() - sl // 2, bottom), QPoint(r.left() + sl // 2, bottom)])
            p.setBrush(white if i == cur else grey)
            p.setPen(black)
            p.drawPolygon(poly)
            if i == cur:  # the current tab is open to the sheet above it
                p.setPen(white)
                p.drawLine(r.left() - sl // 2 + 1, 0, r.right() + sl // 2 - 1, 0)
            else:
                p.setPen(dark)
                p.drawLine(r.left() + sl // 2 + 1, bottom - 1, r.right() - sl // 2 - 1, bottom - 1)
            f = QFont("Microsoft Sans Serif")
            f.setPixelSize(11)
            f.setBold(i == cur)
            f.setStyleStrategy(QFont.NoAntialias)
            p.setFont(f)
            p.setPen(black)
            p.drawText(r, Qt.AlignCenter, obj.tabText(i))
        p.end()
        return True


def _tab_painter(win):
    tp = getattr(win, "_skin_tab_painter", None)
    if tp is None:
        tp = _TabPainter(win)
        win._skin_tab_painter = tp
    return tp
