"""Mouse pointer themes (View > Cursor): Excel's normal cursors, or any image file the user picks
(e.g. a game item sprite they saved), shown at its exact size and pixels.

The chosen image replaces the normal pointer and the white cell plus. Resize, fill-handle and
text cursors stay as they are, so they still show what a drag will do. Saved as
"cursor_theme" (+ "cursor_file" for an image) in settings.json; "default" = Excel's cursors."""
import os
import shutil

from PySide6.QtCore import Qt
from PySide6.QtGui import QCursor, QImage, QPixmap

THEMES = {"default": ("Excel (normal)", None)}  # name -> (menu text, unused)
FILE = "file"  # theme name for a user-picked image

IMAGE_EXTS = (".png", ".gif", ".ico", ".cur", ".bmp", ".jpg", ".jpeg", ".webp")

_cache = {}
_current = "default"
_file = None
_loaded = False


def load_once(settings):
    """Apply the saved theme the first time any window asks (later windows share it)."""
    global _loaded
    if not _loaded:
        set_current(settings.get("cursor_theme", "default"), settings.get("cursor_file"))
        _loaded = True


def library(folder):
    """The cursor images kept in `folder`, as [(menu label, path)]: 'Dragon_scimitar.png' -> 'Dragon scimitar'."""
    try:
        names = sorted(f for f in os.listdir(folder) if f.lower().endswith(IMAGE_EXTS))
    except OSError:
        return []
    return [(os.path.splitext(f)[0].replace("_", " ").strip(), os.path.join(folder, f)) for f in names]


def current():
    return _current


def current_file():
    return _file


def set_current(name, file=None):
    global _current, _file
    if name == FILE and file and os.path.exists(file):
        _current, _file = FILE, file
    else:
        _current = "default"
    _cache.pop(FILE, None)


def pointer():
    """The picked image as a cursor, or None for the normal cursors."""
    if _current != FILE:
        return None
    if FILE not in _cache:
        _cache[FILE] = cursor_from_file(_file)
    return _cache[FILE]


def keep_copy(path, folder):
    """Copy a picked image into Ekxel's settings folder, so moving the original doesn't break it."""
    os.makedirs(folder, exist_ok=True)
    dest = os.path.join(folder, os.path.basename(path))
    if os.path.normcase(os.path.abspath(path)) != os.path.normcase(os.path.abspath(dest)):
        shutil.copyfile(path, dest)
    return dest


def cursor_from_file(path):
    """A cursor from an image (png, gif, ico, cur, bmp...), or None if it can't be read.
    Shown exactly as the file is: same pixels, same size, never scaled or smoothed (only a
    giant image over 256 px is shrunk, with hard pixel edges). The click point is the tip:
    the visible pixel nearest the top-left corner, or the top-right one if the picture
    points that way."""
    img = QImage(path) if path else QImage()
    if img.isNull():
        return None
    img = img.convertToFormat(QImage.Format_ARGB32)
    if max(img.width(), img.height()) > 256:
        img = img.scaled(256, 256, Qt.KeepAspectRatio, Qt.FastTransformation)
    w, h = img.width(), img.height()
    opaque = [(x, y) for y in range(h) for x in range(w) if img.pixelColor(x, y).alpha() > 64]
    if not opaque:
        return None
    top_left = min(opaque, key=lambda q: q[0] + q[1])
    top_right = min(opaque, key=lambda q: (w - 1 - q[0]) + q[1])
    hot = top_left if sum(top_left) <= (w - 1 - top_right[0]) + top_right[1] else top_right
    pm = QPixmap.fromImage(img)
    pm.setDevicePixelRatio(1.0)  # 1 image pixel = 1 screen pixel
    return QCursor(pm, *hot)
