"""
gui/theme.py
============
Applies the app's light/dark theme: native OS-style rendering forced via
QStyleHints, plus a small QSS layer for custom-attributed widgets
(iconButton, flatList) that native styling doesn't cover.
"""
from pathlib import Path
from PySide6.QtWidgets import QApplication, QPushButton
from PySide6.QtGui import QIcon, QPixmap, QPainter, QColor
from PySide6.QtCore import Qt
from pathlib import Path

_STYLES_DIR = Path(__file__).parent / "ui" / "res" / "styles"
DARK_QSS_PATH = _STYLES_DIR / "dark.qss"
LIGHT_QSS_PATH = _STYLES_DIR / "light.qss"

ICON_COLOR_DARK = QColor("#f3f3f3")
ICON_COLOR_LIGHT = QColor("#1a1a1a")


def refresh_icons(main_window, mode: str):
    window = main_window.window

    for button in window.findChildren(QPushButton):
        base = button.property("iconBase")
        if base:
            button.setIcon(QIcon(f":/res/icons/{base}_{mode}.svg"))

    for hook in main_window._theme_refresh_hooks:
        hook(mode)


def apply_theme(mode: str):
    app = QApplication.instance()

    scheme = Qt.ColorScheme.Light if mode == "light" else Qt.ColorScheme.Dark
    app.styleHints().setColorScheme(scheme)

    qss_path = LIGHT_QSS_PATH if mode == "light" else DARK_QSS_PATH
    app.setStyleSheet(qss_path.read_text(encoding="utf-8"))