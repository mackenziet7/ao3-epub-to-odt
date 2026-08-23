"""
gui/main_window.py
==================
MainWindow: loads main_window.ui, wires all signals/slots, and hosts the
custom PagePreviewWidget for the Page Setup screen.

Entry point: call run() from main.py.
"""

from pathlib import Path

from PySide6.QtWidgets import *
from PySide6.QtUiTools import QUiLoader
from PySide6.QtGui import QIcon

from gui.config import resolve_lo_python, load_config
from gui.first_run_dialog import LOPathDialog

from gui.ui import resources_rc  # noqa: F401 — registers Qt resources
from gui.wiring.page0_wiring import wire_page0
from gui.wiring.page1_wiring import wire_page1
from gui.wiring.page2_wiring import wire_page2
from gui.wiring.page3_wiring import wire_page3
from gui.wiring.page4_wiring import wire_page4
from gui.wiring.page5_wiring import wire_page5

_UI_PATH = Path(__file__).parent / "ui" / "screens" / "main_window.ui"
_QSS_PATH = Path(__file__).parent / "ui" / "res" / "styles" / "dark.qss"


class MainWindow:
    """
    Wraps the QUiLoader window and owns all signal/slot wiring.

    Not a QMainWindow subclass — we load the window from .ui and store it
    on self.window, matching the established pattern in this project.
    """

    def __init__(self):
        loader = QUiLoader()
        self.window = loader.load(str(_UI_PATH))

        self._lo_python = self._resolve_lo_on_startup()

        self._updating_page_size_controls = False
        self._page_size_aspect_ratio = 1.0
        self._tb_margin_ratio = 1.0
        self._lr_margin_ratio = 1.0
        wire_page0(self)
        wire_page1(self)
        wire_page2(self)
        wire_page3(self)
        wire_page4(self)
        wire_page5(self)
        self._wire_navigation()

    def _resolve_lo_on_startup(self):
        lo = resolve_lo_python()
        if lo is not None:
            return lo
        cfg = load_config()
        invalid = cfg.get("lo_python")
        dialog = LOPathDialog(self.window, invalid_path=invalid)
        if dialog.exec():
            return Path(load_config()["lo_python"])
        return None
    # ------------------------------------------------------------------
    # Navigation
    # ------------------------------------------------------------------
    def _wire_navigation(self):
        w     = self.window
        stack = w.findChild(QStackedWidget, "stackedWidget")
        stack.setCurrentIndex(0)  # always start on the main screen

        nav_pairs = [
            ("buttonConvert", 5), # main → convert
            ("buttonNext",   1),  # main → page setup
            ("buttonBack",   0),  # page setup → main

            ("buttonSettings", 6), # main → settings
            ("buttonBack_5", 0),  # settings → main

            ("buttonNext_2", 2),  # page setup → typography basic
            ("buttonBack_2", 1),  # typography basic → page setup

            ("buttonMoreOptions", 3),  # typography basic → typography more options
            ("buttonNext_3", 4),  # typography basic → additional options
            ("buttonBack_3", 2),  # typography more options → typography basic

            ("buttonNext_5", 5),  # typography more options → additional options
            ("buttonBack_4", 3),  # additional options → typography #TODO store last typography page state so this takes to appropriate page
            
            ("buttonComplete", 0) 
        ]

        for button_name, target_index in nav_pairs:
            button = w.findChild(QPushButton, button_name)
            if button is None:
                continue
            button.clicked.connect(
                lambda checked=False, idx=target_index: stack.setCurrentIndex(idx)
            )

    # ------------------------------------------------------------------
    # Show
    # ------------------------------------------------------------------
    def show(self):
        self.window.show()


def run():
    import sys
    app = QApplication.instance() or QApplication(sys.argv)
    app.setStyleSheet(_QSS_PATH.read_text(encoding="utf-8"))

    mw = MainWindow()
    mw.show()

    sys.exit(app.exec())