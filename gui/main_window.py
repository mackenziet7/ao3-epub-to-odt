"""
gui/main_window.py
==================
MainWindow: loads main_window.ui, wires all signals/slots, and hosts the
custom PagePreviewWidget for the Page Setup screen.

Entry point: call run() from main.py.
"""
import sys

from pathlib import Path

from PySide6.QtWidgets import *
from PySide6.QtUiTools import QUiLoader

from gui.config import resolve_lo_python, load_config
from gui.first_run_dialog import LOPathDialog

from gui.ui import resources_rc  # noqa: F401 — registers Qt resources
from gui.wiring.page0_wiring import wire_page0
from gui.wiring.page1_wiring import wire_page1
from gui.wiring.page2_wiring import wire_page2
from gui.wiring.page3_wiring import wire_page3
from gui.wiring.page4_wiring import wire_page4
from gui.wiring.page5_wiring import wire_page5
from gui.wiring.reset import snapshot_wizard_defaults

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
        self.history = [0]

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
        snapshot_wizard_defaults(self)

    def _resolve_lo_on_startup(self):
        lo = resolve_lo_python()
        if lo is not None:
            return lo

        cfg = load_config()
        invalid = cfg.get("lo_python")

        dialog = LOPathDialog(
            self.window,
            invalid_path=invalid
        )

        if dialog.exec() == QDialog.DialogCode.Accepted:
            return Path(load_config()["lo_python"])

        # User closed the LibreOffice dialog with X
        QApplication.quit()
        sys.exit(0)
    
    # ------------------------------------------------------------------
    # Navigation
    # ------------------------------------------------------------------
    def _go_to_page(self, index):
        """Navigate to a page and record it in the history stack."""
        stack = self.window.findChild(QStackedWidget, "stackedWidget")

        current = stack.currentIndex()

        if current == index:
            return

        self.history.append(index)
        stack.setCurrentIndex(index)

    def _go_back(self):
        """Return to the previous page in the navigation history."""
        if len(self.history) <= 1:
            return

        # Remove the current page.
        self.history.pop()

        # The new last item is the previous page.
        previous_page = self.history[-1]

        stack = self.window.findChild(QStackedWidget, "stackedWidget")
        stack.setCurrentIndex(previous_page)

    def _complete_wizard(self):
        """Finish the wizard and return to the main page."""
        self.history = [0]

        stack = self.window.findChild(QStackedWidget, "stackedWidget")
        stack.setCurrentIndex(0)


    def _wire_navigation(self):
        w = self.window
        stack = w.findChild(QStackedWidget, "stackedWidget")
        stack.setCurrentIndex(0)  # always start on the main screen

        # --------------------------------------------------------------
        # Forward navigation
        # --------------------------------------------------------------
        forward_pairs = [
            ("buttonConvert", 5),       # main → convert
            ("buttonNext", 1),          # main → page setup

            ("buttonSettings", 6),      # main → settings

            ("buttonNext_2", 2),        # page setup → typography basic

            ("buttonMoreOptions", 3),   # typography basic → more options
            ("buttonNext_3", 4),        # typography basic → additional options
        ]

        for button_name, target_index in forward_pairs:
            button = w.findChild(QPushButton, button_name)

            if button is None:
                continue

            button.clicked.connect(
                lambda checked=False, idx=target_index:
                    self._go_to_page(idx)
            )

        # Back navigation
        back_buttons = [
            "buttonBack",
            "buttonBack_2",
            "buttonBack_4",
            "buttonBack_5",
        ]

        for button_name in back_buttons:
            button = w.findChild(QPushButton, button_name)
            if button is None:
                continue

            button.clicked.connect(
                lambda checked=False:
                    self._go_back()
            )

        # Complete wizard
        button_complete = w.findChild(QPushButton, "buttonComplete")

        if button_complete is not None:
            button_complete.clicked.connect(self._complete_wizard)

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