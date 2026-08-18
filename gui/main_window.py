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

from gui.ui import resources_rc  # noqa: F401 — registers Qt resources
from gui.wiring.page0_wiring import wire_page0
from gui.wiring.page1_wiring import wire_page1
from gui.wiring.page2_wiring import wire_page2

_UI_PATH = Path(__file__).parent / "ui" / "screens" / "main_window.ui"
_QSS_PATH = Path(__file__).parent / "ui" / "res" / "styles" / "dark.qss"

_CIRCLE_EMPTY_ICON   = ":/res/icons/circle_dark.svg"
_CIRCLE_CHECKED_ICON = ":/res/icons/circle_checked.svg"


class MainWindow:
    """
    Wraps the QUiLoader window and owns all signal/slot wiring.

    Not a QMainWindow subclass — we load the window from .ui and store it
    on self.window, matching the established pattern in this project.
    """

    def __init__(self):
        loader = QUiLoader()
        self.window = loader.load(str(_UI_PATH))

        self._updating_page_size_controls = False
        self._page_size_aspect_ratio = 1.0
        self._tb_margin_ratio = 1.0
        self._lr_margin_ratio = 1.0
        wire_page0(self)
        wire_page1(self)
        wire_page2(self)
        self._wire_page3()
        self._wire_page4()
        self._wire_navigation()

    # ------------------------------------------------------------------
    # Page 3 — Typography Advanced ("More Options") screen
    # ------------------------------------------------------------------
    def _wire_page3(self):
        w = self.window
        splitterPage3 = w.findChild(QSplitter, "splitterTypographyMore")
        splitterPage3.setCollapsible(0, False)
        splitterPage3.setCollapsible(1, False)
        splitterPage3.setSizes([200, 700])

        self._groupList = w.findChild(QListWidget, "groupListWidget")
        self._groupList.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)

        empty_icon = QIcon(_CIRCLE_EMPTY_ICON)
        for i in range(self._groupList.count()):
            self._groupList.item(i).setIcon(empty_icon)

        self._groupContentStack = w.findChild(QStackedWidget, "groupContentStack")
        self._groupContentStack.setCurrentIndex(0)

        self._groupList.setCurrentRow(0)  # start on "Main Book"
        self._groupList.currentRowChanged.connect(self._on_typography_group_changed)

        self._typography_group_checked = [False, False, False]
        self._typography_current_row = 0

        w.findChild(QPushButton, "buttonNext_4").clicked.connect(
            self._on_typography_advanced_next
        )
        w.findChild(QPushButton, "buttonBack_3").clicked.connect(
            self._on_typography_advanced_back
        )

    def _on_typography_group_changed(self, row: int):
        if self._typography_current_row is not None:
            self._mark_typography_group_checked(self._typography_current_row)

        self._typography_current_row = row
        self._groupContentStack.setCurrentIndex(row)

    def _mark_typography_group_checked(self, row: int):
        self._typography_group_checked[row] = True
        self._groupList.item(row).setIcon(QIcon(_CIRCLE_CHECKED_ICON))

    def _on_typography_advanced_next(self):
        current_row = self._groupList.currentRow()

        if current_row < self._groupList.count() - 1:
            self._groupList.setCurrentRow(current_row + 1)
            return

        self._mark_typography_group_checked(current_row)

        if not all(self._typography_group_checked):
            reply = QMessageBox.question(
                self.window,
                "Incomplete Sections",
                "Some sections still have default settings. Would you like to continue anyways?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if reply != QMessageBox.StandardButton.Yes:
                return  # user chose to go back and review — stay on this page

        self.window.findChild(QStackedWidget, "stackedWidget").setCurrentIndex(4)

    def _on_typography_advanced_back(self):
        reply = QMessageBox.question(
            self.window,
            "Discard Changes?",
            "Any per-style changes made here will be lost. Continue?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        # Reset checklist state
        self._typography_group_checked = [False, False, False]
        self._typography_current_row = None
        empty_icon = QIcon(_CIRCLE_EMPTY_ICON)
        for i in range(self._groupList.count()):
            self._groupList.item(i).setIcon(empty_icon)

        self._groupList.setCurrentRow(0)
        self._groupContentStack.setCurrentIndex(0)

        # NOTE: actual per-style field values aren't reset here yet —
        # that requires the overrides/backend system we haven't built.
        # For now this only resets the checklist UI state.

        w = self.window
        stack = w.findChild(QStackedWidget, "stackedWidget")
        stack.setCurrentIndex(2)

    # ------------------------------------------------------------------
    # Page 4 — Additional Options screen
    # ------------------------------------------------------------------
    def _wire_page4(self):
        w = self.window
        self._checkSavePreset = w.findChild(QCheckBox, "checkSavePreset")
        self._presetNameWidget = w.findChild(QWidget, "presetNameWidget")

        self._presetNameWidget.setVisible(self._checkSavePreset.isChecked())
        self._checkSavePreset.toggled.connect(self._presetNameWidget.setVisible)

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