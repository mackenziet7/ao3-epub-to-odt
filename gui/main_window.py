"""
gui/main_window.py
==================
MainWindow: loads main_window.ui, wires all signals/slots, and hosts the
custom PagePreviewWidget for the Page Setup screen.

Entry point: call run() from main.py.
"""

from pathlib import Path

from PySide6.QtWidgets import (
    QApplication, QSplitter, QButtonGroup, QPushButton,
    QCheckBox, QComboBox, QLabel, QDoubleSpinBox, QWidget,
    QStackedWidget, QBoxLayout,
)
from PySide6.QtUiTools import QUiLoader

from gui.ui import resources_rc  # noqa: F401 — registers Qt resources
from gui.widgets.page_preview_widget import PagePreviewWidget


_UI_PATH = Path(__file__).parent / "ui" / "screens" / "main_window.ui"
_QSS_PATH = Path(__file__).parent / "ui" / "res" / "styles" / "dark.qss"

CM_TO_IN = 1 / 2.54
IN_TO_CM = 2.54

MARGIN_SPIN_NAMES = [
    "spinMarginTop", "spinMarginBottom",
    "spinMarginLeft", "spinMarginRight",
]
CUSTOM_SPIN_NAMES = ["spinCustomWidth", "spinCustomHeight"]
ALL_SPIN_NAMES    = MARGIN_SPIN_NAMES + CUSTOM_SPIN_NAMES


class MainWindow:
    """
    Wraps the QUiLoader window and owns all signal/slot wiring.

    Not a QMainWindow subclass — we load the window from .ui and store it
    on self.window, matching the established pattern in this project.
    """

    def __init__(self):
        loader = QUiLoader()
        self.window = loader.load(str(_UI_PATH))

        self._wire_page0()
        self._wire_page1()
        self._wire_navigation()

    # ------------------------------------------------------------------
    # Page 0 — main screen
    # ------------------------------------------------------------------

    def _wire_page0(self):
        w = self.window

        splitter = w.findChild(QSplitter, "splitter")
        splitter.setCollapsible(0, False)
        splitter.setCollapsible(1, False)

        mode_group = QButtonGroup(w)
        for name in ("modeSingleButton", "modeBatchButton", "modeCollectionButton"):
            mode_group.addButton(w.findChild(QPushButton, name))
        mode_group.setExclusive(True)
        self._mode_group = mode_group  # keep reference

    # ------------------------------------------------------------------
    # Page 1 — Page Setup screen
    # ------------------------------------------------------------------

    def _wire_page1(self):
        w = self.window

        splitterSetup = w.findChild(QSplitter, "splitterSetup")
        splitterSetup.setCollapsible(0, False)
        splitterSetup.setCollapsible(1, False)

        # Custom size widget visibility
        self._comboPageSize    = w.findChild(QComboBox,      "comboPageSize")
        self._customSizeWidget = w.findChild(QWidget,         "customSizeWidget")
        self._comboPageSize.currentIndexChanged.connect(self._on_page_size_changed)

        # Mirrored margins label swap
        self._labelMarginLeft  = w.findChild(QLabel,    "labelMarginLeft")
        self._labelMarginRight = w.findChild(QLabel,    "labelMarginRight")
        self._checkMirrored    = w.findChild(QCheckBox, "checkMirroredMargins")
        self._checkMirrored.toggled.connect(self._on_mirrored_toggled)
        self._on_mirrored_toggled(True)

        # Units toggle
        self._comboUnits = w.findChild(QComboBox, "comboUnits")
        self._comboUnits.currentIndexChanged.connect(self._on_units_changed)

        # Orientation radios
        from PySide6.QtWidgets import QRadioButton
        self._radioPortrait  = w.findChild(QRadioButton, "radioPortrait")
        self._radioLandscape = w.findChild(QRadioButton, "radioLandscape")

        # Page preview widget — replace placeholder
        self._install_preview()

        # Wire all controls → preview sync
        self._comboPageSize.currentTextChanged.connect(self._sync_preview)
        self._comboUnits.currentTextChanged.connect(self._sync_preview)
        self._checkMirrored.toggled.connect(self._sync_preview)
        self._radioPortrait.toggled.connect(self._sync_preview)
        self._radioLandscape.toggled.connect(self._sync_preview)

        for name in ALL_SPIN_NAMES:
            spin = w.findChild(QDoubleSpinBox, name)
            if spin:
                spin.valueChanged.connect(self._sync_preview)

        self._sync_preview()

    def _install_preview(self):
        """Replace the QWidget placeholder with a live PagePreviewWidget."""
        placeholder = self.window.findChild(QWidget, "pagePreviewWidget")
        container   = placeholder.parentWidget()
        layout      = container.layout()

        if not isinstance(layout, QBoxLayout):
            # Fallback: just overlay if layout type is unexpected
            self._preview = PagePreviewWidget(container)
            placeholder.hide()
            return

        idx = layout.indexOf(placeholder)
        placeholder.hide()
        layout.removeWidget(placeholder)

        self._preview = PagePreviewWidget(container)
        layout.insertWidget(idx, self._preview)

    # ------------------------------------------------------------------
    # Navigation
    # ------------------------------------------------------------------

    def _wire_navigation(self):
        w     = self.window
        stack = w.findChild(QStackedWidget, "stackedWidget")

        w.findChild(QPushButton, "buttonNext").clicked.connect(
            lambda: stack.setCurrentIndex(1)
        )
        w.findChild(QPushButton, "buttonBack").clicked.connect(
            lambda: stack.setCurrentIndex(0)
        )

    # ------------------------------------------------------------------
    # Slots
    # ------------------------------------------------------------------

    def _on_page_size_changed(self):
        is_custom = self._comboPageSize.currentText().startswith("Custom")
        self._customSizeWidget.setVisible(is_custom)

    def _on_mirrored_toggled(self, checked: bool):
        self._labelMarginLeft.setText("Inside" if checked else "Left")
        self._labelMarginRight.setText("Outside" if checked else "Right")

    def _on_units_changed(self):
        w    = self.window
        unit = self._comboUnits.currentText()
        suffix = f" {unit}"

        for name in ALL_SPIN_NAMES:
            spin = w.findChild(QDoubleSpinBox, name)
            if spin is None:
                continue
            old_val  = spin.value()
            is_custom = "Custom" in name

            spin.setSuffix(suffix)
            if unit == "in":
                spin.setDecimals(3)
                spin.setMaximum(40.0 if is_custom else 4.0)
                spin.setSingleStep(0.05)
                spin.setValue(round(old_val * CM_TO_IN, 3))
            else:
                spin.setDecimals(2)
                spin.setMaximum(100.0 if is_custom else 10.0)
                spin.setSingleStep(0.1)
                spin.setValue(round(old_val * IN_TO_CM, 2))

    def _sync_preview(self):
        """Read current UI state and push to PagePreviewWidget."""
        w    = self.window
        unit = self._comboUnits.currentText()

        def to_mm(name: str) -> float:
            spin = w.findChild(QDoubleSpinBox, name)
            val  = spin.value() if spin else 0.0
            return val * 25.4 if unit == "in" else val * 10.0

        self._preview.update_from_settings(
            page_size_name   = self._comboPageSize.currentText(),
            landscape        = self._radioLandscape.isChecked(),
            mirrored         = self._checkMirrored.isChecked(),
            custom_w_mm      = to_mm("spinCustomWidth"),
            custom_h_mm      = to_mm("spinCustomHeight"),
            margin_top_mm    = to_mm("spinMarginTop"),
            margin_bottom_mm = to_mm("spinMarginBottom"),
            margin_inside_mm = to_mm("spinMarginLeft"),   # Left == Inside
            margin_outside_mm= to_mm("spinMarginRight"),  # Right == Outside
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