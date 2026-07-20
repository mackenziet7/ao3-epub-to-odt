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
    QStackedWidget, QBoxLayout, QRadioButton, QAbstractSpinBox,
    QToolTip
)
from PySide6.QtUiTools import QUiLoader
from PySide6.QtGui import QIcon, QCursor
from PySide6.QtCore import QEvent, QObject, Qt

from gui.ui import resources_rc  # noqa: F401 — registers Qt resources
from gui.widgets.page_preview_widget import PagePreviewWidget, PAGE_SIZES_MM


_UI_PATH = Path(__file__).parent / "ui" / "screens" / "main_window.ui"
_QSS_PATH = Path(__file__).parent / "ui" / "res" / "styles" / "dark.qss"

CM_TO_IN = 1 / 2.54
IN_TO_CM = 2.54

MM_TO_IN = 1 / 25.4
MM_TO_CM = 0.1
PAGE_SIZE_MATCH_TOLERANCE_MM = 0.5

MARGIN_SPIN_NAMES = [
    "spinMarginTop", "spinMarginBottom",
    "spinMarginLeft", "spinMarginRight",
]
CUSTOM_SPIN_NAMES = ["spinCustomWidth", "spinCustomHeight"]
ALL_SPIN_NAMES    = MARGIN_SPIN_NAMES + CUSTOM_SPIN_NAMES

_LOCK_ICON   = ":/res/icons/lock_dark.svg"
_UNLOCK_ICON = ":/res/icons/lock_open_dark.svg"


class _SpinRangeNotifier(QObject):
    """Shows a tooltip if the user types a value outside min/max,
    right before Qt silently clamps it via CorrectToNearestValue."""

    def __init__(self, spin: QDoubleSpinBox):
        super().__init__(spin)
        self._spin = spin

    def eventFilter(self, obj, event):
        is_commit = event.type() == QEvent.Type.FocusOut or (
            event.type() == QEvent.Type.KeyPress
            and event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter)
        )
        if is_commit:
            self._check_range()
        return False  # never swallow the event

    def _check_range(self):
        spin = self._spin
        value, ok = spin.locale().toDouble(spin.cleanText())
        if not ok:
            return
        if value < spin.minimum():
            self._hint(f"Minimum is {spin.minimum():g}{spin.suffix()}")
        elif value > spin.maximum():
            self._hint(f"Maximum is {spin.maximum():g}{spin.suffix()}")

    def _hint(self, msg: str):
        QToolTip.showText(QCursor.pos(), msg, self._spin)

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
        self._wire_page0()
        self._wire_page1()
        self._wire_page2()
        self._wire_navigation()

    # ------------------------------------------------------------------
    # Page 0 — main screen
    # ------------------------------------------------------------------

    def _wire_page0(self):
        w = self.window

        splitterPage0 = w.findChild(QSplitter, "splitter")
        splitterPage0.setCollapsible(0, False)
        splitterPage0.setCollapsible(1, False)
        splitterPage0.setSizes([500, 400])

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

        splitterPage1 = w.findChild(QSplitter, "splitterSetup")
        splitterPage1.setCollapsible(0, False)
        splitterPage1.setCollapsible(1, False)
        splitterPage1.setSizes([500, 400])

        # Page size combo (custom width/height are now always visible —
        # the old show/hide customSizeWidget was removed in the layout
        # restructure, so there's nothing to wire here beyond preview sync)
        self._comboPageSize = w.findChild(QComboBox, "comboPageSize")

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
        self._radioPortrait  = w.findChild(QRadioButton, "radioPortrait")
        self._radioLandscape = w.findChild(QRadioButton, "radioLandscape")

        # Margin / custom-size spinboxes
        self._spinTop    = w.findChild(QDoubleSpinBox, "spinMarginTop")
        self._spinBottom = w.findChild(QDoubleSpinBox, "spinMarginBottom")
        self._spinLeft   = w.findChild(QDoubleSpinBox, "spinMarginLeft")
        self._spinRight  = w.findChild(QDoubleSpinBox, "spinMarginRight")
        self._spinCW     = w.findChild(QDoubleSpinBox, "spinCustomWidth")
        self._spinCH     = w.findChild(QDoubleSpinBox, "spinCustomHeight")

        self._spinCW.setKeyboardTracking(False)
        self._spinCH.setKeyboardTracking(False)

        # Lock toggle buttons (icon-swapping checkable QPushButtons)
        self._btnLockTB   = w.findChild(QPushButton, "buttonTopBotLock")
        self._btnLockLR   = w.findChild(QPushButton, "buttonLeftRightLock")
        self._btnLockSize = w.findChild(QPushButton, "buttonLockPageSize")

        for btn in (self._btnLockTB, self._btnLockLR, self._btnLockSize):
            self._wire_lock_button(btn)
        self._btnLockSize.toggled.connect(
            lambda checked: self._remember_page_size_aspect_ratio() if checked else None
        )
        self._btnLockTB.toggled.connect(
            lambda checked: self._remember_tb_margin_ratio() if checked else None
        )
        self._btnLockLR.toggled.connect(
            lambda checked: self._remember_lr_margin_ratio() if checked else None
        )

        # Page preview widget — replace placeholder
        self._install_preview()

        # Wire linked spinboxes (mirror value to partner when locked)
        self._spinTop.valueChanged.connect(self._on_top_changed)
        self._spinBottom.valueChanged.connect(self._on_bottom_changed)
        self._spinLeft.valueChanged.connect(self._on_left_changed)
        self._spinRight.valueChanged.connect(self._on_right_changed)
        self._spinCW.valueChanged.connect(self._on_cw_changed)
        self._spinCH.valueChanged.connect(self._on_ch_changed)

        self._spinCW.editingFinished.connect(
            lambda: self._on_custom_size_edit_finished(self._spinCW)
        )
        self._spinCH.editingFinished.connect(
            lambda: self._on_custom_size_edit_finished(self._spinCH)
        )

        # Wire all controls → preview sync
        self._comboPageSize.currentTextChanged.connect(self._on_page_size_changed)
        self._comboUnits.currentTextChanged.connect(self._sync_preview)
        self._checkMirrored.toggled.connect(self._sync_preview)
        self._radioPortrait.toggled.connect(self._sync_preview)
        self._radioLandscape.toggled.connect(self._sync_preview)

        self._spin_notifiers = []  # keep references alive; add near __init__ or here

        for name in ALL_SPIN_NAMES:
            spin = w.findChild(QDoubleSpinBox, name)
            if spin:
                spin.setCorrectionMode(QAbstractSpinBox.CorrectionMode.CorrectToNearestValue)
                spin.valueChanged.connect(self._sync_preview)

                notifier = _SpinRangeNotifier(spin)
                spin.installEventFilter(notifier)
                self._spin_notifiers.append(notifier)

        self._on_page_size_changed()

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
    # Page 2 — Basic Typography Setup screen
    # ------------------------------------------------------------------

    def _wire_page2(self):
            w = self.window
            splitterPage3 = w.findChild(QSplitter, "splitterPage3")
            splitterPage3.setCollapsible(0, False)
            splitterPage3.setCollapsible(1, False)
            splitterPage3.setSizes([500, 400])

    # ------------------------------------------------------------------
    # Navigation
    # ------------------------------------------------------------------

    def _wire_navigation(self):
            w     = self.window
            stack = w.findChild(QStackedWidget, "stackedWidget")
            stack.setCurrentIndex(0)  # always start on the main screen

            nav_pairs = [
                ("buttonNext",   1),  # main → page setup
                ("buttonBack",   0),  # page setup → main
                ("buttonNext_2", 2),  # page setup → typography basic
                ("buttonBack_2", 1),  # typography basic → page setup
                # buttonNext_3 intentionally not wired yet — no page 3 to go to
            ]

            for button_name, target_index in nav_pairs:
                button = w.findChild(QPushButton, button_name)
                if button is None:
                    continue
                button.clicked.connect(
                    lambda checked=False, idx=target_index: stack.setCurrentIndex(idx)
                )

    # ------------------------------------------------------------------
    # Lock icon toggle
    # ------------------------------------------------------------------

    def _wire_lock_button(self, button: QPushButton):
        """Swap between lock/unlock icons as the button is toggled.
        Defaults to unlocked on startup regardless of what the .ui shows."""
        locked_icon   = QIcon(_LOCK_ICON)
        unlocked_icon = QIcon(_UNLOCK_ICON)

        def apply_icon(checked: bool):
            button.setIcon(locked_icon if checked else unlocked_icon)

        button.setChecked(False)
        apply_icon(False)
        button.toggled.connect(apply_icon)

    # ------------------------------------------------------------------
    # Slots — lock logic
    # (blockSignals prevents the mirror write from triggering another sync)
    # ------------------------------------------------------------------

    def _on_top_changed(self, val):
        if self._btnLockTB.isChecked():
            if self._tb_margin_ratio > 0:
                self._spinBottom.blockSignals(True)
                self._spinBottom.setValue(val / self._tb_margin_ratio)
                self._spinBottom.blockSignals(False)
        else:
            self._remember_tb_margin_ratio()

    def _on_bottom_changed(self, val):
        if self._btnLockTB.isChecked():
            self._spinTop.blockSignals(True)
            self._spinTop.setValue(val * self._tb_margin_ratio)
            self._spinTop.blockSignals(False)
        else:
            self._remember_tb_margin_ratio()  

    def _on_left_changed(self, val):
        if self._btnLockLR.isChecked():
            if self._lr_margin_ratio > 0:
                self._spinRight.blockSignals(True)
                self._spinRight.setValue(val / self._lr_margin_ratio)
                self._spinRight.blockSignals(False)
        else:
            self._remember_lr_margin_ratio()

    def _on_right_changed(self, val):
        if self._btnLockLR.isChecked():
            self._spinLeft.blockSignals(True)
            self._spinLeft.setValue(val * self._lr_margin_ratio)
            self._spinLeft.blockSignals(False)
        else:
            self._remember_lr_margin_ratio()

    def _on_cw_changed(self, val):
        if self._updating_page_size_controls:
            return

        if self._btnLockSize.isChecked() and self._page_size_aspect_ratio > 0:
            self._spinCH.blockSignals(True)
            self._spinCH.setValue(val / self._page_size_aspect_ratio)
            self._spinCH.blockSignals(False)
        else:
            self._remember_page_size_aspect_ratio()

        self._on_custom_size_changed()

    def _on_ch_changed(self, val):
        if self._updating_page_size_controls:
            return

        if self._btnLockSize.isChecked():
            self._spinCW.blockSignals(True)
            self._spinCW.setValue(val * self._page_size_aspect_ratio)
            self._spinCW.blockSignals(False)
        else:
            self._remember_page_size_aspect_ratio()

        self._on_custom_size_changed()

    # ------------------------------------------------------------------
    # Slots
    # ------------------------------------------------------------------

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

    def _page_size_key_for_combo_text(self, text: str) -> str | None:
        if text.startswith("Custom"):
            return None

        for key in PAGE_SIZES_MM:
            if text == key or text.startswith(key):
                return key

        if "Paperback" in text:
            return '5.5×8.5" Paperback'
        if text.startswith("Digest"):
            return "Digest"

        return None
        
    def _remember_page_size_aspect_ratio(self):
        width_mm, height_mm = self._custom_size_mm()
        if width_mm > 0 and height_mm > 0:
            self._page_size_aspect_ratio = width_mm / height_mm   

    def _remember_tb_margin_ratio(self):
        top, bottom = self._spinTop.value(), self._spinBottom.value()
        if bottom > 0:
            self._tb_margin_ratio = top / bottom

    def _remember_lr_margin_ratio(self):
        left, right = self._spinLeft.value(), self._spinRight.value()
        if right > 0:
            self._lr_margin_ratio = left / right

    def _on_custom_size_edit_finished(self, sender):
        if self._updating_page_size_controls:
            return

        if self._btnLockSize.isChecked():
            if sender is self._spinCW and self._page_size_aspect_ratio > 0:
                self._spinCH.blockSignals(True)
                self._spinCH.setValue(
                    self._spinCW.value() / self._page_size_aspect_ratio
                )
                self._spinCH.blockSignals(False)

            elif sender is self._spinCH:
                self._spinCW.blockSignals(True)
                self._spinCW.setValue(
                    self._spinCH.value() * self._page_size_aspect_ratio
                )
                self._spinCW.blockSignals(False)

        else:
            self._remember_page_size_aspect_ratio()

        self._on_custom_size_changed()

    def _combo_index_for_page_size_key(self, key: str | None) -> int:
        for idx in range(self._comboPageSize.count()):
            text = self._comboPageSize.itemText(idx)
            if key is None and text.startswith("Custom"):
                return idx
            if key is not None and self._page_size_key_for_combo_text(text) == key:
                return idx
        return -1

    def _set_combo_page_size(self, key: str | None) -> None:
        idx = self._combo_index_for_page_size_key(key)
        if idx < 0 or idx == self._comboPageSize.currentIndex():
            return

        self._comboPageSize.blockSignals(True)
        self._comboPageSize.setCurrentIndex(idx)
        self._comboPageSize.blockSignals(False)

    def _set_custom_size_spinboxes_from_mm(self, width_mm: float, height_mm: float):
        unit = self._comboUnits.currentText()
        factor = MM_TO_IN if unit == "in" else MM_TO_CM

        self._spinCW.blockSignals(True)
        self._spinCH.blockSignals(True)
        self._spinCW.setValue(round(width_mm * factor, self._spinCW.decimals()))
        self._spinCH.setValue(round(height_mm * factor, self._spinCH.decimals()))
        self._spinCH.blockSignals(False)
        self._spinCW.blockSignals(False)

    def _custom_size_mm(self) -> tuple[float, float]:
        unit = self._comboUnits.currentText()
        factor = 25.4 if unit == "in" else 10.0
        return self._spinCW.value() * factor, self._spinCH.value() * factor

    def _matching_page_size_key(self, width_mm: float, height_mm: float) -> str | None:
        for key, (preset_w_mm, preset_h_mm) in PAGE_SIZES_MM.items():
            if (
                abs(width_mm - preset_w_mm) <= PAGE_SIZE_MATCH_TOLERANCE_MM
                and abs(height_mm - preset_h_mm) <= PAGE_SIZE_MATCH_TOLERANCE_MM
            ):
                return key
        return None

    def _on_page_size_changed(self, *_args):
        if self._updating_page_size_controls:
            return

        key = self._page_size_key_for_combo_text(self._comboPageSize.currentText())
        if key is not None:
            self._updating_page_size_controls = True
            try:
                width_mm, height_mm = PAGE_SIZES_MM[key]
                self._set_custom_size_spinboxes_from_mm(width_mm, height_mm)
                self._remember_page_size_aspect_ratio()
            finally:
                self._updating_page_size_controls = False

        self._sync_preview()

    def _on_custom_size_changed(self):
        if self._updating_page_size_controls:
            return

        width_mm, height_mm = self._custom_size_mm()
        self._set_combo_page_size(self._matching_page_size_key(width_mm, height_mm))
        self._sync_preview()    
    
    def _sync_preview(self, *_args):
        """Read current UI state and push to PagePreviewWidget."""
        w    = self.window
        unit = self._comboUnits.currentText()

        def to_mm(name: str) -> float:
            spin = w.findChild(QDoubleSpinBox, name)
            val  = spin.value() if spin else 0.0
            return val * 25.4 if unit == "in" else val * 10.0

        self._preview.update_from_settings(
            page_size_name   = (
                self._page_size_key_for_combo_text(self._comboPageSize.currentText())
                or "Custom"
            ),
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