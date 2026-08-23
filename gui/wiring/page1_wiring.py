
from PySide6.QtWidgets import (
    QAbstractSpinBox,
    QBoxLayout,
    QLabel,
    QSplitter,
    QCheckBox,
    QRadioButton,
    QComboBox,
    QDoubleSpinBox,
    QPushButton,
    QToolTip,
    QWidget,
)
from PySide6.QtCore import QEvent, QObject, Qt
from PySide6.QtGui import QIcon, QCursor

from gui.widgets.page_preview_widget import PagePreviewWidget
from gui.widgets.page_preview_widget import PAGE_SIZES_MM

_LOCK_ICON   = ":/res/icons/lock_dark.svg"
_UNLOCK_ICON = ":/res/icons/lock_open_dark.svg"

ALL_SPIN_NAMES = (
    "spinMarginTop", "spinMarginBottom",
    "spinMarginLeft", "spinMarginRight",
    "spinCustomWidth", "spinCustomHeight",
)

CM_TO_IN = 1 / 2.54
IN_TO_CM = 2.54
MM_TO_IN = 1 / 25.4
MM_TO_CM = 0.1
PAGE_SIZE_MATCH_TOLERANCE_MM = 0.5


class _SpinRangeNotifier(QObject):
    """Warn before Qt clamps a manually entered value to its valid range."""

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
        return False

    def _check_range(self):
        spin = self._spin
        value, ok = spin.locale().toDouble(spin.cleanText())
        if not ok:
            return
        if value < spin.minimum():
            QToolTip.showText(QCursor.pos(), f"Minimum is {spin.minimum():g}{spin.suffix()}", spin)
        elif value > spin.maximum():
            QToolTip.showText(QCursor.pos(), f"Maximum is {spin.maximum():g}{spin.suffix()}", spin)


# ------------------------------------------------------------------
# Page 1 — Page Setup screen
# ------------------------------------------------------------------
def wire_page1(main_window):
    w = main_window.window

    splitterPage1 = w.findChild(QSplitter, "splitterSetup")
    splitterPage1.setCollapsible(0, False)
    splitterPage1.setCollapsible(1, False)
    splitterPage1.setSizes([500, 400])

    main_window._comboPageSize = w.findChild(QComboBox, "comboPageSize")

    # Mirrored margins label swap
    main_window._labelMarginLeft  = w.findChild(QLabel,    "labelMarginLeft")
    main_window._labelMarginRight = w.findChild(QLabel,    "labelMarginRight")
    main_window._checkMirrored    = w.findChild(QCheckBox, "checkMirroredMargins")
    main_window._checkMirrored.toggled.connect(
        lambda checked: _on_mirrored_toggled(main_window, checked)
    )
    _on_mirrored_toggled(main_window, True)

    # Units toggle
    main_window._comboUnits = w.findChild(QComboBox, "comboUnits")
    main_window._comboUnits.currentIndexChanged.connect(
        lambda _index: _on_units_changed(main_window)
    )

    # Orientation radios
    main_window._radioPortrait  = w.findChild(QRadioButton, "radioPortrait")
    main_window._radioLandscape = w.findChild(QRadioButton, "radioLandscape")

    # Margin / custom-size spinboxes
    main_window._spinTop    = w.findChild(QDoubleSpinBox, "spinMarginTop")
    main_window._spinBottom = w.findChild(QDoubleSpinBox, "spinMarginBottom")
    main_window._spinLeft   = w.findChild(QDoubleSpinBox, "spinMarginLeft")
    main_window._spinRight  = w.findChild(QDoubleSpinBox, "spinMarginRight")
    main_window._spinCW     = w.findChild(QDoubleSpinBox, "spinCustomWidth")
    main_window._spinCH     = w.findChild(QDoubleSpinBox, "spinCustomHeight")

    main_window._spinCW.setKeyboardTracking(False)
    main_window._spinCH.setKeyboardTracking(False)

    # Lock toggle buttons (icon-swapping checkable QPushButtons)
    main_window._btnLockTB   = w.findChild(QPushButton, "buttonTopBotLock")
    main_window._btnLockLR   = w.findChild(QPushButton, "buttonLeftRightLock")
    main_window._btnLockSize = w.findChild(QPushButton, "buttonLockPageSize")

    for btn in (main_window._btnLockTB, main_window._btnLockLR, main_window._btnLockSize):
        _wire_lock_button(btn)

    main_window._btnLockSize.toggled.connect(
        lambda checked: _remember_page_size_aspect_ratio(main_window) if checked else None
    )
    main_window._btnLockTB.toggled.connect(
        lambda checked: _remember_tb_margin_ratio(main_window) if checked else None
    )
    main_window._btnLockLR.toggled.connect(
        lambda checked: _remember_lr_margin_ratio(main_window) if checked else None
    )

    # Wire linked spinboxes (mirror value to partner when locked)
    main_window._spinTop.valueChanged.connect(
        lambda value: _on_top_changed(main_window, value)
    )
    main_window._spinBottom.valueChanged.connect(
        lambda value: _on_bottom_changed(main_window, value)
    )
    main_window._spinLeft.valueChanged.connect(
        lambda value: _on_left_changed(main_window, value)
    )
    main_window._spinRight.valueChanged.connect(
        lambda value: _on_right_changed(main_window, value)
    )
    main_window._spinCW.valueChanged.connect(
        lambda value: _on_cw_changed(main_window, value)
    )
    main_window._spinCH.valueChanged.connect(
        lambda value: _on_ch_changed(main_window, value)
    )

    main_window._spinCW.editingFinished.connect(
        lambda: _on_custom_size_edit_finished(main_window, main_window._spinCW)
    )
    main_window._spinCH.editingFinished.connect(
        lambda: _on_custom_size_edit_finished(main_window, main_window._spinCH)
    )

    _install_preview(main_window)

    # Keep the visual preview in sync with all page-setup controls.
    main_window._comboPageSize.currentTextChanged.connect(
        lambda _text: _on_page_size_changed(main_window)
    )
    main_window._comboUnits.currentTextChanged.connect(
        lambda _text: _sync_preview(main_window)
    )
    main_window._checkMirrored.toggled.connect(
        lambda _checked: _sync_preview(main_window)
    )
    main_window._radioPortrait.toggled.connect(
        lambda _checked: _sync_preview(main_window)
    )
    main_window._radioLandscape.toggled.connect(
        lambda _checked: _sync_preview(main_window)
    )

    main_window._spin_notifiers = []
    for name in ALL_SPIN_NAMES:
        spin = w.findChild(QDoubleSpinBox, name)
        if spin is None:
            continue
        spin.setCorrectionMode(QAbstractSpinBox.CorrectionMode.CorrectToNearestValue)
        spin.valueChanged.connect(lambda _value: _sync_preview(main_window))

        notifier = _SpinRangeNotifier(spin)
        spin.installEventFilter(notifier)
        main_window._spin_notifiers.append(notifier)

    _on_page_size_changed(main_window)


# ------------------------------------------------------------------
# Slots — lock logic
# (blockSignals prevents the mirror write from triggering another sync)
# ------------------------------------------------------------------

def _wire_lock_button(button: QPushButton):
    """Swap between lock/unlock icons as the button is toggled.
    Defaults to unlocked on startup regardless of what the .ui shows."""
    locked_icon   = QIcon(_LOCK_ICON)
    unlocked_icon = QIcon(_UNLOCK_ICON)

    def apply_icon(checked: bool):
        button.setIcon(locked_icon if checked else unlocked_icon)

    button.setChecked(False)
    apply_icon(False)
    button.toggled.connect(apply_icon)


def _install_preview(main_window):
    """Replace the Designer placeholder with the live page preview widget."""
    placeholder = main_window.window.findChild(QWidget, "pagePreviewWidget")
    container = placeholder.parentWidget()
    layout = container.layout()

    if not isinstance(layout, QBoxLayout):
        main_window._preview = PagePreviewWidget(container)
        placeholder.hide()
        return

    idx = layout.indexOf(placeholder)
    placeholder.hide()
    layout.removeWidget(placeholder)

    main_window._preview = PagePreviewWidget(container)
    layout.insertWidget(idx, main_window._preview)

def _on_top_changed(main_window, val):
    if main_window._btnLockTB.isChecked():
        if main_window._tb_margin_ratio > 0:
            main_window._spinBottom.blockSignals(True)
            main_window._spinBottom.setValue(val / main_window._tb_margin_ratio)
            main_window._spinBottom.blockSignals(False)
    else:
        _remember_tb_margin_ratio(main_window)

def _on_bottom_changed(main_window, val):
    if main_window._btnLockTB.isChecked():
        main_window._spinTop.blockSignals(True)
        main_window._spinTop.setValue(val * main_window._tb_margin_ratio)
        main_window._spinTop.blockSignals(False)
    else:
        _remember_tb_margin_ratio(main_window)

def _on_left_changed(main_window, val):
    if main_window._btnLockLR.isChecked():
        if main_window._lr_margin_ratio > 0:
            main_window._spinRight.blockSignals(True)
            main_window._spinRight.setValue(val / main_window._lr_margin_ratio)
            main_window._spinRight.blockSignals(False)
    else:
        _remember_lr_margin_ratio(main_window)

def _on_right_changed(main_window, val):
    if main_window._btnLockLR.isChecked():
        main_window._spinLeft.blockSignals(True)
        main_window._spinLeft.setValue(val * main_window._lr_margin_ratio)
        main_window._spinLeft.blockSignals(False)
    else:
        _remember_lr_margin_ratio(main_window)

def _on_cw_changed(main_window, val):
    if main_window._updating_page_size_controls:
        return

    if main_window._btnLockSize.isChecked() and main_window._page_size_aspect_ratio > 0:
        main_window._spinCH.blockSignals(True)
        main_window._spinCH.setValue(val / main_window._page_size_aspect_ratio)
        main_window._spinCH.blockSignals(False)
    else:
        _remember_page_size_aspect_ratio(main_window)

    _on_custom_size_changed(main_window)

def _on_ch_changed(main_window, val):
    if main_window._updating_page_size_controls:
        return

    if main_window._btnLockSize.isChecked():
        main_window._spinCW.blockSignals(True)
        main_window._spinCW.setValue(val * main_window._page_size_aspect_ratio)
        main_window._spinCW.blockSignals(False)
    else:
        _remember_page_size_aspect_ratio(main_window)

    _on_custom_size_changed(main_window)


def _on_mirrored_toggled(main_window, checked: bool):
    main_window._labelMarginLeft.setText("Inside" if checked else "Left")
    main_window._labelMarginRight.setText("Outside" if checked else "Right")


def _on_units_changed(main_window):
    unit = main_window._comboUnits.currentText()
    suffix = f" {unit}"

    for name in ALL_SPIN_NAMES:
        spin = main_window.window.findChild(QDoubleSpinBox, name)
        if spin is None:
            continue
        old_val = spin.value()
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


def page_size_key_for_combo_text(text: str) -> str | None:
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


def _remember_page_size_aspect_ratio(main_window):
    width_mm, height_mm = _custom_size_mm(main_window)
    if width_mm > 0 and height_mm > 0:
        main_window._page_size_aspect_ratio = width_mm / height_mm


def _remember_tb_margin_ratio(main_window):
    top, bottom = main_window._spinTop.value(), main_window._spinBottom.value()
    if bottom > 0:
        main_window._tb_margin_ratio = top / bottom


def _remember_lr_margin_ratio(main_window):
    left, right = main_window._spinLeft.value(), main_window._spinRight.value()
    if right > 0:
        main_window._lr_margin_ratio = left / right


def _on_custom_size_edit_finished(main_window, sender):
    if main_window._updating_page_size_controls:
        return
    if main_window._btnLockSize.isChecked():
        if sender is main_window._spinCW and main_window._page_size_aspect_ratio > 0:
            main_window._spinCH.blockSignals(True)
            main_window._spinCH.setValue(
                main_window._spinCW.value() / main_window._page_size_aspect_ratio
            )
            main_window._spinCH.blockSignals(False)
        elif sender is main_window._spinCH:
            main_window._spinCW.blockSignals(True)
            main_window._spinCW.setValue(
                main_window._spinCH.value() * main_window._page_size_aspect_ratio
            )
            main_window._spinCW.blockSignals(False)
    else:
        _remember_page_size_aspect_ratio(main_window)
    _on_custom_size_changed(main_window)


def _combo_index_for_page_size_key(main_window, key: str | None) -> int:
    for idx in range(main_window._comboPageSize.count()):
        text = main_window._comboPageSize.itemText(idx)
        if key is None and text.startswith("Custom"):
            return idx
        if key is not None and page_size_key_for_combo_text(text) == key:
            return idx
    return -1


def _set_combo_page_size(main_window, key: str | None):
    idx = _combo_index_for_page_size_key(main_window, key)
    if idx < 0 or idx == main_window._comboPageSize.currentIndex():
        return
    main_window._comboPageSize.blockSignals(True)
    main_window._comboPageSize.setCurrentIndex(idx)
    main_window._comboPageSize.blockSignals(False)


def _set_custom_size_spinboxes_from_mm(main_window, width_mm: float, height_mm: float):
    unit = main_window._comboUnits.currentText()
    factor = MM_TO_IN if unit == "in" else MM_TO_CM
    main_window._spinCW.blockSignals(True)
    main_window._spinCH.blockSignals(True)
    main_window._spinCW.setValue(round(width_mm * factor, main_window._spinCW.decimals()))
    main_window._spinCH.setValue(round(height_mm * factor, main_window._spinCH.decimals()))
    main_window._spinCH.blockSignals(False)
    main_window._spinCW.blockSignals(False)


def _custom_size_mm(main_window) -> tuple[float, float]:
    factor = 25.4 if main_window._comboUnits.currentText() == "in" else 10.0
    return main_window._spinCW.value() * factor, main_window._spinCH.value() * factor


def _matching_page_size_key(width_mm: float, height_mm: float) -> str | None:
    for key, (preset_w_mm, preset_h_mm) in PAGE_SIZES_MM.items():
        if (
            abs(width_mm - preset_w_mm) <= PAGE_SIZE_MATCH_TOLERANCE_MM
            and abs(height_mm - preset_h_mm) <= PAGE_SIZE_MATCH_TOLERANCE_MM
        ):
            return key
    return None


def _on_page_size_changed(main_window):
    if main_window._updating_page_size_controls:
        return
    key = page_size_key_for_combo_text(main_window._comboPageSize.currentText())
    if key is not None:
        main_window._updating_page_size_controls = True
        try:
            width_mm, height_mm = PAGE_SIZES_MM[key]
            _set_custom_size_spinboxes_from_mm(main_window, width_mm, height_mm)
            _remember_page_size_aspect_ratio(main_window)
        finally:
            main_window._updating_page_size_controls = False
    _sync_preview(main_window)


def _on_custom_size_changed(main_window):
    if main_window._updating_page_size_controls:
        return
    width_mm, height_mm = _custom_size_mm(main_window)
    _set_combo_page_size(main_window, _matching_page_size_key(width_mm, height_mm))
    _sync_preview(main_window)


def _sync_preview(main_window):
    unit = main_window._comboUnits.currentText()

    def to_mm(name: str) -> float:
        spin = main_window.window.findChild(QDoubleSpinBox, name)
        value = spin.value() if spin else 0.0
        return value * 25.4 if unit == "in" else value * 10.0

    main_window._preview.update_from_settings(
        page_size_name=page_size_key_for_combo_text(main_window._comboPageSize.currentText()) or "Custom",
        landscape=main_window._radioLandscape.isChecked(),
        mirrored=main_window._checkMirrored.isChecked(),
        custom_w_mm=to_mm("spinCustomWidth"),
        custom_h_mm=to_mm("spinCustomHeight"),
        margin_top_mm=to_mm("spinMarginTop"),
        margin_bottom_mm=to_mm("spinMarginBottom"),
        margin_inside_mm=to_mm("spinMarginLeft"),
        margin_outside_mm=to_mm("spinMarginRight"),
    )
