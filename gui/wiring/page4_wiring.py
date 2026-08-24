from PySide6.QtWidgets import (
    QCheckBox,
    QPushButton,
    QWidget,
)

from gui.config import lo_missing_reason

# ------------------------------------------------------------------
# Page 4 — Additional Options
# ------------------------------------------------------------------
def wire_page4(main_window):
    w = main_window.window
    main_window._checkSavePreset = w.findChild(QCheckBox, "checkSavePreset")
    main_window._presetNameWidget = w.findChild(QWidget, "presetNameWidget")
    main_window._buttonNext_5 = w.findChild(QPushButton, "buttonNext_5")

    main_window._presetNameWidget.setVisible(main_window._checkSavePreset.isChecked())
    main_window._checkSavePreset.toggled.connect(main_window._presetNameWidget.setVisible)

    update_next_button_5(main_window)


def update_next_button_5(main_window):
    reason = lo_missing_reason(main_window)
    main_window._buttonNext_5.setEnabled(reason is None)
    main_window._buttonNext_5.setToolTip(reason or "")
