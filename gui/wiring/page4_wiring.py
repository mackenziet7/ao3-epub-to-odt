from PySide6.QtWidgets import (
    QCheckBox,
    QWidget,
)

# ------------------------------------------------------------------
# Page 4 — Additional Options
# ------------------------------------------------------------------
def wire_page4(main_window):
    w = main_window.window
    main_window._checkSavePreset = w.findChild(QCheckBox, "checkSavePreset")
    main_window._presetNameWidget = w.findChild(QWidget, "presetNameWidget")

    main_window._presetNameWidget.setVisible(main_window._checkSavePreset.isChecked())
    main_window._checkSavePreset.toggled.connect(main_window._presetNameWidget.setVisible)
