from PySide6.QtWidgets import (
    QCheckBox,
    QLineEdit,
    QListWidget,
    QPushButton,
    QStackedWidget,
    QWidget,
)

from gui.config import lo_missing_reason
from gui.wiring.preset_builder import collect_wizard_settings
from gui.wiring.page5_wiring import start_conversion

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

    main_window._buttonNext_5.clicked.connect(lambda: _wizard_convert(main_window))

    update_next_button_5(main_window)


def update_next_button_5(main_window):
    """
    Gate buttonNext_5 on LibreOffice availability so users don't walk
    through the whole wizard only to hit a dead end on the last page.
    """
    reason = lo_missing_reason(main_window)
    main_window._buttonNext_5.setEnabled(reason is None)
    main_window._buttonNext_5.setToolTip(reason or "")


def _wizard_convert(main_window):
    """
    Wizard-path conversion entry point. Mirrors page0_wiring._quick_convert,
    but builds settings from the wizard pages (1-4) instead of loading
    a saved preset from disk.
    """
    w = main_window.window

    settings = collect_wizard_settings(main_window)

    file_list = w.findChild(QListWidget, "listWidgetFileInput")
    epub_paths = [file_list.item(i).text() for i in range(file_list.count())]
    output_folder = w.findChild(QLineEdit, "lineEditOutputFolder").text()

    # TODO: save-as-preset write-to-disk still deferred (gui/config.py) —
    # main_window._checkSavePreset.isChecked() is available here when ready.

    w.findChild(QStackedWidget, "stackedWidget").setCurrentIndex(5)
    start_conversion(main_window, epub_paths, output_folder, settings)