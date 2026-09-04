from PySide6.QtWidgets import (
    QCheckBox,
    QLineEdit,
    QListWidget,
    QMessageBox,
    QPushButton,
    QStackedWidget,
    QWidget,
)

from gui.config import INVALID_FILE_CHARS, lo_missing_reason, preset_name_to_path, save_preset
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
    main_window._checkSavePreset.toggled.connect(lambda _checked: update_next_button_5(main_window))

    main_window._buttonNext_5.clicked.connect(lambda: _on_next_5_clicked(main_window))

    w.findChild(QLineEdit, "lineEditPresetName").textChanged.connect(
        lambda _text: update_next_button_5(main_window)
    )

    update_next_button_5(main_window)


def update_next_button_5(main_window):
    """
    Gate buttonNext_5 on LibreOffice availability so users don't walk
    through the whole wizard only to hit a dead end on the last page.
    """
    reason = lo_missing_reason(main_window)
    next_reasons = []
    if reason is not None:
        next_reasons.append(reason)

    if main_window._presetNameWidget.isVisible():
        text = main_window.window.findChild(QLineEdit, "lineEditPresetName").text().strip()
        if not text:
            next_reasons.append("Please enter a preset name")

        for char in INVALID_FILE_CHARS:
            if char in text:
                next_reasons.append(f"Please enter a valid file name (remove '{char}')")

    main_window._buttonNext_5.setEnabled(not next_reasons)
    main_window._buttonNext_5.setToolTip("\n".join(next_reasons))


def _on_next_5_clicked(main_window):
    settings = collect_wizard_settings(main_window)
    if not _handle_save_preset(main_window, settings):
        return  # user cancelled the overwrite prompt — stop here
    _wizard_convert(main_window, settings)

def _handle_save_preset(main_window, settings: dict) -> bool:
    if not main_window._checkSavePreset.isChecked():
        return True  # nothing to save, let conversion proceed

    name = main_window.window.findChild(QLineEdit, "lineEditPresetName").text().strip()
    path =  preset_name_to_path(name)

    if not path.exists():
        save_preset(name, settings)
        return True

    # file exists — need to ask the user
    reply = QMessageBox.question(
            main_window.window,
            "Preset already exists",
            f"There already exists a preset called {name}. Would you like to override the previous file and continue anyways?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
    if reply != QMessageBox.StandardButton.Yes:
        return False # user chose to go back and review — stay on this page
    else:
        save_preset(name, settings)
        return True

def _wizard_convert(main_window, settings:dict):
    """
    Wizard-path conversion entry point. Mirrors page0_wiring._quick_convert,
    but builds settings from the wizard pages (1-4) instead of loading
    a saved preset from disk.
    """
    w = main_window.window

    file_list = w.findChild(QListWidget, "listWidgetFileInput")
    epub_paths = [file_list.item(i).text() for i in range(file_list.count())]
    output_folder = w.findChild(QLineEdit, "lineEditOutputFolder").text()

    w.findChild(QStackedWidget, "stackedWidget").setCurrentIndex(5)
    start_conversion(main_window, epub_paths, output_folder, settings)