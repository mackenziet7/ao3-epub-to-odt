from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
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

def apply_page4_mode(main_window):
    check = main_window._checkSavePreset
    name_edit = main_window.window.findChild(QLineEdit, "lineEditPresetName")

    if main_window.preset_only:
        check.setChecked(True)          # keeps _handle_save_preset working unchanged
        check.setVisible(False)
        main_window._presetNameWidget.setVisible(True)
        main_window._buttonNext_5.setText("Complete")
        if main_window.editing_preset:
            name_edit.setText(main_window.editing_preset)
    else:
        check.setVisible(True)
        check.setChecked(False)         # also hides the name widget via the toggled signal
        main_window._buttonNext_5.setText("Next")

    update_next_button_5(main_window)

def update_next_button_5(main_window):
    next_reasons = []

    # LibreOffice is only needed if actually converting
    if not main_window.preset_only:
        reason = lo_missing_reason(main_window)
        if reason is not None:
            next_reasons.append(reason)

    if main_window._checkSavePreset.isChecked():
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

    if main_window.preset_only:
        _finish_preset_only(main_window)
        return

    _wizard_convert(main_window, settings)

def _finish_preset_only(main_window):
    from gui.wiring.page6_wiring import populate_preset_combo_boxes
    from gui.wiring.page0_wiring import populate_preset_list

    w = main_window.window
    new_name = w.findChild(QLineEdit, "lineEditPresetName").text().strip()
    old_name = main_window.editing_preset
    if old_name and new_name != old_name:
        preset_name_to_path(old_name).unlink(missing_ok=True)   # rename = replace the old file
    w.findChild(QLineEdit, "lineEditPresetName").clear()

    populate_preset_combo_boxes(w, w.findChild(QComboBox, "comboRemovePreset"))
    populate_preset_list(w)
    main_window._return_to_settings()   # resets history, flag, and page 4 UI

def _handle_save_preset(main_window, settings: dict) -> bool:
    if not main_window._checkSavePreset.isChecked():
        return True  # nothing to save, let conversion proceed

    name = main_window.window.findChild(QLineEdit, "lineEditPresetName").text().strip()
    path =  preset_name_to_path(name)

    if not path.exists():
        save_preset(name, settings)
        return True
    if name == main_window.editing_preset:
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

    main_window._go_to_page(5)
    start_conversion(main_window, epub_paths, output_folder, settings)