from pathlib import Path

from PySide6.QtWidgets import (
    QComboBox, 
    QFileDialog, 
    QLineEdit,
    QMessageBox, 
    QPushButton,
    QWidget
)
from PySide6.QtCore import QUrl, Qt
from PySide6.QtGui import QDesktopServices

from gui.config import config_path, DEFAULT_SAVE_LOCATION, load_config, preset_name_to_path, save_config
from gui.wiring.page0_wiring import populate_preset_list
from gui.wiring.preset_builder import DEFAULT_PRESET_STEM

def wire_page6(main_window):
    w = main_window.window

    button_git = w.findChild(QPushButton, "buttonSeeGithub")
    button_git.clicked.connect(lambda:    QDesktopServices.openUrl(
                QUrl("https://github.com/mackenziet7/ao3-epub-to-odt")
            ))

    button_config_location = w.findChild(QPushButton, "buttonFileLocation")
    button_config_location.clicked.connect(lambda:    QDesktopServices.openUrl(
                QUrl.fromLocalFile(str(config_path().parent))
            ))

    # theme
    combo_theme = w.findChild(QComboBox, "comboTheme")
    combo_theme.setCurrentText("Dark" if main_window.theme == "dark" else "Light")
    combo_theme.currentTextChanged.connect(
        lambda theme: main_window.set_theme(theme)
    )

    # Default save location
    w.findChild(QPushButton, "buttonBrowse_2").clicked.connect(
        lambda: pick_folder(w)
    )

    cfg = load_config()
    w.findChild(QLineEdit, "lineEditDefaultSave").setText(cfg["output_folder"])
    w.findChild(QPushButton, "buttonSaveReset").clicked.connect(
        lambda: save_folder_reset(w)
    )

    # Presets
    remove_combo = w.findChild(QComboBox, "comboRemovePreset")
    populate_preset_combo_boxes(w, remove_combo)
    w.findChild(QPushButton, "buttonRemovePreset").clicked.connect(
        lambda: remove_preset(w, remove_combo)
    )

# ------------------------------------------------------------------
# Default save folder
# ------------------------------------------------------------------
def pick_folder(window):
    folder = QFileDialog.getExistingDirectory(
            window,
            "Select output folder",
            DEFAULT_SAVE_LOCATION,
        )
    if folder:
        cfg = load_config()
        cfg["output_folder"] = folder
        save_config(cfg)
        window.findChild(QLineEdit, "lineEditDefaultSave").setText(folder)
        window.findChild(QLineEdit, "lineEditOutputFolder").setText(folder)

def save_folder_reset(window):
    cfg = load_config()
    cfg["output_folder"] = DEFAULT_SAVE_LOCATION
    save_config(cfg)
    window.findChild(QLineEdit, "lineEditDefaultSave").setText(DEFAULT_SAVE_LOCATION)
    window.findChild(QLineEdit, "lineEditOutputFolder").setText(DEFAULT_SAVE_LOCATION)


# ------------------------------------------------------------------
# Presets
# ------------------------------------------------------------------
def populate_preset_combo_boxes(window, remove_combo):
    edit_combo = window.findChild(QComboBox, "comboEditPreset")

    folder = config_path().parent
    presets = sorted(
        p for p in folder.glob("preset_*") if p.stem != DEFAULT_PRESET_STEM
    )

    for combo in (remove_combo, edit_combo):
        combo.clear()
        for item in presets:
            name = item.stem.removeprefix("preset_").replace("_", " ")
            combo.addItem(name, str(item))

    section = window.findChild(QWidget, "editRemoveWidget")
    if section is not None:
        section.setVisible(bool(presets))

def remove_preset(window,remove_combo):
    reply = QMessageBox.question(
        window,
        "Remove Preset",
        "Removing this preset will delete the preset file, and you will no longer be able to use this preset.\n\nWould you like to delete the preset file?",
        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        QMessageBox.StandardButton.Yes,
    )
    if reply != QMessageBox.StandardButton.Yes:
        return

    preset = preset_name_to_path(remove_combo.currentText())
    preset.unlink(missing_ok=True)
    populate_preset_combo_boxes(window, remove_combo)
    populate_preset_list(window)