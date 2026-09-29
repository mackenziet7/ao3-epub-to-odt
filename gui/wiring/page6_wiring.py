from pathlib import Path

from PySide6.QtWidgets import QComboBox, QFileDialog, QLineEdit, QPushButton
from PySide6.QtCore import QUrl
from PySide6.QtGui import QDesktopServices

from gui.config import config_path, DEFAULT_SAVE_LOCATION, load_config, save_config

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

    combo_theme = w.findChild(QComboBox, "comboTheme")
    combo_theme.setCurrentText("Dark" if main_window.theme == "dark" else "Light")
    combo_theme.currentTextChanged.connect(
        lambda theme: main_window.set_theme(theme)
    )

    w.findChild(QPushButton, "buttonBrowse_2").clicked.connect(
        lambda: pick_folder(w)
    )

    cfg = load_config()

    w.findChild(QLineEdit, "lineEditDefaultSave").setText(cfg["output_folder"])
    w.findChild(QPushButton, "buttonSaveReset").clicked.connect(
        lambda: save_folder_reset(w)
    )

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
