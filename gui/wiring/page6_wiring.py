from PySide6.QtWidgets import QComboBox, QPushButton
from PySide6.QtCore import QUrl
from PySide6.QtGui import QDesktopServices

from gui.config import config_path

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
