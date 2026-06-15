
from pathlib import Path
from PySide6.QtUiTools import QUiLoader
from PySide6.QtWidgets import QApplication

import resources_rc
print(resources_rc.__file__)

from PySide6.QtGui import QIcon

icon = QIcon(":/icons/settings_dark.svg")
print("isNull:", icon.isNull())

app = QApplication([])

ui_path = Path(__file__).parent / "gui" / "ui" / "main_window.ui"

loader = QUiLoader()
window = loader.load(str(ui_path))

window.show()
app.setStyleSheet(open("gui/ui/styles/dark.qss").read())
app.exec()