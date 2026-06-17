from pathlib import Path
from PySide6.QtUiTools import QUiLoader
from PySide6.QtWidgets import QApplication

import resources_rc
print(resources_rc.__file__)

from PySide6.QtGui import QIcon

app = QApplication([])

ui_path = Path(__file__).parent / "gui" / "ui" / "main_window.ui"

loader = QUiLoader()
window = loader.load(str(ui_path))

# Prevent right panel from fully collapsing
splitter = window.findChild(__import__('PySide6.QtWidgets', fromlist=['QSplitter']).QSplitter, 'splitter')
splitter.setCollapsible(0, False)  # 0 = left panel
splitter.setCollapsible(1, False)  # 1 = right panel

# Make mode buttons mutually exclusive
from PySide6.QtWidgets import QButtonGroup, QPushButton
modeGroup = QButtonGroup(window)
modeGroup.addButton(window.findChild(QPushButton, 'modeSingleButton'))
modeGroup.addButton(window.findChild(QPushButton, 'modeBatchButton'))
modeGroup.addButton(window.findChild(QPushButton, 'modeCollectionButton'))
modeGroup.setExclusive(True)

window.show()
app.setStyleSheet(open("gui/ui/styles/dark.qss").read())
app.exec()