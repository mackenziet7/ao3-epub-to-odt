from pathlib import Path
from PySide6.QtUiTools import QUiLoader
from PySide6.QtWidgets import (
    QApplication, QSplitter, QButtonGroup, QPushButton,
    QCheckBox, QComboBox, QLabel, QDoubleSpinBox, QWidget
)
from gui.ui import resources_rc

app = QApplication([])

ui_path = Path(__file__).parent / "gui" / "ui" / "screens" / "main_window.ui"
loader = QUiLoader()
window = loader.load(str(ui_path))

# ── Page 0: main splitter ──────────────────────────────────────────
splitter = window.findChild(QSplitter, 'splitter')
splitter.setCollapsible(0, False)
splitter.setCollapsible(1, False)

# Make mode buttons mutually exclusive
modeGroup = QButtonGroup(window)
modeGroup.addButton(window.findChild(QPushButton, 'modeSingleButton'))
modeGroup.addButton(window.findChild(QPushButton, 'modeBatchButton'))
modeGroup.addButton(window.findChild(QPushButton, 'modeCollectionButton'))
modeGroup.setExclusive(True)

# ── Page 1: page setup splitter ────────────────────────────────────
splitterSetup = window.findChild(QSplitter, 'splitterSetup')
splitterSetup.setCollapsible(0, False)
splitterSetup.setCollapsible(1, False)

# ── Page navigation ────────────────────────────────────────────────
from PySide6.QtWidgets import QStackedWidget
stack = window.findChild(QStackedWidget, 'stackedWidget')

window.findChild(QPushButton, 'buttonNext').clicked.connect(
    lambda: stack.setCurrentIndex(1)
)
window.findChild(QPushButton, 'buttonBack').clicked.connect(
    lambda: stack.setCurrentIndex(0)
)

# ── Custom page size visibility ────────────────────────────────────
comboPageSize = window.findChild(QComboBox, 'comboPageSize')
customSizeWidget = window.findChild(QWidget, 'customSizeWidget')

comboPageSize.currentIndexChanged.connect(
    lambda: customSizeWidget.setVisible(comboPageSize.currentText().startswith('Custom'))
)

# ── Mirrored margins label swap ────────────────────────────────────
labelMarginLeft = window.findChild(QLabel, 'labelMarginLeft')
labelMarginRight = window.findChild(QLabel, 'labelMarginRight')
checkMirrored = window.findChild(QCheckBox, 'checkMirroredMargins')

def on_mirrored_toggled(checked):
    labelMarginLeft.setText('Inside' if checked else 'Left')
    labelMarginRight.setText('Outside' if checked else 'Right')

checkMirrored.toggled.connect(on_mirrored_toggled)
on_mirrored_toggled(True)  # apply initial state

# ── Units toggle ───────────────────────────────────────────────────
comboUnits = window.findChild(QComboBox, 'comboUnits')

SPIN_NAMES = ['spinMarginTop', 'spinMarginBottom', 'spinMarginLeft',
              'spinMarginRight', 'spinCustomWidth', 'spinCustomHeight']

CM_TO_IN = 1 / 2.54
IN_TO_CM = 2.54

def on_units_changed(index):
    unit = comboUnits.currentText()
    suffix = f' {unit}'
    for name in SPIN_NAMES:
        spin = window.findChild(QDoubleSpinBox, name)
        old_val = spin.value()
        spin.setSuffix(suffix)
        if unit == 'in':
            spin.setDecimals(3)
            spin.setMaximum(4.0 if 'Custom' not in name else 40.0)
            spin.setSingleStep(0.05)
            spin.setValue(round(old_val * CM_TO_IN, 3))
        else:
            spin.setDecimals(2)
            spin.setMaximum(10.0 if 'Custom' not in name else 100.0)
            spin.setSingleStep(0.1)
            spin.setValue(round(old_val * IN_TO_CM, 2))

comboUnits.currentIndexChanged.connect(on_units_changed)

# ── Launch ─────────────────────────────────────────────────────────
app.setStyleSheet(open("gui/ui/res/styles/dark.qss").read())
window.show()
app.exec()