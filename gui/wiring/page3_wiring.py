from PySide6.QtWidgets import (
    QSplitter,
    QListWidget,
    QAbstractItemView,
    QStackedWidget,
    QPushButton,
    QMessageBox,
)
from PySide6.QtGui import QIcon

_CIRCLE_EMPTY_ICON   = ":/res/icons/circle_dark.svg"
_CIRCLE_CHECKED_ICON = ":/res/icons/circle_checked.svg"


# ------------------------------------------------------------------
# Page 3 — Typography Advanced ("More Options") screen
# ------------------------------------------------------------------
def wire_page3(main_window):
    w = main_window.window

    splitterPage3 = w.findChild(QSplitter, "splitterTypographyMore")
    splitterPage3.setCollapsible(0, False)
    splitterPage3.setCollapsible(1, False)
    splitterPage3.setSizes([200, 700])

    main_window._groupList = w.findChild(QListWidget, "groupListWidget")
    main_window._groupList.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)

    empty_icon = QIcon(_CIRCLE_EMPTY_ICON)
    for i in range(main_window._groupList.count()):
        main_window._groupList.item(i).setIcon(empty_icon)

    main_window._groupContentStack = w.findChild(QStackedWidget, "groupContentStack")
    main_window._groupContentStack.setCurrentIndex(0)

    main_window._groupList.setCurrentRow(0)  # start on "Main Book"
    main_window._groupList.currentRowChanged.connect(
        lambda row: _on_typography_group_changed(main_window, row)
    )
    main_window._typography_group_checked = [False, False, False]
    main_window._typography_current_row = 0

    w.findChild(QPushButton, "buttonNext_4").clicked.connect(
        lambda: _on_typography_advanced_next(main_window)
    )
    w.findChild(QPushButton, "buttonBack_3").clicked.connect(
        lambda: _on_typography_advanced_back(main_window)
    )


def _on_typography_group_changed(main_window, row: int):
    if main_window._typography_current_row is not None:
        _mark_typography_group_checked(main_window, main_window._typography_current_row)

    main_window._typography_current_row = row
    main_window._groupContentStack.setCurrentIndex(row)

def _mark_typography_group_checked(main_window, row: int):
    main_window._typography_group_checked[row] = True
    main_window._groupList.item(row).setIcon(QIcon(_CIRCLE_CHECKED_ICON))

def _on_typography_advanced_next(main_window):
    current_row = main_window._groupList.currentRow()

    if current_row < main_window._groupList.count() - 1:
        main_window._groupList.setCurrentRow(current_row + 1)
        return

    _mark_typography_group_checked(main_window, current_row)

    if not all(main_window._typography_group_checked):
        reply = QMessageBox.question(
            main_window.window,
            "Incomplete Sections",
            "Some sections still have default settings. Would you like to continue anyways?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return  # user chose to go back and review — stay on this page

    main_window._go_to_page(4)

def _on_typography_advanced_back(main_window):
    reply = QMessageBox.question(
        main_window.window,
        "Discard Changes?",
        "Any per-style changes made here will be lost. Continue?",
        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
    )
    if reply != QMessageBox.StandardButton.Yes:
        return

    # Reset checklist state
    main_window._typography_group_checked = [False, False, False]
    main_window._typography_current_row = None
    empty_icon = QIcon(_CIRCLE_EMPTY_ICON)
    for i in range(main_window._groupList.count()):
        main_window._groupList.item(i).setIcon(empty_icon)

    main_window._groupList.setCurrentRow(0)
    main_window._groupContentStack.setCurrentIndex(0)

    main_window._go_back()
