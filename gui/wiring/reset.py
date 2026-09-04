# gui/wiring/reset.py
from PySide6.QtWidgets import (
    QWidget, QDoubleSpinBox, QComboBox, QCheckBox, QLineEdit,
    QRadioButton, QFontComboBox, QPushButton, QListWidget, QStackedWidget,
)
from PySide6.QtGui import QFont

# Only these four pages hold wizard state that should be reset.
# page_0 (quick convert), page_5 (progress), page_6 (settings) are untouched.
WIZARD_PAGE_NAMES = ["page_1", "page_2", "page_3", "page_4"]

_GETTERS = {
    QDoubleSpinBox: lambda w: w.value(),
    QComboBox: lambda w: w.currentIndex(),
    QCheckBox: lambda w: w.isChecked(),
    QLineEdit: lambda w: w.text(),
    QRadioButton: lambda w: w.isChecked(),
    QFontComboBox: lambda w: w.currentFont().family(),
    QPushButton: lambda w: w.isChecked(),  # only meaningful for the checkable lock buttons
}
_SETTERS = {
    QDoubleSpinBox: lambda w, v: w.setValue(v),
    QComboBox: lambda w, v: w.setCurrentIndex(v),
    QCheckBox: lambda w, v: w.setChecked(v),
    QLineEdit: lambda w, v: w.setText(v),
    QRadioButton: lambda w, v: w.setChecked(v),
    QFontComboBox: lambda w, v: w.setCurrentFont(QFont(v)),
    QPushButton: lambda w, v: w.setChecked(v),
}


def _wizard_containers(main_window):
    containers = []
    for name in WIZARD_PAGE_NAMES:
        page = main_window.window.findChild(QWidget, name)
        if page is not None:
            containers.append(page)
    return containers


def snapshot_wizard_defaults(main_window):
    """Call once at startup, after all wire_pageN() calls, before any user edits."""
    snapshot = {}
    for page in _wizard_containers(main_window):
        for widget_type, getter in _GETTERS.items():
            for widget in page.findChildren(widget_type):
                name = widget.objectName()
                if name:
                    snapshot[name] = (widget_type, getter(widget))

    # Non-generic selection state (list row / stacked page) for the
    # groupListWidget <-> groupContentStack pairing on page_3.
    group_list = main_window.window.findChild(QListWidget, "groupListWidget")
    group_stack = main_window.window.findChild(QStackedWidget, "groupContentStack")
    snapshot["__groupListWidget_row"] = group_list.currentRow() if group_list else 0
    snapshot["__groupContentStack_index"] = group_stack.currentIndex() if group_stack else 0

    main_window._wizard_defaults = snapshot


def restore_wizard_defaults(main_window):
    for name, entry in main_window._wizard_defaults.items():
        if name.startswith("__"):
            continue
        widget_type, value = entry
        widget = main_window.window.findChild(widget_type, name)
        if widget is not None:
            _SETTERS[widget_type](widget, value)

    group_list = main_window.window.findChild(QListWidget, "groupListWidget")
    group_stack = main_window.window.findChild(QStackedWidget, "groupContentStack")
    if group_list is not None:
        group_list.setCurrentRow(main_window._wizard_defaults.get("__groupListWidget_row", 0))
    if group_stack is not None:
        group_stack.setCurrentIndex(main_window._wizard_defaults.get("__groupContentStack_index", 0))

def reset_ui(main_window):
    restore_wizard_defaults(main_window)

    file_list = main_window.window.findChild(QListWidget, "listWidgetFileInput")
    if file_list is not None:
        file_list.clear()

    w = main_window.window
    w.findChild(QPushButton, "buttonNext").setEnabled(False)
    w.findChild(QPushButton, "buttonConvert").setEnabled(False)
