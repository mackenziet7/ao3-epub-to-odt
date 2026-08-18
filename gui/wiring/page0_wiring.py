
from PySide6.QtWidgets import (
    QFileDialog,
    QButtonGroup,
    QListWidget,
    QPushButton,
    QSplitter,
    QLineEdit,
    QListWidgetItem,
)
from PySide6.QtCore import Qt
from pathlib import Path
from gui.config import (
    load_config,
    config_path
)

# ------------------------------------------------------------------
# Page 0 — main screen
# ------------------------------------------------------------------
def wire_page0(main_window):
    """
    Sets up main screen splitter and buttons
    """
    w = main_window.window

    config = load_config()
    output_folder = Path(config.get("output_folder", str(Path.home() / "Downloads")))

    if output_folder.is_dir():
        w.findChild(QLineEdit, "lineEditOutputFolder").setText(str(output_folder))
    else:
        w.findChild(QLineEdit, "lineEditOutputFolder").setText(str(Path.home() / "Downloads"))

    splitterPage0 = w.findChild(QSplitter, "splitter")
    splitterPage0.setCollapsible(0, False)
    splitterPage0.setCollapsible(1, False)
    splitterPage0.setSizes([500, 400])

    mode_group = QButtonGroup(w)
    for name in ("modeSingleButton", "modeBatchButton", "modeCollectionButton"):
        mode_group.addButton(w.findChild(QPushButton, name))
    mode_group.setExclusive(True)
    main_window._mode_group = mode_group

    # Add/remove files
    file_list = w.findChild(QListWidget, "listWidgetFileInput")
    w.findChild(QPushButton, "buttonAdd").clicked.connect(
        lambda: add_files(w, file_list)
    )
    w.findChild(QPushButton, "buttonRemove").clicked.connect(
        lambda: remove_selected_files(w, file_list)
    )

    # Output folder
    w.findChild(QPushButton, "buttonBrowse").clicked.connect(
        lambda: pick_folder(w)
    )

    # Presets
    
    w.findChild(QListWidget, "presetList").itemSelectionChanged.connect(
        lambda: update_convert_button(w)
    )
    populate_preset_list(w)
    update_convert_button(w) #update after populating


# ------------------------------------------------------------------
# Add and remove files
# ------------------------------------------------------------------
def add_files(window, file_list):
    """
    Opens a file picker that allows selecting multiple EPUBs at once.
    getOpenFileNames (plural) returns a list of paths rather than one.
    Duplicate entries are silently ignored.
    """
    paths, _ = QFileDialog.getOpenFileNames(
            window,
            "Select EPUB files",
            str(Path.home() / "Downloads"),
            "EPUB files (*.epub);;All files (*.*)"
        )
    existing = {
        file_list.item(i).text()
        for i in range(file_list.count())
    }

    for path in paths:
        if path not in existing:
            file_list.addItem(path)
            existing.add(path) # for duplicate prevention

    update_mode_buttons(window, file_list)

def remove_selected_files(window, file_list):
    """
    Removes whichever rows the user has selected.
    iterate in reverse order so that removing one item doesn't shift
    the indices of the items below it 
    """
    for item in reversed(file_list.selectedItems()): 
        file_list.takeItem(file_list.row(item))

    update_mode_buttons(window, file_list)

def update_mode_buttons(window, file_list):
    """
    Updates the single/batch/collection buttons based on number 
    of files selected
    """
    count = file_list.count()
    if count == 0:
        return
    elif count == 1:
        selected_button = window.findChild(QPushButton, "modeSingleButton")
    elif count > 1:
        selected_button = window.findChild(QPushButton, "modeBatchButton")

    selected_button.setChecked(True)

# ------------------------------------------------------------------
# Output folder selection
# ------------------------------------------------------------------
def pick_folder(window):
    folder = QFileDialog.getExistingDirectory(
            window,
            "Select output folder",
            str(Path.home() / "Downloads"),
        )
    if folder:
        window.findChild(QLineEdit, "lineEditOutputFolder").setText(folder)

# ------------------------------------------------------------------
# Presets
# ------------------------------------------------------------------
def populate_preset_list(window):
    folder = config_path().parent
    presets = list(folder.glob("preset_*"))
    print(presets)

    preset_list = window.findChild(QListWidget, "presetList")
    preset_list.clear()

    for item in presets:
        name = item.stem.removeprefix("preset_").replace("_", " ")
        preset_item = QListWidgetItem(name)
        preset_item.setData(Qt.ItemDataRole.UserRole, str(item)) # adds path to hidden data on each item
        preset_list.addItem(preset_item)

def update_convert_button(window):
    preset_list = window.findChild(QListWidget, "presetList")
    convert_button = window.findChild(QPushButton, "buttonConvert")
    convert_button.setEnabled(bool(preset_list.selectedItems()))
    