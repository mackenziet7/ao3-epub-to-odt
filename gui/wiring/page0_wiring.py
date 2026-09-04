
import shutil

from PySide6.QtWidgets import (
    QFileDialog,
    QButtonGroup,
    QListWidget,
    QPushButton,
    QSplitter,
    QLineEdit,
    QListWidgetItem,
    QStackedWidget,
)
from PySide6.QtCore import Qt
from pathlib import Path
from gui.config import (
    default_preset_source_path,
    load_config,
    config_path,
    load_preset,
    lo_missing_reason,
    preset_name_to_path,
)

from gui.wiring.page5_wiring import start_conversion

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
    for name in ("modeSingleButton", "modeBatchButton"):
        mode_group.addButton(w.findChild(QPushButton, name))
    mode_group.setExclusive(True)
    main_window._mode_group = mode_group

    # Add/remove files
    file_list = w.findChild(QListWidget, "listWidgetFileInput")
    w.findChild(QPushButton, "buttonAdd").clicked.connect(
        lambda: add_files(w, file_list, main_window)
    )
    w.findChild(QPushButton, "buttonRemove").clicked.connect(
        lambda: remove_selected_files(w, file_list, main_window)
    )

    # Output folder
    w.findChild(QPushButton, "buttonBrowse").clicked.connect(
        lambda: pick_folder(w)
    )
    w.findChild(QLineEdit, "lineEditOutputFolder").textChanged.connect(
        lambda _text: update_action_buttons(w, main_window)
    )

    # Presets
    
    w.findChild(QListWidget, "presetList").itemSelectionChanged.connect(
        lambda: update_action_buttons(w, main_window)
    )
    populate_preset_list(w)
    update_action_buttons(w, main_window)  #update after populating

    w.findChild(QPushButton, "buttonConvert").clicked.connect(
        lambda: _quick_convert(main_window)
    )


# ------------------------------------------------------------------
# Add and remove files
# ------------------------------------------------------------------
def add_files(window, file_list, main_window):
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
    update_action_buttons(window, main_window)

def remove_selected_files(window, file_list, main_window):
    """
    Removes whichever rows the user has selected.
    iterate in reverse order so that removing one item doesn't shift
    the indices of the items below it 
    """
    for item in reversed(file_list.selectedItems()): 
        file_list.takeItem(file_list.row(item))

    update_mode_buttons(window, file_list)
    update_action_buttons(window, main_window)

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
    _ensure_default_preset_exists()

    folder = config_path().parent
    presets = list(folder.glob("preset_*"))

    preset_list = window.findChild(QListWidget, "presetList")
    preset_list.clear()

    for item in presets:
        name = item.stem.removeprefix("preset_").replace("_", " ")
        preset_item = QListWidgetItem(name)
        preset_item.setData(Qt.ItemDataRole.UserRole, str(item)) # adds path to hidden data on each item
        preset_list.addItem(preset_item)

def update_action_buttons(window, main_window):
    file_list = window.findChild(QListWidget, "listWidgetFileInput")
    preset_list = window.findChild(QListWidget, "presetList")
    folder_text = window.findChild(QLineEdit, "lineEditOutputFolder").text()

    has_files = file_list.count() > 0
    has_valid_folder = Path(folder_text).is_dir()
    has_preset = bool(preset_list.selectedItems())

    next_button = window.findChild(QPushButton, "buttonNext")
    convert_button = window.findChild(QPushButton, "buttonConvert")
    lo_reason = lo_missing_reason(main_window)

    next_reasons = []
    if not has_files:
        next_reasons.append("Add at least one EPUB file")
    if not has_valid_folder:
        next_reasons.append("Select a valid output folder")
    if lo_reason is not None:
        next_reasons.append(lo_reason)

    convert_reasons = list(next_reasons)
    if not has_preset:
        convert_reasons.append("Select a preset")

    next_button.setEnabled(not next_reasons)
    next_button.setToolTip("\n".join(next_reasons))

    convert_button.setEnabled(not convert_reasons)
    convert_button.setToolTip("\n".join(convert_reasons))


def _quick_convert(main_window):
    w = main_window.window
    preset_list = w.findChild(QListWidget, "presetList")
    selected = preset_list.selectedItems()
    if not selected:
        return  # button is disabled without a selection, but guard anyway

    settings = load_preset(selected[0].data(Qt.ItemDataRole.UserRole))

    file_list = w.findChild(QListWidget, "listWidgetFileInput")
    epub_paths = [file_list.item(i).text() for i in range(file_list.count())]
    output_folder = w.findChild(QLineEdit, "lineEditOutputFolder").text()

    w.findChild(QStackedWidget, "stackedWidget").setCurrentIndex(5)
    start_conversion(main_window, epub_paths, output_folder, settings) 

def _ensure_default_preset_exists():
    dest = preset_name_to_path("Default Book Layout")
    if not dest.exists():
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy(default_preset_source_path(), dest)