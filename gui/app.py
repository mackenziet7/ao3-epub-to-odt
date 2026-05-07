# gui/app.py
import subprocess
import sys
import tempfile
import urllib.request
from pathlib import Path

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget,
    QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit, QCheckBox, QFileDialog,
    QListWidget, QAbstractItemView, QProgressBar
)
from PyQt6.QtGui import QIcon

sys.path.insert(0, str(Path(__file__).parent.parent))
from gui.worker import ConversionWorker
from gui.config import resolve_lo_python, load_config
from gui.first_run_dialog import LOPathDialog


# ── Constants ──────────────────────────────────────────────────────────────────
NO_WINDOW = subprocess.CREATE_NO_WINDOW


def get_base_path():
    if hasattr(sys, '_MEIPASS'):
        return Path(sys._MEIPASS)
    return Path(__file__).parent.parent


def get_script_path():
    if hasattr(sys, '_MEIPASS'):
        return Path(sys._MEIPASS) / "scripts" / "script_ao3_to_odt.py"
    return Path(__file__).parent.parent / "scripts" / "script_ao3_to_odt.py"


# ── Main window ────────────────────────────────────────────────────────────────

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("AO3 to ODT Converter")
        self.setMinimumSize(600, 500)

        icon_path = get_base_path() / "icon.ico"
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))

        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)

        # ── EPUB queue ────────────────────────────────────────────────────────
        main_layout.addWidget(QLabel("EPUB files to convert:"))
        self.epub_list = QListWidget()

        # ExtendedSelection lets the user select multiple items at once
        self.epub_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        main_layout.addWidget(self.epub_list)

        # Add / Remove buttons sit below the list
        list_btn_row = QHBoxLayout()
        self.btn_add = QPushButton("Add EPUBs")
        self.btn_add.clicked.connect(self.add_epubs)
        list_btn_row.addWidget(self.btn_add)

        self.btn_remove = QPushButton("Remove Selected")
        self.btn_remove.clicked.connect(self.remove_selected)
        list_btn_row.addWidget(self.btn_remove)

        list_btn_row.addStretch()  # pushes buttons to the left
        main_layout.addLayout(list_btn_row)

        # ── Output folder row ─────────────────────────────────────────────────
        folder_row = QHBoxLayout()
        folder_row.addWidget(QLabel("Output folder:"))
        self.folder_input = QLineEdit()
        self.folder_input.setPlaceholderText("No folder selected")
        self.folder_input.setReadOnly(True)
        self.folder_input.setText(str(Path.home() / "Downloads"))
        folder_row.addWidget(self.folder_input)
        self.btn_folder = QPushButton("Browse")
        self.btn_folder.clicked.connect(self.pick_folder)
        folder_row.addWidget(self.btn_folder)
        main_layout.addLayout(folder_row)

        # ── Options row ───────────────────────────────────────────────────────
        options_row = QHBoxLayout()
        self.chk_toc = QCheckBox("Include Table of Contents")
        self.chk_toc.setChecked(True)
        options_row.addWidget(self.chk_toc)
        self.chk_qr = QCheckBox("Include QR code")
        self.chk_qr.setChecked(True)
        options_row.addWidget(self.chk_qr)
        options_row.addStretch()
        main_layout.addLayout(options_row)

        # ── Convert button ──
        self.btn_convert = QPushButton("Convert")
        self.btn_convert.clicked.connect(self.convert)
        main_layout.addWidget(self.btn_convert)

        # ── Log output ──
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        main_layout.addWidget(self.log)

        # ── Progress bar ──────────────────────────────────────────────────────────
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        # This tells Qt to show "1 / 3" style text instead of a percentage
        self.progress.setFormat("%v / %m files")
        main_layout.addWidget(self.progress)

        # ── Conversion state ──────────────────────────────────────────────────
        # These three lists track the queue and results during a conversion run.
        self.queue = []      # full paths of EPUBs still waiting to be processed
        self.succeeded = []  # paths that converted successfully
        self.failed = []     # paths that failed

        # worker holds the current ConversionWorker thread, or None when idle
        self.worker = None

        # ── Resolve LO path on startup ────────────────────────────────────────
        self.lo_python = self._resolve_lo_on_startup()
        if self.lo_python is None:
            self.btn_convert.setEnabled(False)
            self.log.append("LibreOffice path not set — please restart the app.")

    # ── LO path resolution ─────────────────────────────────────────────────────

    def _resolve_lo_on_startup(self) -> Path | None:
        """
        Called once at startup. Tries to find LO's python.exe via the config.
        If it can't, shows the first-run dialog.
        Returns the resolved Path, or None if the user closed the dialog.
        """
        lo = resolve_lo_python()
        if lo is not None:
            return lo

        # resolve_lo_python() returns None in two cases:
        #   1. No config file yet (first run)
        #   2. Config has a path but the file no longer exists (LO moved/uninstalled)
        cfg = load_config()
        invalid = cfg.get("lo_python")  # will be None on first run, a string on case 2

        dialog = LOPathDialog(self, invalid_path=invalid)
        if dialog.exec():
            return Path(load_config()["lo_python"])

        return None

    # ── File list management ───────────────────────────────────────────────────

    def add_epubs(self):
        """
        Opens a file picker that allows selecting multiple EPUBs at once.
        getOpenFileNames (plural) returns a list of paths rather than one.
        Duplicate entries are silently ignored.
        """
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select EPUB files",
            str(Path.home() / "Downloads"),
            "EPUB files (*.epub);;All files (*.*)"
        )
        # Get the set of paths already in the list 
        existing = {
            self.epub_list.item(i).text()
            for i in range(self.epub_list.count())
        }
        for path in paths:
            if path not in existing:
                self.epub_list.addItem(path)

    def remove_selected(self):
        """
        Removes whichever rows the user has selected.
        iterate in reverse order so that removing one item doesn't shift
        the indices of the items below it 
        """
        for item in reversed(self.epub_list.selectedItems()):
            self.epub_list.takeItem(self.epub_list.row(item))

    # ── Output folder ──────────────────────────────────────────────────────────

    def pick_folder(self):
        folder = QFileDialog.getExistingDirectory(
            self,
            "Select output folder",
            str(Path.home() / "Downloads"),
        )
        if folder:
            self.folder_input.setText(folder)

    def suggest_odt_path(self, epub: str, output_folder: str) -> str:
        """
        Builds a non-colliding output path for the ODT file.
        e.g. given epub="my_fic.epub" and folder="C:/output",
        returns "C:/output/my_fic_book.odt", or
                "C:/output/my_fic_book_2.odt" if that already exists, etc.
        """
        base = Path(output_folder) / (Path(epub).stem + "_book.odt")
        if not base.exists():
            return str(base)
        counter = 2
        original_stem = base.stem
        while base.exists():
            base = base.with_stem(original_stem + f"_{counter}")
            counter += 1
        return str(base)

    # ── Dependency installer ───────────────────────────────────────────────────

    def ensure_lo_deps(self) -> bool:
        """
        Makes sure the required Python packages are installed into LO's Python.
        Returns True if everything is ready, False if something went wrong.
        Only runs the full install once — after that the marker file short-circuits it.
        """
        if self.lo_python is None:
            self.log.append("ERROR: No LibreOffice installation found.")
            self.log.append("Please restart the app to set the path.")
            return False

        if not self.lo_python.exists():
            self.log.append("ERROR: LibreOffice could not be found.")
            self.log.append(f"  {self.lo_python}")
            self.log.append(f"Download it from: {LO_DOWNLOAD_URL}")
            return False

        marker = Path(tempfile.gettempdir()) / "ao3_odt_deps_v2.txt"
        if marker.exists():
            return True

        self.log.append("First run detected — installing required libraries into LibreOffice Python...")
        self.log.append("This will only happen once, please wait...")
        QApplication.processEvents()

        try:
            result = subprocess.run(
                [self.lo_python, "-m", "ensurepip", "--upgrade"],
                capture_output=True, timeout=60,
                creationflags=NO_WINDOW
            )

            if result.returncode != 0:
                self.log.append("ensurepip unavailable, trying alternative pip install...")
                self.log.append("(This requires an internet connection)")
                QApplication.processEvents()
                get_pip = Path(tempfile.gettempdir()) / "get-pip.py"
                urllib.request.urlretrieve("https://bootstrap.pypa.io/get-pip.py", get_pip)
                subprocess.run(
                    [self.lo_python, str(get_pip)],
                    capture_output=True, timeout=120,
                    creationflags=NO_WINDOW
                )

            result = subprocess.run(
                [self.lo_python, "-m", "pip", "install",
                 "ebooklib", "beautifulsoup4", "lxml", "qrcode[pil]", "-q"],
                capture_output=True, text=True, timeout=120,
                creationflags=NO_WINDOW
            )
            if result.returncode == 0:
                marker.write_text("installed")
                self.log.append("Libraries installed successfully!")
                return True
            else:
                self.log.append("ERROR: Failed to install libraries.")
                self.log.append(result.stderr)
                return False
        except Exception as e:
            self.log.append(f"ERROR during library installation: {e}")
            return False

    # ── Conversion ─────────────────────────────────────────────────────────────

    def convert(self):
        """
        Validates inputs, builds the queue, and kicks off the first conversion.
        The rest of the queue is chained through on_finished().
        """
        if self.lo_python is None:
            self.log.append("ERROR: LibreOffice path is not set.")
            self.log.append("Please restart the app to set the path.")
            return

        if self.epub_list.count() == 0:
            self.log.append("ERROR: Please add at least one EPUB file.")
            return

        folder = self.folder_input.text()
        if not folder or folder == "No folder selected":
            self.log.append("ERROR: Please select an output folder.")
            return

        self.btn_convert.setEnabled(False)
        self.log.clear()

        if not self.ensure_lo_deps():
            self.btn_convert.setEnabled(True)
            return

        # Build the queue from whatever is in the list widget.
        self.queue = [
            self.epub_list.item(i).text()
            for i in range(self.epub_list.count())
        ]
        self.succeeded = []
        self.failed = []

        self.progress.setMaximum(len(self.queue))  # total number of files
        self.progress.setValue(0)                  # start at zero
        self.progress.setVisible(True)

        self.log.append(f"Starting conversion of {len(self.queue)} file(s)...")
        self._start_next()

    def _start_next(self):
        """
        Starts a ConversionWorker for the first item in self.queue.
        Called by convert() for the first file, then by on_finished() for each
        subsequent file, so we never run two conversions at the same time.
        """
        epub = self.queue[0]
        odt = self.suggest_odt_path(epub, self.folder_input.text())

        self.log.append(f"\nConverting: {Path(epub).name}")

        # Clean up the previous worker if there was one.
        if self.worker is not None:
            self.worker.log_signal.disconnect()
            self.worker.finished_signal.disconnect()

        self.worker = ConversionWorker(
            self.lo_python,
            get_script_path(),
            epub, odt,
            self.chk_toc.isChecked(),
            self.chk_qr.isChecked(),
        )
        self.worker.log_signal.connect(self.log.append)
        self.worker.finished_signal.connect(self.on_finished)
        self.worker.start()

    def on_finished(self, success: bool):
        """
        Called by the worker when a single conversion finishes.
        Records the result, then either starts the next file or shows the summary.
        """
        current = self.queue.pop(0)

        if success:
            self.succeeded.append(current)
        else:
            self.failed.append(current)
            # Immediate per-file error so the user sees it in context in the log
            self.log.append(f"✗ Failed: {Path(current).name}")

        self.progress.setValue(len(self.succeeded) + len(self.failed))

        if self.queue:
            self._start_next()  # more files to process
        else:
            self._show_summary()
            self.btn_convert.setEnabled(True)
            self.progress.setVisible(False)

    def _show_summary(self):
        """
        Prints a tidy summary at the end of a full queue run.
        Shown whether everything succeeded, everything failed, or a mix.
        """
        self.log.append("")
        self.log.append("── Conversion complete ──────────────────")
        self.log.append(f"✓ {len(self.succeeded)} file(s) converted successfully")
        if self.failed:
            self.log.append(f"✗ {len(self.failed)} file(s) failed:")
            for f in self.failed:
                self.log.append(f"  - {Path(f).name}")