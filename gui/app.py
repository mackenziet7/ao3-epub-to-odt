import subprocess
import sys

import time
import tempfile
import urllib.request
from pathlib import Path

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget,
    QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit, QCheckBox, QFileDialog
)
from PyQt6.QtCore import QThread, pyqtSignal
from PyQt6.QtGui import QIcon

from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))
from gui.worker import ConversionWorker


# ── Constants ──────────────────────────────────────────────────────────────────

LO_PYTHON    = r"C:\Program Files\LibreOffice\program\python.exe"
LO_DOWNLOAD_URL = "https://www.libreoffice.org/download/download-libreoffice/"
NO_WINDOW    = subprocess.CREATE_NO_WINDOW

def get_base_path():
    """Return the correct base path whether running as .py or bundled .exe."""
    if hasattr(sys, '_MEIPASS'):
        return Path(sys._MEIPASS)
    return Path(__file__).parent.parent   # repo root


def get_script_path():
    if hasattr(sys, '_MEIPASS'):
        return Path(sys._MEIPASS) / "scripts" / "script_ao3_to_odt.py"
    return Path(__file__).parent.parent / "scripts" / "script_ao3_to_odt.py"


# ── Main window ────────────────────────────────────────────────────────────────

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("AO3 to ODT Converter")
        self.setMinimumSize(600, 400)

        icon_path = get_base_path() / "icon.ico"
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))

        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)

        # ── EPUB row ──
        epub_row = QHBoxLayout()
        epub_row.addWidget(QLabel("EPUB file:"))
        self.epub_input = QLineEdit()
        self.epub_input.setPlaceholderText("No file selected")
        self.epub_input.setReadOnly(True)
        epub_row.addWidget(self.epub_input)
        self.btn_epub = QPushButton("Browse")
        self.btn_epub.clicked.connect(self.pick_epub)
        epub_row.addWidget(self.btn_epub)
        main_layout.addLayout(epub_row)

        # ── Save row ──
        save_row = QHBoxLayout()
        save_row.addWidget(QLabel("Save as:"))
        self.save_input = QLineEdit()
        self.save_input.setPlaceholderText("No save location selected")
        self.save_input.setReadOnly(True)
        save_row.addWidget(self.save_input)
        self.btn_save = QPushButton("Browse")
        self.btn_save.clicked.connect(self.pick_save)
        save_row.addWidget(self.btn_save)
        main_layout.addLayout(save_row)

        # ── Options row ──
        options_row = QHBoxLayout()
        self.chk_toc = QCheckBox("Include Table of Contents")
        self.chk_toc.setChecked(True)
        options_row.addWidget(self.chk_toc)
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

    # ── Path helpers ───────────────────────────────────────────────────────────

    def get_downloads_folder(self):
        return str(Path.home() / "Downloads")

    def suggest_odt_path(self, epub):
        base = Path(epub).parent / (Path(epub).stem + "_book.odt")
        if not base.exists():
            return str(base)
        counter = 2
        original_stem = base.stem  
        while base.exists():
            base = base.with_stem(original_stem + f"_{counter}")  
            counter += 1
        return str(base)

    # ── File dialogs ───────────────────────────────────────────────────────────

    def pick_epub(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Select your EPUB file",
            self.get_downloads_folder(),
            "EPUB files (*.epub);;All files (*.*)"
        )
        if path:
            self.epub_input.setText(path)
            self.save_input.setText(self.suggest_odt_path(path))

    def pick_save(self):
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Save ODT as",
            self.save_input.text() or self.get_downloads_folder(),
            "ODT files (*.odt);;All files (*.*)"
        )
        if path:
            self.save_input.setText(path)

    # ── Dependency installer ───────────────────────────────────────────────────

    def ensure_lo_deps(self):
        if not Path(LO_PYTHON).exists():
            self.log.append("ERROR: LibreOffice does not appear to be installed.")
            self.log.append("")
            self.log.append("This tool requires LibreOffice to be installed at:")
            self.log.append(f"  {LO_PYTHON}")
            self.log.append("")
            self.log.append(f"Download it from: {LO_DOWNLOAD_URL}")
            self.log.append("")
            self.log.append("After installing, restart this application and try again.")
            return False

        marker = Path(tempfile.gettempdir()) / "ao3_odt_deps_installed.txt"
        if marker.exists():
            return True

        self.log.append("First run detected — installing required libraries into LibreOffice Python...")
        self.log.append("This will only happen once, please wait...")
        QApplication.processEvents()

        try:
            result = subprocess.run(
                [LO_PYTHON, "-m", "ensurepip", "--upgrade"],
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
                    [LO_PYTHON, str(get_pip)],
                    capture_output=True, timeout=120,
                    creationflags=NO_WINDOW
                )

            result = subprocess.run(
                [LO_PYTHON, "-m", "pip", "install", "ebooklib", "beautifulsoup4", "lxml", "-q"],
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
        epub = self.epub_input.text()
        odt  = self.save_input.text()

        if not epub or epub == "No file selected":
            self.log.append("ERROR: Please select an EPUB file first.")
            return
        if not odt or odt == "No save location selected":
            self.log.append("ERROR: Please select a save location first.")
            return

        self.btn_convert.setEnabled(False)
        self.log.clear()

        if not self.ensure_lo_deps():
            self.btn_convert.setEnabled(True)
            return

        self.log.append("Starting conversion...")

        self.worker = ConversionWorker(
            LO_PYTHON,
            get_script_path(),
            epub, odt,
            self.chk_toc.isChecked()
        )
        self.worker.log_signal.connect(self.log.append)
        self.worker.finished_signal.connect(self.on_finished)
        self.worker.start()

    def on_finished(self, success):
        if success:
            self.log.append("✓ Done!")
        else:
            self.log.append("Something went wrong. Check the log above.")
        self.btn_convert.setEnabled(True)