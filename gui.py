import subprocess
import sys
import time
import tempfile
import urllib.request
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget,
    QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit, QCheckBox
)
from PyQt6.QtCore import QThread, pyqtSignal
from PyQt6.QtGui import QIcon
from pathlib import Path


# ── Constants ─────────────────────────────────────────────────────────────────

LO_PYTHON = r"C:\Program Files\LibreOffice\program\python.exe"
SCRIPT = Path(__file__).parent / "scripts" / "ao3_to_odt.py"
LO_DOWNLOAD_URL = "https://www.libreoffice.org/download/download-libreoffice/"
NO_WINDOW = subprocess.CREATE_NO_WINDOW  # suppress console windows on Windows


# ── Helpers ───────────────────────────────────────────────────────────────────

def get_base_path():
    """Return the correct base path whether running as .py or bundled .exe."""
    if hasattr(sys, '_MEIPASS'):
        return Path(sys._MEIPASS)
    return Path(__file__).parent


# ── Background worker ─────────────────────────────────────────────────────────

class ConversionWorker(QThread):
    log_signal = pyqtSignal(str)       # emits a line of text to the log
    finished_signal = pyqtSignal(bool) # emits True=success, False=failure

    def __init__(self, epub, odt, include_toc=True):
        super().__init__()
        self.epub = epub
        self.odt = odt
        self.include_toc = include_toc

    def run(self):
        # Kill any existing LO instances from previous runs
        subprocess.run(
            ["taskkill", "/f", "/im", "soffice.exe"],
            capture_output=True,
            creationflags=NO_WINDOW
        )
        time.sleep(2)  # give LO time to fully exit

        cmd = [LO_PYTHON, "-u", str(SCRIPT), self.epub, self.odt]
        if not self.include_toc:
            cmd.append("--no-toc")

        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding='utf-8',
            creationflags=NO_WINDOW
        )

        # Read character by character so partial lines (dots, etc.) show up live
        current_line = ""
        while True:
            char = process.stdout.read(1)
            if not char:
                break
            if char == "\n":
                if current_line.strip():
                    self.log_signal.emit(current_line.strip())
                current_line = ""
            elif char == "\r":
                pass  # ignore carriage returns
            else:
                current_line += char

        # Emit any remaining text that didn't end with a newline
        if current_line.strip():
            self.log_signal.emit(current_line.strip())

        try:
            process.wait(timeout=30)
        except subprocess.TimeoutExpired:
            process.kill()
            process.wait()

        # Clean up LO regardless of outcome
        subprocess.run(
            ["taskkill", "/f", "/im", "soffice.exe"],
            capture_output=True,
            creationflags=NO_WINDOW
        )
        time.sleep(3)  # give LO time to release file handles before PyInstaller cleans up

        self.finished_signal.emit(True)


# ── Main window ───────────────────────────────────────────────────────────────

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("AO3 to ODT Converter")
        self.setMinimumSize(600, 400)

        # Set window icon if available
        icon_path = get_base_path() / "icon.ico"
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))

        # Central widget — QMainWindow requires one as a container
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
        self.chk_toc.setChecked(True)  # on by default
        options_row.addWidget(self.chk_toc)
        options_row.addStretch()  # push checkbox to the left
        main_layout.addLayout(options_row)

        # ── Convert button ──
        self.btn_convert = QPushButton("Convert")
        self.btn_convert.clicked.connect(self.convert)
        main_layout.addWidget(self.btn_convert)

        # ── Log output ──
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        main_layout.addWidget(self.log)

    # ── Path helpers ──────────────────────────────────────────────────────────

    def get_downloads_folder(self):
        return str(Path.home() / "Downloads")

    def suggest_odt_path(self, epub):
        """Mirror the output filename logic from ao3_to_odt.py."""
        base = Path(epub).parent / (Path(epub).stem + "_book.odt")
        if not base.exists():
            return str(base)
        counter = 2
        while base.exists():
            base = base.with_stem(base.stem + f"_{counter}")
            counter += 1
        return str(base)

    # ── File dialogs ──────────────────────────────────────────────────────────

    def pick_epub(self):
        from PyQt6.QtWidgets import QFileDialog
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
        from PyQt6.QtWidgets import QFileDialog
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Save ODT as",
            self.save_input.text() or self.get_downloads_folder(),
            "ODT files (*.odt);;All files (*.*)"
        )
        if path:
            self.save_input.setText(path)

    # ── Dependency installer ──────────────────────────────────────────────────

    def ensure_lo_deps(self):
        """
        Install required libraries into LO's Python on first run.
        Uses a marker file in the temp folder so it only runs once.
        Falls back to downloading get-pip.py if ensurepip is unavailable
        (affects older LibreOffice versions shipping Python 3.8).
        Returns True if deps are ready, False if installation failed.
        """
        # Check LO is actually installed before doing anything else
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
            return True  # already installed on a previous run

        self.log.append("First run detected — installing required libraries into LibreOffice Python...")
        self.log.append("This will only happen once, please wait...")
        QApplication.processEvents()  # force UI to update before blocking

        try:
            # Try ensurepip first
            result = subprocess.run(
                [LO_PYTHON, "-m", "ensurepip", "--upgrade"],
                capture_output=True, timeout=60,
                creationflags=NO_WINDOW
            )

            # If ensurepip failed (common on LO with Python 3.8),
            # fall back to downloading get-pip.py directly
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

            # Now install the required packages
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

    # ── Conversion ────────────────────────────────────────────────────────────

    def convert(self):
        epub = self.epub_input.text()
        odt = self.save_input.text()

        if not epub or epub == "No file selected":
            self.log.append("ERROR: Please select an EPUB file first.")
            return
        if not odt or odt == "No save location selected":
            self.log.append("ERROR: Please select a save location first.")
            return

        self.btn_convert.setEnabled(False)
        self.log.clear()

        # Ensure LO deps are installed before attempting conversion
        if not self.ensure_lo_deps():
            self.btn_convert.setEnabled(True)
            return

        self.log.append("Starting conversion...")

        self.worker = ConversionWorker(epub, odt, self.chk_toc.isChecked())
        self.worker.log_signal.connect(self.log.append)
        self.worker.finished_signal.connect(self.on_finished)
        self.worker.start()

    def on_finished(self, success):
        if success:
            self.log.append("✓ Done!")
        else:
            self.log.append("Something went wrong. Check the log above.")
        self.btn_convert.setEnabled(True)


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())