import subprocess
import sys
import time
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget,
    QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit
)
from PyQt6.QtCore import QThread, pyqtSignal
from pathlib import Path

LO_PYTHON = r"C:\Program Files\LibreOffice\program\python.exe"
def get_base_path():
    # PyInstaller sets sys._MEIPASS when running as an exe
    # Fall back to the script's directory when running normally
    if hasattr(sys, '_MEIPASS'):
        return Path(sys._MEIPASS)
    return Path(__file__).parent

SCRIPT = get_base_path() / "scripts" / "ao3_to_odt.py"

class ConversionWorker(QThread):
    log_signal = pyqtSignal(str)      # emits a string to the log
    finished_signal = pyqtSignal(bool) # emits True=success, False=failure

    def __init__(self, epub, odt):
        super().__init__()
        self.epub = epub
        self.odt = odt

    def run(self):
        # Kill any existing LO first and wait for it to die
        subprocess.run(["taskkill", "/f", "/im", "soffice.exe"], capture_output=True)
        import time
        time.sleep(2)  # give it time to fully exit

        process = subprocess.Popen(
            [LO_PYTHON, "-u", str(SCRIPT), self.epub, self.odt],  # -u = unbuffered
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding='utf-8'
        )

        # Read character by character instead of line by line
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
        
        # Emit any remaining text
        if current_line.strip():
            self.log_signal.emit(current_line.strip())

        try:
            process.wait(timeout=30)
        except subprocess.TimeoutExpired:
            process.kill()
            process.wait()

        subprocess.run(["taskkill", "/f", "/im", "soffice.exe"], capture_output=True)
        self.finished_signal.emit(True)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("AO3 to ODT Converter")
        self.setMinimumSize(600, 400)

        # Central widget — QMainWindow requires one as a container
        central = QWidget()
        self.setCentralWidget(central)

        # Main vertical layout
        main_layout = QVBoxLayout(central)

        # ── EPUB row ──
        epub_row = QHBoxLayout()
        epub_row.addWidget(QLabel("EPUB file:"))
        self.epub_input = QLineEdit()
        self.epub_input.setPlaceholderText("No file selected")
        self.epub_input.setReadOnly(True)
        epub_row.addWidget(self.epub_input)
        self.btn_epub = QPushButton("Browse")
        self.btn_epub.clicked.connect(self.pick_epub)   # connect signal to slot
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
        self.btn_save.clicked.connect(self.pick_save)   # connect signal to slot
        save_row.addWidget(self.btn_save)
        main_layout.addLayout(save_row)

        # ── Convert ──
        self.btn_convert = QPushButton("Convert")
        self.btn_convert.clicked.connect(self.convert)  # connect signal to slot
        main_layout.addWidget(self.btn_convert)

        # ── Log output ──
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        main_layout.addWidget(self.log)

    def get_downloads_folder(self):
        return str(Path.home() / "Downloads")

    def suggest_odt_path(self, epub):
        from pathlib import Path
        base = Path(epub).parent / (Path(epub).stem + "_book.odt")
        if not base.exists():
            return str(base)
        counter = 2
        while base.exists():
            base = base.with_stem(base.stem + f"_{counter}")
            counter += 1
        return str(base)

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
            self.save_input.text() or self.get_downloads_folder(),   # start dialog at current suggested path
            "ODT files (*.odt);;All files (*.*)"
        )
        if path:
            self.save_input.setText(path)

    def convert(self):
        epub = self.epub_input.text()
        odt = self.save_input.text()

        if epub == "No file selected" or epub == "":
            self.log.append("ERROR: Please select an EPUB file first.")
            return
        if odt == "No save location selected" or odt == "":
            self.log.append("ERROR: Please select a save location first.")
            return

        self.btn_convert.setEnabled(False)
        self.log.clear()
        self.log.append("Starting conversion...")

        self.worker = ConversionWorker(epub, odt)
        self.worker.log_signal.connect(self.log.append)
        self.worker.finished_signal.connect(self.on_finished)
        self.worker.start()

    def on_finished(self, success):
        if success:
            self.log.append("✓ Done!")
        else:
            self.log.append("Something went wrong. Check the log above.")
        self.btn_convert.setEnabled(True)


if __name__ == "__main__":

    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())