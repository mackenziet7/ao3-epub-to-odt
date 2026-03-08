import subprocess
import time
from pathlib import Path

from PyQt6.QtCore import QThread, pyqtSignal


class ConversionWorker(QThread):
    log_signal      = pyqtSignal(str)   # emits a line of text to the log
    finished_signal = pyqtSignal(bool)  # emits True=success, False=failure

    def __init__(self, lo_python, script, epub, odt, include_toc=True):
        super().__init__()
        self.lo_python   = lo_python
        self.script      = script
        self.epub        = epub
        self.odt         = odt
        self.include_toc = include_toc

    def run(self):
        subprocess.run(
            ["taskkill", "/f", "/im", "soffice.exe"],
            capture_output=True,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        time.sleep(2)

        cmd = [self.lo_python, "-u", str(self.script), self.epub, self.odt]
        if not self.include_toc:
            cmd.append("--no-toc")

        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding='utf-8',
            creationflags=subprocess.CREATE_NO_WINDOW
        )

        # Read character by character so partial lines show up live
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
                pass
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
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        time.sleep(3)

        self.finished_signal.emit(True)