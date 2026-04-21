import os
import sys
import glob
import subprocess
import time
from pathlib import Path
from PyQt6.QtCore import QThread, pyqtSignal

class ConversionWorker(QThread):
    log_signal      = pyqtSignal(str)   # emits a line of text to the log
    finished_signal = pyqtSignal(bool)  # emits True=success, False=failure

    def __init__(self, lo_python, script, epub, odt, include_toc=True, include_qr=True):
        super().__init__()
        self.lo_python   = lo_python
        self.script      = script
        self.epub        = epub
        self.odt         = odt
        self.include_toc = include_toc
        self.include_qr = include_qr

    def run(self):
        subprocess.run(
            ["taskkill", "/f", "/im", "soffice.exe"],
            capture_output=True,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        time.sleep(2)

        if hasattr(sys, '_MEIPASS'):
            # Rename Python DLLs and extension modules that conflict with LO's Python 3.12
            conflicting = ['python314.dll', 'python3.dll', '_socket.pyd', 
                          '_ssl.pyd', '_hashlib.pyd', 'select.pyd',
                          '_bz2.pyd', '_decimal.pyd', '_lzma.pyd', '_zstd.pyd',
                          'unicodedata.pyd']
            for name in conflicting:
                target = os.path.join(sys._MEIPASS, name)
                if os.path.exists(target):
                    try:
                        os.rename(target, target + '.bak')
                    except OSError:
                        pass

        cmd = [self.lo_python, "-u", str(self.script), self.epub, self.odt]
        if not self.include_toc:
            cmd.append("--no-toc")
        if not self.include_qr:
            cmd.append("--no-qr")

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

        # Wait for process to fully exit — the script kills LO itself
        # so we just need to wait for the python subprocess to finish
        try:
            process.wait(timeout=60)
        except subprocess.TimeoutExpired:
            self.log_signal.emit("Warning: conversion script timed out, forcing stop.")
            process.kill()
            process.wait()

        # Final LO cleanup in case script exited without killing it
        subprocess.run(
            ["taskkill", "/f", "/im", "soffice.exe"],
            capture_output=True,
            creationflags=subprocess.CREATE_NO_WINDOW
        )

        self.finished_signal.emit(True)  