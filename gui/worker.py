import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
import warnings
from pathlib import Path
from PySide6.QtCore import QThread, Signal
from bs4 import XMLParsedAsHTMLWarning

from scripts.ao3_to_odt.epub.parser import parse_epub
from scripts.ao3_to_odt.epub.models import book_to_dict
from scripts.ao3_to_odt.qr import generate_qr_png
from scripts.ao3_to_odt.preset_schema import SCHEMA_VERSION, validate_preset

warnings.filterwarnings("ignore", category=XMLParsedAsHTMLWarning)

NO_WINDOW = subprocess.CREATE_NO_WINDOW
PROGRESS_PREFIX = "@@PROGRESS|"

class ConversionWorker(QThread):
    log_signal      = Signal(str)   # emits a line of text to the log
    progress_signal = Signal(int, str)
    finished_signal = Signal(bool)  # emits True=success, False=failure

    def __init__(self, lo_python, script, epub, odt, preset):
        super().__init__()
        self.lo_python = lo_python
        self.script    = script
        self.epub      = epub
        self.odt       = odt
        self.preset = preset

    def run(self):
        work_dir = Path(tempfile.mkdtemp(prefix="ao3toodt_"))
        ok = False
        try:
            ok = self._convert(work_dir)
        except Exception as e:
            self.log_signal.emit(f"ERROR: {e}")
        finally:
            shutil.rmtree(work_dir, ignore_errors=True)
        self.finished_signal.emit(ok)

    def _convert(self, work_dir: Path) -> bool:
        errors = validate_preset(self.preset)
        if errors:
            self.log_signal.emit("Preset problem: " + "; ".join(errors))
            return False
        self.log_signal.emit("Parsing EPUB...")
        self.progress_signal.emit(2, "Parsing EPUB")
        book = parse_epub(self.epub)
        self.log_signal.emit(
            f"  {book.metadata.title} by {book.metadata.author} "
            f"({len(book.chapters)} chapters)"
        )

        payload = book_to_dict(book)
        payload["qr_png"] = None
        opts = self.preset.get("additional_options", {})
        if opts.get("include_qr_code", True) and book.metadata.ao3_url:
            qr_file = work_dir / "qr.png"
            try:
                generate_qr_png(book.metadata.ao3_url, qr_file)
                payload["qr_png"] = str(qr_file)
            except Exception as e:
                self.log_signal.emit(f"  QR code skipped: {e}")

        book_json = work_dir / "book.json"
        preset_json = work_dir / "preset.json"
        book_json.write_text(json.dumps(payload), encoding="utf-8")
        preset_json.write_text(json.dumps(self.preset), encoding="utf-8")
        self.progress_signal.emit(10, "Starting LibreOffice")

        # ── 2. Existing cleanup + DLL workaround (unchanged) ─────────────────
        subprocess.run(
            ["taskkill", "/f", "/im", "soffice.exe"],
            capture_output=True,
            creationflags=NO_WINDOW
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

        # ── 3. Run the LO-Python script ──────────────────────────────────────
        cmd = [self.lo_python, "-u", str(self.script),
               "--book", str(book_json),
               "--preset", str(preset_json),
               "--out", self.odt]

        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding='utf-8',
            creationflags=NO_WINDOW
        )

        # Read character by character so partial lines show up live
        current_line = ""
        while True:
            char = process.stdout.read(1)
            if not char:
                break
            if char == "\n":
                if current_line.strip():
                    self._handle_line(current_line.strip())
                current_line = ""
            elif char == "\r":
                pass
            else:
                current_line += char

        # Emit any remaining text that didn't end with a newline
        if current_line.strip():
            self._handle_line(current_line.strip())

        # Wait for process to fully exit — the script kills LO itself
        timed_out = False
        try:
            process.wait(timeout=60)
        except subprocess.TimeoutExpired:
            timed_out = True
            self.log_signal.emit("Warning: conversion script timed out, forcing stop.")
            process.kill()
            process.wait()

        # Final LO cleanup in case script exited without killing it
        subprocess.run(
            ["taskkill", "/f", "/im", "soffice.exe"],
            capture_output=True,
            creationflags=NO_WINDOW
        )

        # ── 4. Success = exit code 0 AND the output file exists ──────────────
        if timed_out or process.returncode != 0:
            self.log_signal.emit(f"Converter exited with code {process.returncode}")
            return False
        if not Path(self.odt).exists():
            self.log_signal.emit("Converter finished but produced no output file.")
            return False
        return True

    def _handle_line(self, line: str):
        if line.startswith(PROGRESS_PREFIX):
            try:
                _, pct, msg = line.split("|", 2)
                self.progress_signal.emit(int(pct), msg)
                return
            except ValueError:
                pass                      # malformed: fall through and show it in the log
        self.log_signal.emit(line)
