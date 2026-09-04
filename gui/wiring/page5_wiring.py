# gui/wiring/page5_wiring.py
import subprocess
import sys
from pathlib import Path

from PySide6.QtWidgets import (
    QTextBrowser,
    QProgressBar, 
    QPushButton, 
    QFileDialog, 
    QLabel, 
    QStackedWidget
)

from gui.worker import ConversionWorker

NO_WINDOW = subprocess.CREATE_NO_WINDOW

def get_script_path() -> Path:
    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / "scripts" / "script_ao3_to_odt.py"
    return Path(__file__).parent.parent.parent / "scripts" / "script_ao3_to_odt.py"


def wire_page5(main_window):
    w = main_window.window

    main_window._logOutput = w.findChild(QTextBrowser, "textBrowser")
    main_window._progress = w.findChild(QProgressBar, "progressBar")
    main_window._progress.setVisible(False)
    main_window._progress.setFormat("%v / %m files")

    w.findChild(QPushButton, "buttonExport").clicked.connect(
        lambda: _export_log(main_window)
    )
    w.findChild(QPushButton, "buttonComplete").clicked.connect(
        lambda: _on_complete(main_window)
    )

# ------------------------------------------------------------------
# Conversion
# ------------------------------------------------------------------
def start_conversion(main_window, epub_paths, output_folder, settings):
    """
    settings: dict of resolved conversion options (toc, qr, margins, fonts, etc.)
    regardless of whether it came from a preset file or live wizard state.
    """
    if not epub_paths:
        main_window._logOutput.append("No files to convert.")
        return
    
    main_window._queue = list(epub_paths)
    main_window._succeeded = []
    main_window._failed = []
    main_window._worker = None
    main_window._output_folder = output_folder
    main_window._conversion_settings = settings

    main_window._progress.setMaximum(len(main_window._queue))
    main_window._progress.setValue(0)
    main_window._progress.setVisible(True)

    main_window._logOutput.append(f"Starting conversion of {len(main_window._queue)} file(s)...")
    _start_next(main_window)


def _start_next(main_window):
    epub = main_window._queue[0]
    odt = suggest_odt_path(epub, main_window._output_folder)
    main_window._logOutput.append(f"\nConverting: {Path(epub).name}")

    if main_window._worker is not None:
        main_window._worker.log_signal.disconnect()
        main_window._worker.finished_signal.disconnect()

    opts = main_window._conversion_settings.get("additional_options", {})
    main_window._worker = ConversionWorker(
        main_window._lo_python,
        get_script_path(),
        epub, odt,
        opts.get("include_table_of_contents", True),
        opts.get("include_qr_code", True),
    )
    main_window._worker.log_signal.connect(main_window._logOutput.append)
    main_window._worker.finished_signal.connect(
        lambda success: on_finished(main_window, success)
    )
    main_window._worker.start()


def on_finished(main_window, success: bool):
    current = main_window._queue.pop(0)

    if success:
        main_window._succeeded.append(current)
    else:
        main_window._failed.append(current)
        main_window._logOutput.append(f"✗ Failed: {Path(current).name}")

    main_window._progress.setValue(len(main_window._succeeded) + len(main_window._failed))

    if main_window._queue:
        _start_next(main_window)
    else:
        _show_summary(main_window)
        main_window._progress.setVisible(False)
        complete_text = main_window.window.findChild(QLabel, "labelCompletionProgress")
        complete_text.setText("Conversion Complete, you are free to exit the program.")

        complete_button = main_window.window.findChild(QPushButton, "buttonComplete")
        complete_button.setEnabled(True)

def _show_summary(main_window):
    main_window._logOutput.append("")
    main_window._logOutput.append("── Conversion complete ──────────────────")
    main_window._logOutput.append(f"✓ {len(main_window._succeeded)} file(s) converted successfully")
    if main_window._failed:
        main_window._logOutput.append(f"✗ {len(main_window._failed)} file(s) failed:")
        for f in main_window._failed:
            main_window._logOutput.append(f"  - {Path(f).name}")

def suggest_odt_path(epub: str, output_folder: str) -> str:
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

# ------------------------------------------------------------------
# Export / Complete
# ------------------------------------------------------------------
def _export_log(main_window):
    path, _ = QFileDialog.getSaveFileName(
        main_window.window, "Export Log", "conversion_log.txt", "Text files (*.txt)"
    )
    if path:
        Path(path).write_text(main_window._logOutput.toPlainText(), encoding="utf-8")


def _on_complete(main_window):
    main_window.window.findChild(QStackedWidget, "stackedWidget").setCurrentIndex(0)