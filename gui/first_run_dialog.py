from pathlib import Path

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QFileDialog
)
from PySide6.QtGui import QDesktopServices
from PySide6.QtCore import QUrl

from gui.config import save_config

class LOPathDialog(QDialog):
    """
    Dialog for locating LO's python.exe.
    Shows on first run or if the saved path is no longer valid.
    Reusable later from a settings button.
    """

    def __init__(self, parent = None, invalid_path: str = None):
        super().__init__(parent)
        self.setWindowTitle("Locate LibreOffice")
        self.setMinimumWidth(500)

        layout = QVBoxLayout(self)

        # Message
        if invalid_path:
            msg = (
                f"LibreOffice could not be found at the saved location:\n"
                f"  {invalid_path}\n\n"
                f"It may have been moved or uninstalled. "
                f"Please locate python.exe inside your LibreOffice installation."
            )
        else:
            msg = (
                "LibreOffice is required but could not be found at the default "
                "install location.\n\n"
                "Please browse to find LibreOffice, usually located at:\n"
                r"  C:\Program Files\LibreOffice"
            )

        label = QLabel(msg)
        label.setWordWrap(True)
        layout.addWidget(label)

        btn_download = QPushButton("Don't have LibreOffice? Download it here")
        btn_download.setFlat(True)  # makes it look more like a link than a button
        btn_download.clicked.connect(lambda: QDesktopServices.openUrl(
            QUrl("https://www.libreoffice.org/download/download-libreoffice/")
        ))
        layout.addWidget(btn_download)

        # ── Path row ─────────────────────────────────────────────────────────
        path_row = QHBoxLayout()
        self.path_input = QLineEdit()
        self.path_input.setPlaceholderText(r"C:\Program Files\LibreOffice")
        path_row.addWidget(self.path_input)

        btn_browse = QPushButton("Browse")
        btn_browse.clicked.connect(self._browse)
        path_row.addWidget(btn_browse)
        layout.addLayout(path_row)

        # ── Confirm ───────────────────────────────────────────────────────
        self.btn_confirm = QPushButton("Confirm")
        self.btn_confirm.setEnabled(False)  # disabled until a valid path is entered
        self.btn_confirm.clicked.connect(self._on_confirm)
        layout.addWidget(self.btn_confirm)

    def _browse(self):
        folder = QFileDialog.getExistingDirectory(
            self,
            "Select your LibreOffice installation folder",
            r"C:\Program Files",
        )
        if folder:
            python_path = Path(folder) / "program" / "python.exe"
            if python_path.exists():
                self.path_input.setText(str(python_path))
                self.btn_confirm.setEnabled(True)  # ← unlock confirm
            else:
                self.path_input.setPlaceholderText(
                    "⚠ python.exe not found in that folder — are you sure this is LibreOffice?"
                )
                self.path_input.setText("")
                self.btn_confirm.setEnabled(False)

    def _on_confirm(self):
        path = self.path_input.text().strip()
        if not path:
            return  # don't close if empty

        p = Path(path)
        if not p.exists():
            self.path_input.setPlaceholderText("⚠ File not found — please check the path")
            self.path_input.setText("")
            return

        # Save to config and close successfully
        save_config({"lo_python": str(p)})
        self.accept()  # signals to the caller that a valid path was saved
        