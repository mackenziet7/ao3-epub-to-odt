import json
from datetime import datetime, timedelta, timezone

from PySide6.QtCore import QEvent, QObject, QUrl, Qt, Signal
from PySide6.QtGui import QColor, QPainter
from PySide6.QtNetwork import QNetworkAccessManager, QNetworkReply, QNetworkRequest
from PySide6.QtWidgets import QWidget

from gui.config import load_config, update_config

__version__ = "1.3.0"

API_URL = "https://api.github.com/repos/mackenziet7/ao3-epub-to-odt/releases/latest"
CHECK_INTERVAL = timedelta(hours=24)
BADGE_COLOR = QColor("#4c8df6")

def _parse(v: str) -> tuple[int, ...]:
    parts = [int(p) for p in v.strip().lstrip("vV").split(".") if p.isdigit()]
    return tuple(parts + [0] * (3 - len(parts)))

def update_state(cfg: dict) -> tuple[bool, bool]:
    """Returns (newer_available, show_badge) from the cached values."""
    latest = cfg.get("latest_known_version", "")
    newer = bool(latest) and _parse(latest) > _parse(__version__)
    badge = newer and latest != cfg.get("skipped_version", "")
    return newer, badge


class UpdateChecker(QObject):
    finished = Signal(bool)  # True if GitHub responded successfully

    def __init__(self, parent=None):
        super().__init__(parent)
        self._nam = QNetworkAccessManager(self)
        self._nam.finished.connect(self._on_reply)

    def _due(self, cfg: dict) -> bool:
        try:
            last = datetime.fromisoformat(cfg["last_update_check"])
        except (KeyError, ValueError):
            return True
        return datetime.now(timezone.utc) - last >= CHECK_INTERVAL

    def check(self, force: bool = False):
        if not force and not self._due(load_config()):
            return
        req = QNetworkRequest(QUrl(API_URL))
        req.setRawHeader(b"User-Agent", b"AO3toODT-update-check")
        req.setRawHeader(b"Accept", b"application/vnd.github+json")
        req.setTransferTimeout(8000)
        self._nam.get(req)

    def _on_reply(self, reply: QNetworkReply):
        ok = False
        try:
            if reply.error() == QNetworkReply.NetworkError.NoError:
                data = json.loads(bytes(reply.readAll()).decode("utf-8"))
                tag = data.get("tag_name", "").lstrip("vV")
                url = data.get("html_url", "")
                if tag and url.startswith("https://github.com/"):
                    # timestamp only on success, so an offline launch doesn't burn the daily check
                    update_config(
                        last_update_check=datetime.now(timezone.utc).isoformat(),
                        latest_known_version=tag,
                        latest_known_url=url,
                    )
                    ok = True
        except (ValueError, AttributeError):
            pass
        finally:
            reply.deleteLater()
            self.finished.emit(ok)


class BadgeDot(QWidget):
    """Small dot overlaid on the top-right of a parent button. No subclassing or .ui changes needed."""

    def __init__(self, button: QWidget):
        super().__init__(button)
        self.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        self.setFixedSize(12, 12)
        button.installEventFilter(self)
        self._reposition()
        self.hide()

    def _reposition(self):
        self.move(self.parent().width() - self.width() - 3, 0)
        self.raise_()

    def eventFilter(self, obj, ev):
        if obj is self.parent() and ev.type() == QEvent.Type.Resize:
            self._reposition()
        return False

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(BADGE_COLOR)
        p.drawEllipse(self.rect())