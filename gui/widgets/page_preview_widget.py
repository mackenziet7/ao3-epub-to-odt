"""
PagePreviewWidget
=================
Draws a two-page spread (verso + recto) with margin guidelines. Makes the
mirrored/non-mirrored margin distinction immediately visible — inside margins
face each other at the spine.
"""

from PySide6.QtWidgets import QWidget, QSizePolicy
from PySide6.QtGui import QPainter, QColor, QPen
from PySide6.QtCore import Qt, QRectF


PAGE_SIZES_MM: dict[str, tuple[float, float]] = {
    "A4":                  (210.0, 297.0),
    "A5":                  (148.0, 210.0),
    '5.5×8.5" Paperback':  (139.7, 215.9),
    "Digest":              (140.0, 216.0),
}

_GUTTER_MM = 8.0   # gap between pages in the spread


class PagePreviewWidget(QWidget):
    """Draws a scaled two-page spread with margin guidelines."""

    _BG_COLOUR     = QColor("#1e1e1e")
    _SHADOW_COLOUR = QColor(0, 0, 0, 120)
    _PAGE_COLOUR   = QColor("#ffffff")
    _SPINE_COLOUR  = QColor("#888888")
    _GUIDE_COLOUR  = QColor("#4a90d9")

    _PADDING       = 24
    _SHADOW_OFFSET = 4

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.setMinimumSize(160, 120)

        # State — all in mm
        self._page_w_mm         = 210.0
        self._page_h_mm         = 297.0
        self._landscape         = False
        self._mirrored          = True
        self._margin_top_mm     = 20.0
        self._margin_bottom_mm  = 20.0
        self._margin_inside_mm  = 20.0
        self._margin_outside_mm = 20.0

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def update_from_settings(
        self,
        *,
        page_size_name: str,
        landscape: bool,
        custom_w_mm: float,
        custom_h_mm: float,
        margin_top_mm: float,
        margin_bottom_mm: float,
        margin_inside_mm: float,
        margin_outside_mm: float,
        mirrored: bool,
    ) -> None:
        """Push new Page Setup state and trigger a repaint."""
        if page_size_name.startswith("Custom"):
            self._page_w_mm = custom_w_mm
            self._page_h_mm = custom_h_mm
        else:
            w, h = PAGE_SIZES_MM.get(page_size_name, (210.0, 297.0))
            self._page_w_mm = w
            self._page_h_mm = h

        self._landscape         = landscape
        self._mirrored          = mirrored
        self._margin_top_mm     = margin_top_mm
        self._margin_bottom_mm  = margin_bottom_mm
        self._margin_inside_mm  = margin_inside_mm
        self._margin_outside_mm = margin_outside_mm

        self.update()

    # ------------------------------------------------------------------
    # paintEvent
    # ------------------------------------------------------------------

    def paintEvent(self, _event) -> None:
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        painter.fillRect(self.rect(), self._BG_COLOUR)

        # Page dimensions — swap axes for landscape
        pw_mm = self._page_w_mm
        ph_mm = self._page_h_mm
        if self._landscape:
            pw_mm, ph_mm = ph_mm, pw_mm

        avail_w = self.width()  - 2 * self._PADDING
        avail_h = self.height() - 2 * self._PADDING
        if pw_mm <= 0 or ph_mm <= 0 or avail_w <= 0 or avail_h <= 0:
            return

        # Scale so the full two-page spread fits
        spread_mm_w = 2 * pw_mm + _GUTTER_MM
        scale       = min(avail_w / spread_mm_w, avail_h / ph_mm)

        page_px_w = pw_mm      * scale
        page_px_h = ph_mm      * scale
        gutter_px = _GUTTER_MM * scale

        spread_px_w = 2 * page_px_w + gutter_px
        ox = (self.width()  - spread_px_w) / 2
        oy = (self.height() - page_px_h)   / 2

        verso_x = ox
        recto_x = ox + page_px_w + gutter_px

        for side, sx in (("verso", verso_x), ("recto", recto_x)):
            self._draw_page(painter, sx, oy, page_px_w, page_px_h)
            margin_rect = self._margin_rect(side, sx, oy, page_px_w, page_px_h, scale)
            self._draw_guide(painter, margin_rect)

        # Spine line
        spine_x = ox + page_px_w + gutter_px / 2
        pen = QPen(self._SPINE_COLOUR, 1.0)
        pen.setCosmetic(True)
        painter.setPen(pen)
        painter.drawLine(int(spine_x), int(oy), int(spine_x), int(oy + page_px_h))

        painter.end()

    # ------------------------------------------------------------------
    # Drawing helpers
    # ------------------------------------------------------------------

    def _draw_page(self, painter, ox, oy, w, h):
        rect = QRectF(ox, oy, w, h)
        painter.fillRect(
            rect.translated(self._SHADOW_OFFSET, self._SHADOW_OFFSET),
            self._SHADOW_COLOUR,
        )
        painter.fillRect(rect, self._PAGE_COLOUR)

    def _draw_guide(self, painter, rect: QRectF):
        pen = QPen(self._GUIDE_COLOUR, 1.0, Qt.PenStyle.DashLine)
        pen.setCosmetic(True)
        painter.setPen(pen)
        painter.drawRect(rect)

    def _margin_rect(
        self, side: str, ox, oy, page_px_w, page_px_h, scale
    ) -> QRectF:
        top_px    = self._margin_top_mm    * scale
        bottom_px = self._margin_bottom_mm * scale

        if self._mirrored:
            inside_px  = self._margin_inside_mm  * scale
            outside_px = self._margin_outside_mm * scale
            # Verso (left page):  binding on the right → outside left, inside right
            # Recto (right page): binding on the left  → inside left, outside right
            left_px  = outside_px if side == "verso" else inside_px
            right_px = inside_px  if side == "verso" else outside_px
        else:
            # Non-mirrored: inside = left, outside = right on both pages
            left_px  = self._margin_inside_mm  * scale
            right_px = self._margin_outside_mm * scale

        # Guard against guides crossing each other
        if top_px + bottom_px >= page_px_h:
            top_px = bottom_px = page_px_h * 0.25
        if left_px + right_px >= page_px_w:
            left_px = right_px = page_px_w * 0.25

        return QRectF(
            ox + left_px,
            oy + top_px,
            page_px_w - left_px - right_px,
            page_px_h - top_px  - bottom_px,
        )