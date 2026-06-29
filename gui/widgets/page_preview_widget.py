"""
PagePreviewWidget
=================
A QWidget subclass that draws a two-page spread (verso + recto) with margin
guidelines. Makes the mirrored/non-mirrored margin distinction immediately
visible — inside margins face each other at the spine.
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

# Gap between the two pages (in mm, treated as part of the spread width)
_GUTTER_MM = 8.0


class PagePreviewWidget(QWidget):
    """Draws a two-page spread with margin guidelines."""

    _BG_COLOUR      = QColor("#1e1e1e")
    _SHADOW_COLOUR  = QColor(0, 0, 0, 120)
    _PAGE_COLOUR    = QColor("#ffffff")
    _SPINE_COLOUR   = QColor("#888888")   # thin line down the centre join

    _PADDING        = 24   # px gap between widget edge and the spread
    _SHADOW_OFFSET  = 4    # px

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

        # Page dimensions (swap for landscape)
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

        page_px_w   = pw_mm   * scale
        page_px_h   = ph_mm   * scale
        gutter_px   = _GUTTER_MM * scale

        spread_px_w = 2 * page_px_w + gutter_px
        origin_x    = (self.width()  - spread_px_w) / 2
        origin_y    = (self.height() - page_px_h)   / 2

        verso_x = origin_x
        recto_x = origin_x + page_px_w + gutter_px

        # ── draw each page ──────────────────────────────────────────
        for side in ("verso", "recto"):
            ox = verso_x if side == "verso" else recto_x
            page_rect = QRectF(ox, origin_y, page_px_w, page_px_h)

            # Shadow
            painter.fillRect(
                page_rect.translated(self._SHADOW_OFFSET, self._SHADOW_OFFSET),
                self._SHADOW_COLOUR,
            )
            # Body
            painter.fillRect(page_rect, self._PAGE_COLOUR)

            # Margin guide
            margin_rect = self._margin_rect(
                side, ox, origin_y, page_px_w, page_px_h, scale
            )
            if margin_rect is not None:
                pen = QPen(QColor("#4a90d9"), 1.0, Qt.PenStyle.DashLine)
                pen.setCosmetic(True)
                painter.setPen(pen)
                painter.drawRect(margin_rect)

        # ── spine line ──────────────────────────────────────────────
        spine_x = origin_x + page_px_w + gutter_px / 2
        pen = QPen(self._SPINE_COLOUR, 1.0, Qt.PenStyle.SolidLine)
        pen.setCosmetic(True)
        painter.setPen(pen)
        painter.drawLine(
            int(spine_x), int(origin_y),
            int(spine_x), int(origin_y + page_px_h),
        )

        painter.end()

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def _margin_rect(
        self,
        side: str,
        ox: float,
        oy: float,
        page_px_w: float,
        page_px_h: float,
        scale: float,
    ) -> QRectF | None:
        top_px    = self._margin_top_mm    * scale
        bottom_px = self._margin_bottom_mm * scale

        if self._mirrored:
            inside_px  = self._margin_inside_mm  * scale
            outside_px = self._margin_outside_mm * scale
            # Verso (left page): outside on left, inside on right
            # Recto (right page): inside on left, outside on right
            left_px  = outside_px if side == "verso" else inside_px
            right_px = inside_px  if side == "verso" else outside_px
        else:
            # Non-mirrored: inside = left margin, outside = right margin
            # Both pages are identical
            left_px  = self._margin_inside_mm  * scale
            right_px = self._margin_outside_mm * scale

        # Guard against crossing guides
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