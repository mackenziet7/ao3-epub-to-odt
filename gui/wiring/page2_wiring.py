
from PySide6.QtWidgets import (
    QSplitter,
    QFontComboBox,
    QDoubleSpinBox,
    QLabel,

)

# ------------------------------------------------------------------
# Page 2 — Basic Typography Setup screen
# ------------------------------------------------------------------
def wire_page2(main_window):
    w = main_window.window

    splitterPage3 = w.findChild(QSplitter, "splitterPage3")
    splitterPage3.setCollapsible(0, False)
    splitterPage3.setCollapsible(1, False)
    splitterPage3.setSizes([500, 400])

    # Typography controls
    main_window._headerFontCombo = w.findChild(QFontComboBox, "comboHeaderFont")
    main_window._bodyFontCombo = w.findChild(QFontComboBox, "comboBodyFont")
    main_window._headerSizeSpin = w.findChild(QDoubleSpinBox, "spinHeaderFontSize")
    main_window._bodySizeSpin = w.findChild(QDoubleSpinBox, "spinBodyFontSize")

    # Preview Labels
    main_window._headerPreviewLabel = w.findChild(QLabel, "labelHeaderPreview")
    main_window._bodyPreviewLabel = w.findChild(QLabel, "labelBodyPreview")

    # Wire preview
    main_window._headerFontCombo.currentFontChanged.connect(lambda: _update_typography_preview(main_window))
    main_window._bodyFontCombo.currentFontChanged.connect(lambda: _update_typography_preview(main_window))
    main_window._headerSizeSpin.valueChanged.connect(lambda: _update_typography_preview(main_window))
    main_window._bodySizeSpin.valueChanged.connect(lambda: _update_typography_preview(main_window))

    _update_typography_preview(main_window)

def _update_typography_preview(main_window):
    header_font = main_window._headerFontCombo.currentFont()
    header_font.setPointSize(main_window._headerSizeSpin.value())
    main_window._headerPreviewLabel.setFont(header_font)

    body_font = main_window._bodyFontCombo.currentFont()
    body_font.setPointSize(main_window._bodySizeSpin.value())
    main_window._bodyPreviewLabel.setFont(body_font)