# gui/wiring/preset_builder.py
from PySide6.QtWidgets import (
    QComboBox, QDoubleSpinBox, QFontComboBox, QCheckBox, QRadioButton,
)

CM_TO_IN = 1 / 2.54

_SIZE_PRESET_SLUGS = {
    "A4": "a4",
    "A5": "a5",
    '5.5×8.5" Paperback': "5.5x8.5_paperback",
    "Digest": "digest",
}


def _to_inches(value: float, unit: str) -> float:
    return value * CM_TO_IN if unit == "cm" else value


def _alignment(w, name: str) -> str:
    return w.findChild(QComboBox, name).currentText().lower()


def _font(w, name: str) -> str:
    return w.findChild(QFontComboBox, name).currentFont().family()


def _size(w, name: str) -> float:
    return w.findChild(QDoubleSpinBox, name).value()


def _checked(w, name: str) -> bool:
    return w.findChild(QCheckBox, name).isChecked()


def collect_wizard_settings(main_window) -> dict:
    """
    Reads the current state of pages 1-4 and returns a dict matching the
    preset schema (minus preset_name/schema_version, which are added by
    whatever saves this as a named preset).
    """
    w = main_window.window

    # ---- page_setup (page 1) ----
    unit = w.findChild(QComboBox, "comboUnits").currentText()
    size_key = None
    combo_text = w.findChild(QComboBox, "comboPageSize").currentText()
    if not combo_text.startswith("Custom"):
        # re-derive the same key to avoid a cross-module import.
        from gui.wiring.page1_wiring import page_size_key_for_combo_text
        size_key = page_size_key_for_combo_text(combo_text)

    mirrored = _checked(w, "checkMirroredMargins")
    left_val = round(_to_inches(_size(w, "spinMarginLeft"), unit), 3)
    right_val = round(_to_inches(_size(w, "spinMarginRight"), unit), 3)

    margins = {
        "top": round(_to_inches(_size(w, "spinMarginTop"), unit), 3),
        "bottom": round(_to_inches(_size(w, "spinMarginBottom"), unit), 3),
    }
    if mirrored:
        margins["inside"] = left_val
        margins["outside"] = right_val
    else:
        margins["left"] = left_val
        margins["right"] = right_val

    page_setup = {
        "size_preset": _SIZE_PRESET_SLUGS.get(size_key, "custom"),
        "width_in": round(_to_inches(_size(w, "spinCustomWidth"), unit), 3),
        "height_in": round(_to_inches(_size(w, "spinCustomHeight"), unit), 3),
        "display_unit": "in",
        "orientation": "landscape" if w.findChild(QRadioButton, "radioLandscape").isChecked() else "portrait",
        "mirrored_margins": mirrored,
        "margins": margins,
    }

    # ---- typography_basic (page 2) ----
    typography_basic = {
        "header_font": _font(w, "comboHeaderFont"),
        "header_size_pt": _size(w, "spinHeaderFontSize"),
        "body_font": _font(w, "comboBodyFont"),
        "body_size_pt": _size(w, "spinBodyFontSize"),
    }

    # ---- typography_advanced (page 3) ----
    typography_advanced = {
        "main_book": {
            "chapter_headers": {
                "font": _font(w, "comboChapHeadFont"),
                "size_pt": _size(w, "spinChapHeaderFontSize"),
                "bold": _checked(w, "checkBold"),
                "alignment": _alignment(w, "comboHeaderAlignment"),
                "display_unit": "in",
                "top_margin_in": _size(w, "spinChapHeaderTopMargin"),
                "bottom_margin_in": _size(w, "spinChapHeaderBottomMargin"),
            },
            "body": {
                "font": _font(w, "comboMainBookBodyFont"),
                "size_pt": _size(w, "spinMainBookBodyFontSize"),
                "alignment": _alignment(w, "comboBodyAlignment"),
                "display_unit": "in",
                "first_line_indent_in": _size(w, "spinFirstLineIndent"),
                "line_spacing_in": _size(w, "spinBodyLineSpacing"),
                "no_indent_on_first_paragraph_after_heading": _checked(w, "checkFirstLineIndent"),
            },
        },
        "front_matter": {
            "head": {
                "font": _font(w, "comboFrontMatterHeadFont"),
                "size_pt": _size(w, "spinFrontMatterHeadSize"),
                "bold": _checked(w, "checkFrontMatterBold"),
                "alignment": _alignment(w, "comboFrontMatterHeadAlignment"),
                "display_unit": "in",
                "top_margin_in": _size(w, "spinFrontMatterHeaderTopMargin"),
                "bottom_margin_in": _size(w, "spinFrontMatterHeaderBotMargin"),
            },
            "body": {
                "font": _font(w, "comboFrontMatterFont"),
                "size_pt": _size(w, "spinFrontMatterSize"),
                "alignment": _alignment(w, "comboFrontMatterAlignment"),
            },
            "qr_code": {
                "alignment": _alignment(w, "comboQrAlignment"),
                "display_unit": "in",
                "top_margin_in": _size(w, "spinQrTopMargin"),
                "bottom_margin_in": _size(w, "spinQrBotMargin"),
            },
            "qr_caption": {
                "font": _font(w, "comboQrCaptionFont"),
                "size_pt": _size(w, "spinQrCaptionSize"),
                "italic": _checked(w, "checkItalic"),
            },
        },
        "appendix": {
            "head": {
                "font": _font(w, "comboAppendixHeaderFont"),
                "size_pt": _size(w, "spinAppendixHeadSize"),
                "bold": _checked(w, "checkAppendixBold"),
                "alignment": _alignment(w, "comboAppendixHeadAlignment"),
                "display_unit": "in",
                "top_margin_in": _size(w, "spinAppendixHeaderTopMargin"),
                "bottom_margin_in": _size(w, "spinAppendixHeaderBotMargin"),
            },
            "note_label": {
                "font": _font(w, "comboAppendixNoteLabelFont"),
                "size_pt": _size(w, "spinAppendixNoteLabelSize"),
                "alignment": _alignment(w, "comboAppendixNoteLabelAlignment"),
            },
            "note": {
                "font": _font(w, "comboAppendixNoteFont"),
                "size_pt": _size(w, "spinAppendixNoteSize"),
                "alignment": _alignment(w, "comboAppendixNoteAlignment"),
                "display_unit": "in",
                "left_margin_in": _size(w, "spinAppendixNoteLeftMargin"),
            },
        },
    }

    # ---- additional_options (page 4) ----
    additional_options = {
        "include_table_of_contents": _checked(w, "checkTableOfContents"),
        "include_qr_code": _checked(w, "checkQr"),
    }

    return {
        "page_setup": page_setup,
        "typography_basic": typography_basic,
        "typography_advanced": typography_advanced,
        "additional_options": additional_options,
    }