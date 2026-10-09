# gui/wiring/preset_builder.py
"""
Translate between the wizard widgets (pages 1-4) and the preset dict.

    collect_wizard_settings()   widgets -> dict
    apply_preset_to_wizard()    dict    -> widgets

Page 1 is hand-written (units, mirrored margins, size-combo interplay).
Everything else is described once in the field tables below and read and
written by the same code, so the two directions can't drift apart.
"""
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QCheckBox, QComboBox, QDoubleSpinBox, QFontComboBox, QRadioButton,
)

from gui.wiring import page1_wiring as p1
from scripts.ao3_to_odt.preset_schema import SCHEMA_VERSION

DEFAULT_PRESET_STEM = "preset_Default_Book_Layout"

_SIZE_PRESET_SLUGS = {
    "A4": "a4",
    "A5": "a5",
    '5.5×8.5" Paperback': "5.5x8.5_paperback",
    "Digest": "digest",
}
_SLUG_TO_SIZE_KEY = {slug: key for key, slug in _SIZE_PRESET_SLUGS.items()}

# (widget, source, ratio): source is "header" or "body"
_DERIVED_SIZES = [
    ("spinChapHeaderFontSize",        "header", 1.0),
    ("spinFrontMatterHeadSize",       "header", 1.0),
    ("spinAppendixHeadSize",          "body",   11.0 / 11.5),
    ("spinMainBookBodyFontSize",      "body",   1.0),
    ("spinFrontMatterSize",           "body",   9.0 / 11.5),
    ("spinQrCaptionSize",             "body",   10.0 / 11.5),
    ("spinAppendixNoteLabelSize",     "body",   8.0 / 11.5),
    ("spinAppendixNoteSize",          "body",   8.0 / 11.5),
]

_DERIVED_FONTS = [
    ("comboChapHeadFont",            "header"),
    ("comboFrontMatterHeadFont",     "header"),
    ("comboAppendixHeaderFont",      "header"),
    ("comboMainBookBodyFont",        "body"),
    ("comboFrontMatterFont",         "body"),
    ("comboQrCaptionFont",           "body"),
    ("comboAppendixNoteLabelFont",   "body"),
    ("comboAppendixNoteFont",        "body"),
]

# ------------------------------------------------------------------
# Widget access
# ------------------------------------------------------------------
def _widget(w, cls, name):
    widget = w.findChild(cls, name)
    if widget is None:
        raise LookupError(f"{cls.__name__} '{name}' not found in main_window.ui")
    return widget


def _get_font(w, name):  return _widget(w, QFontComboBox, name).currentFont().family()
def _get_spin(w, name):  return _widget(w, QDoubleSpinBox, name).value()
def _get_check(w, name): return _widget(w, QCheckBox, name).isChecked()
def _get_align(w, name): return _widget(w, QComboBox, name).currentText().lower()


# Setters ignore None so missing keys (older, hand-edited, or partial
# presets) leave the widget at its current value instead of clobbering it.
def _set_font(w, name, family):
    if family:
        _widget(w, QFontComboBox, name).setCurrentFont(QFont(family))


def _set_spin(w, name, value):
    if value is not None:
        _widget(w, QDoubleSpinBox, name).setValue(value)


def _set_check(w, name, value):
    if value is not None:
        _widget(w, QCheckBox, name).setChecked(bool(value))


def _set_align(w, name, value):
    if value is None:
        return
    combo = _widget(w, QComboBox, name)
    idx = combo.findText(value, Qt.MatchFlag.MatchFixedString)  # case-insensitive
    if idx >= 0:
        combo.setCurrentIndex(idx)


_READERS = {"font": _get_font, "spin": _get_spin, "check": _get_check, "align": _get_align}
_WRITERS = {"font": _set_font, "spin": _set_spin, "check": _set_check, "align": _set_align}


# ------------------------------------------------------------------
# Field tables: (preset key, widget name, kind)
# ------------------------------------------------------------------
_BASIC_FIELDS = [
    ("header_font", "comboHeaderFont", "font"),
    ("header_size_pt", "spinHeaderFontSize", "spin"),
    ("body_font", "comboBodyFont", "font"),
    ("body_size_pt", "spinBodyFontSize", "spin"),
]

_OPTION_FIELDS = [
    ("include_table_of_contents", "checkTableOfContents", "check"),
    ("include_qr_code", "checkQr", "check"),
]

# Keyed by path inside typography_advanced.
_ADVANCED_FIELDS = {
    ("main_book", "chapter_headers"): [
        ("font", "comboChapHeadFont", "font"),
        ("size_pt", "spinChapHeaderFontSize", "spin"),
        ("bold", "checkBold", "check"),
        ("alignment", "comboHeaderAlignment", "align"),
        ("top_margin_in", "spinChapHeaderTopMargin", "spin"),
        ("bottom_margin_in", "spinChapHeaderBottomMargin", "spin"),
    ],
    ("main_book", "body"): [
        ("font", "comboMainBookBodyFont", "font"),
        ("size_pt", "spinMainBookBodyFontSize", "spin"),
        ("alignment", "comboBodyAlignment", "align"),
        ("first_line_indent_in", "spinFirstLineIndent", "spin"),
        ("line_spacing_in", "spinBodyLineSpacing", "spin"),
        ("no_indent_on_first_paragraph_after_heading", "checkFirstLineIndent", "check"),
    ],
    ("front_matter", "head"): [
        ("font", "comboFrontMatterHeadFont", "font"),
        ("size_pt", "spinFrontMatterHeadSize", "spin"),
        ("bold", "checkFrontMatterBold", "check"),
        ("alignment", "comboFrontMatterHeadAlignment", "align"),
        ("top_margin_in", "spinFrontMatterHeaderTopMargin", "spin"),
        ("bottom_margin_in", "spinFrontMatterHeaderBotMargin", "spin"),
    ],
    ("front_matter", "body"): [
        ("font", "comboFrontMatterFont", "font"),
        ("size_pt", "spinFrontMatterSize", "spin"),
        ("alignment", "comboFrontMatterAlignment", "align"),
    ],
    ("front_matter", "qr_code"): [
        ("alignment", "comboQrAlignment", "align"),
        ("top_margin_in", "spinQrTopMargin", "spin"),
        ("bottom_margin_in", "spinQrBotMargin", "spin"),
    ],
    ("front_matter", "qr_caption"): [
        ("font", "comboQrCaptionFont", "font"),
        ("size_pt", "spinQrCaptionSize", "spin"),
        ("italic", "checkItalic", "check"),
    ],
    ("appendix", "head"): [
        ("font", "comboAppendixHeaderFont", "font"),
        ("size_pt", "spinAppendixHeadSize", "spin"),
        ("bold", "checkAppendixBold", "check"),
        ("alignment", "comboAppendixHeadAlignment", "align"),
        ("top_margin_in", "spinAppendixHeaderTopMargin", "spin"),
        ("bottom_margin_in", "spinAppendixHeaderBotMargin", "spin"),
    ],
    ("appendix", "note_label"): [
        ("font", "comboAppendixNoteLabelFont", "font"),
        ("size_pt", "spinAppendixNoteLabelSize", "spin"),
        ("alignment", "comboAppendixNoteLabelAlignment", "align"),
    ],
    ("appendix", "note"): [
        ("font", "comboAppendixNoteFont", "font"),
        ("size_pt", "spinAppendixNoteSize", "spin"),
        ("alignment", "comboAppendixNoteAlignment", "align"),
        ("left_margin_in", "spinAppendixNoteLeftMargin", "spin"),
    ],
}


def _read_fields(w, fields) -> dict:
    block = {}
    for key, widget, kind in fields:
        # Blocks with inch-valued fields declare their unit, just before the first one.
        if key.endswith("_in") and "display_unit" not in block:
            block["display_unit"] = "in"
        block[key] = _READERS[kind](w, widget)
    return block


def _write_fields(w, fields, block: dict) -> None:
    for key, widget, kind in fields:
        _WRITERS[kind](w, widget, block.get(key))


def _dig(data: dict, path) -> dict:
    for key in path:
        data = data.get(key, {}) if isinstance(data, dict) else {}
    return data


# ------------------------------------------------------------------
# Units
# ------------------------------------------------------------------
def _to_inches(value: float, unit: str) -> float:
    return value * p1.CM_TO_IN if unit == "cm" else value


def _from_inches(value: float, unit: str) -> float:
    return value * p1.IN_TO_CM if unit == "cm" else value


# ------------------------------------------------------------------
# Wizard -> preset dict
# ------------------------------------------------------------------
def _collect_page_setup(w) -> dict:
    unit = _widget(w, QComboBox, "comboUnits").currentText()

    def inches(spin_name):
        return round(_to_inches(_get_spin(w, spin_name), unit), 3)

    size_key = p1.page_size_key_for_combo_text(
        _widget(w, QComboBox, "comboPageSize").currentText()
    )
    mirrored = _get_check(w, "checkMirroredMargins")
    left_key, right_key = ("inside", "outside") if mirrored else ("left", "right")
    landscape = _widget(w, QRadioButton, "radioLandscape").isChecked()

    return {
        "size_preset": _SIZE_PRESET_SLUGS.get(size_key, "custom"),
        "width_in": inches("spinCustomWidth"),
        "height_in": inches("spinCustomHeight"),
        "display_unit": "in",
        "orientation": "landscape" if landscape else "portrait",
        "mirrored_margins": mirrored,
        "margins": {
            "top": inches("spinMarginTop"),
            "bottom": inches("spinMarginBottom"),
            left_key: inches("spinMarginLeft"),
            right_key: inches("spinMarginRight"),
        },
    }


def _collect_advanced(w) -> dict:
    result = {}
    for path, fields in _ADVANCED_FIELDS.items():
        *parents, leaf = path
        node = result
        for key in parents:
            node = node.setdefault(key, {})
        node[leaf] = _read_fields(w, fields)
    return result


def collect_wizard_settings(main_window) -> dict:
    """
    Reads the current state of pages 1-4 and returns a dict matching the
    preset schema (minus preset_name/schema_version, which are added by
    whatever saves this as a named preset).
    """
    w = main_window.window
    return {
        "schema_version": SCHEMA_VERSION,
        "page_setup": _collect_page_setup(w),
        "typography_basic": _read_fields(w, _BASIC_FIELDS),
        "typography_advanced": _collect_advanced(w),
        "additional_options": _read_fields(w, _OPTION_FIELDS),
    }


# ------------------------------------------------------------------
# Preset dict -> wizard
# ------------------------------------------------------------------
def _apply_page_setup(main_window, ps: dict) -> None:
    w = main_window.window
    unit = _widget(w, QComboBox, "comboUnits").currentText()  # keep the user's display unit

    # A checked lock would overwrite the partner spinbox mid-load.
    for btn in (main_window._btnLockTB, main_window._btnLockLR, main_window._btnLockSize):
        btn.setChecked(False)

    if "mirrored_margins" in ps:
        main_window._checkMirrored.setChecked(bool(ps["mirrored_margins"]))
    mirrored = main_window._checkMirrored.isChecked()

    orientation = ps.get("orientation")
    if orientation == "landscape":
        main_window._radioLandscape.setChecked(True)
    elif orientation == "portrait":
        main_window._radioPortrait.setChecked(True)

    # Only touch the size combo if the preset says something about it;
    # "custom" (or an unknown slug) selects the Custom entry.
    slug = ps.get("size_preset")
    if slug is not None:
        p1._set_combo_page_size(main_window, _SLUG_TO_SIZE_KEY.get(slug))

    main_window._updating_page_size_controls = True
    try:
        for name, key in (("spinCustomWidth", "width_in"), ("spinCustomHeight", "height_in")):
            value = ps.get(key)
            if value is not None:
                _set_spin(w, name, _from_inches(value, unit))
    finally:
        main_window._updating_page_size_controls = False
    p1._remember_page_size_aspect_ratio(main_window)

    margins = ps.get("margins", {})
    left_key, right_key = ("inside", "outside") if mirrored else ("left", "right")
    for name, key in (
        ("spinMarginTop", "top"), ("spinMarginBottom", "bottom"),
        ("spinMarginLeft", left_key), ("spinMarginRight", right_key),
    ):
        value = margins.get(key)
        if value is not None:
            _set_spin(w, name, _from_inches(value, unit))

    p1._sync_preview(main_window)


def apply_preset_to_wizard(main_window, data: dict) -> None:
    """Inverse of collect_wizard_settings: push a loaded preset into pages 1-4."""
    w = main_window.window

    _apply_page_setup(main_window, data.get("page_setup", {}))

    # Basic before advanced: if basic fonts propagate into advanced widgets
    # via signals, the preset's advanced values then overwrite that.
    _write_fields(w, _BASIC_FIELDS, data.get("typography_basic", {}))

    advanced = data.get("typography_advanced", {})
    for path, fields in _ADVANCED_FIELDS.items():
        _write_fields(w, fields, _dig(advanced, path))

    _write_fields(w, _OPTION_FIELDS, data.get("additional_options", {}))
    main_window._advanced_edited = True

# ------------------------------------------------------------------
# Derived values for preset
# ------------------------------------------------------------------
def derive_advanced_from_basic(main_window) -> None:
    """Fill page 3 font/size widgets from the basic fonts/sizes (page 2).
    No-op once the user has edited page 3."""
    if main_window._advanced_edited:
        return
    w = main_window.window

    fonts = {"header": _get_font(w, "comboHeaderFont"), "body": _get_font(w, "comboBodyFont")}
    sizes = {"header": _get_spin(w, "spinHeaderFontSize"), "body": _get_spin(w, "spinBodyFontSize")}

    was_edited = main_window._advanced_edited      # derivation must not count as an edit
    try:
        for widget, source in _DERIVED_FONTS:
            _set_font(w, widget, fonts[source])
        for widget, source, ratio in _DERIVED_SIZES:
            _set_spin(w, widget, round(sizes[source] * ratio * 2) / 2)
    finally:
        main_window._advanced_edited = was_edited


def track_advanced_edits(main_window) -> None:
    """Call once from wire_page3, after the widgets exist."""
    w = main_window.window
    mark = lambda *_: setattr(main_window, "_advanced_edited", True)
    signal_for = {
        "spin":  lambda x: x.valueChanged,
        "check": lambda x: x.toggled,
        "align": lambda x: x.currentIndexChanged,
        "font":  lambda x: x.currentFontChanged,
    }
    cls_for = {"spin": QDoubleSpinBox, "check": QCheckBox, "align": QComboBox, "font": QFontComboBox}
    for fields in _ADVANCED_FIELDS.values():
        for _key, widget, kind in fields:
            signal_for[kind](_widget(w, cls_for[kind], widget)).connect(mark)