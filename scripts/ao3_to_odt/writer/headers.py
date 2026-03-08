import uno
from .uno_utils import inches, pt
from .styles import get_default_page_style


def setup_headers(doc, meta):
    """
    Running headers using HeaderTextRight (odd/right pages) and HeaderTextLeft
    (even/left pages) — the correct UNO property names for mirrored page styles.

    Right (odd) pages:  TITLE <tab> page#
    Left  (even) pages: page# <tab> AUTHOR
    """
    ps = get_default_page_style(doc)

    title_upper  = meta.title.upper()
    author_upper = meta.author.upper()

    def make_tab_stop(position):
        tab = uno.createUnoStruct("com.sun.star.style.TabStop")
        tab.Position    = position
        tab.Alignment   = 2        # RIGHT
        tab.FillChar    = ord(' ')
        tab.DecimalChar = ord('.')
        return tab

    tab = make_tab_stop(inches(4.0))

    # ── Right page header (odd pages): TITLE <tab> pagenum ──────────────────
    rh_text = ps.getPropertyValue("HeaderTextRight")
    rc = rh_text.createTextCursor()
    rc.gotoStart(False)
    rc.gotoEnd(True)
    uno.invoke(rc, "setPropertyValue", ("ParaTabStops", uno.Any(
        "[]com.sun.star.style.TabStop", (tab,))))
    rc.setPropertyValue("CharHeight", 8.0)
    rc.gotoEnd(False)
    rh_text.insertString(rc, title_upper + "\t", False)
    pf = doc.createInstance("com.sun.star.text.TextField.PageNumber")
    pf.setPropertyValue("SubType", 1)       # CURRENT
    pf.setPropertyValue("NumberingType", 4) # ARABIC
    rh_text.insertTextContent(rc, pf, False)

    # ── Left page header (even pages): pagenum <tab> AUTHOR ─────────────────
    lh_text = ps.getPropertyValue("HeaderTextLeft")
    lc = lh_text.createTextCursor()
    lc.gotoStart(False)
    lc.gotoEnd(True)
    uno.invoke(lc, "setPropertyValue", ("ParaTabStops", uno.Any(
        "[]com.sun.star.style.TabStop", (tab,))))
    lc.setPropertyValue("CharHeight", 8.0)
    lc.gotoEnd(False)
    pf2 = doc.createInstance("com.sun.star.text.TextField.PageNumber")
    pf2.setPropertyValue("SubType", 1)
    pf2.setPropertyValue("NumberingType", 4)
    lh_text.insertTextContent(lc, pf2, False)
    lh_text.insertString(lc, "\t" + author_upper, False)

    print("  [✓] Running headers")