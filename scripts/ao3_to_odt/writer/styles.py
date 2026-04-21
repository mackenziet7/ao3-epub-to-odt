from com.sun.star.style.PageStyleLayout import MIRRORED, ALL
import uno

from .uno_utils import inches, pt, prop, fixed_ls, prop_ls

def get_or_create_style(doc, name, parent="Standard"):
    styles = doc.getStyleFamilies().getByName("ParagraphStyles")
    if not styles.hasByName(name):
        s = doc.createInstance("com.sun.star.style.ParagraphStyle")
        styles.insertByName(name, s)
    s = styles.getByName(name)
    if parent and styles.hasByName(parent):
        s.setParentStyle(parent)
    return s

def get_default_page_style(doc):
    page_styles = doc.getStyleFamilies().getByName("PageStyles")
    for name in ("Default Page Style", "Default", "Standard", "Стандартный"):
        if page_styles.hasByName(name):
            return page_styles.getByName(name)
    return page_styles.getByIndex(0)

def get_or_create_page_style(doc, name):
    page_styles = doc.getStyleFamilies().getByName("PageStyles")
    if not page_styles.hasByName(name):
        s = doc.createInstance("com.sun.star.style.PageStyle")
        page_styles.insertByName(name, s)
    return page_styles.getByName(name)

# Apply 5.5x8.5 mirrored layout to a page style.
def apply_book_page_dims(ps):
    ps.IsLandscape     = False
    ps.Width           = inches(5.5)
    ps.Height          = inches(8.5)
    ps.PageStyleLayout = MIRRORED
    ps.TopMargin       = inches(0.64)
    ps.BottomMargin    = inches(0.60)
    ps.LeftMargin      = inches(0.90)
    ps.RightMargin     = inches(0.60)
    ps.FooterIsOn      = False

def apply_frontmatter_page_dims(ps, is_verso=False):
    ps.IsLandscape     = False
    ps.Width           = inches(5.5)
    ps.Height          = inches(8.5)
    ps.PageStyleLayout = ALL          # no recto/verso enforcement → no auto blanks
    ps.TopMargin       = inches(0.64)
    ps.BottomMargin    = inches(0.60)
    # Manually mirror: spine (wider) margin swaps side depending on recto/verso
    if is_verso:
        ps.LeftMargin  = inches(0.60)   # outer
        ps.RightMargin = inches(0.90)   # spine
    else:
        ps.LeftMargin  = inches(0.90)   # spine
        ps.RightMargin = inches(0.60)   # outer
    ps.FooterIsOn      = False

def setup_page_style(doc):
    # ── Default Page Style: running headers on, mirrored ──────────────────
    ps = get_default_page_style(doc)
    apply_book_page_dims(ps)
    ps.HeaderIsOn         = True
    ps.HeaderIsShared     = False
    ps.HeaderBodyDistance = pt(18)

    # ── ChapterFirstPage: same dims, NO header ─────────────────────────────
    cfp = get_or_create_page_style(doc, "ChapterFirstPage")
    apply_book_page_dims(cfp)
    cfp.HeaderIsOn  = False
    default_name    = get_default_page_style(doc).Name
    cfp.FollowStyle = default_name

    # ── FrontMatterRecto: odd/right pages — no header ─────────────────────
    fmr = get_or_create_page_style(doc, "FrontMatterRecto")
    apply_frontmatter_page_dims(fmr, is_verso=False)
    fmr.HeaderIsOn  = False

    # ── FrontMatterVerso: even/left pages — no header ─────────────────────
    fmv = get_or_create_page_style(doc, "FrontMatterVerso")
    apply_frontmatter_page_dims(fmv, is_verso=True)
    fmv.HeaderIsOn  = False

    # ── AppendixPage: same dims, NO header ────────────────────────────────────
    ap = get_or_create_page_style(doc, "AppendixPage")
    apply_book_page_dims(ap)
    ap.HeaderIsOn  = False
    ap.FollowStyle = "AppendixPage"  # stays in appendix mode for all subsequent pages

    # ── Disable forced recto starts to prevent automatic blank pages ───────
    for style in [ps, cfp, fmr, fmv]:
        try:
            style.setPropertyValue("FirstIsRightPage", False)
        except Exception:
            pass  # property may not exist in this LO version

    print("  [✓] Page: 5.5×8.5\", mirrored margins, page styles created")

def create_para_styles(doc):
    body = get_or_create_style(doc, "MyBody")
    body.CharHeight           = 11.5
    body.CharFontName         = "Garamond"
    body.ParaAdjust           = 0        # left
    body.ParaFirstLineIndent  = inches(0.30)
    body.ParaLineSpacing      = fixed_ls(pt(16))
    body.ParaTopMargin        = 0
    body.ParaBottomMargin     = 0
    body.ParaOrphans          = 2
    body.ParaWidows           = 2

    first = get_or_create_style(doc, "MyBodyFirst", "MyBody")
    first.ParaFirstLineIndent = 0

    front = get_or_create_style(doc, "FrontMatter")
    front.CharHeight          = 9.0
    front.CharFontName        = "Garamond"
    front.ParaAdjust          = 0        # left
    front.ParaFirstLineIndent = 0
    front.ParaLineSpacing     = prop_ls(100)

    # ── ChapHeads: chapter titles only — OutlineLevel=1 so TOC picks them up
    chap = get_or_create_style(doc, "ChapHeads")
    chap.CharHeight           = 18.0
    chap.CharFontName         = "Garamond"
    chap.CharWeight           = 150      # bold
    chap.ParaAdjust           = 3        # center (block center)
    chap.ParaFirstLineIndent  = 0
    chap.ParaTopMargin        = pt(24)
    chap.ParaBottomMargin     = pt(18)
    chap.OutlineLevel         = 1        # ← indexed by TOC

    # ── FrontMatterHead: same visual style as ChapHeads but OutlineLevel=0
    fmhead = get_or_create_style(doc, "FrontMatterHead")
    fmhead.CharHeight          = 18.0
    fmhead.CharFontName        = "Garamond"
    fmhead.CharWeight          = 150
    fmhead.ParaAdjust          = 3
    fmhead.ParaFirstLineIndent = 0
    fmhead.ParaTopMargin       = pt(24)
    fmhead.ParaBottomMargin    = pt(18)
    fmhead.OutlineLevel        = 0

    qr = get_or_create_style(doc, "QRCodeBlock")
    qr.CharHeight          = 12.0
    qr.CharFontName        = "Garamond"
    qr.ParaAdjust          = 3              # CENTER
    qr.ParaFirstLineIndent = 0
    qr.ParaTopMargin       = pt(18)         # space above QR
    qr.ParaBottomMargin    = pt(6)          # tight gap to caption
    qr.OutlineLevel        = 0

    qrcap = get_or_create_style(doc, "QRCodeCaption")
    qrcap.CharHeight          = 10.0         # slightly smaller
    qrcap.CharFontName        = "Garamond"
    qrcap.CharPosture         = 2            # italic (optional, but very “book-like”)
    qrcap.ParaAdjust          = 3            # CENTER
    qrcap.ParaFirstLineIndent = 0
    qrcap.ParaTopMargin       = pt(0)
    qrcap.ParaBottomMargin    = pt(18)       # space below block
    qrcap.OutlineLevel        = 0

    note_label = get_or_create_style(doc, "AppendixNoteLabel")
    note_label.CharHeight          = 8.0
    note_label.CharFontName        = "Garamond"
    note_label.ParaAdjust          = 0
    note_label.ParaFirstLineIndent = 0
    note_label.ParaLeftMargin      = 0
    note_label.ParaLineSpacing     = prop_ls(100)

    note = get_or_create_style(doc, "AppendixNote")
    note.CharHeight           = 8.0
    note.CharFontName         = "Garamond"
    note.ParaAdjust           = 0
    note.ParaFirstLineIndent  = 0
    note.ParaLeftMargin       = inches(0.25)
    note.ParaLineSpacing      = prop_ls(100)

    ahead = get_or_create_style(doc, "AppendixHead")
    ahead.CharHeight          = 11.0
    ahead.CharFontName        = "Garamond"
    ahead.CharWeight          = 150
    ahead.ParaTopMargin       = pt(12)
    ahead.ParaBottomMargin    = pt(4)

    print("  [✓] Paragraph styles created")
