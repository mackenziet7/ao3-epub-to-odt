from com.sun.star.style.PageStyleLayout import MIRRORED, ALL
import uno

from .uno_utils import inches, pt, prop, fixed_ls, prop_ls

ALIGN = {"left": 0, "right": 1, "justify": 2, "center": 3}

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

def _apply_dims(ps, page):
    w, h = page["width_in"], page["height_in"]
    landscape = page["orientation"] == "landscape"
    if landscape and w < h:          # preset stores portrait dims
        w, h = h, w
    ps.IsLandscape  = landscape
    ps.Width        = inches(w)
    ps.Height       = inches(h)
    ps.TopMargin    = inches(page["margins"]["top"])
    ps.BottomMargin = inches(page["margins"]["bottom"])
    ps.FooterIsOn   = False

def apply_book_page_dims(ps, page):
    _apply_dims(ps, page)
    m = page["margins"]
    ps.PageStyleLayout = MIRRORED if page["mirrored_margins"] else ALL
    ps.LeftMargin  = inches(m["inside"])    # in MIRRORED, Left = inside
    ps.RightMargin = inches(m["outside"])

def apply_frontmatter_page_dims(ps, page, is_verso=False):
    _apply_dims(ps, page)
    m = page["margins"]
    ps.PageStyleLayout = ALL                # no recto/verso enforcement → no auto blanks
    if page["mirrored_margins"] and is_verso:
        ps.LeftMargin, ps.RightMargin = inches(m["outside"]), inches(m["inside"])
    else:
        ps.LeftMargin, ps.RightMargin = inches(m["inside"]), inches(m["outside"])

def setup_page_style(doc,  page):
    # ── Default Page Style: running headers on, mirrored ──────────────────
    ps = get_default_page_style(doc)
    apply_book_page_dims(ps, page)
    ps.HeaderIsOn         = True
    ps.HeaderIsShared     = False
    ps.HeaderBodyDistance = pt(18)

    # ── ChapterFirstPage: same dims, NO header ─────────────────────────────
    cfp = get_or_create_page_style(doc, "ChapterFirstPage")
    apply_book_page_dims(cfp, page)
    cfp.HeaderIsOn  = False
    default_name    = get_default_page_style(doc).Name
    cfp.FollowStyle = default_name

    # ── FrontMatterRecto: odd/right pages — no header ─────────────────────
    fmr = get_or_create_page_style(doc, "FrontMatterRecto")
    apply_frontmatter_page_dims(fmr, page, is_verso=False)
    fmr.HeaderIsOn  = False

    # ── FrontMatterVerso: even/left pages — no header ─────────────────────
    fmv = get_or_create_page_style(doc, "FrontMatterVerso")
    apply_frontmatter_page_dims(fmv, page, is_verso=True)
    fmv.HeaderIsOn  = False

    # ── AppendixPage: same dims, NO header ────────────────────────────────────
    ap = get_or_create_page_style(doc, "AppendixPage")
    apply_book_page_dims(ap, page)
    ap.HeaderIsOn  = False
    ap.FollowStyle = "AppendixPage"  # stays in appendix mode for all subsequent pages

    # ── Disable forced recto starts to prevent automatic blank pages ───────
    for style in [ps, cfp, fmr, fmv]:
        try:
            style.setPropertyValue("FirstIsRightPage", False)
        except Exception:
            pass  # property may not exist in this LO version

    print(f"  [✓] Page: {page['size_preset']}, mirrored={page['mirrored_margins']}, page styles created")

def _apply(s, d):
    """Apply preset keys to a style; missing/None keys are skipped (inherit)."""
    def has(k): return d.get(k) is not None
    if has("font"):                 s.CharFontName = d["font"]
    if has("size_pt"):              s.CharHeight = d["size_pt"]
    if has("bold"):                 s.CharWeight = 150 if d["bold"] else 100
    if has("italic"):               s.CharPosture = 2 if d["italic"] else 0
    if has("alignment"):            s.ParaAdjust = ALIGN[d["alignment"]]
    if has("top_margin_in"):        s.ParaTopMargin = inches(d["top_margin_in"])
    if has("bottom_margin_in"):     s.ParaBottomMargin = inches(d["bottom_margin_in"])
    if has("left_margin_in"):       s.ParaLeftMargin = inches(d["left_margin_in"])
    if has("first_line_indent_in"): s.ParaFirstLineIndent = inches(d["first_line_indent_in"])
    if has("line_spacing_in"):      s.ParaLineSpacing = fixed_ls(inches(d["line_spacing_in"]))

def create_para_styles(doc, adv):
    mb, fm, ap = adv["main_book"], adv["front_matter"], adv["appendix"]

    body = get_or_create_style(doc, "MyBody")
    _apply(body, mb["body"])
    body.ParaTopMargin = 0
    body.ParaBottomMargin = 0
    body.ParaOrphans = 2
    body.ParaWidows = 2

    first = get_or_create_style(doc, "MyBodyFirst", "MyBody")
    first.ParaFirstLineIndent = (
        0 if mb["body"]["no_indent_on_first_paragraph_after_heading"]
        else inches(mb["body"]["first_line_indent_in"])
    )

    sb = get_or_create_style(doc, "SceneBreak", "MyBody")
    sb.ParaAdjust = 3
    sb.ParaFirstLineIndent = 0

    front = get_or_create_style(doc, "FrontMatter")
    _apply(front, fm["body"])
    front.ParaFirstLineIndent = 0
    front.ParaLineSpacing = prop_ls(100)

    chap = get_or_create_style(doc, "ChapHeads")
    _apply(chap, mb["chapter_headers"])
    chap.ParaFirstLineIndent = 0
    chap.OutlineLevel = 1

    fmhead = get_or_create_style(doc, "FrontMatterHead")
    _apply(fmhead, fm["head"])
    fmhead.ParaFirstLineIndent = 0
    fmhead.OutlineLevel = 0

    qr = get_or_create_style(doc, "QRCodeBlock")
    _apply(qr, fm["qr_code"])
    qr.CharHeight = 12.0
    qr.ParaFirstLineIndent = 0
    qr.OutlineLevel = 0

    qrcap = get_or_create_style(doc, "QRCodeCaption")
    _apply(qrcap, fm["qr_caption"])
    qrcap.ParaAdjust = 3
    qrcap.ParaFirstLineIndent = 0
    qrcap.ParaTopMargin = 0
    qrcap.ParaBottomMargin = pt(18)
    qrcap.OutlineLevel = 0

    note_label = get_or_create_style(doc, "AppendixNoteLabel")
    _apply(note_label, ap["note_label"])
    note_label.ParaFirstLineIndent = 0
    note_label.ParaLeftMargin = 0
    note_label.ParaLineSpacing = prop_ls(100)

    note = get_or_create_style(doc, "AppendixNote")
    _apply(note, ap["note"])
    note.ParaFirstLineIndent = 0
    note.ParaLineSpacing = prop_ls(100)

    ahead = get_or_create_style(doc, "AppendixHead")
    _apply(ahead, ap["head"])

    print("  [✓] Paragraph styles created")
