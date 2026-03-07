"""
ao3_to_odt.py  —  AO3 EPUB → Print-ready ODT
==============================================
Single script, run it with LibreOffice's Python directly.

STEP 1 — find LO's Python (one-time, run in regular PowerShell):
    dir "C:/Program Files/LibreOffice/program/python.exe"

STEP 2 — run this script with LO's Python every time:
    & "C:/Program Files/LibreOffice/program/python.exe" ao3_to_odt.py your_fic.epub

Output: your_fic_book.odt in the same folder as the epub.

Requires ebooklib + beautifulsoup4 installed into LO's Python.
On first run this script installs them automatically.
"""

import sys
import os
import re
import time
import socket
import subprocess
from pathlib import Path

# Force UTF-8 output so special characters like ✓ work on Windows
sys.stdout.reconfigure(encoding='utf-8')

# Make repo root importable during migration
sys.path.insert(0, str(Path(__file__).parent.parent))

# UNO — only available in LO's Python
try:
    import uno
    from com.sun.star.beans import PropertyValue
    from com.sun.star.text.ControlCharacter import PARAGRAPH_BREAK
    from com.sun.star.style.PageStyleLayout import MIRRORED
except ImportError:
    print("\nERROR: 'uno' module not found.")
    print("You must run this script with LibreOffice's Python, not your system Python.")
    print()
    print("Run it like this:")
    print('  & "C:\\Program Files\\LibreOffice\\program\\python.exe" ao3_to_odt.py your_fic.epub')
    sys.exit(1)

# Third-party
import ebooklib
from ebooklib import epub
# Suppress BeautifulSoup XML-parsed-as-HTML warnings (AO3 epubs are XHTML)
from bs4 import BeautifulSoup, XMLParsedAsHTMLWarning
import warnings
warnings.filterwarnings("ignore", category=XMLParsedAsHTMLWarning)

# Local
from scripts.ao3_to_odt.epub.models import (
    AO3Metadata, Run, Paragraph, Chapter, AO3Book
)

from scripts.ao3_to_odt.epub.text_utils import (
    clean_text, clean_run, SCENE_BREAK_RE
)

from scripts.ao3_to_odt.epub.parser import parse_epub

# ═══════════════════════════════════════════════════════════════════════════════
# PART 2 — LIBREOFFICE UNO DOCUMENT BUILDER
# ═══════════════════════════════════════════════════════════════════════════════

def inches(n): return int(n * 2540)
def pt(n):     return int(n * 35.3)

def prop(name, value):
    p = PropertyValue()
    p.Name = name; p.Value = value
    return p

def fixed_ls(h):
    ls = uno.createUnoStruct("com.sun.star.style.LineSpacing")
    ls.Mode = 1; ls.Height = h; return ls

def prop_ls(pct):
    ls = uno.createUnoStruct("com.sun.star.style.LineSpacing")
    ls.Mode = 0; ls.Height = pct; return ls


# ── LO connection ──────────────────────────────────────────────────────────────

def find_soffice():
    import shutil, glob
    # Hard-coded paths first (prefer .exe over .COM on Windows)
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/usr/lib/libreoffice/program/soffice",
        "/usr/local/bin/soffice",
    ]
    # Also search for any LO version installed under Program Files
    for pattern in [r"C:\Program Files\LibreOffice*\program\soffice.exe"]:
        found = glob.glob(pattern)
        candidates.extend(sorted(found))
    for c in candidates:
        if Path(c).exists(): return c
    # Last resort: whatever shutil finds (may be .COM on Windows)
    which = shutil.which("soffice")
    if which: return which
    return None

def is_port_open(port):
    try:
        with socket.create_connection(("localhost", port), timeout=1): return True
    except: return False

def clear_lo_locks():
    """
    Remove LO lock files that cause it to crash on startup after a previous
    unclean exit. These live in the LO user profile directory.
    """
    import glob
    profile_dirs = []
    if sys.platform == "win32":
        appdata = os.environ.get("APPDATA", "")
        if appdata:
            profile_dirs.append(os.path.join(appdata, "LibreOffice", "4", "user"))
    else:
        home = os.path.expanduser("~")
        profile_dirs.append(os.path.join(home, ".config", "libreoffice", "4", "user"))

    removed = []
    for profile in profile_dirs:
        for lock_pattern in [".~lock.*", "*.lock"]:
            for f in glob.glob(os.path.join(profile, lock_pattern)):
                try:
                    os.remove(f)
                    removed.append(f)
                except OSError:
                    pass
    if removed:
        print(f"  Cleared {len(removed)} stale lock file(s)")


def start_lo_listener(soffice_path, port=2002):
    # Clean up any stale locks before starting
    clear_lo_locks()

    accept = f"socket,host=localhost,port={port};urp;StarOffice.ServiceManager"

    # Use a temporary isolated user profile in the user's temp folder
    tmp_profile = os.path.join(os.environ.get("TEMP", os.path.expanduser("~")), "lo_ao3_profile")
    os.makedirs(tmp_profile, exist_ok=True)
    user_install = f"-env:UserInstallation={uno.systemPathToFileUrl(tmp_profile)}"

    cmd = [
        soffice_path,
        "--headless",
        "--norestore",
        "--nofirststartwizard",
        "--nologo",
        user_install,
        f"--accept={accept}",
    ]
    print(f"  Profile: {tmp_profile}")

    kwargs = {}
    if sys.platform == "win32":
        kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
    proc = subprocess.Popen(
        cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, **kwargs)
    return proc

def connect_uno(port=2002, retries=12, delay=3):
    localContext = uno.getComponentContext()
    resolver = localContext.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", localContext)
    url = f"uno:socket,host=localhost,port={port};urp;StarOffice.ComponentContext"
    last_err = None
    for attempt in range(1, retries + 1):
        try:
            print(f"  Connecting to LO UNO (attempt {attempt}/{retries})...", end="", flush=True)
            ctx = resolver.resolve(url)
            desktop = ctx.ServiceManager.createInstanceWithContext(
                "com.sun.star.frame.Desktop", ctx)
            print(" ok")
            return desktop
        except Exception as e:
            last_err = e
            print(f" waiting {delay}s...")
            time.sleep(delay)
    raise ConnectionError(f"Could not connect after {retries} attempts: {last_err}")


# ── Document construction ──────────────────────────────────────────────────────

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
    """Get the default page style regardless of LO language/locale."""
    page_styles = doc.getStyleFamilies().getByName("PageStyles")
    # Try common names first
    for name in ("Default Page Style", "Default", "Standard", "Стандартный"):
        if page_styles.hasByName(name):
            return page_styles.getByName(name)
    # Fall back to the first page style (always the default)
    return page_styles.getByIndex(0)

def get_or_create_page_style(doc, name):
    page_styles = doc.getStyleFamilies().getByName("PageStyles")
    if not page_styles.hasByName(name):
        s = doc.createInstance("com.sun.star.style.PageStyle")
        page_styles.insertByName(name, s)
    return page_styles.getByName(name)

def apply_book_page_dims(ps):
    """Apply 5.5x8.5 mirrored layout to a page style."""
    ps.IsLandscape     = False
    ps.Width           = inches(5.5)
    ps.Height          = inches(8.5)
    ps.PageStyleLayout = MIRRORED
    ps.TopMargin       = inches(0.64)
    ps.BottomMargin    = inches(0.60)
    ps.LeftMargin      = inches(0.90)   # inner/binding
    ps.RightMargin     = inches(0.60)   # outer
    ps.FooterIsOn      = False

def setup_page_style(doc):
    # ── Default Page Style: running headers on, mirrored ──────────────────
    ps = get_default_page_style(doc)
    apply_book_page_dims(ps)
    ps.HeaderIsOn           = True
    ps.HeaderIsShared       = False
    ps.HeaderBodyDistance   = pt(18)

    # ── ChapterFirstPage: same dims, NO header ─────────────────────────────
    # Used for the first page of every chapter (title page of chapter).
    # FollowStyle = default so subsequent pages automatically get headers.
    cfp = get_or_create_page_style(doc, "ChapterFirstPage")
    apply_book_page_dims(cfp)
    cfp.HeaderIsOn    = False
    default_name = get_default_page_style(doc).Name
    cfp.FollowStyle   = default_name

    # ── FrontMatterPage: same dims, NO header, NO page number ─────────────
    # Used for half-title, full-title, copyright, TOC pages.
    fmp = get_or_create_page_style(doc, "FrontMatterPage")
    apply_book_page_dims(fmp)
    fmp.HeaderIsOn    = False
    fmp.FollowStyle   = "FrontMatterPage"  # stays front matter until we switch

    print("  [\u2713] Page: 5.5\u00d78.5\", mirrored margins, page styles created")

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
    # Used for half-title, full title, and "Contents" heading so they are
    # NOT picked up by the TOC.
    fmhead = get_or_create_style(doc, "FrontMatterHead")
    fmhead.CharHeight          = 18.0
    fmhead.CharFontName        = "Garamond"
    fmhead.CharWeight          = 150
    fmhead.ParaAdjust          = 3
    fmhead.ParaFirstLineIndent = 0
    fmhead.ParaTopMargin       = pt(24)
    fmhead.ParaBottomMargin    = pt(18)
    fmhead.OutlineLevel        = 0       # ← NOT indexed by TOC

    note = get_or_create_style(doc, "AppendixNote")
    note.CharHeight           = 8.0
    note.CharFontName         = "Garamond"
    note.ParaAdjust           = 0
    note.ParaFirstLineIndent  = 0
    note.ParaLeftMargin       = inches(0.5)
    note.ParaLineSpacing      = prop_ls(100)

    ahead = get_or_create_style(doc, "AppendixHead")
    ahead.CharHeight          = 11.0
    ahead.CharFontName        = "Garamond"
    ahead.CharWeight          = 150
    ahead.ParaTopMargin       = pt(12)
    ahead.ParaBottomMargin    = pt(4)

    print("  [✓] Paragraph styles created")

def ins(text_obj, cursor, content, style, page_break=False, page_style=None, page_number=None):
    """
    Insert one paragraph.
    page_style: if set, start a new page using this page style
    page_number: if set, reset the page counter to this number
    """
    if page_style:
        # Setting PageDescName triggers a page break and switches page style.
        # This is the correct LO UNO way — do NOT combine with BreakType.
        cursor.setPropertyValue("PageDescName", page_style)
        cursor.setPropertyValue("PageNumberOffset", page_number if page_number is not None else 0)
    elif page_break:
        cursor.setPropertyValue("BreakType", 4)   # PAGE_BEFORE
    cursor.setPropertyValue("ParaStyleName", style)
    text_obj.insertString(cursor, content, False)
    text_obj.insertControlCharacter(cursor, PARAGRAPH_BREAK, False)
    # Clear page desc so subsequent paragraphs don't keep triggering breaks
    if page_style:
        cursor.setPropertyValue("PageDescName", "")

def blank(text_obj, cursor):
    """Insert a blank paragraph. Never triggers page breaks."""
    cursor.setPropertyValue("ParaStyleName", "Standard")
    cursor.setPropertyValue("PageDescName", "")
    text_obj.insertControlCharacter(cursor, PARAGRAPH_BREAK, False)

def blank_with_style(text_obj, cursor, page_style, page_number=1):
    """Insert a blank paragraph that also starts a new page with page_style."""
    cursor.setPropertyValue("PageDescName", page_style)
    cursor.setPropertyValue("PageNumberOffset", page_number)
    cursor.setPropertyValue("ParaStyleName", "Standard")
    text_obj.insertControlCharacter(cursor, PARAGRAPH_BREAK, False)
    cursor.setPropertyValue("PageDescName", "")

def build_content(doc, book, include_toc=True, toc_objects=None):
    text   = doc.getText()
    cursor = text.createTextCursor()
    cursor.gotoStart(False)

    m = book.metadata

    # ── Half title (p.1) — FrontMatterHead so it doesn't appear in TOC
    # Page style set on the first blank, then the rest are plain blanks
    blank_with_style(text, cursor, "FrontMatterPage")
    for _ in range(5): blank(text, cursor)
    ins(text, cursor, m.title.upper(), "FrontMatterHead")

    # ── Full title + author (p.3) ──────────────────────────────────────────────
    ins(text, cursor, "", "Standard", page_style="FrontMatterPage")
    for _ in range(5): blank(text, cursor)
    ins(text, cursor, m.title,           "FrontMatterHead")
    ins(text, cursor, f"by {m.author}",  "FrontMatter")
    blank(text, cursor); blank(text, cursor)
    if m.ao3_url:
        ins(text, cursor, m.ao3_url, "FrontMatter")

    # ── Front matter (p.4 verso) ───────────────────────────────────────────────
    ins(text, cursor, "", "Standard", page_style="FrontMatterPage")
    fields = []
    if m.rating:         fields.append(f"Rating: {m.rating}")
    if m.warnings:       fields.append(f"Warnings: {', '.join(m.warnings)}")
    if m.category:       fields.append(f"Category: {m.category}")
    if m.fandom:         fields.append(f"Fandom: {', '.join(m.fandom)}")
    if m.relationships:  fields.append(f"Relationships: {', '.join(m.relationships)}")
    if m.characters:     fields.append(f"Characters: {', '.join(m.characters)}")
    if m.tags:           fields.append(f"Tags: {', '.join(m.tags)}")
    if m.language:       fields.append(f"Language: {m.language}")
    if m.published:      fields.append(f"Published: {m.published}")
    if m.completed:      fields.append(f"Completed: {m.completed}")
    if m.words:
        wf = f"{int(m.words):,}" if m.words.isdigit() else m.words
        fields.append(f"Words: {wf}")
    for line in fields:
        ins(text, cursor, line, "FrontMatter")
    if m.summary:
        blank(text, cursor)
        ins(text, cursor, "Summary:", "FrontMatter")
        ins(text, cursor, m.summary,  "FrontMatter")
    print("  [✓] Front matter")

    # ── TOC — "Contents" heading uses FrontMatterHead (not indexed by TOC)
    if include_toc:
        ins(text, cursor, "", "Standard", page_style="FrontMatterPage")
        ins(text, cursor, "Contents", "FrontMatterHead")
        toc = doc.createInstance("com.sun.star.text.ContentIndex")
        toc.CreateFromOutline = True
        toc.CreateFromChapter = False
        toc.Level             = 1
        toc.Title             = ""
        text.insertTextContent(cursor, toc, False)
        text.insertControlCharacter(cursor, PARAGRAPH_BREAK, False)
        if toc_objects is not None:
            toc_objects.append(toc)  # save reference for refresh after content is built
        print("  [✓] TOC")

    # ── Chapters ──────────────────────────────────────────────────────────────
    # LO API limitation: PageDescName (page style switch) ALWAYS pairs with
    # PageNumberOffset as an explicit reset — there is no "continue" value.
    # Solution: only use PageDescName for chapter 1 (resets to p.1).
    # Chapters 2+ use a plain BreakType=4 page break which continues the count.
    # Trade-off: chapters 2+ show a header on their title page.
    for i, ch in enumerate(book.chapters):
        if i == 0:
            # Reset page count to 1 at chapter 1, using the default page style
            default_style_name = get_default_page_style(doc).Name
            ins(text, cursor, ch.title, "ChapHeads",
                page_style=default_style_name, page_number=1)
        else:
            ins(text, cursor, ch.title, "ChapHeads", page_break=True)
        first_para = True
        for para in ch.body:
            if para.type == 'break':
                cursor.setPropertyValue("ParaStyleName", "MyBody")
                cursor.setPropertyValue("ParaAdjust", 2)
                cursor.setPropertyValue("ParaFirstLineIndent", 0)
                text.insertString(cursor, "* * *", False)
                text.insertControlCharacter(cursor, PARAGRAPH_BREAK, False)
                # first_para stays False — paragraph after scene break is indented
                # (only the very first para of a chapter skips the indent)
            else:
                style = "MyBodyFirst" if first_para else "MyBody"
                cursor.setPropertyValue("ParaStyleName", style)
                for run in para.runs:
                    cursor.setPropertyValue("CharWeight",  150 if run.bold   else 100)
                    cursor.setPropertyValue("CharPosture",   2 if run.italic else 0)
                    text.insertString(cursor, run.text, False)
                cursor.setPropertyValue("CharWeight", 100)
                cursor.setPropertyValue("CharPosture", 0)
                text.insertControlCharacter(cursor, PARAGRAPH_BREAK, False)
                first_para = False
        print(f"  [✓] Chapter {ch.index}: {ch.title}")

    # ── Appendix ──────────────────────────────────────────────────────────────
    noted = [c for c in book.chapters if c.prenotes or c.endnotes]
    print(f"  Appendix: {len(noted)} chapters with notes")
    for c in noted:
        print(f"    - Ch {c.index} '{c.title}': prenotes={len(c.prenotes)} chars, endnotes={len(c.endnotes)} chars")
    if noted:
        ins(text, cursor, "Appendix: Notes", "ChapHeads", page_break=True)
        for ch in noted:
            ins(text, cursor, ch.title, "AppendixHead")
            for label, notes_text in [("Note:", ch.prenotes), ("End Note:", ch.endnotes)]:
                if notes_text:
                    ins(text, cursor, label, "AppendixNote")
                    for line in notes_text.split('\n'):
                        line = line.strip()
                        if line: ins(text, cursor, line, "AppendixNote")
        print(f"  [✓] Appendix ({len(noted)} chapters with notes)")

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
    uno.invoke(rc, "setPropertyValue", ("ParaTabStops", uno.Any("[]com.sun.star.style.TabStop", (tab,))))
    rc.setPropertyValue("CharHeight", 8.0)
    rc.gotoEnd(False)
    rh_text.insertString(rc, title_upper + "\t", False)
    pf = doc.createInstance("com.sun.star.text.TextField.PageNumber")
    pf.setPropertyValue("SubType", 1)        # CURRENT
    pf.setPropertyValue("NumberingType", 4)  # ARABIC
    rh_text.insertTextContent(rc, pf, False)

    # ── Left page header (even pages): pagenum <tab> AUTHOR ─────────────────
    lh_text = ps.getPropertyValue("HeaderTextLeft")
    lc = lh_text.createTextCursor()
    lc.gotoStart(False)
    lc.gotoEnd(True)
    uno.invoke(lc, "setPropertyValue", ("ParaTabStops", uno.Any("[]com.sun.star.style.TabStop", (tab,))))
    lc.setPropertyValue("CharHeight", 8.0)
    lc.gotoEnd(False)
    pf2 = doc.createInstance("com.sun.star.text.TextField.PageNumber")
    pf2.setPropertyValue("SubType", 1)
    pf2.setPropertyValue("NumberingType", 4)
    lh_text.insertTextContent(lc, pf2, False)
    lh_text.insertString(lc, "\t" + author_upper, False)

    print("  [✓] Running headers")


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════════


def save_odt(doc, out_path):
    url = uno.systemPathToFileUrl(os.path.abspath(out_path))
    doc.storeToURL(url, [prop("FilterName", "writer8"), prop("Overwrite", True)])
    print(f"  [✓] Saved: {out_path}")

def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    epub_path = str(Path(sys.argv[1]).resolve())
    if not Path(epub_path).exists():
        print(f"ERROR: File not found: {epub_path}")
        sys.exit(1)

    # --debug-chapter N  dumps the raw HTML of document N (0-indexed) so you
    # can see exactly how notes are structured in a specific chapter
    if '--debug-chapter' in sys.argv:
        idx = sys.argv.index('--debug-chapter')
        doc_num = int(sys.argv[idx + 1]) if idx + 1 < len(sys.argv) else 0
        import warnings, ebooklib
        from ebooklib import epub as _epub
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            ebook = _epub.read_epub(epub_path)
        items = list(ebook.get_items_of_type(ebooklib.ITEM_DOCUMENT))
        print(f"\nAll {len(items)} documents:")
        for i, item in enumerate(items):
            print(f"  [{i}] {item.get_name()}")
        print(f"\n--- Full HTML of document [{doc_num}]: {items[doc_num].get_name()} ---")
        print(items[doc_num].get_content().decode('utf-8', errors='replace'))
        sys.exit(0)

    # --debug-notes  parses the epub and shows all notes found per chapter
    if '--debug-notes' in sys.argv:
        book = parse_epub(epub_path)
        print(f"\n{len(book.chapters)} chapters parsed:")
        for ch in book.chapters:
            pre = f"{len(ch.prenotes)} chars" if ch.prenotes else "none"
            end = f"{len(ch.endnotes)} chars" if ch.endnotes else "none"
            print(f"  Ch {ch.index}: {ch.title}")
            print(f"    prenotes: {pre}")
            if ch.prenotes: print(f"      {ch.prenotes[:200]!r}")
            print(f"    endnotes: {end}")
            if ch.endnotes: print(f"      {ch.endnotes[:200]!r}")
        sys.exit(0)

    if len(sys.argv) > 2 and not sys.argv[2].startswith('--'):
        out_path = sys.argv[2]
    else:
        base = Path(epub_path).parent / (Path(epub_path).stem + "_book.odt")
        out_path = str(base)
        # If file already exists, add a counter suffix
        if Path(out_path).exists():
            counter = 2
            while Path(out_path).exists():
                out_path = str(base.with_stem(base.stem + f"_{counter}"))
                counter += 1
            print(f"  Output file already exists, saving as: {Path(out_path).name}")

    PORT = 2002

    # ── Parse epub ────────────────────────────────────────────────────────────
    print(f"\n{'='*60}\nParsing EPUB\n{'='*60}")
    book = parse_epub(epub_path)
    print(f"  Title:    {book.metadata.title}")
    print(f"  Author:   {book.metadata.author}")
    print(f"  Chapters: {len(book.chapters)}")
    print(f"  Words:    {book.metadata.words}")

    # ── Start LO listener ─────────────────────────────────────────────────────
    print(f"\n{'='*60}\nStarting LibreOffice\n{'='*60}")
    lo_process = None
    if is_port_open(PORT):
        print("  Already running, connecting...")
    else:
        soffice = find_soffice()
        if not soffice:
            print("ERROR: Cannot find soffice executable.")
            sys.exit(1)
        print(f"  Launching: {soffice}")
        lo_process = start_lo_listener(soffice, PORT)
        print("  Waiting for LO to start", end="", flush=True)
        for _ in range(40):
            if is_port_open(PORT):
                break
            # Check if process died early
            if lo_process.poll() is not None:
                stdout, stderr = lo_process.communicate()
                print(f"\n  ERROR: LO exited with code {lo_process.returncode}")
                if stderr: print(f"  stderr: {stderr.decode(errors='replace')[:500]}")
                sys.exit(1)
            print(".", end="", flush=True)
            time.sleep(1)
        print()
        # Wait for UNO bridge to be fully initialised (port open != UNO ready)
        print("  Port open, waiting for UNO bridge...", end="", flush=True)
        for _ in range(8):
            time.sleep(1)
            print(".", end="", flush=True)
        print(" ready!")

    # ── Build document ────────────────────────────────────────────────────────
    print(f"\n{'='*60}\nBuilding document\n{'='*60}")
    try:
        print("  Connecting...")
        desktop = connect_uno(PORT)
        print("  Connected. Creating document...")
        doc = None
        for attempt in range(6):
            try:
                doc = desktop.loadComponentFromURL(
                    "private:factory/swriter", "_blank", 0, [
                        prop("Hidden", True),
                        prop("MacroExecutionMode", 4),
                    ])
                break
            except Exception as e:
                if attempt < 5:
                    print(f"  Attempt {attempt+1} failed ({e}), retrying in 3s...")
                    time.sleep(3)
                else:
                    raise
        print("  Document created.")
        setup_page_style(doc)
        print("  Page style done.")
        create_para_styles(doc)
        print("  Para styles done.")
        include_toc = '--no-toc' not in sys.argv
        toc_objects = []
        build_content(doc, book, include_toc, toc_objects)
        print("  Content built.")
        if toc_objects:
            try:
                toc_objects[0].update()
                print("  [✓] TOC refreshed")
            except Exception as e:
                print(f"  TOC refresh failed (open in LO and press F9): {e}")
        setup_headers(doc, book.metadata)
        print("  Headers done.")
        save_odt(doc, out_path)
        time.sleep(2)      # let OS flush file buffers before closing
        doc.close(True)
        print("  Document closed.")
    except Exception as e:
        print(f"\nERROR during document build:\n  {e}")
        import traceback; traceback.print_exc()
        raise
    finally:
        if lo_process:
            lo_process.kill()
            try: lo_process.wait(timeout=5)
            except: pass
            subprocess.run(["taskkill", "/f", "/im", "soffice.exe"], capture_output=True)
            print("  LO shut down.")

    print(f"\n{'='*60}\nDONE\n{'='*60}")
    print(f"\n  Output: {out_path}")
    print("\n  In LibreOffice Writer:")
    print("  1. File > Export as PDF when ready to print")

if __name__ == "__main__":
    main()