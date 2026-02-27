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


# ─── Auto-install dependencies into LO's Python ───────────────────────────────

def ensure_deps():
    packages = {"ebooklib": "ebooklib", "beautifulsoup4": "bs4", "lxml": "lxml"}
    lo_python = sys.executable
    missing = []
    for pip_name, import_name in packages.items():
        try:
            __import__(import_name)
        except ImportError:
            missing.append(pip_name)

    if not missing:
        return  # all good

    print(f"Installing missing packages into LO Python: {missing}")
    subprocess.run([lo_python, "-m", "ensurepip", "--upgrade"], capture_output=True)
    for pkg in missing:
        print(f"  pip install {pkg}...", end="", flush=True)
        r = subprocess.run([lo_python, "-m", "pip", "install", pkg, "-q"],
                           capture_output=True, text=True)
        if r.returncode == 0:
            print(" ok")
        else:
            print(f" FAILED:\n{r.stderr}")
            sys.exit(1)
    print()

ensure_deps()

# Suppress BeautifulSoup XML-parsed-as-HTML warnings (AO3 epubs are XHTML)
from bs4 import XMLParsedAsHTMLWarning
import warnings
warnings.filterwarnings("ignore", category=XMLParsedAsHTMLWarning)

# Now safe to import
import ebooklib
from ebooklib import epub
from bs4 import BeautifulSoup

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


# ═══════════════════════════════════════════════════════════════════════════════
# PART 1 — EPUB PARSER
# ═══════════════════════════════════════════════════════════════════════════════

from dataclasses import dataclass, field
from typing import Optional

@dataclass
class AO3Metadata:
    title: str = ""
    author: str = ""
    rating: str = ""
    warnings: list = field(default_factory=list)
    category: str = ""
    fandom: list = field(default_factory=list)
    relationships: list = field(default_factory=list)
    characters: list = field(default_factory=list)
    tags: list = field(default_factory=list)
    language: str = ""
    published: str = ""
    completed: str = ""
    words: str = ""
    summary: str = ""
    ao3_url: str = ""

@dataclass
class Run:
    text: str = ""
    italic: bool = False
    bold: bool = False

@dataclass
class Paragraph:
    type: str = "para"
    runs: list = field(default_factory=list)

    @property
    def text(self):
        return "".join(r.text for r in self.runs)

@dataclass
class Chapter:
    index: int = 0
    title: str = ""
    prenotes: str = ""
    body: list = field(default_factory=list)
    endnotes: str = ""

@dataclass
class AO3Book:
    metadata: AO3Metadata = field(default_factory=AO3Metadata)
    chapters: list = field(default_factory=list)
    cover_image: Optional[bytes] = None
    cover_image_ext: str = "jpg"


def clean_text(text):
    """Full clean including strip — use for whole paragraphs/fields."""
    text = text.replace(' ', ' ').replace('​', '').replace('﻿', '')
    text = re.sub(r'(?<!\-)\-\-(?!\-)', '—', text)
    return re.sub(r'  +', ' ', text).strip()

def clean_run(text):
    """Light clean for inline runs — preserves leading/trailing spaces
    so italic spans don't eat the space before/after them."""
    text = text.replace(' ', ' ').replace('​', '').replace('﻿', '')
    text = re.sub(r'(?<!\-)\-\-(?!\-)', '—', text)
    text = re.sub(r'  +', ' ', text)
    return text  # no .strip()


# Matches text-based scene breaks: *  *  *, ---, ~~~, ###, etc.
SCENE_BREAK_RE = re.compile(r'^[\*\-~=#_ ]{2,}$')

def parse_body_html(soup):
    """
    Parse body HTML into a list of Paragraphs.

    Scene breaks come from four sources:
      1. <hr> tags — the standard AO3 HTML scene break
      2. <p> containing only break characters (* - ~ = _ #)
      3. <p class="...separator..."> or similar class names
      4. Images used as decorative dividers (treated as breaks)
    All are normalised to Paragraph(type='break').
    """
    def extract_runs(tag, italic=False, bold=False):
        runs = []
        for child in tag.children:
            if isinstance(child, str):
                t = clean_run(str(child))
                if t.strip(): runs.append(Run(text=t, italic=italic, bold=bold))
            elif hasattr(child, 'name'):
                if child.name in ('em', 'i'):
                    runs.extend(extract_runs(child, italic=True, bold=bold))
                elif child.name in ('strong', 'b'):
                    runs.extend(extract_runs(child, italic=italic, bold=True))
                elif child.name == 'br':
                    runs.append(Run(text='\n', italic=italic, bold=bold))
                else:
                    runs.extend(extract_runs(child, italic=italic, bold=bold))
        return runs

    def is_scene_break_tag(tag):
        """Return True if this tag represents a scene break."""
        if tag.name == 'hr':
            return True
        if tag.name == 'p':
            # Class-based: separator, scene-break, divider, etc.
            classes = ' '.join(tag.get('class', []))
            if re.search(r'sep|break|divid|rule', classes, re.I):
                return True
            # Content-based: only break characters
            text = tag.get_text(strip=True)
            if text and SCENE_BREAK_RE.match(text):
                return True
            # Empty <p> with only an <img> inside (decorative divider)
            children = [c for c in tag.children
                        if hasattr(c, 'name') or str(c).strip()]
            if len(children) == 1 and hasattr(children[0], 'name') and children[0].name == 'img':
                return True
        return False

    paragraphs = []
    # Iterate over all direct and nested block elements
    for tag in soup.find_all(['p', 'hr']):
        if is_scene_break_tag(tag):
            # Avoid consecutive duplicate breaks
            if not paragraphs or paragraphs[-1].type != 'break':
                paragraphs.append(Paragraph(type='break'))
        elif tag.name == 'p':
            runs = extract_runs(tag)
            if not runs: continue
            full = "".join(r.text for r in runs).strip()
            if not full: continue
            paragraphs.append(Paragraph(type='para', runs=runs))
    return paragraphs

def parse_info_page(soup, meta):
    def dd_values(dt_tag):
        dd = dt_tag.find_next_sibling('dd')
        if not dd: return []
        links = dd.find_all('a')
        if links: return [clean_text(a.get_text()) for a in links if a.get_text(strip=True)]
        text = clean_text(dd.get_text(separator=', '))
        return [t.strip() for t in text.split(',') if t.strip()]

    def get_field(label, root=None):
        for dt in (root or soup).find_all('dt'):
            if label.lower() in dt.get_text().lower():
                return dd_values(dt)
        return []

    def get_single(label, root=None):
        v = get_field(label, root)
        return v[0] if v else ""

    tags_dl = soup.find('dl', class_=re.compile(r'tags|meta', re.I)) or soup.find('dl')
    if not tags_dl: return meta

    r = get_single('Rating', tags_dl);      meta.rating = r if r else meta.rating
    w = get_field('Archive Warning', tags_dl) or get_field('Warning', tags_dl)
    if w: meta.warnings = w
    c = get_single('Categor', tags_dl);     meta.category = c if c else meta.category
    f = get_field('Fandom', tags_dl);       meta.fandom = f if f else meta.fandom
    rel = get_field('Relationship', tags_dl); meta.relationships = rel if rel else meta.relationships
    ch = get_field('Character', tags_dl);   meta.characters = ch if ch else meta.characters
    t = get_field('Additional Tag', tags_dl) or get_field('Freeform', tags_dl)
    if t: meta.tags = t
    lang = get_single('Language', tags_dl); meta.language = lang if lang else meta.language

    # Stats — Format A: nested dl.stats | Format B: plain text dd
    stats_dl = tags_dl.find('dl', class_=re.compile(r'stats', re.I))
    if stats_dl:
        pub = get_single('Published', stats_dl)
        if pub: meta.published = re.sub(r'T.*', '', pub)
        comp = get_single('Completed', stats_dl) or get_single('Updated', stats_dl)
        if comp: meta.completed = re.sub(r'T.*', '', comp)
        words = get_single('Words', stats_dl)
        if words: meta.words = words.replace(',', '')
    else:
        for dt in tags_dl.find_all('dt'):
            if 'stats' in dt.get_text().lower():
                dd = dt.find_next_sibling('dd')
                if dd:
                    txt = dd.get_text(separator='\n')
                    m = re.search(r'Published:\s*(\d{4}-\d{2}-\d{2})', txt)
                    if m and not meta.published: meta.published = m.group(1)
                    m = re.search(r'Words:\s*([\d,]+)', txt)
                    if m: meta.words = m.group(1).replace(',', '')
                    m = re.search(r'Completed:\s*(\d{4}-\d{2}-\d{2})', txt)
                    if m: meta.completed = m.group(1)
                    m = re.search(r'Chapters:\s*(\d+)/(\d+)', txt)
                    if m and not meta.completed and m.group(2) != '?' and m.group(1) == m.group(2):
                        meta.completed = meta.published
                break

    if not meta.ao3_url:
        for a in soup.find_all('a', href=re.compile(r'archiveofourown\.org/works/\d+')):
            meta.ao3_url = a['href'].split('?')[0]; break

    if not meta.summary:
        div = soup.find(class_=re.compile(r'summary', re.I))
        if div:
            bq = div.find('blockquote') or div
            meta.summary = clean_text(bq.get_text(separator=' '))

    return meta

def _bq_text(tag):
    """Get clean text from a blockquote or its first blockquote child."""
    bq = tag.find('blockquote') if tag.name != 'blockquote' else tag
    if bq:
        return clean_text(bq.get_text(separator='\n'))
    return clean_text(tag.get_text(separator='\n'))

NOTE_LABELS = ('chapter notes', 'chapter summary',
               'author note', "author's note", 'notes:')

def parse_chapter(item, index):
    """
    Parse a chapter document. Works on both AO3 native and Calibre exports.

    Strategy: search the whole document for known patterns rather than
    relying on container navigation, which breaks across epub variants.
    """
    ch = Chapter(index=index)
    soup = BeautifulSoup(item.get_content(), 'lxml')

    # ── Title ──
    t = (soup.find(class_=re.compile(r'toc-heading|heading', re.I)) or
         soup.find('h2') or soup.find('h3'))
    if t: ch.title = clean_text(t.get_text())

    # ── End notes ──
    # Search whole document for id="endnotes", "endnotes1", "endnotes2", etc.
    end_div = soup.find(id=re.compile(r'endnotes', re.I))
    if end_div:
        ch.endnotes = _bq_text(end_div)
        end_div.decompose()

    # Remove "See end of chapter for more notes" link divs
    for d in soup.find_all(class_='endnote-link'):
        d.decompose()

    # ── Pre-notes ──
    # Format A: <div id="notes"> or <div class="notes">
    nd = soup.find(id='notes') or soup.find(class_=re.compile(r'^notes$', re.I))
    if nd:
        ch.prenotes = _bq_text(nd)
        nd.decompose()
    else:
        # Format B: <p>Chapter Notes</p> immediately before <blockquote>
        # These can be nested inside calibre divs — search all <p> in document
        note_parts = []
        for p in soup.find_all('p'):
            label = p.get_text(strip=True).lower()
            if any(x in label for x in NOTE_LABELS):
                # blockquote may be a sibling of p, OR a sibling of p's parent
                bq = p.find_next_sibling('blockquote')
                if not bq:
                    # try parent's next sibling
                    bq = p.parent.find_next_sibling('blockquote') if p.parent else None
                if not bq:
                    # try the very next blockquote anywhere after this p
                    for sib in p.next_elements:
                        if hasattr(sib, 'name') and sib.name == 'blockquote':
                            bq = sib
                            break
                        if hasattr(sib, 'name') and sib.name == 'p':
                            break  # hit next paragraph, stop
                if bq:
                    note_parts.append(clean_text(bq.get_text(separator='\n')))
                    bq.decompose()
                p.decompose()
        if note_parts:
            ch.prenotes = '\n\n'.join(note_parts)

    # ── Body ──
    # Find the main content area — prefer userstuff1/2, fall back to body
    body = (soup.find(id='chapters') or
            soup.find(class_='userstuff2') or
            soup.find(class_='userstuff1') or
            soup.find('body'))
    if body:
        ch.body = parse_body_html(body)

    return ch
def parse_epub(epub_path):
    book = AO3Book()
    import warnings
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        ebook = epub.read_epub(epub_path)

    # OPF base metadata
    meta = AO3Metadata()
    titles = ebook.get_metadata('DC', 'title')
    if titles: meta.title = titles[0][0]
    creators = ebook.get_metadata('DC', 'creator')
    if creators: meta.author = creators[0][0]
    langs = ebook.get_metadata('DC', 'language')
    if langs: meta.language = langs[0][0]
    descs = ebook.get_metadata('DC', 'description')
    if descs:
        meta.summary = clean_text(BeautifulSoup(descs[0][0], 'lxml').get_text(' '))
    dates = ebook.get_metadata('DC', 'date')
    if dates: meta.published = re.sub(r'T.*', '', dates[0][0])
    book.metadata = meta

    # Cover image
    for item in ebook.get_items_of_type(ebooklib.ITEM_IMAGE):
        if 'cover' in item.get_name().lower():
            book.cover_image = item.get_content()
            book.cover_image_ext = item.media_type.split('/')[-1].replace('jpeg','jpg')
            break

    # Process documents
    chapter_count = 0
    for item in ebook.get_items_of_type(ebooklib.ITEM_DOCUMENT):
        content = item.get_content().decode('utf-8', errors='replace')
        soup = BeautifulSoup(content, 'lxml')

        has_meta_dl   = bool(soup.find('dl', class_=re.compile(r'tags|meta', re.I)))
        has_userstuff = bool(soup.find(class_=re.compile(r'userstuff', re.I)))
        is_preface    = bool(soup.find(id='preface'))
        userstuff_div = soup.find(class_=re.compile(r'userstuff', re.I))
        has_body      = userstuff_div and len(userstuff_div.find_all('p')) > 2 and not is_preface

        if has_meta_dl:
            book.metadata = parse_info_page(soup, book.metadata)
            if not book.metadata.ao3_url:
                for a in soup.find_all('a', href=re.compile(r'archiveofourown\.org/works/\d+')):
                    book.metadata.ao3_url = a['href'].split('?')[0]; break
        elif is_preface and not has_body:
            if not book.metadata.summary:
                bq = soup.find('blockquote', class_='userstuff')
                if bq: book.metadata.summary = clean_text(bq.get_text(' '))
            if not book.metadata.ao3_url:
                for a in soup.find_all('a', href=re.compile(r'archiveofourown\.org/works/\d+')):
                    book.metadata.ao3_url = a['href'].split('?')[0]; break
        elif has_body or has_userstuff:
            chapter_count += 1
            ch = parse_chapter(item, chapter_count)
            if not ch.title: ch.title = f"Chapter {chapter_count}"
            book.chapters.append(ch)

    return book


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

    chap = get_or_create_style(doc, "ChapHeads")
    chap.CharHeight           = 18.0
    chap.CharFontName         = "Garamond"
    chap.CharWeight           = 150      # bold
    chap.ParaAdjust           = 3        # center (block center)
    chap.ParaFirstLineIndent  = 0
    chap.ParaTopMargin        = pt(24)
    chap.ParaBottomMargin     = pt(18)
    chap.OutlineLevel         = 1

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

def build_content(doc, book):
    text   = doc.getText()
    cursor = text.createTextCursor()
    cursor.gotoStart(False)

    m = book.metadata

    # ── Half title (p.1) — FrontMatterPage style (set once), no header ──────
    # Page style set on the first blank, then the rest are plain blanks
    blank_with_style(text, cursor, "FrontMatterPage")
    for _ in range(5): blank(text, cursor)
    ins(text, cursor, m.title.upper(), "ChapHeads")

    # ── Full title + author (p.3) ──────────────────────────────────────────────
    ins(text, cursor, "", "Standard", page_style="FrontMatterPage")
    for _ in range(5): blank(text, cursor)
    ins(text, cursor, m.title,           "ChapHeads")
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

    # ── TOC ── (still FrontMatterPage, no header) ────────────────────────────
    ins(text, cursor, "", "Standard", page_style="FrontMatterPage")
    ins(text, cursor, "Contents", "ChapHeads")
    toc = doc.createInstance("com.sun.star.text.ContentIndex")
    toc.CreateFromOutline = True
    toc.CreateFromChapter = False
    toc.Level             = 1
    toc.Title             = ""
    text.insertTextContent(cursor, toc, False)
    text.insertControlCharacter(cursor, PARAGRAPH_BREAK, False)
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
        build_content(doc, book)
        print("  Content built.")
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
            time.sleep(3)  # give LO time to finish writing before kill
            lo_process.terminate()
            try: lo_process.wait(timeout=10)
            except: lo_process.kill()
            print("  LO shut down.")

    print(f"\n{'='*60}\nDONE\n{'='*60}")
    print(f"\n  Output: {out_path}")
    print("\n  In LibreOffice Writer:")
    print("  1. Press F9 to refresh the Table of Contents")
    print("  2. File > Export as PDF when ready to print")

if __name__ == "__main__":
    main()