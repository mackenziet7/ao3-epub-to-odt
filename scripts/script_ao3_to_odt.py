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
import time
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

# Suppress BeautifulSoup XML-parsed-as-HTML warnings (AO3 epubs are XHTML)
from bs4 import XMLParsedAsHTMLWarning
import warnings
warnings.filterwarnings("ignore", category=XMLParsedAsHTMLWarning)

# Local
from scripts.ao3_to_odt.epub.parser import parse_epub
from scripts.ao3_to_odt.writer.connection import find_soffice, is_port_open, start_lo_listener, connect_uno
from scripts.ao3_to_odt.writer.uno_utils import prop
from scripts.ao3_to_odt.writer.styles import setup_page_style, create_para_styles
from scripts.ao3_to_odt.writer.content import build_content
from scripts.ao3_to_odt.writer.headers import setup_headers

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