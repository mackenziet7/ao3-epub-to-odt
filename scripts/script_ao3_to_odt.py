"""
script_ao3_to_odt.py — pre-parsed book data → print-ready ODT
Run with LibreOffice's Python (needs only uno + the standard library):

    python.exe script_ao3_to_odt.py --book book.json --preset preset.json --out out.odt

The GUI parses the EPUB and writes book.json. This script never touches the EPUB.
"""
import argparse
import json
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
except ImportError:
    print("\nERROR: 'uno' module not found.")
    print("You must run this script with LibreOffice's Python, not your system Python.")
    sys.exit(1)

# Local
from scripts.ao3_to_odt.epub.models import book_from_dict
from scripts.ao3_to_odt.writer.connection import find_soffice, is_port_open, start_lo_listener, connect_uno
from scripts.ao3_to_odt.writer.uno_utils import prop
from scripts.ao3_to_odt.writer.styles import setup_page_style, create_para_styles
from scripts.ao3_to_odt.writer.content import build_content
from scripts.ao3_to_odt.writer.headers import setup_headers
from scripts.ao3_to_odt.preset_schema import validate_preset

# ═══════════════════════════════════════════════════════════════════════════════
# Helpers
# ═══════════════════════════════════════════════════════════════════════════════

def progress(pct, msg):
    print(f"@@PROGRESS|{pct}|{msg}", flush=True)

def on_chapter(done, total):
    progress(35 + int(50 * done / total), f"Chapter {done}/{total}")

def load_preset(path):
    with open(path, encoding="utf-8") as f:
        preset = json.load(f)
    errors = validate_preset(preset)
    if errors:
        raise ValueError("Invalid preset: " + "; ".join(errors))
    return preset

def save_odt(doc, out_path):
    url = uno.systemPathToFileUrl(os.path.abspath(out_path))
    doc.storeToURL(url, [prop("FilterName", "writer8"), prop("Overwrite", True)])
    print(f"  [✓] Saved: {out_path}")

# ═══════════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════════
def convert_epub(book, qr_path, out_path, preset, port=2002):
    opts = preset["additional_options"]
    include_toc = opts["include_table_of_contents"]
    include_qr = opts["include_qr_code"]

    # ── Book summary ───────────────────────── 
    print(f"\n{'='*60}\nBook\n{'='*60}")
    print(f"  Title:    {book.metadata.title}")
    print(f"  Author:   {book.metadata.author}")
    print(f"  Chapters: {len(book.chapters)}")
    print(f"  Words:    {book.metadata.words}")

    # ── Start LO listener ─────────────────────────────────────────────────────
    print(f"\n{'='*60}\nStarting LibreOffice\n{'='*60}")
    progress(10, "Starting LibreOffice")
    lo_process = None
    if is_port_open(port):
        print("  Already running, connecting...")
        progress(30, "LibreOffice ready")
    else:
        soffice = find_soffice()
        if not soffice:
            print("ERROR: Cannot find soffice executable.")
            sys.exit(1)
        print(f"  Launching: {soffice}")
        lo_process = start_lo_listener(soffice, port)
        print("  Waiting for LO to start", end="", flush=True)
        progress(25, "Waiting for LibreOffice")
        for _ in range(40):
            if is_port_open(port):
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
        progress(30, "LibreOffice ready")

    # ── Build document ────────────────────────────────────────────────────────
    print(f"\n{'='*60}\nBuilding document\n{'='*60}")
    try:
        print("  Connecting...")
        desktop = connect_uno(port)
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
        progress(32, "Creating document")
        setup_page_style(doc, preset["page_setup"])
        print("  Page style done.")
        create_para_styles(doc, preset["typography_advanced"])
        print("  Para styles done.")
        progress(35, "Building chapters")

        toc_objects = []
        build_content(doc, book, include_toc, toc_objects, include_qr,
                      qr_path=qr_path, on_chapter=on_chapter)
        print("  Content built.")
        progress(86, "Updating contents")
        if toc_objects:
            try:
                toc_objects[0].update()
                print("  [✓] TOC refreshed")
            except Exception as e:
                print(f"  TOC refresh failed (open in LO and press F9): {e}")
        progress(91, "Adding running headers")
        setup_headers(doc, book.metadata, preset["typography_advanced"]["main_book"]["body"]["font"])
        print("  Headers done.")
        progress(94, "Saving file")
        save_odt(doc, out_path)
        progress(98, "Finishing")
        time.sleep(2)
        print("  Closing document...")
        import threading
        close_done = threading.Event()

        def close_doc():
            try:
                doc.close(True)
            except Exception:
                pass
            finally:
                close_done.set()

        t = threading.Thread(target=close_doc, daemon=True)
        t.start()
        if close_done.wait(timeout=10):
            print("  Document closed.")
        else:
            print("  Document close timed out, continuing anyway.")

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

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--book", required=True)
    ap.add_argument("--preset", required=True)
    ap.add_argument("--out", required=True)
    args = ap.parse_args()

    with open(args.book, encoding="utf-8") as f:
        data = json.load(f)
    qr_path = data.pop("qr_png", None)
    book = book_from_dict(data)
    preset = load_preset(args.preset)

    convert_epub(book, qr_path, args.out, preset)

if __name__ == "__main__":
    main()