# AO3 EPUB to ODT Converter

A Windows tool that converts AO3 fanfiction EPUBs into print-ready ODT documents, formatted for 5.5×8.5" book layout with running headers, a table of contents, and an appendix for author notes.

> **Note:** This tool does its best to produce clean output, but it is not perfect. Please review the generated document before printing or sharing — check for any unusual formatting, missing chapters, or oddly placed scene breaks. The ODT is fully editable in LibreOffice Writer so any issues can be fixed manually.

This project is based on the formatting guide by [Kira](https://docs.google.com/document/u/0/d/11JyVxeRS8yEWgCYrNMUPlNrEbR5AAD3Z2aDP-QXEP3Y/mobilebasic). All code in this repository is original.

---

## Requirements

- **Windows 10 or 11**
- **LibreOffice** installed at the default location (`C:\Program Files\LibreOffice`)
  - Download from [libreoffice.org](https://www.libreoffice.org/download/download-libreoffice/)

That's it for the `.exe`. If running from source, you also need Python 3.10+ with PyQt6 installed.

---

## Installation

### Option A — Executable (recommended for most users)

1. Go to the [Releases](../../releases) page
2. Download `AO3toODT.exe`
3. Place it anywhere you like — no installation needed
4. Double-click to run

### Option B — Run from source (for developers)

1. Clone the repository:
   ```
   git clone https://github.com/yourusername/ao3-epub-to-odt.git
   ```
2. Install dependencies:
   ```
   pip install PyQt6
   ```
3. Double-click `launch.bat`, or run:
   ```
   python gui.py
   ```

---

## Usage

1. Click **Browse** next to "EPUB file" and select your downloaded AO3 EPUB
2. The save location is automatically suggested based on the EPUB filename — change it with the second **Browse** button if you want
3. Click **Convert** and wait — the log window shows progress
4. When you see **✓ Done!**, your ODT file is ready

### After conversion — finishing the document for print

The tool automates the heavy lifting but a few manual steps are needed before printing. Open the ODT in LibreOffice Writer and work through these in order:

**1. Refresh the Table of Contents**
Press **F9** (or right-click the TOC and select *Update Index*) to populate it with the real chapter titles and page numbers. The TOC is inserted as a placeholder and will appear blank or incorrect until you do this.

**2. Check odd/even page balance**
Chapter 1 should start on an odd-numbered (right-hand) page — this is the standard book convention. Check the page number at the bottom of the screen when your cursor is on the first chapter page. If it's even, insert a blank page before it via *Insert > More Breaks > Manual Break* and set the page style to match.

**3. Review the document**
Scroll through and check for anything that looks off — unusual spacing, missing scene breaks, chapters running together, or oddly placed author notes in the appendix. The ODT is fully editable so fix anything you find directly.

**4. Export to PDF**
Go to **File > Export as PDF**. Make sure *Export as PDF/A* is unchecked, and under the *General* tab set the page range to all pages.

**5. Print your book with Bookbinder**
Upload your PDF to [bookbinder.app](https://bookbinder.app/) — this free tool reorders and arranges your pages into printer-ready signatures for bookbinding. Follow the instructions on the site for your chosen binding style.

The ODT format is also supported by Microsoft Word (Windows and Mac) and Google Docs, but LibreOffice Writer is recommended for the most accurate rendering of the styles this tool creates.

---

## What it produces

- **Half title and full title pages**
- **Front matter** with fic metadata (rating, tags, relationships, word count, etc.)
- **Table of contents**
- **Chapters** with proper book typography (Garamond 11.5pt, 16pt leading, first-line indents, scene breaks)
- **Running headers** — title on right pages, author on left pages
- **Appendix** collecting all author notes from chapters that have them

---

## Known limitations

- **Windows only** — the tool uses LibreOffice's UNO API via a local socket connection which is currently only configured for Windows paths
- **LibreOffice must be installed at the default path** — `C:\Program Files\LibreOffice`. Non-standard install locations will cause the conversion to fail
- **One conversion at a time** — clicking Convert while a conversion is running is not supported; wait for the current one to finish
- **AO3 EPUBs only** — the parser is written specifically for AO3's EPUB structure and will likely not work correctly on EPUBs from other sources
- **The TOC requires a manual F9 refresh** in LibreOffice Writer to populate page numbers after opening — this cannot be automated
- **Running headers appear on chapter title pages** — proper book typography suppresses the header on the first page of each chapter; this is currently not implemented and must be removed manually if desired
- **Odd/even page balance is not checked** — the tool does not verify that chapters start on odd-numbered pages; this must be checked and corrected manually after conversion (see After Conversion steps above)

---

## Troubleshooting

**Conversion gets stuck at "Connecting to LO UNO"**
A leftover LibreOffice process from a previous run may be blocking the port. Open Task Manager, find any `soffice.exe` processes, and end them, then try again.

**Unicode or encoding errors in the log**
Make sure you are using the latest version of both `gui.py` and `ao3_to_odt.py` — earlier versions had a UTF-8 encoding issue on Windows that has since been fixed.

**The ODT looks wrong or chapters are missing**
Some AO3 EPUBs exported via Calibre have a slightly different internal structure. Please open an issue and attach the EPUB (or a link to the fic) so it can be investigated.

---

## Contributing

Bug reports and pull requests are welcome! If an EPUB isn't converting correctly, please open an issue and include:
- A link to the AO3 fic (or attach the EPUB if the fic is locked/deleted)
- The full log output from the conversion window — copy and paste everything that appeared in the log, including any error messages

---

## Future plans

- **Customisable formatting in the GUI** — choose page size, margins, font, font size, and line spacing without editing any code
- **Suppress headers on chapter title pages** — proper book typography has no running header on the first page of each chapter; this should be automatic
- **Automatic TOC refresh** — trigger the F9 refresh programmatically so the TOC is populated without any manual steps
- **Odd/even page balancing** — automatically insert blank pages where needed so chapter 1 starts on an odd-numbered (right-hand) page
- **Mac and Linux support**
- **Support for non-default LibreOffice install locations**
- **Batch conversion** — convert multiple EPUBs at once
- **Fic collections** — combine multiple EPUBs into a single bound book, with each fic as its own section and a shared table of contents

---

## Credits

Formatting approach based on the print binding guide by [Kira](https://docs.google.com/document/u/0/d/11JyVxeRS8yEWgCYrNMUPlNrEbR5AAD3Z2aDP-QXEP3Y/mobilebasic).

---

## License

MIT