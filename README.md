# AO3 EPUB to ODT Converter

A Windows tool that converts AO3 fanfiction EPUBs into print-ready ODT documents, formatted with running headers, a table of contents, and an appendix for author notes.

> **Note:** This tool does its best to produce clean output, but it is not perfect. Please review the generated document before printing or sharing — check for any unusual formatting, missing chapters, or oddly placed scene breaks. The ODT is fully editable in LibreOffice Writer so any issues can be fixed manually.

This project is based on the formatting guide by [Kira](https://docs.google.com/document/u/0/d/11JyVxeRS8yEWgCYrNMUPlNrEbR5AAD3Z2aDP-QXEP3Y/mobilebasic). All code in this repository is original.


<img width="900" height="675" alt="ao3toODT" src="https://github.com/user-attachments/assets/7110a2b1-7ce0-4252-819f-7b9e821b83cb" />

---

## Requirements

- **Windows 10 or 11**
- **LibreOffice** installed
  - Download from [libreoffice.org](https://www.libreoffice.org/download/download-libreoffice/)
  - This program works best with the newest version of LibreOffice

That's it for the `.exe`. If running from source, you also need Python 3.10+ with PySide6 installed.

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
   git clone https://github.com/mackenziet7/ao3-epub-to-odt.git
   ```
2. Install dependencies:
   ```
   pip install PySide6
   ```
3. Run:
   ```
   python main.py
   ```


## Uninstalling

The app itself leaves no registry entries and needs no formal uninstaller — just delete the `.exe` (or source folder) when you're done.

However, one additional location is written to during normal use and should be removed manually if you want a clean uninstall:

### App data (config and settings)

The app stores your settings, theme choice, and saved presets at:

```
%APPDATA%\AO3toODT\
```

To delete it, paste that path into File Explorer or a Run dialog (`Win+R`) and delete the folder.

---

## Usage

The app has two ways to convert: the **step-by-step** wizard for full control, and **quick convert** for one-click conversions with a saved preset.

#### Using the wizard
1. Input — click Browse and select your downloaded AO3 EPUB. The save location is automatically suggested based on the EPUB filename (or your default save location, if you've set one in Settings)
2. Layout — choose page size, margins, and other page layout options
3. Typography — set your font, font size, and line spacing on the Basic screen. For per-style control, use the Advanced Typography page. If you skip Advanced, those values are derived from your basic font and size
4. Options — toggle optional extras such as the QR code to AO3 and the table of contents
5. Convert — start the conversion and watch the progress on the results page. When you see ✓ Done!, your ODT file is ready

#### Presets and quick convert
Once you have settings you like, save them as a named preset from the wizard. Quick convert lets you pick an EPUB and a preset and convert in one go, without walking through the wizard again. A default book-layout preset is included.

### After conversion — finishing the document for print

The tool automates the heavy lifting but a few manual steps are needed before printing. Open the ODT in LibreOffice Writer and work through these in order:

**1. Review the document**
Scroll through and check for anything that looks off — unusual spacing, missing scene breaks, chapters running together, or oddly placed author notes in the appendix. The ODT is fully editable so fix anything you find directly.

**2. Export to PDF**
Go to **File > Export as PDF**. Make sure *Export as PDF/A* is unchecked, and under the *General* tab set the page range to all pages.

**3. Print your book with Bookbinder**
Upload your PDF to [bookbinder.app](https://bookbinder.app/) — this free tool reorders and arranges your pages into printer-ready signatures for bookbinding. Follow the instructions on the site for your chosen binding style.

The ODT format is also supported by Microsoft Word (Windows and Mac) and Google Docs, but LibreOffice Writer is recommended for the most accurate rendering of the styles this tool creates.

---

## What it produces

- **Half title and full title pages** with proper blank verso pages between them
- **Front matter** with fic metadata (rating, tags, relationships, word count, etc.)
- **Optional QR code to AO3** takes you right to the fanfic on AO3
- **Optional Table of contents** starting on a recto (right-hand) page
- **Chapters** with proper book typography based on chosen settings
- **Suppressed headers** on chapter opening pages, with running headers (title on right pages, author on left pages) on all other body pages
- **Appendix** collecting all author notes from chapters that have them, with suppressed headers throughout

<img width="1340" height="588" alt="Screenshot 2026-10-12 153323" src="https://github.com/user-attachments/assets/96aa1f72-86e2-46d9-87ca-214b03b3412c" />


---

## Known limitations

- **Windows only** — the tool uses LibreOffice's UNO API via a local socket connection which is currently only configured for Windows paths
- **AO3 EPUBs only** — the parser is written specifically for AO3's EPUB structure and will likely not work correctly on EPUBs from other sources

---

## Troubleshooting

**Conversion gets stuck at "Connecting to LO UNO"**
A leftover LibreOffice process from a previous run may be blocking the port. Open Task Manager, find any `soffice.exe` processes, and end them, then try again.

**The ODT looks wrong or chapters are missing**
Some AO3 EPUBs exported via Calibre have a slightly different internal structure. Please open an issue and attach the EPUB (or a link to the fic) so it can be investigated.

---

## Contributing

Bug reports and pull requests are welcome! If an EPUB isn't converting correctly, please open an issue and include:
- A link to the AO3 fic (or attach the EPUB if the fic is locked/deleted)
- The full log output from the conversion window — copy and paste everything that appeared in the log, including any error messages
- The preset you used. To get this go to the config file location (available through settings page) and find the correct .json file to add
---

## Future plans

- **Fic collections** — combine multiple EPUBs into a single bound book, with each fic as its own section and a shared table of contents
- **Scene break handling** — detect scene breaks and allow users to insert their own image or character string to customize the typeset

---

## Credits

Formatting approach based on the print binding guide by [Kira](https://docs.google.com/document/u/0/d/11JyVxeRS8yEWgCYrNMUPlNrEbR5AAD3Z2aDP-QXEP3Y/mobilebasic).

---

## License

MIT
