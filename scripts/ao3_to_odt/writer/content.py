from com.sun.star.text.ControlCharacter import PARAGRAPH_BREAK
import uno

from .uno_utils import inches, pt
from .styles import get_default_page_style

"""
Insert one paragraph.
page_style: if set, start a new page using this page style
page_number: if set, reset the page counter to this number
"""
def ins(text_obj, cursor, content, style, page_break=False, page_style=None, page_number=None):
    if page_style:
        cursor.setPropertyValue("PageDescName", page_style)
        if page_number is not None:
            cursor.setPropertyValue("PageNumberOffset", page_number)
    elif page_break:
        cursor.setPropertyValue("BreakType", 4)   # PAGE_BEFORE
    cursor.setPropertyValue("ParaStyleName", style)
    text_obj.insertString(cursor, content, False)
    text_obj.insertControlCharacter(cursor, PARAGRAPH_BREAK, False)
    # Clear page desc so subsequent paragraphs don't keep triggering breaks
    if page_style:
        cursor.setPropertyValue("PageDescName", "")

"""Insert a blank paragraph. Never triggers page breaks."""
def blank(text_obj, cursor):
    cursor.setPropertyValue("ParaStyleName", "Standard")
    cursor.setPropertyValue("PageDescName", "")
    text_obj.insertControlCharacter(cursor, PARAGRAPH_BREAK, False)

"""Insert a blank paragraph that also starts a new page with page_style."""
def blank_with_style(text_obj, cursor, page_style, page_number=1):
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

    # ── p.i   Half title (recto) ──────────────────────────────────────────────
    blank_with_style(text, cursor, "FrontMatterRecto", page_number=1)
    for _ in range(5): blank(text, cursor) # Adding a blank page
    ins(text, cursor, m.title.upper(), "FrontMatterHead")

    # ── p.ii  Blank verso ─────────────────────────────────────────────────────
    ins(text, cursor, "", "Standard", page_style="FrontMatterVerso")

    # ── p.iii Full title + author (recto) ─────────────────────────────────────
    ins(text, cursor, "", "Standard", page_style="FrontMatterRecto")
    for _ in range(5): blank(text, cursor)
    ins(text, cursor, m.title,           "FrontMatterHead")
    ins(text, cursor, f"by {m.author}",  "FrontMatter")
    blank(text, cursor); blank(text, cursor)
    if m.ao3_url:
        ins(text, cursor, m.ao3_url, "FrontMatter")

    # ── p.iv  Copyright / front matter (verso) ────────────────────────────────
    ins(text, cursor, "", "Standard", page_style="FrontMatterVerso")
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

    # ── p.v   TOC (recto) ─────────────────────────────────────────────────────
    if include_toc:
        ins(text, cursor, "Contents", "FrontMatterHead", page_style="FrontMatterRecto")
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
    for i, ch in enumerate(book.chapters):
        if i == 0:
            # Chapter 1 — suppress header, reset counter to 1
            ins(text, cursor, ch.title, "ChapHeads", page_style="ChapterFirstPage", page_number=1)
        else:
            # Chapters 2+ — switch to ChapterFirstPage but continue page count
            # page_number=0 tells LO "use the current count, don't reset"
            ins(text, cursor, ch.title, "ChapHeads", page_style="ChapterFirstPage")
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
        ins(text, cursor, "Appendix: Notes", "FrontMatterHead", page_style="AppendixPage")
        for ch in noted:
            ins(text, cursor, ch.title, "AppendixHead")
            for label, notes_text in [("Note:", ch.prenotes), ("End Note:", ch.endnotes)]:
                if notes_text:
                    ins(text, cursor, label, "AppendixNoteLabel")
                    for line in notes_text.split('\n'):
                        line = line.strip()
                        if line: ins(text, cursor, line, "AppendixNote")
        print(f"  [✓] Appendix ({len(noted)} chapters with notes)")