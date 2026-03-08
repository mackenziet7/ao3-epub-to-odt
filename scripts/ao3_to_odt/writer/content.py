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

    # ── Half title (p.1) ──────────────────────────────────────────────────────
    blank_with_style(text, cursor, "FrontMatterPage")
    for _ in range(5): blank(text, cursor) # Adding a blank page
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

    text   = doc.getText()
    cursor = text.createTextCursor()
    cursor.gotoStart(False)

    m = book.metadata

    # ── Half title (p.1) ──────────────────────────────────────────────────────
    blank_with_style(text, cursor, "FrontMatterPage")
    for _ in range(5): blank(text, cursor)
    ins(text, cursor, m.title.upper(), "FrontMatterHead")

    # ── Full title + author (p.3) ─────────────────────────────────────────────
    ins(text, cursor, "", "Standard", page_style="FrontMatterPage")
    for _ in range(5): blank(text, cursor)
    ins(text, cursor, m.title,          "FrontMatterHead")
    ins(text, cursor, f"by {m.author}", "FrontMatter")
    blank(text, cursor); blank(text, cursor)
    if m.ao3_url:
        ins(text, cursor, m.ao3_url, "FrontMatter")

    # ── Front matter / copyright (p.4) ───────────────────────────────────────
    ins(text, cursor, "", "Standard", page_style="FrontMatterPage")
    fields = []
    if m.rating:        fields.append(f"Rating: {m.rating}")
    if m.warnings:      fields.append(f"Warnings: {', '.join(m.warnings)}")
    if m.category:      fields.append(f"Category: {m.category}")
    if m.fandom:        fields.append(f"Fandom: {', '.join(m.fandom)}")
    if m.relationships: fields.append(f"Relationships: {', '.join(m.relationships)}")
    if m.characters:    fields.append(f"Characters: {', '.join(m.characters)}")
    if m.tags:          fields.append(f"Tags: {', '.join(m.tags)}")
    if m.language:      fields.append(f"Language: {m.language}")
    if m.published:     fields.append(f"Published: {m.published}")
    if m.completed:     fields.append(f"Completed: {m.completed}")
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

    # ── TOC ───────────────────────────────────────────────────────────────────
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
            toc_objects.append(toc)
        print("  [✓] TOC")

    # ── Chapters ──────────────────────────────────────────────────────────────
    for i, ch in enumerate(book.chapters):
        if i == 0:
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