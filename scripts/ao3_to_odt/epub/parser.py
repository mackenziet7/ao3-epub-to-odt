import re
import warnings
import ebooklib
from ebooklib import epub
from bs4 import BeautifulSoup

from .models import AO3Metadata, Run, Paragraph, Chapter, AO3Book
from .text_utils import clean_text, clean_run, SCENE_BREAK_RE

NOTE_LABELS = ('chapter notes', 'chapter summary',
               'author note', "author's note", 'notes:')

# Individual paragraphs and character level
def parse_body_html(soup):
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
        if tag.name == 'hr':
            return True
        if tag.name == 'p':
            classes = ' '.join(tag.get('class', []))
            if re.search(r'sep|break|divid|rule', classes, re.I):
                return True
            text = tag.get_text(strip=True)
            if text and SCENE_BREAK_RE.match(text):
                return True
            children = [c for c in tag.children
                        if hasattr(c, 'name') or str(c).strip()]
            if len(children) == 1 and hasattr(children[0], 'name') and children[0].name == 'img':
                return True
        return False

    paragraphs = []
    for tag in soup.find_all(['p', 'hr']):
        if is_scene_break_tag(tag):
            if not paragraphs or paragraphs[-1].type != 'break':
                paragraphs.append(Paragraph(type='break'))
        elif tag.name == 'p':
            runs = extract_runs(tag)
            if not runs: continue
            full = "".join(r.text for r in runs).strip()
            if not full: continue
            paragraphs.append(Paragraph(type='para', runs=runs))
    return paragraphs

def _bq_text(tag):
    """Get clean text from a blockquote or its first blockquote child."""
    bq = tag.find('blockquote') if tag.name != 'blockquote' else tag
    if bq:
        return clean_text(bq.get_text(separator='\n'))
    return clean_text(tag.get_text(separator='\n'))

# For meta data pages
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

# Extracts a chapter's title, notes, and body text
def parse_chapter(item, index):
    ch = Chapter(index=index)
    soup = BeautifulSoup(item.get_content(), 'lxml')

    t = (soup.find(class_=re.compile(r'toc-heading|heading', re.I)) or
         soup.find('h2') or soup.find('h3'))
    if t: ch.title = clean_text(t.get_text())

    end_div = soup.find(id=re.compile(r'endnotes', re.I))
    if end_div:
        ch.endnotes = _bq_text(end_div)
        end_div.decompose()

    for d in soup.find_all(class_='endnote-link'):
        d.decompose()

    nd = soup.find(id='notes') or soup.find(class_=re.compile(r'^notes$', re.I))
    if nd:
        ch.prenotes = _bq_text(nd)
        nd.decompose()
    else:
        note_parts = []
        for p in soup.find_all('p'):
            label = p.get_text(strip=True).lower()
            if any(x in label for x in NOTE_LABELS):
                bq = p.find_next_sibling('blockquote')
                if not bq:
                    bq = p.parent.find_next_sibling('blockquote') if p.parent else None
                if not bq:
                    for sib in p.next_elements:
                        if hasattr(sib, 'name') and sib.name == 'blockquote':
                            bq = sib
                            break
                        if hasattr(sib, 'name') and sib.name == 'p':
                            break
                if bq:
                    note_parts.append(clean_text(bq.get_text(separator='\n')))
                    bq.decompose()
                p.decompose()
        if note_parts:
            ch.prenotes = '\n\n'.join(note_parts)

    body = (soup.find(id='chapters') or
            soup.find(class_='userstuff2') or
            soup.find(class_='userstuff1') or
            soup.find('body'))
    if body:
        ch.body = parse_body_html(body)

    return ch

def parse_epub(epub_path):
    book = AO3Book()
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        ebook = epub.read_epub(epub_path)

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

    for item in ebook.get_items_of_type(ebooklib.ITEM_IMAGE):
        if 'cover' in item.get_name().lower():
            book.cover_image = item.get_content()
            book.cover_image_ext = item.media_type.split('/')[-1].replace('jpeg', 'jpg')
            break

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
            if ch.title.strip().lower() in ("afterword", "foreword", "preface"):
                continue
            book.chapters.append(ch)

    return book