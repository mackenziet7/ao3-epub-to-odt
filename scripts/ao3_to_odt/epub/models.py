from dataclasses import asdict, dataclass, field
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

# String formatting flags
@dataclass
class Run:
    text: str = ""
    italic: bool = False
    bold: bool = False

@dataclass
class Paragraph:
    type: str = "para" # Paragraphy type
    runs: list = field(default_factory=list) # List of Run

    @property
    def text(self):
        return "".join(r.text for r in self.runs)

@dataclass
class Chapter:
    index: int = 0
    title: str = ""
    prenotes: str = ""
    body: list = field(default_factory=list) # List of Paragraph
    endnotes: str = ""

@dataclass
class AO3Book:
    metadata: AO3Metadata = field(default_factory=AO3Metadata)
    chapters: list = field(default_factory=list)
    cover_image: Optional[bytes] = None
    cover_image_ext: str = "jpg"

def book_to_dict(book):
    d = asdict(book)
    d.pop("cover_image", None)
    return d

def book_from_dict(d):
    chapters = []
    for c in d["chapters"]:
        body = [
            Paragraph(type=p["type"], runs=[Run(**r) for r in p["runs"]])
            for p in c["body"]
        ]
        chapters.append(Chapter(
            index=c["index"], title=c["title"],
            prenotes=c["prenotes"], endnotes=c["endnotes"], body=body,
        ))
    return AO3Book(
        metadata=AO3Metadata(**d["metadata"]),
        chapters=chapters,
        cover_image_ext=d.get("cover_image_ext", "jpg"),
    )