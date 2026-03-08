import re

# Matches text-based scene breaks: *  *  *, ---, ~~~, ###, etc.
SCENE_BREAK_RE = re.compile(r'^[\*\-~=#_ ]{2,}$')

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
