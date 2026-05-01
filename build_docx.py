#!/usr/bin/env python3
"""Build book.docx from book.md.

Handles:
- `# Title` -> Heading 1 (top-level book title)
- `# Chapter N: Foo` -> Heading 1 (chapter)
- Inline `*italic*` runs -> italic
- Blank lines between paragraphs in Markdown become paragraph boundaries
- Smart-quote lines (chapter 9) and straight-quote lines both pass through
"""
import re
import sys
from docx import Document
from docx.shared import Pt


ITALIC_RE = re.compile(r"\*([^*\n]+)\*")


def add_runs(paragraph, text: str) -> None:
    """Split text on `*...*` italic markers and add runs to the paragraph."""
    pos = 0
    for m in ITALIC_RE.finditer(text):
        if m.start() > pos:
            paragraph.add_run(text[pos:m.start()])
        run = paragraph.add_run(m.group(1))
        run.italic = True
        pos = m.end()
    if pos < len(text):
        paragraph.add_run(text[pos:])


def build(md_path: str, out_path: str) -> None:
    with open(md_path, encoding="utf-8") as fh:
        lines = fh.read().splitlines()

    doc = Document()

    # Set default body style to a readable serif size.
    style = doc.styles["Normal"]
    style.font.name = "Garamond"
    style.font.size = Pt(12)

    saw_book_title = False
    for raw in lines:
        line = raw.rstrip()
        if not line:
            continue
        if line.startswith("# "):
            heading_text = line[2:].strip()
            if not saw_book_title:
                # First H1 is the volume title.
                p = doc.add_heading(heading_text, level=0)
                saw_book_title = True
            else:
                # Force a page break before each chapter so the DOCX paginates cleanly.
                doc.add_page_break()
                p = doc.add_heading(heading_text, level=1)
            continue
        p = doc.add_paragraph()
        add_runs(p, line)

    doc.save(out_path)


if __name__ == "__main__":
    md = sys.argv[1] if len(sys.argv) > 1 else "book.md"
    out = sys.argv[2] if len(sys.argv) > 2 else "book.docx"
    build(md, out)
    print(f"Wrote {out}")
