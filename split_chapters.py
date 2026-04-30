"""Split the source .docx into one Markdown file per chapter under chapters/."""
import re
from pathlib import Path

import docx

SOURCE = Path("AI Slop Book 1 (5).docx")
OUT_DIR = Path("chapters")

CHAPTER_RE = re.compile(
    r"^\s*CHAPTER\s+"
    r"(ONE|TWO|THREE|FOUR|FIVE|SIX|SEVEN|EIGHT|NINE|TEN|ELEVEN|TWELVE)\s*$",
    re.IGNORECASE,
)
WORD_TO_NUM = {
    "ONE": 1, "TWO": 2, "THREE": 3, "FOUR": 4, "FIVE": 5, "SIX": 6,
    "SEVEN": 7, "EIGHT": 8, "NINE": 9, "TEN": 10, "ELEVEN": 11, "TWELVE": 12,
}


def slugify(text: str) -> str:
    text = text.lower().strip()
    text = re.sub(r"[^a-z0-9]+", "-", text)
    return text.strip("-")


def render_paragraph(para) -> str:
    """Render a paragraph to Markdown, merging adjacent same-format runs."""
    parts: list[str] = []
    cur_text = ""
    cur_italic = False
    cur_bold = False

    def flush():
        nonlocal cur_text, cur_italic, cur_bold
        if not cur_text:
            return
        s = cur_text
        if cur_bold and cur_italic:
            s = f"***{s}***"
        elif cur_bold:
            s = f"**{s}**"
        elif cur_italic:
            s = f"*{s}*"
        parts.append(s)
        cur_text = ""

    for run in para.runs:
        text = run.text
        if not text:
            continue
        italic = bool(run.italic)
        bold = bool(run.bold)
        if italic == cur_italic and bold == cur_bold:
            cur_text += text
        else:
            flush()
            cur_text = text
            cur_italic = italic
            cur_bold = bold
    flush()
    return "".join(parts)


def main() -> None:
    doc = docx.Document(SOURCE)
    paragraphs = doc.paragraphs

    # Locate chapter boundaries by matching paragraph text, regardless of style.
    boundaries: list[tuple[int, int]] = []  # (paragraph_index, chapter_number)
    for i, para in enumerate(paragraphs):
        m = CHAPTER_RE.match(para.text.strip())
        if m:
            boundaries.append((i, WORD_TO_NUM[m.group(1).upper()]))

    if not boundaries:
        raise SystemExit("No chapter markers found.")

    OUT_DIR.mkdir(exist_ok=True)
    # Wipe existing chapter files so reruns are clean.
    for old in OUT_DIR.glob("*.md"):
        old.unlink()

    index_lines: list[str] = ["# AI Slop, Book 1 — Chapters", ""]

    for idx, (start, num) in enumerate(boundaries):
        end = boundaries[idx + 1][0] if idx + 1 < len(boundaries) else len(paragraphs)

        # The chapter title is the first non-empty paragraph after the marker.
        title = ""
        body_start = start + 1
        for j in range(start + 1, end):
            t = paragraphs[j].text.strip()
            if t:
                title = t
                body_start = j + 1
                break

        slug = slugify(title) or f"chapter-{num:02d}"
        filename = f"{num:02d}-{slug}.md"
        path = OUT_DIR / filename

        lines: list[str] = [f"# Chapter {num}: {title}", ""]
        prev_blank = True  # collapse leading blanks
        for j in range(body_start, end):
            rendered = render_paragraph(paragraphs[j])
            if rendered.strip() == "":
                if not prev_blank:
                    lines.append("")
                    prev_blank = True
            else:
                lines.append(rendered)
                prev_blank = False
        # Trim trailing blank lines.
        while lines and lines[-1] == "":
            lines.pop()
        lines.append("")  # final newline

        path.write_text("\n".join(lines), encoding="utf-8")
        body_para_count = end - body_start
        print(f"Wrote {path} ({body_para_count} paragraphs)")
        index_lines.append(f"- [Chapter {num}: {title}]({filename})")

    index_lines.append("")
    (OUT_DIR / "README.md").write_text("\n".join(index_lines), encoding="utf-8")
    print(f"Wrote {OUT_DIR / 'README.md'}")


if __name__ == "__main__":
    main()
