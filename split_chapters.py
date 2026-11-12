"""Split each extracted textbook markdown into one file per chapter (unit).

Reads every `<stem>.md` produced by `extract_text.py` (i.e. every markdown in
`textbook_z_1_converted_to_pdf` that still contains `## Page N (book page M)`
markers) and writes one file per unit into `textbook_z_2_converted_to_pdf_splited`,
so each chapter can be processed individually.

A unit starts on the first page that
  * is not front matter (front matter pages carry a roman-numeral page number
    or a "Contents" / "How to use this book" / "Science Skills" header), and
  * contains a "In this topic you will" box, and
  * contains the unit's first topic header, e.g. `1.1 ...`.

Output files (front matter and any unassigned pages are kept too):
    <stem>_front_matter.md
    <stem>_unit01_respiration.md
    <stem>_unit02_properties_of_materials.md

Usage:
    python split_chapters.py
"""

from __future__ import annotations

import re
from pathlib import Path

SRC_DIR = Path("textbook_z_1_converted_to_pdf")
DST_DIR = Path("textbook_z_2_converted_to_pdf_splited")

PAGE_RE = re.compile(r"(?m)^## Page (\d+)(?: \(book page ([^)]+)\))?\s*$")
TOPIC_1_RE = re.compile(r"^\s*(\d{1,2})\.1\s+\S")          # first topic of a unit
UNIT_HEADER_RE = re.compile(r"^\s*(\d{1,2})\s{2,}([A-Za-z].*)$")  # running header
ROMAN_RE = re.compile(r"^[ivx]{1,5}$", re.IGNORECASE)
FRONT_MATTER_HEADERS = ("contents", "how to use this book", "science skills")
TOPIC_BOX = "in this topic you will"
TITLE_STOPWORDS = {"of", "and", "the", "in", "to", "for", "with", "a", "an", "on"}


def parse_pages(md_text: str) -> tuple[str, list[tuple[int, str | None, str]]]:
    """Return the document header and a list of (pdf page, book page, text)."""
    parts = PAGE_RE.split(md_text)
    pages = [
        (int(parts[i]), parts[i + 1] or None, parts[i + 2].strip("\n"))
        for i in range(1, len(parts), 3)
    ]
    return parts[0].strip(), pages


def is_front_matter(lines: list[str]) -> bool:
    """Front matter pages are numbered with roman numerals or have a known header."""
    head = [line.strip() for line in lines if line.strip()][:3]
    for line in head:
        if ROMAN_RE.match(line):
            return True
        if any(marker in line.lower() for marker in FRONT_MATTER_HEADERS):
            return True
    return False


def unit_title(lines: list[str]) -> str:
    """Best-effort unit title from the running header on the unit's first page."""
    for index, line in enumerate(lines):
        match = UNIT_HEADER_RE.match(line)
        if not match:
            continue
        title = match.group(2).strip()
        # The header can wrap; keep appending while it ends on a connector word.
        following = [l.strip() for l in lines[index + 1 :] if l.strip()]
        while title.split()[-1].lower() in TITLE_STOPWORDS and following:
            title += " " + following.pop(0)
        return title
    return ""


def slugify(text: str) -> str:
    return re.sub(r"_+", "_", re.sub(r"[^a-z0-9]+", "_", text.lower())).strip("_")


def find_units(pages: list[tuple[int, str | None, str]]) -> list[dict]:
    """Locate unit start pages and their titles."""
    units = []
    for number, _book, text in pages:
        lines = text.splitlines()
        if is_front_matter(lines) or TOPIC_BOX not in text.lower():
            continue
        match = next((m for line in lines if (m := TOPIC_1_RE.match(line))), None)
        if match:
            units.append({"num": int(match.group(1)), "page": number, "title": unit_title(lines)})
    return units


def split_document(md_text: str) -> list[dict]:
    """Turn one markdown document into labelled chapter segments."""
    _, pages = parse_pages(md_text)
    if not pages:
        return []

    units = find_units(pages)
    segments = []

    def segment(label: str, start: int, stop: int) -> dict:
        chosen = [p for p in pages if start <= p[0] < stop]
        body = "\n\n".join(
            f"## Page {pdf}" + (f" (book page {book})" if book else "") + f"\n\n{text}"
            for pdf, book, text in chosen
        )
        books = [b for _, b, _ in chosen if b]
        return {
            "label": label,
            "first": chosen[0][0],
            "last": chosen[-1][0],
            "first_book": books[0] if books else None,
            "last_book": books[-1] if books else None,
            "body": body,
        }

    first_unit_page = units[0]["page"] if units else pages[-1][0] + 1
    if first_unit_page > pages[0][0]:
        segments.append(segment("front_matter", pages[0][0], first_unit_page))

    for i, unit in enumerate(units):
        stop = units[i + 1]["page"] if i + 1 < len(units) else pages[-1][0] + 1
        segments.append(segment(f"unit{unit['num']:02d}", unit["page"], stop))
        segments[-1]["title"] = unit["title"]

    return segments


def process(md_path: Path) -> list[Path]:
    """Split one document into per-unit markdown files; return the paths written."""
    stem = md_path.stem
    segments = split_document(md_path.read_text(encoding="utf-8"))
    if not segments:
        return []

    # Remove stale outputs from a previous run before writing.
    for old in DST_DIR.glob(f"{stem}_unit*.md"):
        old.unlink()

    written = []
    for seg in segments:
        title = seg.get("title") or seg["label"].replace("_", " ").title()
        if seg["label"] == "front_matter":
            suffix = "front_matter"
        else:
            suffix = f"{seg['label']}_{slugify(seg['title'])}"
        out_path = DST_DIR / f"{stem}_{suffix}.md"
        heading = f"# {stem} — {title}"
        pages_txt = f"PDF pages {seg['first']}–{seg['last']}"
        if seg["first_book"]:
            pages_txt += f" · book pages {seg['first_book']}–{seg['last_book']}"
        source = f"> Source: `{md_path.name}` · {pages_txt}"
        out_path.write_text(f"{heading}\n\n{source}\n\n---\n\n{seg['body']}\n", encoding="utf-8")
        written.append(out_path)
    return written


def main() -> None:
    DST_DIR.mkdir(parents=True, exist_ok=True)
    inputs = [
        md
        for md in sorted(SRC_DIR.glob("*.md"))
        if not re.search(r"_unit\d+_", md.name) and not md.name.endswith("_front_matter.md")
    ]

    if not inputs:
        print(f"No source markdown found in {SRC_DIR}/ — run extract_text.py first.")
        return

    for md_path in inputs:
        written = process(md_path)
        if not written:
            print(f"{md_path.name}: no units found, skipped")
            continue
        for path in written:
            print(f"{md_path.name} -> {path.name}")


if __name__ == "__main__":
    main()
