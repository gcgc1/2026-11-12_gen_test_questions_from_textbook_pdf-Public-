"""Extract text from every PDF in `textbook_pdf`, chapter by chapter.

Each PDF is treated as one chapter/part. The text is written to a markdown file
with the same name in `textbook_z_1_converted_to_pdf`, with page markers that
carry both the PDF page number and the book page number, so the content can be
processed individually downstream.

The book page number is the folio printed at the top of each page. Pages whose
folio did not survive text extraction are filled in from the most common
PDF-to-book page offset; front matter keeps its roman numerals.

Usage:
    python extract_text.py
"""

from __future__ import annotations

import logging
import re
from collections import Counter
from pathlib import Path

from pypdf import PdfReader

SRC_DIR = Path("textbook_pdf")
DST_DIR = Path("textbook_z_1_converted_to_pdf")

# pypdf is noisy about missing fontTools; keep the console output clean.
logging.getLogger("pypdf").setLevel(logging.ERROR)


# pypdf sometimes fails to decode ligatures; repair the common cases.
FIXES = {
    "\ufb00": "ff", "\ufb01": "fi", "\ufb02": "fl",
    "/f_i": "fi", "/f_l": "fl", "/f_f": "ff",
    "\u00a0": " ",
}

ARABIC_RE = re.compile(r"^\d{1,3}$")            # printed book page number
ROMAN_RE = re.compile(r"^[ivxlcdm]{1,7}$", re.IGNORECASE)  # front matter folio
HEAD_LINES = 12                                 # running headers/folios live here


def clean_page_text(text: str) -> str:
    """Tidy a raw page dump without altering the wording."""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    for bad, good in FIXES.items():
        text = text.replace(bad, good)
    text = re.sub(r"[ \t]+\n", "\n", text)      # trailing whitespace
    text = re.sub(r"\n{3,}", "\n\n", text)      # collapse blank runs
    return text.strip()


def head_lines(text: str, count: int = HEAD_LINES) -> list[str]:
    """First non-empty lines of a page, where running headers/folios live."""
    return [line.strip() for line in text.splitlines() if line.strip()][:count]


def arabic_folio(lines: list[str]) -> int | None:
    """The printed book page number, if an Arabic folio is visible."""
    for line in lines:
        if ARABIC_RE.match(line) and int(line) >= 1:
            return int(line)
    return None


def roman_folio(lines: list[str]) -> str | None:
    """The roman front-matter folio, if visible."""
    for line in lines[:5]:
        if ROMAN_RE.match(line) and line.lower() != "i":  # skip stray pronoun "I"
            return line.lower()
    return None


def book_pages(pages: list[tuple[int, str]]) -> dict[int, int | str | None]:
    """Map each PDF page number to its book page number.

    Content pages take the folio printed on them; the most common PDF-to-book
    offset fills in pages where the folio is missing. Pages before the first
    content page are treated as front matter and keep their roman numerals.
    """
    folios = {n: arabic_folio(head_lines(text)) for n, text in pages}

    offsets = Counter(n - f for n, f in folios.items() if f is not None)
    offset = offsets.most_common(1)[0][0] if offsets else None

    first_content = next(
        (n for n, f in folios.items() if f is not None and n - f == offset), None
    )

    mapping: dict[int, int | str | None] = {}
    for n, text in pages:
        if first_content is not None and n >= first_content:
            mapping[n] = folios[n] if folios[n] is not None else n - offset
        else:
            mapping[n] = roman_folio(head_lines(text))
    return mapping


def page_heading(number: int, book: int | str | None) -> str:
    """Markdown heading carrying both page numbers."""
    if book is None:
        return f"## Page {number}"
    return f"## Page {number} (book page {book})"


def convert(pdf_path: Path) -> Path:
    """Extract one PDF into a markdown file and return its path."""
    reader = PdfReader(str(pdf_path))
    pages = [
        (number, clean_page_text(page.extract_text() or ""))
        for number, page in enumerate(reader.pages, start=1)
    ]
    book = book_pages(pages)

    lines = [f"# {pdf_path.stem}", ""]
    for number, text in pages:
        lines += [page_heading(number, book[number]), "", text, ""]

    out_path = DST_DIR / f"{pdf_path.stem}.md"
    out_path.write_text("\n".join(lines), encoding="utf-8")
    return out_path


def main() -> None:
    DST_DIR.mkdir(parents=True, exist_ok=True)
    pdfs = sorted(SRC_DIR.glob("*.pdf"))

    if not pdfs:
        print(f"No PDF files found in {SRC_DIR}/")
        return

    for pdf_path in pdfs:
        out_path = convert(pdf_path)
        print(f"{pdf_path.name} -> {out_path}")


if __name__ == "__main__":
    main()
