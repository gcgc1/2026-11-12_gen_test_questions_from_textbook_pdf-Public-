"""Regenerate converted markdown only when a source PDF is newer than its outputs.

For every PDF in `10_textbook_pdf`, this checks whether the PDF is newer than
its markdown in `20_converted_to_md`, or newer than its split files in
`30_splited_by_chapter` (or whether those outputs are simply missing). Only the
stale PDFs are re-extracted and re-split; up-to-date ones are left untouched.

It reuses `extract_text_md.py` and `split_chapters_md.py`, so the output format
is identical (each page is marked with both the PDF and the book page number).

Usage:
    python update_converted.py
"""

from __future__ import annotations

from pathlib import Path

import extract_text_md as extract
import split_chapters_md as split

PDF_DIR = extract.SRC_DIR
MD_DIR = extract.DST_DIR
SPLIT_DIR = split.DST_DIR


def newest_mtime(paths) -> float:
    """Newest modification time among `paths`, or 0.0 when there are none."""
    return max((p.stat().st_mtime for p in paths), default=0.0)


def main() -> None:
    pdfs = sorted(PDF_DIR.glob("*.pdf"))
    if not pdfs:
        print(f"No PDF files found in {PDF_DIR}/")
        return

    for pdf in pdfs:
        md = MD_DIR / f"{pdf.stem}.md"
        splits = list(SPLIT_DIR.glob(f"{pdf.stem}_*.md"))
        pdf_mtime = pdf.stat().st_mtime
        md_mtime = md.stat().st_mtime if md.exists() else 0.0
        split_mtime = newest_mtime(splits)

        extract_needed = (md_mtime == 0.0) or (pdf_mtime > md_mtime)
        split_needed = (
            extract_needed
            or not splits
            or pdf_mtime > split_mtime
            or md_mtime > split_mtime
        )

        if not (extract_needed or split_needed):
            print(f"{pdf.name}: up to date, skipped")
            continue

        if extract_needed:
            print(f"{pdf.name} -> {extract.convert(pdf)}")
        for path in split.process(md):
            print(f"{md.name} -> {path.name}")


if __name__ == "__main__":
    main()
