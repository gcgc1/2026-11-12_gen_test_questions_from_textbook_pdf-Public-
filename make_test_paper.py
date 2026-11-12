"""Generate a Year 9 cloze test for Chapter 1 (Respiration).

Five cloze questions are produced for each section of the chapter, based on the
content of the split markdown in `textbook_z_2_converted_to_pdf_splited`. The
paper is written as an A4 .docx (font size 15) with:

  * a shared word bank at the top,
  * the questions grouped by section,
  * a separate answer-key page at the end.

Words in the word bank may be used once, more than once, or not at all.

Usage:
    python make_test_paper.py
"""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt

DST_DIR = Path("z_test_paper")
FILE_NAME = "1 - Respiration.docx"
FONT_SIZE = Pt(15)
BLANK = "__________"
BLANK_WIDTH = Mm(70)              # width of each cloze fill-in line
BLANK_FONT = "Courier New"        # monospaced: exactly 0.6 em per character
BLANK_CHARS = round(BLANK_WIDTH.pt / (0.6 * FONT_SIZE.pt))
# Words offered to students. Some are distractors and are never used.
WORD_BANK = [
    "air sacs", "alveoli", "antibodies", "blood plasma", "bronchi",
    "bronchioles", "capillaries", "cartilage", "cell", "contract",
    "decreases", "diaphragm", "diffusion", "energy", "expired",
    "gas exchange", "glucose", "haemoglobin", "heat", "increases",
    "inspired", "intercostal", "limewater", "mitochondria", "nucleus",
    "oxyhaemoglobin", "oxygen", "pathogens", "phagocytosis", "pressure",
    "relax", "respiration", "trachea", "vocal cords", "water",
]

# Each section: title, the five cloze sentences and the answers (in blank order).
SECTIONS = [
    {
        "title": "1.1  The human respiratory system",
        "questions": [
            ("Respiration is a series of chemical reactions that takes place "
             "inside every living __________ and releases energy from food.",
             ["cell"]),
            ("Air passes from the nose or mouth, down the __________ (also "
             "called the windpipe) and into the lungs.",
             ["trachea"]),
            ("Rings of __________ around the trachea keep it open and stop it "
             "from collapsing.",
             ["cartilage"]),
            ("The trachea divides into two __________, one to each lung, and "
             "each of these divides into many smaller tubes called __________.",
             ["bronchi", "bronchioles"]),
            ("At the end of the bronchioles are tiny __________ where gas "
             "exchange happens; sound is made when air makes the __________ "
             "in your larynx vibrate.",
             ["air sacs", "vocal cords"]),
        ],
    },
    {
        "title": "1.2  Gas exchange",
        "questions": [
            ("The air sacs in the lungs are also called __________, and the "
             "very thin blood vessels wrapped around them are __________.",
             ["alveoli", "capillaries"]),
            ("Oxygen moves from the air into the blood by __________, from a "
             "place where there is a lot of it to a place where there is less.",
             ["diffusion"]),
            ("Once oxygen is inside a red blood cell it combines with "
             "__________ to form oxyhaemoglobin.",
             ["haemoglobin"]),
            ("The movement of oxygen into the blood and carbon dioxide out of "
             "the blood in the lungs is called __________.",
             ["gas exchange"]),
            ("The air you breathe in is called __________ air and the air you "
             "breathe out is called __________ air; __________ is the chemical "
             "used to test for carbon dioxide.",
             ["inspired", "expired", "limewater"]),
        ],
    },
    {
        "title": "1.3  Breathing",
        "questions": [
            ("The muscles between the ribs are called the __________ muscles.",
             ["intercostal"]),
            ("When you breathe in, the intercostal muscles and the diaphragm "
             "__________, which makes the chest cavity bigger.",
             ["contract"]),
            ("When the volume of the chest cavity increases, the pressure "
             "inside it __________, so air is pushed into the lungs.",
             ["decreases"]),
            ("When you breathe out, the intercostal muscles and the diaphragm "
             "__________, and the volume of the chest cavity gets smaller.",
             ["relax"]),
            ("Breathing moves air into and out of the lungs, but __________ is "
             "a series of chemical reactions that happens inside cells.",
             ["respiration"]),
        ],
    },
    {
        "title": "1.4  Respiration",
        "questions": [
            ("Carbohydrates in our food are broken down to a sugar called "
             "__________, which is carried to every cell in the blood.",
             ["glucose"]),
            ("Aerobic respiration happens inside tiny structures in the cell "
             "called __________.",
             ["mitochondria"]),
            ("The word equation for aerobic respiration is: glucose + "
             "__________ \u2192 carbon dioxide + __________.",
             ["oxygen", "water"]),
            ("Respiration releases useful __________ from glucose in a very "
             "controlled way.",
             ["energy"]),
            ("Some of the energy transferred during respiration is released as "
             "__________, so respiring cells become warmer than their "
             "surroundings.",
             ["heat"]),
        ],
    },
    {
        "title": "1.5  Blood",
        "questions": [
            ("The pale yellow liquid part of blood, which carries dissolved "
             "substances, is called __________.",
             ["blood plasma"]),
            ("Red blood cells contain a red pigment called __________, which "
             "carries oxygen.",
             ["haemoglobin"]),
            ("When oxygen joins with haemoglobin, a bright red compound called "
             "__________ is formed.",
             ["oxyhaemoglobin"]),
            ("White blood cells defend the body against microorganisms called "
             "__________.",
             ["pathogens"]),
            ("Some white blood cells produce chemicals called __________ that "
             "stick to pathogens, while others engulf and digest pathogens in "
             "a process called __________.",
             ["antibodies", "phagocytosis"]),
        ],
    },
]


def add_para(doc, text="", *, bold=False, italic=False, size=FONT_SIZE,
             align=None, space_after=6, space_before=0):
    """Add a paragraph whose runs all use the requested formatting."""
    para = doc.add_paragraph()
    if align is not None:
        para.alignment = align
    para.paragraph_format.space_after = Pt(space_after)
    para.paragraph_format.space_before = Pt(space_before)
    run = para.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.size = size
    return para


def add_field(paragraph, instruction, size=FONT_SIZE):
    """Append a Word field (e.g. PAGE, NUMPAGES) to a paragraph."""
    run = paragraph.add_run()
    run.font.size = size
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = instruction
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    for element in (begin, instr, end):
        run._r.append(element)
    return run


def add_footer(doc):
    """Put 'Page X of Y' in the footer, centred."""
    para = doc.sections[0].footer.paragraphs[0]
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para.add_run("Page ").font.size = FONT_SIZE
    add_field(para, "PAGE")
    para.add_run(" of ").font.size = FONT_SIZE
    add_field(para, "NUMPAGES")
    return para


def add_word_bank(doc, words):
    """Render the word bank inside a single bordered cell."""
    table = doc.add_table(rows=1, cols=1)
    table.style = "Table Grid"
    para = table.cell(0, 0).paragraphs[0]
    para.paragraph_format.space_after = Pt(0)
    run = para.add_run("   ".join(words))
    run.font.size = FONT_SIZE
    return table


def add_blank(paragraph, size=FONT_SIZE):
    """Append a fixed-width (~7 cm) underline as a run of underscores.

    Underscore characters in a monospaced font join into a continuous line, so
    the fill-in area is visible in every viewer (unlike underlined spaces).
    """
    run = paragraph.add_run("_" * BLANK_CHARS)
    run.font.name = BLANK_FONT
    run.font.size = size
    return run


def add_question(doc, number, text, space_after=8):
    """Add a numbered cloze question, drawing each blank as a fill-in line."""
    para = doc.add_paragraph()
    para.paragraph_format.space_after = Pt(space_after)
    para.add_run(f"{number}.  ").font.size = FONT_SIZE
    for index, chunk in enumerate(text.split(BLANK)):
        if index:
            add_blank(para)
        if chunk:
            para.add_run(chunk).font.size = FONT_SIZE
    return para


def build_paper(doc):
    add_para(doc, "Chapter 1 — Respiration", bold=True,
             align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    add_para(doc, "Cloze test (Year 9)", italic=True,
             align=WD_ALIGN_PARAGRAPH.CENTER, space_after=12)

    add_para(doc, "Name: ______________________     "
                  "Class: ____________     Date: ____________", space_after=12)

    add_para(doc, "Instructions", bold=True, space_after=2)
    add_para(doc,
             "Complete each sentence using words from the word bank below. "
             "Each word may be used once, more than once, or not at all. "
             "There are five questions in each section.", space_after=12)

    add_para(doc, "Word bank", bold=True, space_after=4)
    add_word_bank(doc, WORD_BANK)
    add_para(doc, "", space_after=12)

    for section in SECTIONS:
        add_para(doc, section["title"], bold=True, space_before=6, space_after=6)
        for q_number, (text, _answers) in enumerate(section["questions"], start=1):
            add_question(doc, q_number, text)


def build_answer_key(doc):
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    add_para(doc, "Answer key", bold=True,
             align=WD_ALIGN_PARAGRAPH.CENTER, space_after=12)

    for section in SECTIONS:
        add_para(doc, section["title"], bold=True, space_before=6, space_after=4)
        for q_number, (_text, answers) in enumerate(section["questions"], start=1):
            answer = ", ".join(answers)
            add_para(doc, f"{q_number}.  {answer}", space_after=6)


def main():
    doc = Document()

    # A4 page, and 15 pt for every style that could contribute text.
    section = doc.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.left_margin = Mm(15)
    section.right_margin = Mm(15)
    normal = doc.styles["Normal"]
    normal.font.size = FONT_SIZE
    normal.paragraph_format.line_spacing = 1.5
    normal.paragraph_format.space_after = Pt(6)

    build_paper(doc)
    build_answer_key(doc)
    add_footer(doc)

    DST_DIR.mkdir(parents=True, exist_ok=True)
    out_path = DST_DIR / FILE_NAME
    doc.save(str(out_path))
    print(f"Wrote {out_path}")


if __name__ == "__main__":
    main()
