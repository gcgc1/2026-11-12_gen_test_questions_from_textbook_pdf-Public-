"""Generate Year 9 cloze tests for the textbook chapters listed in CHAPTERS.

Content is taken from the split chapter markdown in `30_splited_by_chapter`.
One A4 .docx is written per chapter into `40_test_papers_docx`, named
"<chapter no.> - <topic>.docx", with:

  * a shared word bank at the top,
  * five cloze questions per section,
  * a separate answer-key page at the end,
  * font size 15, 1.5 line spacing,
  * left/right margins of 1.5 cm,
  * a footer showing the chapter, topic, book page range and "Page X of Y",
  * fill-in lines about 7 cm wide (underscored).

Words in the word bank may be used once, more than once, or not at all.

Usage:
    python make_test_papers.py        # every chapter
    python make_test_papers.py 5      # only chapter 5
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt

DST_DIR = Path("40_test_papers_docx")
SPLIT_DIR = Path("30_splited_by_chapter")
BOOK_RANGE_RE = re.compile(r"book pages (\d+)\s*[\u2013-]\s*(\d+)")
FONT_SIZE = Pt(15)
FOOTER_SIZE = Pt(12)
ANSWER_SIZE = Pt(12)
ANSWER_COLUMNS = 2
BLANK = "__________"
BLANK_WIDTH = Mm(70)              # width of each cloze fill-in line
BLANK_FONT = "Courier New"        # monospaced: exactly 0.6 em per character
BLANK_CHARS = round(BLANK_WIDTH.pt / (0.6 * FONT_SIZE.pt))
SIDE_MARGIN = Mm(15)              # 1.5 cm left and right
LINE_SPACING = 1.5

# Chapter 1 — Respiration (from unit01_respiration.md)
RESPIRATION_WORD_BANK = [
    "air sacs", "alveoli", "antibodies", "blood plasma", "bronchi",
    "bronchioles", "capillaries", "cartilage", "cell", "contract",
    "decreases", "diaphragm", "diffusion", "energy", "expired",
    "gas exchange", "glucose", "haemoglobin", "heat", "increases",
    "inspired", "intercostal", "limewater", "mitochondria", "nucleus",
    "oxyhaemoglobin", "oxygen", "pathogens", "phagocytosis", "pressure",
    "relax", "respiration", "trachea", "vocal cords", "water",
]

# Chapter 2 — Properties of materials (from unit02_properties_of_materials.md)
MATERIALS_WORD_BANK = [
    "conserved", "dissolving", "melting", "mixture", "opaque",
    "particles", "solute", "solution", "solvent", "transparent",
]

# Chapter 5 — Materials and their structure (from unit05_materials_and_their_structure.md)
MATERIALS_STRUCTURE_WORD_BANK = [
    "alloy", "analogy", "atmosphere", "atoms", "biodegradable", "bioplastics",
    "boulder", "carat", "climate", "climatology", "collaboration", "compound",
    "cycled", "deflected", "deforestation", "electrical charge", "electrons",
    "electrostatic attraction", "element", "emissions", "evaporation", "fossil",
    "fossil fuels", "glacial period", "glaciers", "global warming",
    "greenhouse effect", "humidity", "ice ages", "impurities",
    "interglacial period", "locked up", "meteorology", "mixture", "neutrons",
    "nucleus", "oxygen", "peat bog", "peer review", "photosynthesis",
    "protons", "pure", "recycled", "renewable", "seawater", "statistics",
    "sub-atomic particles", "translucent", "visibility", "weather",
]

CHAPTERS = [
    {
        "number": 1,
        "topic": "Respiration",
        "word_bank": RESPIRATION_WORD_BANK,
        "sections": [
            {
                "title": "1.1  The human respiratory system",
                "questions": [
                    ("Respiration is a series of chemical reactions that takes "
                     "place inside every living __________ and releases energy "
                     "from food.", ["cell"]),
                    ("Air passes from the nose or mouth, down the __________ "
                     "(also called the windpipe) and into the lungs.", ["trachea"]),
                    ("Rings of __________ around the trachea keep it open and "
                     "stop it from collapsing.", ["cartilage"]),
                    ("The trachea divides into two __________, one to each lung, "
                     "and each of these divides into many smaller tubes called "
                     "__________.", ["bronchi", "bronchioles"]),
                    ("At the end of the bronchioles are tiny __________ where gas "
                     "exchange happens; sound is made when air makes the "
                     "__________ in your larynx vibrate.", ["air sacs", "vocal cords"]),
                ],
            },
            {
                "title": "1.2  Gas exchange",
                "questions": [
                    ("The air sacs in the lungs are also called __________, and "
                     "the very thin blood vessels wrapped around them are "
                     "__________.", ["alveoli", "capillaries"]),
                    ("Oxygen moves from the air into the blood by __________, "
                     "from a place where there is a lot of it to a place where "
                     "there is less.", ["diffusion"]),
                    ("Once oxygen is inside a red blood cell it combines with "
                     "__________ to form oxyhaemoglobin.", ["haemoglobin"]),
                    ("The movement of oxygen into the blood and carbon dioxide "
                     "out of the blood in the lungs is called __________.",
                     ["gas exchange"]),
                    ("The air you breathe in is called __________ air and the air "
                     "you breathe out is called __________ air; __________ is the "
                     "chemical used to test for carbon dioxide.",
                     ["inspired", "expired", "limewater"]),
                ],
            },
            {
                "title": "1.3  Breathing",
                "questions": [
                    ("The muscles between the ribs are called the __________ "
                     "muscles.", ["intercostal"]),
                    ("When you breathe in, the intercostal muscles and the "
                     "diaphragm __________, which makes the chest cavity bigger.",
                     ["contract"]),
                    ("When the volume of the chest cavity increases, the pressure "
                     "inside it __________, so air is pushed into the lungs.",
                     ["decreases"]),
                    ("When you breathe out, the intercostal muscles and the "
                     "diaphragm __________, and the volume of the chest cavity "
                     "gets smaller.", ["relax"]),
                    ("Breathing moves air into and out of the lungs, but "
                     "__________ is a series of chemical reactions that happens "
                     "inside cells.", ["respiration"]),
                ],
            },
            {
                "title": "1.4  Respiration",
                "questions": [
                    ("Carbohydrates in our food are broken down to a sugar called "
                     "__________, which is carried to every cell in the blood.",
                     ["glucose"]),
                    ("Aerobic respiration happens inside tiny structures in the "
                     "cell called __________.", ["mitochondria"]),
                    ("The word equation for aerobic respiration is: glucose + "
                     "__________ \u2192 carbon dioxide + __________.",
                     ["oxygen", "water"]),
                    ("Respiration releases useful __________ from glucose in a "
                     "very controlled way.", ["energy"]),
                    ("Some of the energy transferred during respiration is "
                     "released as __________, so respiring cells become warmer "
                     "than their surroundings.", ["heat"]),
                ],
            },
            {
                "title": "1.5  Blood",
                "questions": [
                    ("The pale yellow liquid part of blood, which carries "
                     "dissolved substances, is called __________.", ["blood plasma"]),
                    ("Red blood cells contain a red pigment called __________, "
                     "which carries oxygen.", ["haemoglobin"]),
                    ("When oxygen joins with haemoglobin, a bright red compound "
                     "called __________ is formed.", ["oxyhaemoglobin"]),
                    ("White blood cells defend the body against microorganisms "
                     "called __________.", ["pathogens"]),
                    ("Some white blood cells produce chemicals called __________ "
                     "that stick to pathogens, while others engulf and digest "
                     "pathogens in a process called __________.",
                     ["antibodies", "phagocytosis"]),
                ],
            },
        ],
    },
    {
        "number": 2,
        "topic": "Properties of materials",
        "word_bank": MATERIALS_WORD_BANK,
        "sections": [
            {
                "title": "2.1  Dissolving",
                "questions": [
                    ("When a substance dissolves in a liquid, the substance that "
                     "dissolves is called the __________, and the liquid it "
                     "dissolves into is called the __________.",
                     ["solute", "solvent"]),
                    ("A __________ is a mixture formed when the particles of a "
                     "solute spread out among the particles of a solvent.",
                     ["solution"]),
                    ("Solutions are always __________, which means you can see "
                     "through them; this does not mean they are colourless.",
                     ["transparent"]),
                    ("A liquid such as milk is __________, so you cannot see "
                     "through it; this shows that milk is not a solution.",
                     ["opaque"]),
                    ("When salt dissolves in water, no mass is lost: the mass of "
                     "the solution equals the mass of the solute plus the "
                     "solvent, so we say that mass has been __________.",
                     ["conserved"]),
                ],
            },
        ],
    },
    {
        "number": 5,
        "topic": "Materials and their structure",
        "word_bank": MATERIALS_STRUCTURE_WORD_BANK,
        "sections": [
            {
                "title": "5.1  The structure of the atom",
                "questions": [
                    ("Atoms are made up of three kinds of __________: protons, "
                     "neutrons and electrons.", ["sub-atomic particles"]),
                    ("The protons and neutrons are grouped closely together in "
                     "the centre of the atom to form the __________.", ["nucleus"]),
                    ("Protons have a positive __________, neutrons have no charge "
                     "and electrons have a negative charge.", ["electrical charge"]),
                    ("The positive charge on the protons and the negative charge "
                     "on the electrons produce an __________ that holds individual "
                     "atoms together.", ["electrostatic attraction"]),
                    ("In Rutherford's gold foil experiment, most particles passed "
                     "straight through the foil but a few were __________; other "
                     "scientists then examine each other's work carefully, a "
                     "process called __________.", ["deflected", "peer review"]),
                ],
            },
            {
                "title": "5.2  Purity",
                "questions": [
                    ("If an element is __________, every one of its atoms is "
                     "exactly the same and made of just one type of atom.", ["pure"]),
                    ("Gold jewellery is often an __________ that contains other "
                     "metals such as copper or silver as well as gold.", ["alloy"]),
                    ("The purity of gold is measured in __________; 24 carat gold "
                     "is pure gold, while 18 carat gold is only 75% gold.",
                     ["carat"]),
                    ("Pure diamonds are colourless and __________, which means "
                     "they let light through, but if other elements are mixed in "
                     "they can be different colours.", ["translucent"]),
                    ("Seawater contains sodium chloride and other salts; to obtain "
                     "a pure product the water must be removed by __________ and "
                     "the other __________ must be separated out.",
                     ["evaporation", "impurities"]),
                ],
            },
            {
                "title": "5.3  Weather and climate",
                "questions": [
                    ("The state of the atmosphere, and how it changes from minute "
                     "to minute, hour to hour and day to day, is called the "
                     "__________.", ["weather"]),
                    ("The average weather of a place over a long period, usually "
                     "more than 30 years, is called its __________.", ["climate"]),
                    ("The study of weather is called __________, while the study "
                     "of climate is called __________.",
                     ["meteorology", "climatology"]),
                    ("Weather is described using factors such as temperature, "
                     "precipitation, cloudiness, wind and __________, which is a "
                     "measure of the amount of water vapour in the air.",
                     ["humidity"]),
                    ("Climate information includes the __________ of the weather, "
                     "which describe the normal weather and the range of extreme "
                     "weather at a place.", ["statistics"]),
                ],
            },
            {
                "title": "5.4  Climate and ice ages",
                "questions": [
                    ("A very cold period, when much of the Earth was frozen, is "
                     "called a __________, and the warmer period between them is "
                     "called an __________.",
                     ["glacial period", "interglacial period"]),
                    ("Over the last 450 000 years the Earth's climate has "
                     "__________ between glacial and interglacial periods.",
                     ["cycled"]),
                    ("__________ are rivers of ice that move slowly downhill and "
                     "carry rocks with them.", ["glaciers"]),
                    ("A large rock that was carried by a glacier and left behind "
                     "when the ice melted is called a __________.", ["boulder"]),
                    ("Long cold periods that contained both glacial and "
                     "interglacial phases are called __________; evidence for "
                     "these cycles comes from pollen preserved in a __________.",
                     ["ice ages", "peat bog"]),
                ],
            },
            {
                "title": "5.5  Atmosphere and climate",
                "questions": [
                    ("Plants use carbon dioxide and water to make glucose and "
                     "oxygen by the process of __________.", ["photosynthesis"]),
                    ("When organisms die and do not rot, their carbon can be "
                     "turned into __________ such as oil and coal, where the "
                     "carbon is __________ until the fuels are burned.",
                     ["fossil fuels", "locked up"]),
                    ("Burning fossil fuels and __________ (cutting down forests) "
                     "increase the __________ of carbon dioxide in the "
                     "atmosphere.", ["deforestation", "emissions"]),
                    ("The extra carbon dioxide strengthens the __________ and "
                     "leads to __________, the warming of the Earth's climate.",
                     ["greenhouse effect", "global warming"]),
                    ("A resource that does not run out, or can be replaced within "
                     "a human lifetime, is called a __________ resource, and a "
                     "__________ material such as bioplastic breaks down "
                     "naturally.", ["renewable", "biodegradable"]),
                ],
            },
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


def add_footer(doc, chapter):
    """Footer: chapter and book page range, then 'Page X of Y', centred."""
    footer = doc.sections[0].footer

    label = f"Chapter {chapter['number']} — {chapter['topic']}"
    book_range = book_page_range(chapter["number"])
    if book_range:
        label += f" · Book pages {book_range[0]}\u2013{book_range[1]}"
    first = footer.paragraphs[0]
    first.alignment = WD_ALIGN_PARAGRAPH.CENTER
    first.add_run(label).font.size = FOOTER_SIZE

    page_para = footer.add_paragraph()
    page_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    page_para.add_run("Page ").font.size = FOOTER_SIZE
    add_field(page_para, "PAGE", size=FOOTER_SIZE)
    page_para.add_run(" of ").font.size = FOOTER_SIZE
    add_field(page_para, "NUMPAGES", size=FOOTER_SIZE)


def book_page_range(unit_number):
    """Start/end book page for a unit, read from its split markdown.

    `30_splited_by_chapter` may hold more than one file for a unit (e.g. a
    partial sample and the full textbook); the widest book-page range wins.
    """
    best = None
    for md in SPLIT_DIR.glob(f"*_unit{unit_number:02d}_*.md"):
        match = BOOK_RANGE_RE.search(md.read_text(encoding="utf-8", errors="ignore"))
        if match:
            start, end = int(match.group(1)), int(match.group(2))
            if best is None or end > best[1]:
                best = (start, end)
    return best


def add_word_bank(doc, words):
    """Render the word bank inside a single bordered cell."""
    table = doc.add_table(rows=1, cols=1)
    table.style = "Table Grid"
    para = table.cell(0, 0).paragraphs[0]
    para.paragraph_format.space_after = Pt(0)
    run = para.add_run("   ".join(words))
    run.font.size = FONT_SIZE


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


def set_columns(section, count, space_twips=425):
    """Set the number of text columns for a section (Word `w:cols`)."""
    sect_pr = section._sectPr
    cols = sect_pr.find(qn("w:cols"))
    if cols is None:
        cols = OxmlElement("w:cols")
        pg_mar = sect_pr.find(qn("w:pgMar"))   # w:cols must follow w:pgMar
        if pg_mar is not None:
            pg_mar.addnext(cols)
        else:
            sect_pr.insert(0, cols)
    cols.set(qn("w:num"), str(count))
    cols.set(qn("w:space"), str(space_twips))


def build_document(chapter) -> Path:
    """Build one chapter's test paper and return the path written."""
    doc = Document()

    # A4 page, 1.5 cm side margins, 15 pt, 1.5 line spacing.
    section = doc.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.left_margin = SIDE_MARGIN
    section.right_margin = SIDE_MARGIN
    normal = doc.styles["Normal"]
    normal.font.size = FONT_SIZE
    normal.paragraph_format.line_spacing = LINE_SPACING
    normal.paragraph_format.space_after = Pt(6)

    title = f"Chapter {chapter['number']} — {chapter['topic']}"
    add_para(doc, title, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
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
    add_word_bank(doc, chapter["word_bank"])
    add_para(doc, "", space_after=12)

    for sec in chapter["sections"]:
        add_para(doc, sec["title"], bold=True, space_before=6, space_after=6)
        for q_number, (text, _answers) in enumerate(sec["questions"], start=1):
            add_question(doc, q_number, text)

    # Answer key on its own page: two columns, font size 12.
    key_section = doc.add_section(WD_SECTION.NEW_PAGE)
    key_section.page_width = Mm(210)
    key_section.page_height = Mm(297)
    key_section.left_margin = SIDE_MARGIN
    key_section.right_margin = SIDE_MARGIN
    set_columns(key_section, ANSWER_COLUMNS)

    add_para(doc, "Answer key", bold=True, size=ANSWER_SIZE,
             align=WD_ALIGN_PARAGRAPH.CENTER, space_after=12)
    for sec in chapter["sections"]:
        add_para(doc, sec["title"], bold=True, size=ANSWER_SIZE,
                 space_before=6, space_after=4)
        for q_number, (_text, answers) in enumerate(sec["questions"], start=1):
            add_para(doc, f"{q_number}.  {', '.join(answers)}",
                     size=ANSWER_SIZE, space_after=6)

    add_footer(doc, chapter)

    DST_DIR.mkdir(parents=True, exist_ok=True)
    out_path = DST_DIR / f"{chapter['number']} - {chapter['topic']}.docx"
    doc.save(str(out_path))
    return out_path


def main() -> None:
    """Build every chapter, or only the numbers given on the command line."""
    wanted = {int(arg) for arg in sys.argv[1:]} if len(sys.argv) > 1 else None
    for chapter in CHAPTERS:
        if wanted and chapter["number"] not in wanted:
            continue
        print(f"Wrote {build_document(chapter)}")


if __name__ == "__main__":
    main()
