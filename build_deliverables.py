"""Build the Phase 2 PowerPoint deck and Word report into deliverables/.

Re-runnable: `python3 build_deliverables.py`. Requires charts/ to already exist
(run phase2_charts.py first if needed).
"""
from __future__ import annotations

import os
from datetime import date

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches as DocxInches
from docx.shared import Pt as DocxPt
from docx.shared import RGBColor as DocxRGB
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

# ------------------------------------------------------------------
# Branding
# ------------------------------------------------------------------
NAVY = RGBColor(0x01, 0x21, 0x69)
GOLD = RGBColor(0xF2, 0xA9, 0x00)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK_TEXT = RGBColor(0x22, 0x22, 0x22)

DOCX_NAVY = DocxRGB(0x01, 0x21, 0x69)
DOCX_GOLD = DocxRGB(0xF2, 0xA9, 0x00)

TITLE = "LGS Master's Programs: Tuition Pricing & Scholarship Analysis"
SUBTITLE = "Emory Laney Graduate School — Phase 2"
REPORT_DATE = date(2026, 4, 22).strftime("%B %d, %Y")

# ------------------------------------------------------------------
# Content (insights keyed to each chart)
# ------------------------------------------------------------------
CHART_SLIDES = [
    ("Discount rates vary six-fold across programs — MDP tops the list at 55% while Data Science and MBID give nothing",
     "charts/01_discount_rate_by_program.png"),
    ("MDP's ~$60k average discount is the biggest absolute tuition give — and its net tuition is still the highest in the portfolio",
     "charts/02_avg_gross_vs_net_tuition.png"),
    ("High international share doesn't come with lower aid — Computer Science is 86% international AND 49% discounted",
     "charts/03_intl_pct_vs_discount_rate.png"),
    ("Computer Science enrollment has roughly doubled since 2022 — driving nearly all of Emory's LGS master's volume growth",
     "charts/04_enrollment_trend_by_program.png"),
    ("Discount rates are broadly stable year-over-year — no program is systematically tightening aid",
     "charts/05_discount_rate_trend_by_program.png"),
    ("Emory's annualized tuition ($47–50k) undercuts Columbia/NYU by $20–40k and sits above Georgia Tech's in-state rates — mid-market positioning",
     "charts/06_peer_benchmark_emory_vs_peers.png"),
]

TOP5 = [
    ("MDP is Emory's costliest and most-discounted master's",
     "$108k average gross, $49k net, 55% discount rate across 135 students over the 8-year panel ($8.4M in LGS scholarships)."),
    ("Two programs operate at 0% discount",
     "Data Science (n=11) and MBID (n=8) show gross = net for every student. Obvious candidates for a pricing review."),
    ("Enrollment is bimodal",
     "Computer Science (154) and MDP (135) together account for 59% of the 493-student sample; eight other programs share the remaining 41%."),
    ("Computer Science is 86% international — and still 49% discounted",
     "Emory competes for these students on price as well as visa-friendliness; aid does not scale down as international share rises."),
    ("Emory sits mid-pack against peers",
     "Annualized gross tuition ~$47–50k across Data/CS, Economics, and General buckets — below Columbia/NYU ($65–91k) and above Georgia Tech's in-state rates ($31–41k)."),
]

APPENDIX_BULLETS = [
    "Source data: institutional Excel exports of gross tuition billed and LGS scholarship awards; 8-year panel spanning AY 2019 through AY 2026.",
    "Net tuition is clipped at zero. In rare cases a kept scholarship exceeds billed tuition for a partial-term student; we report that student as paying $0 rather than a negative amount.",
    "Scholarship filter retains descriptions containing 'lgs' and drops external/federal awards (Pell, GRFP, Yellow Ribbon, Americorps, etc.) plus generic LGS 'special' awards, so discount rates reflect institutional aid only.",
    "Small-n programs are merged: ECONMS + ECON4P1MS → Economics; DATASCIMS + QTMMS → Data Science; DEVPRACMDP + HUMANRTCRT → MDP.",
    "Annualization for peer comparison uses program-specific terms-per-year (from the Information Sheet plus empirical seasonal distribution) rather than a blanket 2-term assumption. This corrects a ~33% understatement for 3-term programs such as MDP and Data Science.",
    "Programs with zero kept scholarship (Data Science, MBID) are retained in the summary with a 0% discount rate — this is intentional and flags potential pricing discrepancies.",
]

REPORT_SECTIONS = [
    ("Discount Rates",
     ("Discount rates range from 0% to 55% across Emory's LGS master's programs. MDP, Computer Science, "
      "and Computer Science 4+1 all discount roughly half of sticker price, while Data Science and MBID "
      "give no institutional scholarship at all. The 8-year trend is broadly stable — year-over-year "
      "fluctuations sit within normal cohort noise and no program is systematically tightening aid."),
     ["charts/01_discount_rate_by_program.png", "charts/05_discount_rate_trend_by_program.png"]),
    ("Gross vs Net Tuition",
     ("Sticker price and net price diverge most sharply for MDP ($108k gross → $49k net) and Computer "
      "Science ($66k → $33k). Programs with low or zero discount rates (Bioethics, Data Science, MBID) "
      "collect close to full sticker from every student. Gross-to-net gap is the clearest single view of "
      "where Emory is spending its tuition-discount dollars."),
     ["charts/02_avg_gross_vs_net_tuition.png"]),
    ("International Mix",
     ("Bubble size shows program enrollment; position shows international share vs discount rate. Computer "
      "Science is the clear outlier — 86% international yet still 49% discounted. Across the portfolio there "
      "is no inverse relationship between international share and aid: high-international programs are not "
      "systematically less aided, which contradicts the common assumption that international students "
      "primarily subsidize domestic ones."),
     ["charts/03_intl_pct_vs_discount_rate.png"]),
    ("Enrollment Trends",
     ("The 8-year panel (AY 2019–2026) shows Computer Science climbing from the mid-30s to 60+ students per "
      "year, accounting for most of the overall growth in LGS master's enrollment. MDP remains Emory's other "
      "large program, steady in the 30–40 range per year. Smaller programs (Data Science, MBID, Economics) "
      "sit below 10 students per year throughout the window."),
     ["charts/04_enrollment_trend_by_program.png"]),
    ("Peer Benchmarking",
     ("Using the corrected program-specific annualization, Emory's gross tuition sits firmly mid-pack: "
      "roughly $47–50k across Data/CS, Economics, and General program buckets. Columbia leads the peer set "
      "at $65–91k depending on discipline, with NYU close behind. Georgia Tech's in-state rates ($31–41k) "
      "undercut Emory in every bucket. The market splits into a high-tier (Columbia/NYU), a mid-tier (Emory), "
      "and a value tier (Georgia Tech)."),
     ["charts/06_peer_benchmark_emory_vs_peers.png"]),
]


# ------------------------------------------------------------------
# PowerPoint helpers
# ------------------------------------------------------------------
def add_rect(slide, x, y, w, h, fill_color, line=False):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if not line:
        shape.line.fill.background()
    return shape


def add_text(slide, x, y, w, h, text, *, color=DARK_TEXT, size=18,
             bold=False, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = "Calibri"
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    return box


def build_title_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    sw, sh = prs.slide_width, prs.slide_height
    # Full-bleed navy background
    add_rect(slide, 0, 0, sw, sh, NAVY)
    # Gold accent bar
    add_rect(slide, Inches(1), Inches(3.0), Inches(1.5), Inches(0.1), GOLD)
    # Title
    add_text(slide, Inches(1), Inches(3.2), sw - Inches(2), Inches(2),
             TITLE, color=WHITE, size=40, bold=True, align=PP_ALIGN.LEFT)
    # Subtitle
    add_text(slide, Inches(1), Inches(5.0), sw - Inches(2), Inches(0.6),
             SUBTITLE, color=GOLD, size=20, bold=True, align=PP_ALIGN.LEFT)
    # Date
    add_text(slide, Inches(1), Inches(6.8), sw - Inches(2), Inches(0.4),
             REPORT_DATE, color=WHITE, size=14, align=PP_ALIGN.LEFT)


def build_exec_summary_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    sw = prs.slide_width
    # Navy heading bar
    add_rect(slide, 0, 0, sw, Inches(1.0), NAVY)
    add_text(slide, Inches(0.5), Inches(0.2), sw - Inches(1), Inches(0.7),
             "Executive Summary — Top 5 Findings",
             color=WHITE, size=28, bold=True)
    # Gold accent under heading
    add_rect(slide, 0, Inches(1.0), sw, Inches(0.08), GOLD)

    y = Inches(1.4)
    line_h = Inches(1.1)
    for i, (headline, detail) in enumerate(TOP5, start=1):
        # Number
        add_text(slide, Inches(0.5), y, Inches(0.7), line_h,
                 f"{i}.", color=GOLD, size=28, bold=True, align=PP_ALIGN.CENTER)
        # Headline + detail
        box = slide.shapes.add_textbox(Inches(1.2), y, sw - Inches(1.7), line_h)
        tf = box.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        r1 = p1.add_run()
        r1.text = headline
        r1.font.name = "Calibri"
        r1.font.size = Pt(16)
        r1.font.bold = True
        r1.font.color.rgb = NAVY
        p2 = tf.add_paragraph()
        r2 = p2.add_run()
        r2.text = detail
        r2.font.name = "Calibri"
        r2.font.size = Pt(13)
        r2.font.color.rgb = DARK_TEXT
        y += line_h


def build_chart_slide(prs, insight, chart_path, page_num):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    sw, sh = prs.slide_width, prs.slide_height
    # Navy headline strip
    add_rect(slide, 0, 0, sw, Inches(1.3), NAVY)
    add_text(slide, Inches(0.5), Inches(0.25), sw - Inches(1), Inches(0.85),
             insight, color=WHITE, size=20, bold=True)
    # Gold accent under headline
    add_rect(slide, 0, Inches(1.3), sw, Inches(0.06), GOLD)
    # Chart centered below
    img_y = Inches(1.6)
    img_h = sh - img_y - Inches(0.4)
    slide.shapes.add_picture(
        chart_path,
        left=Inches(0.5),
        top=img_y,
        width=sw - Inches(1.0),
        height=img_h,
    )
    # Page number bottom right
    add_text(slide, sw - Inches(0.8), sh - Inches(0.35), Inches(0.6), Inches(0.3),
             str(page_num), color=NAVY, size=11, align=PP_ALIGN.RIGHT)


def build_appendix_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    sw = prs.slide_width
    add_rect(slide, 0, 0, sw, Inches(1.0), NAVY)
    add_text(slide, Inches(0.5), Inches(0.2), sw - Inches(1), Inches(0.7),
             "Data Notes & Methodology", color=WHITE, size=26, bold=True)
    add_rect(slide, 0, Inches(1.0), sw, Inches(0.08), GOLD)

    y = Inches(1.4)
    for bullet in APPENDIX_BULLETS:
        box = slide.shapes.add_textbox(Inches(0.5), y, sw - Inches(1), Inches(0.9))
        tf = box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        r1 = p.add_run()
        r1.text = "• "
        r1.font.name = "Calibri"
        r1.font.size = Pt(14)
        r1.font.bold = True
        r1.font.color.rgb = GOLD
        r2 = p.add_run()
        r2.text = bullet
        r2.font.name = "Calibri"
        r2.font.size = Pt(13)
        r2.font.color.rgb = DARK_TEXT
        y += Inches(0.85)


def build_pptx(path: str) -> None:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    build_title_slide(prs)
    build_exec_summary_slide(prs)
    for i, (insight, chart) in enumerate(CHART_SLIDES, start=3):
        build_chart_slide(prs, insight, chart, i)
    build_appendix_slide(prs)

    prs.save(path)
    print(f"  wrote {path}")


# ------------------------------------------------------------------
# Word helpers
# ------------------------------------------------------------------
def _shading(color_hex: str):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color_hex)
    return shd


def _page_break(doc):
    p = doc.add_paragraph()
    run = p.add_run()
    run.add_break(WD_BREAK.PAGE)


def add_gold_rule(doc):
    """Thin gold horizontal rule via a 1-row, 1-col table with bottom border."""
    table = doc.add_table(rows=1, cols=1)
    cell = table.cell(0, 0)
    cell.text = ""
    # Set bottom border to gold, remove others
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "right"):
        b = OxmlElement(f"w:{edge}")
        b.set(qn("w:val"), "nil")
        tcBorders.append(b)
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "18")  # eighths of a point ≈ 2.25pt
    bottom.set(qn("w:color"), "F2A900")
    tcBorders.append(bottom)
    tcPr.append(tcBorders)
    # Zero row height so the rule is just the border
    for p in cell.paragraphs:
        p.paragraph_format.space_before = DocxPt(0)
        p.paragraph_format.space_after = DocxPt(0)


def add_heading(doc, text, level=1, size=20):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = DocxPt(12)
    p.paragraph_format.space_after = DocxPt(4)
    run = p.add_run(text)
    run.font.name = "Calibri"
    run.font.size = DocxPt(size)
    run.font.bold = True
    run.font.color.rgb = DOCX_NAVY


def add_body(doc, text, size=11):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = DocxPt(8)
    run = p.add_run(text)
    run.font.name = "Calibri"
    run.font.size = DocxPt(size)
    run.font.color.rgb = DocxRGB(0x22, 0x22, 0x22)


def add_numbered(doc, n, headline, detail):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = DocxPt(6)
    rn = p.add_run(f"{n}. ")
    rn.font.name = "Calibri"
    rn.font.size = DocxPt(11)
    rn.font.bold = True
    rn.font.color.rgb = DOCX_GOLD
    rh = p.add_run(headline + " — ")
    rh.font.name = "Calibri"
    rh.font.size = DocxPt(11)
    rh.font.bold = True
    rh.font.color.rgb = DOCX_NAVY
    rd = p.add_run(detail)
    rd.font.name = "Calibri"
    rd.font.size = DocxPt(11)
    rd.font.color.rgb = DocxRGB(0x22, 0x22, 0x22)


def add_bullet(doc, text):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = DocxPt(4)
    run = p.runs[0] if p.runs else p.add_run()
    run.text = text
    run.font.name = "Calibri"
    run.font.size = DocxPt(11)
    run.font.color.rgb = DocxRGB(0x22, 0x22, 0x22)


def build_docx(path: str) -> None:
    doc = Document()

    # Global default font
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = DocxPt(11)

    # ---------- Cover page ----------
    for _ in range(6):
        doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(TITLE)
    r.font.name = "Calibri"
    r.font.size = DocxPt(28)
    r.font.bold = True
    r.font.color.rgb = DOCX_NAVY

    add_gold_rule(doc)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(SUBTITLE)
    r.font.name = "Calibri"
    r.font.size = DocxPt(16)
    r.font.color.rgb = DOCX_GOLD

    for _ in range(2):
        doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(REPORT_DATE)
    r.font.name = "Calibri"
    r.font.size = DocxPt(14)
    r.font.color.rgb = DocxRGB(0x44, 0x44, 0x44)

    _page_break(doc)

    # ---------- Executive Summary ----------
    add_heading(doc, "Executive Summary", size=22)
    add_gold_rule(doc)
    add_body(
        doc,
        "This report summarizes pricing, scholarship discounting, and peer benchmarking for Emory's "
        "Laney Graduate School master's programs, using an 8-year panel of student-level tuition and "
        "LGS-scholarship records (AY 2019–2026).",
    )
    for i, (headline, detail) in enumerate(TOP5, start=1):
        add_numbered(doc, i, headline, detail)

    _page_break(doc)

    # ---------- Sections ----------
    for title, narrative, charts in REPORT_SECTIONS:
        add_heading(doc, title, size=18)
        add_gold_rule(doc)
        add_body(doc, narrative)
        for chart in charts:
            doc.add_picture(chart, width=DocxInches(6.5))
            cap = doc.paragraphs[-1]
            cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _page_break(doc)

    # ---------- Appendix ----------
    add_heading(doc, "Data & Methodology", size=18)
    add_gold_rule(doc)
    add_body(
        doc,
        "The analysis pipeline is implemented in analysis.ipynb and produces three Excel outputs "
        "(program_summary_output.xlsx, student_program_level_output.xlsx, and "
        "student_program_year_output.xlsx).",
    )
    for bullet in APPENDIX_BULLETS:
        add_bullet(doc, bullet)

    add_heading(doc, "Terms-Per-Year Lookup", size=14)
    rows = [
        ("COMPSCIMS, CS4P1MS, MATHMS, ECON4P1MS", "2", "Information Sheet (Fall-Spring)"),
        ("BIOETHMA, BIOETH4P1, BMIDMS", "2", "empirical (Fall-Spring dominant)"),
        ("ECONMS", "3", "Information Sheet (Summer-Fall-Spring)"),
        ("DATASCIMS", "3", "Information Sheet (Fall-Spring-Summer)"),
        ("QTMMS, DEVPRACMDP, HUMANRTCRT, BBS4P1MS", "3", "empirical (all three seasons ~equal)"),
    ]
    table = doc.add_table(rows=1 + len(rows), cols=3)
    table.style = "Light Grid Accent 1"
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(("acad_plan", "terms/yr", "source")):
        hdr_cells[i].text = ""
        p = hdr_cells[i].paragraphs[0]
        r = p.add_run(h)
        r.font.bold = True
        r.font.color.rgb = DOCX_NAVY
    for ri, (a, b, c) in enumerate(rows, start=1):
        table.rows[ri].cells[0].text = a
        table.rows[ri].cells[1].text = b
        table.rows[ri].cells[2].text = c

    doc.save(path)
    print(f"  wrote {path}")


def main() -> None:
    os.makedirs("deliverables", exist_ok=True)
    build_pptx("deliverables/emory_grad_pricing.pptx")
    build_docx("deliverables/emory_grad_pricing_report.docx")


if __name__ == "__main__":
    main()
