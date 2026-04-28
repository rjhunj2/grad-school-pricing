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
SUBTITLE = "Emory Laney Graduate School — Phase 2 (v2)"
REPORT_DATE = date(2026, 4, 23).strftime("%B %d, %Y")

# ------------------------------------------------------------------
# Content (insights keyed to each chart)
# ------------------------------------------------------------------
SEMESTER_SLIDE = {
    "headline": "Per-semester billing is consistent Fall/Spring — Summer reveals MDP's heaviest discounting",
    "subtext":  "Charts sorted by Fall gross tuition for easy cross-semester comparison. df_pricing population (gross tuition ≥ $20k).",
    "charts":   [
        "charts/gross_net_fall.png",
        "charts/gross_net_spring.png",
        "charts/gross_net_summer.png",
    ],
}

# Standard one-chart slides. The gross-vs-net slot is built separately as a
# left-panel + three-chart stack (see SEMESTER_SLIDE / build_three_stack_slide).
CHART_SLIDES = [
    ("Discount rates vary six-fold across programs — MDP tops the list at 55% while Data Science and MBID give nothing",
     "charts/01_discount_rate_by_program.png"),
    None,  # slot 4 — replaced by the semester-stack slide
    ("High international share doesn't come with lower aid — Computer Science is 86% international AND 49% discounted",
     "charts/03_intl_pct_vs_discount_rate.png"),
    ("Computer Science enrollment has roughly doubled since 2022 — driving nearly all of Emory's LGS master's volume growth",
     "charts/04_enrollment_trend_by_program.png"),
    ("Discount rates are broadly stable year-over-year — no program is systematically tightening aid",
     "charts/05_discount_rate_trend_by_program.png"),
    ("Emory's annualized tuition ($47–51k) undercuts Columbia/NYU by $20–40k and sits above Georgia Tech's in-state rates — mid-market positioning",
     "charts/06_peer_benchmark_emory_vs_peers.png"),
]

# v2: pricing metrics below (avg gross/net tuition, discount rate, peer
# annualization) are computed on df_pricing — student-program rows with
# total billed tuition of at least $20,000. This drops 11 of 493 rows
# (partial-term or single-course billings) that were pulling the averages
# down and compressing discount rates. Enrollment counts and
# international-mix figures still use the full population.
TOP5 = [
    ("MDP is Emory's costliest and most-discounted master's",
     "$108k average gross, $49k net, 55% discount rate across 135 full-load students over the 8-year panel ($8.4M in LGS scholarships)."),
    ("Two programs operate at 0% discount",
     "Data Science (n=11) and MBID (n=8) show gross = net for every full-load student. Obvious candidates for a pricing review."),
    ("Enrollment is bimodal",
     "Computer Science (154) and MDP (135) together account for 59% of the 493-student sample; eight other programs share the remaining 41%."),
    ("Computer Science is 86% international — and still ~49% discounted",
     "Emory competes for these students on price as well as visa-friendliness; aid does not scale down as international share rises."),
    ("Emory sits mid-pack against peers",
     "Annualized gross tuition ~$47–51k across Data/CS, Economics, and General buckets — below Columbia/NYU ($65–91k) and above Georgia Tech's in-state rates ($31–41k)."),
]

APPENDIX_BULLETS = [
    "Source data: institutional Excel exports of gross tuition billed and LGS scholarship awards; 8-year panel spanning AY 2019 through AY 2026.",
    "Net tuition is clipped at zero. In rare cases a kept scholarship exceeds billed tuition for a partial-term student; we report that student as paying $0 rather than a negative amount.",
    "Scholarship filter retains descriptions containing 'lgs' and drops external/federal awards (Pell, GRFP, Yellow Ribbon, Americorps, etc.) plus generic LGS 'special' awards, so discount rates reflect institutional aid only.",
    "Small-n programs are merged: ECONMS + ECON4P1MS → Economics; DATASCIMS + QTMMS → Data Science; DEVPRACMDP + HUMANRTCRT → MDP.",
    "Annualization for peer comparison uses program-specific terms-per-year (from the Information Sheet plus empirical seasonal distribution) rather than a blanket 2-term assumption. This corrects a ~33% understatement for 3-term programs such as MDP and Data Science.",
    "Programs with zero kept scholarship (Data Science, MBID) are retained in the summary with a 0% discount rate — this is intentional and flags potential pricing discrepancies.",
    "v2 pricing filter (df_pricing): avg gross/net tuition, discount rate, and peer benchmark exclude 11 of 493 student-program rows with total billed tuition under $20,000. These reflect single-course or partial-term billings with ~$0 scholarship that were depressing averages and compressing the gross-vs-net gap. Enrollment counts, international mix, and per-year trend charts continue to use the full population.",
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
      "Science ($67k → $34k). Programs with low or zero discount rates (Bioethics, Data Science, MBID) "
      "collect close to full sticker from every full-load student. Gross-to-net gap is the clearest "
      "single view of where Emory is spending its tuition-discount dollars."),
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
     ("Using the corrected program-specific annualization applied to full-load students, Emory's gross "
      "tuition sits firmly mid-pack: $47,353 Data/CS, $46,719 Economics, $50,654 General. Columbia leads "
      "the peer set at $65–91k depending on discipline, with NYU close behind at ~$70–76k. Georgia Tech's "
      "in-state rates ($31–41k) undercut Emory in every bucket. The market splits into a high-tier "
      "(Columbia/NYU), a mid-tier (Emory), and a value tier (Georgia Tech)."),
     ["charts/06_peer_benchmark_emory_vs_peers.png"]),
]

# ------------------------------------------------------------------
# Speaker notes (one section per slide in the deck)
# ------------------------------------------------------------------
SPEAKER_NOTES = [
    ("Slide 1 — Title",
     [
         "Scope: Emory LGS master's programs, AY 2019 through AY 2026, 10 programs, 493 student-program enrollments.",
         "This is the v2 cut: pricing metrics now exclude 11 partial-term rows under $20,000 in total billed tuition. Enrollment counts and the international mix still reflect the full 493.",
     ]),
    ("Slide 2 — Executive Summary",
     [
         "Lead with MDP: costliest program on both gross and net basis, and the single biggest absolute tuition give.",
         "Flag Data Science (n=11) and MBID (n=8) as 0%-discount anomalies — obvious targets for a pricing review.",
         "Bimodal enrollment: CS + MDP = 59% of the sample; this is a portfolio of two big programs plus a long tail.",
         "High international share doesn't bring aid down — CS is 86% international AND ~49% discounted.",
         "Emory is mid-pack: $47–51k annualized, below Columbia/NYU and above Georgia Tech's in-state rates.",
     ]),
    ("Slide 3 — Discount rate by program",
     [
         "Top: MDP 55%, CS 4+1 49.4%, Computer Science 49.2%. These three are the institutional-aid heavyweights.",
         "Middle band: Math 41%, Economics 34%, Cancer Biology 4+1 26%.",
         "Bottom: Bioethics 9.5% and Bioethics 4+1 16%; Data Science and MBID at 0%.",
         "Uses the pricing filter (482 full-load students). Dropping the 11 sub-$20k rows barely moves any individual bar — the comparisons are robust to that choice — but it keeps the metric defensible.",
     ]),
    ("Slide 4 — Avg Gross vs Net Tuition",
     [
         "MDP is off the top: ~$108k gross, ~$49k net — ~$60k of institutional aid per student on average.",
         "Computer Science: ~$67k gross, ~$34k net (v2 values are ~$1k higher than the v1 deck because the filter excluded four partial-term CS rows).",
         "Bioethics moved the most under the filter: gross $54k → $57k, net $48k → $52k, because six low-billing rows were dropped.",
         "MBID and Data Science: gross = net; the two bars are identical.",
     ]),
    ("Slide 5 — International % vs Discount Rate",
     [
         "Bubble size = enrollment; x = international %, y = discount rate.",
         "CS is the big navy bubble upper-right: 86% international, 49% discount, 154 students.",
         "No downward relationship between international share and aid — contradicts the intuition that international students subsidize domestic ones.",
         "Uses the full population (no pricing filter) — we don't want to undercount international students on a count-based chart.",
     ]),
    ("Slide 6 — Enrollment Trends",
     [
         "Full panel, AY 2019 through AY 2026, full population.",
         "Computer Science climbs from the mid-30s into the 60s — roughly doubled since 2022 — this is the main volume story.",
         "MDP steady around 30–40 per year.",
         "Smaller programs (Data Science, MBID, Economics) stay below ~10 per year throughout.",
     ]),
    ("Slide 7 — Discount Rate Trend",
     [
         "By-program, year-over-year. Shapes tell you stability; no program is systematically tightening or loosening aid.",
         "MDP hovers at 55–60%. CS holds around 49%. 0%-discount programs stay flat at zero.",
         "Year-to-year wobble sits within cohort-size noise; don't over-read single-year moves.",
     ]),
    ("Slide 8 — Peer Benchmark",
     [
         "Emory's three program-group averages on pricing-filtered data: Data/CS $47,353, Economics $46,719, General $50,654.",
         "Columbia leads Data/CS at $64,800 and Economics at $90,732.",
         "NYU sits in the $70,000–$75,750 range across buckets.",
         "Georgia Tech in-state is the floor: $31,210–$41,390.",
         "Positioning: Emory is priced like a mid-tier private — well above a public flagship, well below the urban elites. That's a deliberate pricing decision, not an accident.",
         "Peer numbers are published sticker rates; Emory numbers use the pricing filter so the comparison is like-for-like against a typical full-load student.",
     ]),
    ("Slide 9 — Data & Methodology",
     [
         "Walk through the pipeline if asked: institutional Excel exports → decode term codes → filter to LGS programs → keep institutional-only scholarships → collapse to student-program level.",
         "The sub-$20k pricing filter: 11 rows under $20,000 in total billed tuition were excluded from pricing metrics only. These are partial-term billings with ~$0 kept scholarship that distort averages.",
         "Enrollment counts, international mix, and the per-year trend charts continue to run on the full 493 rows.",
         "Scholarship filter strips federal/external awards (Pell, GRFP, Yellow Ribbon, etc.) so the discount rate reflects institutional aid only.",
     ]),
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


def build_three_stack_slide(prs, headline, subtext, chart_paths, page_num):
    """Slide with a left navy panel (headline + subtext) and three charts
    stacked vertically on the right, each taking roughly equal vertical space."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    sw, sh = prs.slide_width, prs.slide_height

    # Left navy panel (full height) + gold vertical accent.
    panel_w = Inches(4.5)
    add_rect(slide, 0, 0, panel_w, sh, NAVY)
    add_rect(slide, panel_w, 0, Inches(0.06), sh, GOLD)

    # Headline + subtext inside the panel.
    add_text(
        slide, Inches(0.4), Inches(1.0), panel_w - Inches(0.7), Inches(4.0),
        headline, color=WHITE, size=20, bold=True, align=PP_ALIGN.LEFT,
    )
    add_text(
        slide, Inches(0.4), Inches(5.4), panel_w - Inches(0.7), Inches(1.6),
        subtext, color=GOLD, size=13, align=PP_ALIGN.LEFT,
    )

    # Right-side stack: equal vertical thirds.
    right_x = panel_w + Inches(0.25)
    right_w = sw - right_x - Inches(0.3)
    top_pad = Inches(0.3)
    bot_pad = Inches(0.5)  # leaves room for page number
    avail_h = sh - top_pad - bot_pad
    gap = Inches(0.08)
    chart_h = (avail_h - gap * 2) / 3

    for i, path in enumerate(chart_paths):
        slide.shapes.add_picture(
            path,
            left=right_x,
            top=top_pad + (chart_h + gap) * i,
            width=right_w,
            height=chart_h,
        )

    add_text(
        slide, sw - Inches(0.8), sh - Inches(0.35), Inches(0.6), Inches(0.3),
        str(page_num), color=NAVY, size=11, align=PP_ALIGN.RIGHT,
    )


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
    for i, entry in enumerate(CHART_SLIDES, start=3):
        if entry is None:
            build_three_stack_slide(
                prs,
                SEMESTER_SLIDE["headline"],
                SEMESTER_SLIDE["subtext"],
                SEMESTER_SLIDE["charts"],
                i,
            )
        else:
            insight, chart = entry
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


def build_speaker_notes_docx(path: str) -> None:
    """Per-slide talking points for the v2 deck."""
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = DocxPt(11)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Speaker Notes — LGS Master's Programs Pricing Analysis")
    r.font.name = "Calibri"
    r.font.size = DocxPt(22)
    r.font.bold = True
    r.font.color.rgb = DOCX_NAVY

    add_gold_rule(doc)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(f"{SUBTITLE} · {REPORT_DATE}")
    r.font.name = "Calibri"
    r.font.size = DocxPt(13)
    r.font.color.rgb = DOCX_GOLD

    doc.add_paragraph()
    add_body(
        doc,
        "One section per deck slide. v2 updates reflect the sub-$20k pricing "
        "filter applied to df_pricing: 11 partial-term rows excluded from "
        "discount-rate, gross-vs-net, and peer-benchmark metrics only. "
        "Full population retained for enrollment counts and international mix.",
    )

    for slide_title, bullets in SPEAKER_NOTES:
        add_heading(doc, slide_title, size=15)
        for b in bullets:
            add_bullet(doc, b)
        doc.add_paragraph()

    doc.save(path)
    print(f"  wrote {path}")


def main() -> None:
    os.makedirs("deliverables", exist_ok=True)
    build_pptx("deliverables/emory_grad_pricing_v2.pptx")
    build_docx("deliverables/emory_grad_pricing_report_v2.docx")
    build_speaker_notes_docx("deliverables/emory_grad_pricing_speaker_notes_v2.docx")


if __name__ == "__main__":
    main()
