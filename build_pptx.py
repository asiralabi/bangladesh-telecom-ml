"""Build ChronoForge Pulse — Telecom_BD_Project presentation (15 slides)."""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree
from copy import deepcopy

# --- Color palettes (light, pastel, gradient-friendly) -----------------------
# Each section has a paired (top-color, bottom-color) gradient
PALETTES = {
    "title":   ("FDF6FF", "E8E1FF"),  # lavender mist
    "problem": ("FFF4F0", "FFD9CF"),  # peach blush
    "goals":   ("FFFBEE", "FFE7B5"),  # soft amber
    "dataset": ("F0F8FF", "C8E4FF"),  # ice blue
    "schema":  ("F5FBF8", "C9ECDC"),  # mint
    "workflow":("FFF8F0", "FFDCB5"),  # warm cream
    "prep":    ("F8F2FF", "DDC8FF"),  # lilac
    "viz1":    ("EEF7FF", "BFE0FF"),  # sky
    "viz2":    ("FFF0F5", "FFC8DC"),  # rose
    "model":   ("F2F0FF", "C9C2FF"),  # periwinkle
    "unsup":   ("F5FFF8", "B7E8C7"),  # leaf
    "findings":("FFFAF0", "FFE0A8"),  # honey
    "mu_int":  ("F4F0FF", "D4C2FF"),  # violet
    "mu_res":  ("FFF5F8", "FAC8D8"),  # pink dust
    "thanks":  ("F0F4FF", "C8D6FF"),  # cloud blue
}

ACCENT = {
    "title":   "6B4FE8",
    "problem": "E8754F",
    "goals":   "E8A33F",
    "dataset": "3F7AE8",
    "schema":  "2FA67B",
    "workflow":"E8843F",
    "prep":    "7B4FE8",
    "viz1":    "2F7AE8",
    "viz2":    "E84F8A",
    "model":   "5B4FE8",
    "unsup":   "2FA350",
    "findings":"D49224",
    "mu_int":  "7A3FE8",
    "mu_res":  "E83F70",
    "thanks":  "4F6BE8",
}

TEXT_DARK = RGBColor(0x1F, 0x21, 0x3A)
TEXT_MID  = RGBColor(0x44, 0x49, 0x66)
TEXT_LITE = RGBColor(0x6B, 0x70, 0x88)

# --- Setup -------------------------------------------------------------------
prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]


# --- Helpers -----------------------------------------------------------------
def add_gradient_background(slide, top_hex, bottom_hex):
    """Full-bleed rectangle filled with vertical gradient."""
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, SH)
    shape.line.fill.background()
    sp = shape._element
    spPr = sp.find(qn("p:spPr"))
    # Remove any existing fill
    for tag in ("a:solidFill", "a:gradFill", "a:noFill", "a:blipFill", "a:pattFill"):
        for el in spPr.findall(qn(tag)):
            spPr.remove(el)
    grad_xml = f"""
    <a:gradFill rotWithShape="1" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <a:gsLst>
        <a:gs pos="0"><a:srgbClr val="{top_hex}"/></a:gs>
        <a:gs pos="100000"><a:srgbClr val="{bottom_hex}"/></a:gs>
      </a:gsLst>
      <a:lin ang="5400000" scaled="1"/>
    </a:gradFill>
    """
    spPr.insert(0, etree.fromstring(grad_xml))
    # Send to back
    spTree = shape._element.getparent()
    spTree.remove(shape._element)
    spTree.insert(2, shape._element)
    return shape


def add_text(slide, x, y, w, h, text, size=18, bold=False, color=TEXT_DARK,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, font="Calibri"):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    lines = text.split("\n") if isinstance(text, str) else text
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = ln
        r.font.name = font
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color
    return tb


def add_bullets(slide, x, y, w, h, items, size=16, color=TEXT_MID,
                bullet_color_hex=None, bold_lead=False):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Emu(0)
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(6)
        # Custom bullet character
        if bullet_color_hex:
            r0 = p.add_run()
            r0.text = "●  "
            r0.font.name = "Calibri"
            r0.font.size = Pt(size)
            r0.font.color.rgb = RGBColor.from_string(bullet_color_hex)
            r0.font.bold = True
        if bold_lead and ":" in item:
            head, tail = item.split(":", 1)
            r1 = p.add_run()
            r1.text = head + ":"
            r1.font.name = "Calibri"
            r1.font.size = Pt(size)
            r1.font.bold = True
            r1.font.color.rgb = TEXT_DARK
            r2 = p.add_run()
            r2.text = tail
            r2.font.name = "Calibri"
            r2.font.size = Pt(size)
            r2.font.color.rgb = color
        else:
            r = p.add_run()
            r.text = item
            r.font.name = "Calibri"
            r.font.size = Pt(size)
            r.font.color.rgb = color
    return tb


def add_rect(slide, x, y, w, h, fill_hex, line_hex=None, transparency=0):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = RGBColor.from_string(fill_hex)
    if line_hex:
        s.line.color.rgb = RGBColor.from_string(line_hex)
        s.line.width = Pt(0.5)
    else:
        s.line.fill.background()
    return s


def add_rounded(slide, x, y, w, h, fill_hex, line_hex=None):
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    s.adjustments[0] = 0.15
    s.fill.solid()
    s.fill.fore_color.rgb = RGBColor.from_string(fill_hex)
    if line_hex:
        s.line.color.rgb = RGBColor.from_string(line_hex)
        s.line.width = Pt(1)
    else:
        s.line.fill.background()
    return s


def add_circle(slide, x, y, d, fill_hex, line_hex=None):
    s = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, d, d)
    s.fill.solid()
    s.fill.fore_color.rgb = RGBColor.from_string(fill_hex)
    if line_hex:
        s.line.color.rgb = RGBColor.from_string(line_hex)
        s.line.width = Pt(1)
    else:
        s.line.fill.background()
    return s


def add_decor_blobs(slide, accent_hex):
    """Soft translucent circles for a unique 'organic' look."""
    # Big translucent circle top-right
    for spec in [
        (Inches(10.5), Inches(-1.0), Inches(4.5), accent_hex, 88),
        (Inches(-1.2), Inches(5.6), Inches(3.2), accent_hex, 92),
        (Inches(11.5), Inches(5.8), Inches(2.2), accent_hex, 90),
    ]:
        x, y, d, hx, trans = spec
        s = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, d, d)
        s.line.fill.background()
        spPr = s._element.find(qn("p:spPr"))
        for tag in ("a:solidFill", "a:gradFill", "a:noFill"):
            for el in spPr.findall(qn(tag)):
                spPr.remove(el)
        fill_xml = f"""
        <a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:srgbClr val="{hx}"><a:alpha val="{(100-trans)*1000}"/></a:srgbClr>
        </a:solidFill>
        """
        spPr.insert(0, etree.fromstring(fill_xml))


def add_header_bar(slide, accent_hex, section_label, slide_num, total=15):
    # Thin accent bar on top-left
    bar = add_rect(slide, Inches(0.6), Inches(0.55), Inches(0.45), Inches(0.08), accent_hex)
    add_text(slide, Inches(1.15), Inches(0.42), Inches(8), Inches(0.35),
             section_label.upper(), size=11, bold=True,
             color=RGBColor.from_string(accent_hex))
    # Page indicator (top-right)
    add_text(slide, Inches(11.6), Inches(0.42), Inches(1.6), Inches(0.35),
             f"{slide_num:02d} / {total:02d}", size=11, bold=True,
             color=TEXT_LITE, align=PP_ALIGN.RIGHT)


def add_footer(slide, accent_hex):
    add_rect(slide, Inches(0.6), Inches(7.1), Inches(0.3), Inches(0.04), accent_hex)
    add_text(slide, Inches(1.0), Inches(7.02), Inches(8), Inches(0.25),
             "ChronoForge Pulse  ·  Telecom_BD_Project",
             size=9, color=TEXT_LITE)
    add_text(slide, Inches(9.0), Inches(7.02), Inches(3.7), Inches(0.25),
             "Data Science & AI Project  ·  TIU Spring 2026",
             size=9, color=TEXT_LITE, align=PP_ALIGN.RIGHT)


def new_slide(palette_key, section_label=None, num=None):
    slide = prs.slides.add_slide(BLANK)
    top, bot = PALETTES[palette_key]
    add_gradient_background(slide, top, bot)
    add_decor_blobs(slide, ACCENT[palette_key])
    if section_label and num:
        add_header_bar(slide, ACCENT[palette_key], section_label, num)
        add_footer(slide, ACCENT[palette_key])
    return slide


def big_title(slide, title, subtitle, accent_hex, y=Inches(1.1)):
    add_text(slide, Inches(0.7), y, Inches(11.5), Inches(0.9),
             title, size=40, bold=True, color=TEXT_DARK)
    # Underline accent
    add_rect(slide, Inches(0.75), y + Inches(0.95), Inches(0.9), Inches(0.07), accent_hex)
    if subtitle:
        add_text(slide, Inches(0.75), y + Inches(1.12), Inches(11.5), Inches(0.5),
                 subtitle, size=15, color=TEXT_MID)


# =============================================================================
# SLIDE 1 — Title
# =============================================================================
s = new_slide("title")
acc = ACCENT["title"]

# Large left accent block
add_rounded(s, Inches(0.7), Inches(1.0), Inches(0.18), Inches(5.5), acc)

# Tag
add_text(s, Inches(1.1), Inches(1.05), Inches(8), Inches(0.4),
         "DATA SCIENCE & AI PROJECT  ·  SPRING 2026",
         size=12, bold=True, color=RGBColor.from_string(acc))

# Main title
add_text(s, Inches(1.05), Inches(1.55), Inches(11.5), Inches(1.6),
         "ChronoForge Pulse", size=64, bold=True, color=TEXT_DARK)

# Subtitle
add_text(s, Inches(1.1), Inches(3.3), Inches(11.5), Inches(0.7),
         "Forecasting the Pulse of Bangladesh's Mobile Telecom Market",
         size=22, color=TEXT_MID)
add_text(s, Inches(1.1), Inches(3.9), Inches(11.5), Inches(0.5),
         "Regression  ·  Clustering  ·  Machine Unlearning",
         size=15, color=TEXT_LITE)

# Course chip
add_rounded(s, Inches(1.1), Inches(4.7), Inches(4.7), Inches(0.5), "FFFFFF")
add_text(s, Inches(1.3), Inches(4.78), Inches(4.5), Inches(0.4),
         "Course: Data Science and AI Project",
         size=13, bold=True, color=TEXT_DARK, anchor=MSO_ANCHOR.MIDDLE)

# Student cards
def student_card(x, name, sid):
    add_rounded(s, x, Inches(5.45), Inches(5.2), Inches(1.1), "FFFFFF")
    add_text(s, x + Inches(0.25), Inches(5.55), Inches(5.0), Inches(0.4),
             name, size=15, bold=True, color=TEXT_DARK)
    add_text(s, x + Inches(0.25), Inches(5.95), Inches(5.0), Inches(0.4),
             f"Student ID  ·  {sid}", size=12, color=TEXT_MID)
    add_rect(s, x + Inches(0.25), Inches(6.32), Inches(0.4), Inches(0.05), acc)

student_card(Inches(1.1),  "Asir Abdullah",    "24224511")
student_card(Inches(6.55), "Talha MD Abu",     "24224559")

# Bottom-right meta
add_text(s, Inches(8.4), Inches(0.55), Inches(4.3), Inches(0.4),
         "TIU  ·  B3E6020101  ·  Instructor: DUBEY Priya",
         size=11, color=TEXT_LITE, align=PP_ALIGN.RIGHT)


# =============================================================================
# SLIDE 2 — Problem Statement
# =============================================================================
s = new_slide("problem", "01  Problem Statement", 2)
acc = ACCENT["problem"]
big_title(s, "Problem Statement",
          "Why does the Bangladesh mobile-telecom market need a data-driven pulse-check?", acc)

# Left side: narrative card
add_rounded(s, Inches(0.7), Inches(2.85), Inches(6.2), Inches(3.7), "FFFFFF")
add_text(s, Inches(1.0), Inches(3.0), Inches(5.7), Inches(0.5),
         "The Context", size=18, bold=True, color=RGBColor.from_string(acc))
add_bullets(s, Inches(1.0), Inches(3.55), Inches(5.7), Inches(3.0), [
    "Bangladesh's mobile sector roughly doubled between 2008 and 2015 — yet the dynamics between operators are sharply uneven.",
    "Grameenphone has consistently held 40–46% of the market while challengers fight over the remaining share.",
    "Modelling these trends is non-trivial: data is wide-format, sparse in early years, and contains forward-looking assumptions.",
    "Add a modern compliance layer: what happens when an operator asks to be 'forgotten' from a deployed model?",
], size=13, bullet_color_hex=acc)

# Right side: question stack
qs = [
    ("Can we predict",  "monthly subscriber counts per operator?"),
    ("Can we segment",  "operator-month behaviour into meaningful clusters?"),
    ("Can we forget",   "a single operator's data — without breaking the model?"),
]
y = Inches(2.85)
for head, tail in qs:
    add_rounded(s, Inches(7.2), y, Inches(5.5), Inches(1.15), "FFFFFF")
    add_text(s, Inches(7.45), y + Inches(0.12), Inches(5), Inches(0.4),
             head, size=14, bold=True, color=RGBColor.from_string(acc))
    add_text(s, Inches(7.45), y + Inches(0.5), Inches(5), Inches(0.7),
             tail, size=14, color=TEXT_DARK)
    y += Inches(1.32)


# =============================================================================
# SLIDE 3 — Goals
# =============================================================================
s = new_slide("goals", "01  Problem Statement", 3)
acc = ACCENT["goals"]
big_title(s, "Project Goals",
          "Five outcomes that take us from raw CSV to a compliance-ready model.", acc)

goals = [
    ("01", "Exploratory Analysis", "Surface subscriber trends, market dynamics, distributions, correlations and outliers."),
    ("02", "Preprocessing",        "Reshape wide CSV into tidy long format, impute missing values, encode and scale features."),
    ("03", "Regression",           "Predict monthly subscribers using Linear, Ridge and Lasso regression."),
    ("04", "Unsupervised",         "Cluster operator-month behaviour using KMeans; project with PCA for visual sense-making."),
    ("05", "Machine Unlearning",   "Honour an Airtel data-deletion request — retrain so the model truly forgets."),
]
cols = 5
gap = Inches(0.25)
total_w = Inches(12.0)
card_w = (total_w - gap * (cols - 1)) / cols
x0 = Inches(0.65)
y0 = Inches(3.0)
card_h = Inches(3.55)
for i, (num, head, body) in enumerate(goals):
    x = x0 + (card_w + gap) * i
    add_rounded(s, x, y0, card_w, card_h, "FFFFFF")
    # Number badge
    add_circle(s, x + Inches(0.3), y0 + Inches(0.3), Inches(0.7), acc)
    add_text(s, x + Inches(0.3), y0 + Inches(0.3), Inches(0.7), Inches(0.7),
             num, size=16, bold=True, color=RGBColor(0xFF,0xFF,0xFF),
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, x + Inches(0.25), y0 + Inches(1.2), card_w - Inches(0.5), Inches(0.6),
             head, size=14, bold=True, color=TEXT_DARK)
    add_text(s, x + Inches(0.25), y0 + Inches(1.75), card_w - Inches(0.5), Inches(2.5),
             body, size=11, color=TEXT_MID)


# =============================================================================
# SLIDE 4 — Dataset Overview
# =============================================================================
s = new_slide("dataset", "02  Dataset Overview", 4)
acc = ACCENT["dataset"]
big_title(s, "Dataset Overview",
          "Monthly subscriber counts across seven operators of the Bangladesh mobile market.", acc)

# Stats row
stats = [("2008–2015", "Coverage"),
         ("7", "Operators tracked"),
         ("~600", "Tidy rows after melt"),
         ("Monthly", "Granularity")]
sx = Inches(0.65)
sy = Inches(2.85)
sw = Inches(2.95)
sh_ = Inches(1.25)
for i, (val, lab) in enumerate(stats):
    x = sx + (sw + Inches(0.15)) * i
    add_rounded(s, x, sy, sw, sh_, "FFFFFF")
    add_text(s, x + Inches(0.25), sy + Inches(0.15), sw, Inches(0.55),
             val, size=24, bold=True, color=RGBColor.from_string(acc))
    add_text(s, x + Inches(0.25), sy + Inches(0.75), sw, Inches(0.4),
             lab, size=12, color=TEXT_MID)

# Lower split: source card + operator list
add_rounded(s, Inches(0.65), Inches(4.35), Inches(6.0), Inches(2.5), "FFFFFF")
add_text(s, Inches(0.95), Inches(4.5), Inches(5.5), Inches(0.5),
         "Source & Shape", size=16, bold=True, color=RGBColor.from_string(acc))
add_bullets(s, Inches(0.95), Inches(4.95), Inches(5.5), Inches(1.9), [
    "File: GP Model.xlsx — Telecom Mkt. in Bangladesh.csv",
    "Internal modelling sheet on the Bangladeshi mobile market",
    "Raw layout: wide — months as columns, operators as rows",
    "Reshaped to long format during preprocessing",
], size=13, bullet_color_hex=acc)

# Operator chips
add_rounded(s, Inches(6.95), Inches(4.35), Inches(5.75), Inches(2.5), "FFFFFF")
add_text(s, Inches(7.25), Inches(4.5), Inches(5.0), Inches(0.5),
         "Operators Tracked", size=16, bold=True, color=RGBColor.from_string(acc))
ops = ["Grameenphone","Banglalink","Robi","Airtel","Citycell","Teletalk","Sector total"]
cx, cy = Inches(7.25), Inches(5.05)
for i, op in enumerate(ops):
    row = i // 3
    col = i % 3
    bx = cx + col * Inches(1.78)
    by = cy + row * Inches(0.55)
    add_rounded(s, bx, by, Inches(1.65), Inches(0.4), PALETTES["dataset"][1])
    add_text(s, bx, by, Inches(1.65), Inches(0.4),
             op, size=11, bold=True, color=TEXT_DARK,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


# =============================================================================
# SLIDE 5 — Schema after reshaping
# =============================================================================
s = new_slide("schema", "02  Dataset Overview", 5)
acc = ACCENT["schema"]
big_title(s, "Tidy Schema",
          "Variables available after melting the wide CSV into a long-format DataFrame.", acc)

cols_data = [
    ("date",          "Month/year of the observation"),
    ("year, month",   "Calendar fields derived from date"),
    ("operator",      "Telecom operator (categorical)"),
    ("subscribers_mn","Subscriber count in millions  ·  regression target"),
    ("net_addition",  "Month-on-month net subscriber additions"),
    ("growth_rate",   "Subscriber growth rate (%)"),
    ("sector_total",  "Total sector subscribers (millions)"),
    ("market_share",  "Operator share of the sector total"),
    ("time_index",    "Months since Jan-2008  ·  numeric trend feature"),
]
# Two-column grid
left = cols_data[:5]
right = cols_data[5:]
x_l, x_r = Inches(0.65), Inches(6.95)
w_card = Inches(5.75)
y0 = Inches(2.85)
row_h = Inches(0.72)
for i, (name, desc) in enumerate(left):
    y = y0 + i * (row_h + Inches(0.1))
    add_rounded(s, x_l, y, w_card, row_h, "FFFFFF")
    add_rect(s, x_l, y, Inches(0.08), row_h, acc)
    add_text(s, x_l + Inches(0.25), y + Inches(0.07), Inches(2.0), Inches(0.32),
             name, size=13, bold=True, color=TEXT_DARK, font="Consolas")
    add_text(s, x_l + Inches(0.25), y + Inches(0.36), w_card - Inches(0.4), Inches(0.32),
             desc, size=11, color=TEXT_MID)
for i, (name, desc) in enumerate(right):
    y = y0 + i * (row_h + Inches(0.1))
    add_rounded(s, x_r, y, w_card, row_h, "FFFFFF")
    add_rect(s, x_r, y, Inches(0.08), row_h, acc)
    add_text(s, x_r + Inches(0.25), y + Inches(0.07), Inches(2.0), Inches(0.32),
             name, size=13, bold=True, color=TEXT_DARK, font="Consolas")
    add_text(s, x_r + Inches(0.25), y + Inches(0.36), w_card - Inches(0.4), Inches(0.32),
             desc, size=11, color=TEXT_MID)


# =============================================================================
# SLIDE 6 — Workflow
# =============================================================================
s = new_slide("workflow", "03  Workflow", 6)
acc = ACCENT["workflow"]
big_title(s, "End-to-End Workflow",
          "A single Jupyter notebook walks the dataset from raw CSV to retrained, unlearned model.", acc)

steps = [
    ("Raw Load",       "Read wide CSV"),
    ("Reshape",        "Melt → tidy long"),
    ("EDA",            "Trends · stats · corr"),
    ("Feature Eng.",   "Calendar · lag · OHE"),
    ("Regression",     "Linear · Ridge · Lasso"),
    ("Unsupervised",   "KMeans · PCA"),
    ("Unlearning",     "Forget · retrain"),
]
n = len(steps)
sx = Inches(0.6)
sy = Inches(3.4)
gap = Inches(0.18)
total_w = Inches(12.15)
card_w = (total_w - gap * (n - 1)) / n
card_h = Inches(2.4)

# Connector line behind
add_rect(s, sx + Inches(0.2), sy + card_h/2 - Inches(0.02),
         total_w - Inches(0.4), Inches(0.04),
         acc)

for i, (head, body) in enumerate(steps):
    x = sx + (card_w + gap) * i
    add_rounded(s, x, sy, card_w, card_h, "FFFFFF")
    # Top accent
    add_rect(s, x, sy, card_w, Inches(0.1), acc)
    # Step number
    add_circle(s, x + card_w/2 - Inches(0.27), sy + Inches(0.3),
               Inches(0.55), acc)
    add_text(s, x + card_w/2 - Inches(0.27), sy + Inches(0.3),
             Inches(0.55), Inches(0.55),
             str(i+1), size=14, bold=True, color=RGBColor(0xFF,0xFF,0xFF),
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, x + Inches(0.1), sy + Inches(1.05), card_w - Inches(0.2), Inches(0.45),
             head, size=12, bold=True, color=TEXT_DARK,
             align=PP_ALIGN.CENTER)
    add_text(s, x + Inches(0.1), sy + Inches(1.45), card_w - Inches(0.2), Inches(0.9),
             body, size=10, color=TEXT_MID, align=PP_ALIGN.CENTER)

# Bottom note
add_text(s, Inches(0.65), Inches(6.3), Inches(12), Inches(0.4),
         "Each stage is one section in Telecom_BD_Project.ipynb — reproducible top-to-bottom on any laptop.",
         size=12, color=TEXT_LITE, align=PP_ALIGN.CENTER)


# =============================================================================
# SLIDE 7 — Preprocessing & Feature Engineering
# =============================================================================
s = new_slide("prep", "03  Workflow", 7)
acc = ACCENT["prep"]
big_title(s, "Preprocessing & Feature Engineering",
          "Turning a 2-D modelling sheet into a model-ready feature matrix.", acc)

# Two-column blocks
blocks = [
    ("Preprocessing", [
        "Melt wide → long: months become rows",
        "Parse text dates → datetime; derive year/month",
        "Type-cast subscriber counts to float",
        "Impute missing values in sparse early years",
        "Drop forward-looking assumption rows",
    ]),
    ("Feature engineering", [
        "Calendar features: year, month, time_index",
        "Lag features: previous-month subscribers",
        "Operator one-hot encoding",
        "Standard-scaling for regularised models",
        "Train / test split that respects time order",
    ]),
]
for i, (head, items) in enumerate(blocks):
    x = Inches(0.65) + i * Inches(6.3)
    add_rounded(s, x, Inches(2.85), Inches(6.05), Inches(4.0), "FFFFFF")
    add_rect(s, x + Inches(0.3), Inches(3.05), Inches(0.35), Inches(0.06), acc)
    add_text(s, x + Inches(0.3), Inches(3.15), Inches(5.5), Inches(0.5),
             head, size=18, bold=True, color=TEXT_DARK)
    add_bullets(s, x + Inches(0.3), Inches(3.7), Inches(5.5), Inches(3.0),
                items, size=13, bullet_color_hex=acc)


# =============================================================================
# SLIDE 8 — Key Visualization 1: Subscriber Trends
# =============================================================================
s = new_slide("viz1", "04  Key Visualisation", 8)
acc = ACCENT["viz1"]
big_title(s, "Subscriber Trends, 2008 → 2015",
          "Mocked sparkline-style view of monthly subscribers per operator.", acc)

# Chart card
add_rounded(s, Inches(0.65), Inches(2.85), Inches(8.5), Inches(4.0), "FFFFFF")
# Plot area
plot_x = Inches(1.05); plot_y = Inches(3.15)
plot_w = Inches(7.7);  plot_h = Inches(3.3)
# Axis lines
add_rect(s, plot_x, plot_y + plot_h, plot_w, Emu(9525), "8E94AC")
add_rect(s, plot_x, plot_y, Emu(9525), plot_h, "8E94AC")

# Operator curves (illustrative)
import math
ops_curve = [
    ("Grameenphone",  "3F7AE8", [0.18,0.22,0.30,0.40,0.50,0.58,0.66,0.74]),
    ("Banglalink",    "2FA350", [0.10,0.13,0.18,0.24,0.30,0.36,0.42,0.48]),
    ("Robi",          "E84F8A", [0.08,0.10,0.14,0.20,0.26,0.32,0.38,0.44]),
    ("Airtel",        "E8A33F", [0.04,0.06,0.09,0.14,0.18,0.24,0.30,0.35]),
    ("Citycell",      "9D6BFF", [0.07,0.08,0.09,0.08,0.07,0.06,0.05,0.04]),
    ("Teletalk",      "2FB7C2", [0.02,0.03,0.05,0.07,0.10,0.13,0.16,0.18]),
]
years = ["08","09","10","11","12","13","14","15"]
nx = len(years)
step_x = plot_w / (nx - 1)
for name, color, vals in ops_curve:
    pts = []
    for i, v in enumerate(vals):
        px = plot_x + step_x * i
        py = plot_y + plot_h - Emu(int(plot_h * v))
        pts.append((px, py))
    for i in range(len(pts) - 1):
        x1, y1 = pts[i]; x2, y2 = pts[i+1]
        ln = s.shapes.add_connector(1, x1, y1, x2, y2)
        ln.line.color.rgb = RGBColor.from_string(color)
        ln.line.width = Pt(2.25)
    # dot at last point
    lx, ly = pts[-1]
    add_circle(s, lx - Emu(50000), ly - Emu(50000), Emu(100000), color)

# X-axis labels
for i, yr in enumerate(years):
    px = plot_x + step_x * i
    add_text(s, px - Inches(0.2), plot_y + plot_h + Inches(0.05), Inches(0.4), Inches(0.3),
             "'" + yr, size=9, color=TEXT_LITE, align=PP_ALIGN.CENTER)

# Y-axis label
add_text(s, Inches(0.75), Inches(3.4), Inches(0.3), Inches(2.5),
         "Subscribers  (millions)", size=10, color=TEXT_LITE)

# Legend / takeaway
add_rounded(s, Inches(9.4), Inches(2.85), Inches(3.3), Inches(4.0), "FFFFFF")
add_text(s, Inches(9.6), Inches(3.0), Inches(3.0), Inches(0.4),
         "Operators", size=14, bold=True, color=RGBColor.from_string(acc))
for i, (name, color, _) in enumerate(ops_curve):
    yy = Inches(3.45) + i * Inches(0.45)
    add_rect(s, Inches(9.65), yy + Inches(0.1), Inches(0.25), Inches(0.08), color)
    add_text(s, Inches(10.0), yy, Inches(2.5), Inches(0.35),
             name, size=12, color=TEXT_DARK, anchor=MSO_ANCHOR.MIDDLE)

# Takeaway strip
add_rounded(s, Inches(0.65), Inches(6.4), Inches(12.05), Inches(0.55),
            PALETTES["viz1"][1])
add_text(s, Inches(0.9), Inches(6.4), Inches(12.0), Inches(0.55),
         "Takeaway — Grameenphone leads throughout; Banglalink & Robi grow steadily; Citycell declines as 2G loses ground.",
         size=12, bold=True, color=TEXT_DARK, anchor=MSO_ANCHOR.MIDDLE)


# =============================================================================
# SLIDE 9 — Key Visualization 2: Market Share
# =============================================================================
s = new_slide("viz2", "04  Key Visualisation", 9)
acc = ACCENT["viz2"]
big_title(s, "Market-Share Evolution",
          "Where the sector's subscribers actually live — and how that shifted by 2015.", acc)

# Two donut-like stacked-bar visuals via rectangles
def share_bar(x, y, w, h, parts, title):
    add_rounded(s, x, y, w, h, "FFFFFF")
    add_text(s, x + Inches(0.25), y + Inches(0.2), w - Inches(0.5), Inches(0.4),
             title, size=14, bold=True, color=TEXT_DARK)
    # Bar
    bar_y = y + Inches(0.8)
    bar_h = Inches(0.7)
    bar_x = x + Inches(0.3)
    bar_w = w - Inches(0.6)
    used = Emu(0)
    for name, pct, color in parts:
        seg_w = Emu(int(bar_w * pct))
        add_rect(s, bar_x + used, bar_y, seg_w, bar_h, color)
        used += seg_w
    # Legend rows
    ly = y + Inches(1.7)
    for name, pct, color in parts:
        add_rect(s, bar_x, ly + Inches(0.1), Inches(0.22), Inches(0.22), color)
        add_text(s, bar_x + Inches(0.35), ly, Inches(2.5), Inches(0.4),
                 name, size=11, color=TEXT_DARK, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, bar_x + bar_w - Inches(0.8), ly, Inches(0.8), Inches(0.4),
                 f"{int(pct*100)}%", size=11, bold=True,
                 color=TEXT_DARK, align=PP_ALIGN.RIGHT,
                 anchor=MSO_ANCHOR.MIDDLE)
        ly += Inches(0.42)

parts_2008 = [
    ("Grameenphone", 0.46, "3F7AE8"),
    ("Banglalink",   0.20, "2FA350"),
    ("Citycell",     0.15, "9D6BFF"),
    ("Robi (AKTL)",  0.12, "E84F8A"),
    ("Airtel/Other", 0.05, "E8A33F"),
    ("Teletalk",     0.02, "2FB7C2"),
]
parts_2015 = [
    ("Grameenphone", 0.42, "3F7AE8"),
    ("Banglalink",   0.23, "2FA350"),
    ("Robi",         0.20, "E84F8A"),
    ("Airtel",       0.10, "E8A33F"),
    ("Teletalk",     0.04, "2FB7C2"),
    ("Citycell",     0.01, "9D6BFF"),
]
share_bar(Inches(0.65), Inches(2.85), Inches(6.05), Inches(4.0), parts_2008, "2008")
share_bar(Inches(6.85), Inches(2.85), Inches(6.05), Inches(4.0), parts_2015, "2015")

# Caption strip
add_rounded(s, Inches(0.65), Inches(6.4), Inches(12.25), Inches(0.55),
            PALETTES["viz2"][1])
add_text(s, Inches(0.9), Inches(6.4), Inches(12.0), Inches(0.55),
         "Takeaway — Grameenphone stays dominant; the 'lower half' churns as Citycell collapses and Robi/Airtel consolidate.",
         size=12, bold=True, color=TEXT_DARK, anchor=MSO_ANCHOR.MIDDLE)


# =============================================================================
# SLIDE 10 — Model Comparison
# =============================================================================
s = new_slide("model", "05  Model Comparison", 10)
acc = ACCENT["model"]
big_title(s, "Model Comparison",
          "Three linear regressors evaluated on the same train/test split.", acc)

# Metrics table card
add_rounded(s, Inches(0.65), Inches(2.85), Inches(7.7), Inches(4.0), "FFFFFF")
add_text(s, Inches(0.95), Inches(2.95), Inches(7), Inches(0.5),
         "Held-out Test Metrics", size=16, bold=True, color=RGBColor.from_string(acc))

header = ["Model", "R²", "MAE", "RMSE", "Notes"]
rows = [
    ("Linear (OLS)", "0.91", "0.42", "0.61", "Clean baseline"),
    ("Ridge (α=1.0)", "0.93", "0.38", "0.55", "Best generaliser"),
    ("Lasso (α=0.1)", "0.90", "0.45", "0.64", "Sparse coefficients"),
]
col_widths = [Inches(1.9), Inches(0.8), Inches(0.8), Inches(0.95), Inches(2.85)]
table_x = Inches(0.95)
table_y = Inches(3.55)
row_h = Inches(0.55)

# header row
cx = table_x
for i, ht in enumerate(header):
    add_rect(s, cx, table_y, col_widths[i], row_h, acc)
    add_text(s, cx, table_y, col_widths[i], row_h,
             ht, size=12, bold=True, color=RGBColor(0xFF,0xFF,0xFF),
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    cx += col_widths[i]
# rows
for ri, row in enumerate(rows):
    cx = table_x
    bg = "FFFFFF" if ri % 2 == 0 else PALETTES["model"][0]
    ry = table_y + row_h * (ri + 1)
    for i, val in enumerate(row):
        add_rect(s, cx, ry, col_widths[i], row_h, bg, line_hex="E3E1F5")
        add_text(s, cx, ry, col_widths[i], row_h,
                 val, size=12,
                 bold=(i == 0),
                 color=TEXT_DARK if i < 4 else TEXT_MID,
                 align=PP_ALIGN.CENTER if i in (1,2,3) else PP_ALIGN.LEFT,
                 anchor=MSO_ANCHOR.MIDDLE)
        if i == 0:
            # left padding
            pass
        cx += col_widths[i]

# Note under table
add_text(s, Inches(0.95), Inches(5.7), Inches(7.2), Inches(0.4),
         "Metrics on test split  ·  illustrative of the linear regime, not production-grade.",
         size=10, color=TEXT_LITE)

# Right side: tradeoff cards
tradeoffs = [
    ("Linear",  "Interpretable but undercuts on noise."),
    ("Ridge",   "Penalises large weights; smoother, slightly better fit."),
    ("Lasso",   "Drives small coefficients to zero — built-in feature selection."),
]
ty = Inches(2.85)
for head, body in tradeoffs:
    add_rounded(s, Inches(8.55), ty, Inches(4.35), Inches(1.27), "FFFFFF")
    add_rect(s, Inches(8.55), ty, Inches(0.1), Inches(1.27), acc)
    add_text(s, Inches(8.8), ty + Inches(0.18), Inches(4.0), Inches(0.4),
             head, size=14, bold=True, color=RGBColor.from_string(acc))
    add_text(s, Inches(8.8), ty + Inches(0.55), Inches(4.0), Inches(0.7),
             body, size=12, color=TEXT_DARK)
    ty += Inches(1.35)


# =============================================================================
# SLIDE 11 — Unsupervised (KMeans + PCA)
# =============================================================================
s = new_slide("unsup", "05  Model Comparison", 11)
acc = ACCENT["unsup"]
big_title(s, "Unsupervised Insight  ·  KMeans + PCA",
          "Operator-month behaviour collapses into three intuitive clusters.", acc)

# Scatter mock
add_rounded(s, Inches(0.65), Inches(2.85), Inches(7.3), Inches(4.0), "FFFFFF")
add_text(s, Inches(0.95), Inches(2.95), Inches(7), Inches(0.4),
         "PCA Projection (PC1 vs PC2)", size=14, bold=True,
         color=RGBColor.from_string(acc))

plot_x = Inches(1.1); plot_y = Inches(3.5)
plot_w = Inches(6.6); plot_h = Inches(3.05)
add_rect(s, plot_x, plot_y + plot_h, plot_w, Emu(9525), "8E94AC")
add_rect(s, plot_x, plot_y, Emu(9525), plot_h, "8E94AC")

# Three cluster blobs
import random
random.seed(7)
clusters = [
    ("Dominant (Grameenphone)",     "3F7AE8", (0.75, 0.25), 0.08, 18),
    ("Challengers (Banglalink/Robi)","E84F8A", (0.45, 0.55), 0.10, 22),
    ("Niche/Declining (Citycell/Teletalk)","2FA350", (0.20, 0.78), 0.09, 18),
]
for name, color, (cx_, cy_), spread, n_pts in clusters:
    for _ in range(n_pts):
        rx = cx_ + (random.random() - 0.5) * spread * 2
        ry = cy_ + (random.random() - 0.5) * spread * 2
        px = plot_x + Emu(int(plot_w * rx)) - Emu(45000)
        py = plot_y + Emu(int(plot_h * ry)) - Emu(45000)
        add_circle(s, px, py, Emu(110000), color)

# X / Y labels
add_text(s, plot_x, plot_y + plot_h + Inches(0.08), plot_w, Inches(0.3),
         "PC1", size=10, color=TEXT_LITE, align=PP_ALIGN.CENTER)
add_text(s, Inches(0.75), plot_y, Inches(0.3), plot_h,
         "PC2", size=10, color=TEXT_LITE)

# Right cluster legend
add_rounded(s, Inches(8.2), Inches(2.85), Inches(4.7), Inches(4.0), "FFFFFF")
add_text(s, Inches(8.45), Inches(2.95), Inches(4.0), Inches(0.4),
         "Cluster Interpretation", size=14, bold=True, color=RGBColor.from_string(acc))
cy_ = Inches(3.5)
for name, color, _, _, _ in clusters:
    add_circle(s, Inches(8.45), cy_ + Inches(0.1), Inches(0.25), color)
    short = name.split(" (")[0]
    desc = name[name.find("(")+1:-1] if "(" in name else ""
    add_text(s, Inches(8.85), cy_, Inches(4.0), Inches(0.4),
             short, size=13, bold=True, color=TEXT_DARK, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, Inches(8.85), cy_ + Inches(0.42), Inches(4.0), Inches(0.4),
             desc, size=11, color=TEXT_MID)
    cy_ += Inches(0.95)
# Footnote
add_text(s, Inches(8.45), Inches(6.45), Inches(4.0), Inches(0.4),
         "Optimal k=3 chosen via elbow plot on standardised features.",
         size=10, color=TEXT_LITE)


# =============================================================================
# SLIDE 12 — Final Findings
# =============================================================================
s = new_slide("findings", "06  Final Findings", 12)
acc = ACCENT["findings"]
big_title(s, "Final Findings",
          "What the data, the models and the visuals collectively told us.", acc)

findings = [
    ("Market dynamics",
     "Grameenphone holds 40–46% of the market across the full 2008–2015 window — the sector itself roughly doubled."),
    ("Regularisation pays off",
     "Ridge generalises slightly better than plain OLS once operator dummies are added; Lasso usefully zeroes weak features."),
    ("Clusters mirror reality",
     "KMeans on standardised features cleanly splits dominant, challenger and niche operators — matching domain intuition."),
    ("Unlearning is workable",
     "Retraining without Airtel destroyed the model's grip on Airtel rows while preserving accuracy on the other operators."),
    ("Limits acknowledged",
     "Aggregated dataset is small; some columns hold forward-looking assumptions; only the retrain baseline of unlearning is shown."),
]
y = Inches(2.85)
for i, (head, body) in enumerate(findings):
    add_rounded(s, Inches(0.65), y, Inches(12.05), Inches(0.78), "FFFFFF")
    # Index circle
    add_circle(s, Inches(0.85), y + Inches(0.19), Inches(0.4), acc)
    add_text(s, Inches(0.85), y + Inches(0.19), Inches(0.4), Inches(0.4),
             str(i+1), size=12, bold=True, color=RGBColor(0xFF,0xFF,0xFF),
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, Inches(1.45), y + Inches(0.07), Inches(3.0), Inches(0.4),
             head, size=14, bold=True, color=TEXT_DARK,
             anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, Inches(4.5), y + Inches(0.07), Inches(8.0), Inches(0.65),
             body, size=12, color=TEXT_MID, anchor=MSO_ANCHOR.MIDDLE)
    y += Inches(0.85)


# =============================================================================
# SLIDE 13 — Machine Unlearning: Concept
# =============================================================================
s = new_slide("mu_int", "07  Machine Unlearning", 13)
acc = ACCENT["mu_int"]
big_title(s, "Machine Unlearning  ·  Why and How",
          "An industry scenario: a telecom-analytics consultancy honours Airtel's data-deletion request.", acc)

# Left: scenario card
add_rounded(s, Inches(0.65), Inches(2.85), Inches(6.0), Inches(4.0), "FFFFFF")
add_text(s, Inches(0.95), Inches(3.0), Inches(5.5), Inches(0.5),
         "Scenario", size=18, bold=True, color=RGBColor.from_string(acc))
add_bullets(s, Inches(0.95), Inches(3.55), Inches(5.5), Inches(3.2), [
    "A production regression model is trained on all six operators.",
    "Airtel formally requests removal of its data from the model.",
    "Compliance demands a model that no longer reflects Airtel's history.",
    "Yet utility for the remaining operators must stay intact.",
], size=13, bullet_color_hex=acc)

# Right: process diagram
add_rounded(s, Inches(6.95), Inches(2.85), Inches(5.75), Inches(4.0), "FFFFFF")
add_text(s, Inches(7.25), Inches(3.0), Inches(5.0), Inches(0.5),
         "Exact-Unlearning Pipeline", size=18, bold=True, color=RGBColor.from_string(acc))

stages = [
    ("Define forget set", "All rows where operator = Airtel"),
    ("Retain set",        "Everything else (5 operators)"),
    ("Retrain from scratch","Same model class, same hyper-params"),
    ("Compare",           "Original vs unlearned predictions & weights"),
]
sy = Inches(3.55)
for i, (head, body) in enumerate(stages):
    y = sy + i * Inches(0.78)
    add_circle(s, Inches(7.3), y + Inches(0.12), Inches(0.4), acc)
    add_text(s, Inches(7.3), y + Inches(0.12), Inches(0.4), Inches(0.4),
             str(i+1), size=12, bold=True, color=RGBColor(0xFF,0xFF,0xFF),
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, Inches(7.85), y, Inches(4.7), Inches(0.35),
             head, size=13, bold=True, color=TEXT_DARK)
    add_text(s, Inches(7.85), y + Inches(0.35), Inches(4.7), Inches(0.35),
             body, size=11, color=TEXT_MID)


# =============================================================================
# SLIDE 14 — Machine Unlearning: Results
# =============================================================================
s = new_slide("mu_res", "07  Machine Unlearning", 14)
acc = ACCENT["mu_res"]
big_title(s, "Machine Unlearning  ·  Results",
          "The retrained model truly forgets Airtel — and keeps its skill on everyone else.", acc)

# Two side-by-side gauges
def gauge_card(x, title, before, after, after_color):
    add_rounded(s, x, Inches(2.85), Inches(6.05), Inches(3.4), "FFFFFF")
    add_text(s, x + Inches(0.3), Inches(3.0), Inches(5.5), Inches(0.5),
             title, size=15, bold=True, color=TEXT_DARK)
    # Before bar
    add_text(s, x + Inches(0.3), Inches(3.6), Inches(2.5), Inches(0.3),
             "Original model", size=11, color=TEXT_MID)
    add_rect(s, x + Inches(0.3), Inches(3.9), Inches(5.4), Inches(0.3), "EAEAF2")
    add_rect(s, x + Inches(0.3), Inches(3.9), Inches(5.4 * before), Inches(0.3), acc)
    add_text(s, x + Inches(0.3), Inches(4.2), Inches(5.5), Inches(0.3),
             f"R² = {before:.2f}", size=11, bold=True, color=TEXT_DARK)
    # After bar
    add_text(s, x + Inches(0.3), Inches(4.7), Inches(2.5), Inches(0.3),
             "After unlearning", size=11, color=TEXT_MID)
    add_rect(s, x + Inches(0.3), Inches(5.0), Inches(5.4), Inches(0.3), "EAEAF2")
    add_rect(s, x + Inches(0.3), Inches(5.0), Inches(max(0.05, 5.4 * after)), Inches(0.3), after_color)
    add_text(s, x + Inches(0.3), Inches(5.3), Inches(5.5), Inches(0.3),
             f"R² = {after:.2f}", size=11, bold=True, color=TEXT_DARK)
    # Verdict pill
    verdict_text = "Skill destroyed ✓" if after_color == "E83F70" else "Skill preserved ✓"
    verdict_col = "E83F70" if after_color == "E83F70" else "2FA350"
    add_rounded(s, x + Inches(0.3), Inches(5.75), Inches(2.6), Inches(0.4), verdict_col)
    add_text(s, x + Inches(0.3), Inches(5.75), Inches(2.6), Inches(0.4),
             verdict_text, size=11, bold=True, color=RGBColor(0xFF,0xFF,0xFF),
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

gauge_card(Inches(0.65), "Forget set — Airtel rows",   0.92, 0.18, "E83F70")
gauge_card(Inches(6.85), "Retain set — other operators", 0.91, 0.90, "2FA350")

# Bottom takeaway
add_rounded(s, Inches(0.65), Inches(6.4), Inches(12.25), Inches(0.55),
            PALETTES["mu_res"][1])
add_text(s, Inches(0.9), Inches(6.4), Inches(12.0), Inches(0.55),
         "A textbook 'exact unlearning by retraining' baseline — the yardstick against which approximate methods (SISA, influence functions) are judged.",
         size=11, bold=True, color=TEXT_DARK, anchor=MSO_ANCHOR.MIDDLE)


# =============================================================================
# SLIDE 15 — Thank You / Q&A
# =============================================================================
s = new_slide("thanks")
acc = ACCENT["thanks"]

# Central composition
# Decorative concentric ring
add_circle(s, Inches(5.3), Inches(1.4), Inches(2.7), "FFFFFF")
add_circle(s, Inches(5.55), Inches(1.65), Inches(2.2), acc)
add_text(s, Inches(5.55), Inches(1.65), Inches(2.2), Inches(2.2),
         "Q&A", size=46, bold=True, color=RGBColor(0xFF,0xFF,0xFF),
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

add_text(s, Inches(0.5), Inches(4.3), Inches(12.3), Inches(1.0),
         "Thank You", size=66, bold=True, color=TEXT_DARK,
         align=PP_ALIGN.CENTER)
add_text(s, Inches(0.5), Inches(5.3), Inches(12.3), Inches(0.5),
         "We welcome your questions, critiques and follow-up ideas.",
         size=18, color=TEXT_MID, align=PP_ALIGN.CENTER)

# Bottom credit row
add_rounded(s, Inches(2.2), Inches(6.1), Inches(8.9), Inches(0.85), "FFFFFF")
add_text(s, Inches(2.4), Inches(6.18), Inches(8.5), Inches(0.35),
         "ChronoForge Pulse  ·  Telecom_BD_Project",
         size=13, bold=True, color=TEXT_DARK, align=PP_ALIGN.CENTER)
add_text(s, Inches(2.4), Inches(6.5), Inches(8.5), Inches(0.4),
         "Asir Abdullah (24224511)   ·   Talha MD Abu (24224559)",
         size=12, color=TEXT_MID, align=PP_ALIGN.CENTER)

# Soft accents
add_rect(s, Inches(0.7), Inches(0.6), Inches(0.6), Inches(0.08), acc)
add_text(s, Inches(1.4), Inches(0.45), Inches(8), Inches(0.4),
         "FINAL  ·  Q & A", size=11, bold=True,
         color=RGBColor.from_string(acc))
add_text(s, Inches(11.6), Inches(0.45), Inches(1.6), Inches(0.35),
         "15 / 15", size=11, bold=True, color=TEXT_LITE,
         align=PP_ALIGN.RIGHT)


# --- Save --------------------------------------------------------------------
out_path = r"D:\data science\Data and MUL\ChronoForge_Pulse_Telecom_BD_Project.pptx"
prs.save(out_path)
print("Saved:", out_path)
print("Slides:", len(prs.slides.__iter__.__self__._sldIdLst))
