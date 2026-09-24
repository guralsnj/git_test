import copy, sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.oxml.ns import qn
from lxml import etree

SK = "/root/.claude/skills/synced/5dd94d0b-87b7-4407-be58-13dace0ecad4_8e7c9b18-ec18-4b4c-8b20-f72a48e99518/steel-king-ppt-template/assets/template.pptx"
GREEN = RGBColor(0x3C, 0x7D, 0x4D); STEEL = RGBColor(0x5B, 0x66, 0x70); INK = RGBColor(0x15, 0x18, 0x1A)
HAIR = RGBColor(0xE2, 0xE6, 0xE3); TINT = RGBColor(0xF2, 0xF5, 0xF2); WHITE = RGBColor(0xFF, 0xFF, 0xFF)
L = dict(COVER_MASCOT=15, TITLE_CONTENT=4, TWO_CONTENT=5, THREE_CARDS=6)

# ---------- helpers ----------
def ph(slide, idx):
    for sh in slide.placeholders:
        if sh.placeholder_format.idx == idx:
            return sh
    raise KeyError(idx)

def fill(slide, idx, text):
    p = ph(slide, idx).text_frame.paragraphs[0]
    if p.runs: p.runs[0].text = text
    else: p.add_run().text = text

def fill_lines(slide, idx, lines, size=None):
    tf = ph(slide, idx).text_frame
    for extra in tf.paragraphs[1:]:
        extra._p.getparent().remove(extra._p)
    p0 = tf.paragraphs[0]
    if p0.runs: p0.runs[0].text = lines[0]
    else: p0.add_run().text = lines[0]
    for r in p0.runs[1:]: r._r.getparent().remove(r._r)
    for line in lines[1:]:
        p = tf.add_paragraph(); p.add_run().text = line
    if size:
        for p in tf.paragraphs:
            for r in p.runs: r.font.size = Pt(size)

def drop(slide, idx):
    sh = ph(slide, idx); sh._element.getparent().remove(sh._element)

def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text

def textbox(slide, x, y, w, h, text, size=13, bold=False, color=INK, fill_tint=False, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor; tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = Inches(0.10 if fill_tint else 0)
    tf.margin_top = tf.margin_bottom = Inches(0.06 if fill_tint else 0)
    if fill_tint:
        tb.fill.solid(); tb.fill.fore_color.rgb = TINT
    lines = text if isinstance(text, list) else [text]
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        r = p.add_run(); r.text = line
        r.font.name = "Arial"; r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = color
    return tb

def heading(slide, x, y, w, text):
    return textbox(slide, x, y, w, 0.26, text.upper(), size=11, bold=True, color=GREEN)

def _border(cell, color="E2E6E3", width=6350):
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        for old in tcPr.findall(qn(tag)): tcPr.remove(old)
        ln = etree.SubElement(tcPr, qn(tag), w=str(width), cap="flat", cmpd="sng", algn="ctr")
        sf = etree.SubElement(ln, qn("a:solidFill")); etree.SubElement(sf, qn("a:srgbClr"), val=color)
        etree.SubElement(ln, qn("a:prstDash"), val="solid")

def table(slide, x, y, w, col_w, header, rows, row_h=0.30, size=10, body_size=None, header_h=None):
    body_size = body_size or size
    header_h = header_h or row_h
    n = len(rows) + 1
    gf = slide.shapes.add_table(n, len(header), Inches(x), Inches(y), Inches(w), Inches(header_h + row_h * len(rows)))
    t = gf.table
    tblPr = t._tbl.tblPr
    tblPr.set("firstRow", "0"); tblPr.set("bandRow", "0")
    for st in tblPr.findall(qn("a:tableStyleId")): tblPr.remove(st)
    for i, cw in enumerate(col_w): t.columns[i].width = Inches(cw)
    t.rows[0].height = Inches(header_h)
    for r in range(1, n): t.rows[r].height = Inches(row_h)
    for ci, txt in enumerate(header):
        c = t.cell(0, ci); c.fill.solid(); c.fill.fore_color.rgb = GREEN
        c.margin_left = c.margin_right = Inches(0.08); c.margin_top = c.margin_bottom = Inches(0.03)
        c.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = c.text_frame.paragraphs[0]; r = p.add_run(); r.text = txt
        r.font.name = "Arial"; r.font.size = Pt(size); r.font.bold = True; r.font.color.rgb = WHITE
        _border(c, "3C7D4D")
    for ri, row in enumerate(rows, start=1):
        for ci, txt in enumerate(row):
            c = t.cell(ri, ci); c.fill.solid(); c.fill.fore_color.rgb = WHITE
            c.margin_left = c.margin_right = Inches(0.08); c.margin_top = c.margin_bottom = Inches(0.03)
            c.vertical_anchor = MSO_ANCHOR.TOP
            first = ci == 0
            p = c.text_frame.paragraphs[0]; r = p.add_run(); r.text = txt
            r.font.name = "Arial"; r.font.size = Pt(body_size); r.font.bold = first
            r.font.color.rgb = INK if first else STEEL
            _border(c)
    return t

def chart(slide, x, y, w, h):
    data = CategoryChartData()
    data.categories = ["Q4 2025", "Q1 2026", "Q2 2026", "Q3 2026"]
    data.add_series("Bookings", (850, 920, 780, 1010))
    data.add_series("Quotes", (2400, 2650, 2100, 2900))
    gf = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(x), Inches(y), Inches(w), Inches(h), data)
    ch = gf.chart; ch.has_title = False
    ch.has_legend = True; ch.legend.position = XL_LEGEND_POSITION.BOTTOM; ch.legend.include_in_layout = False
    ch.legend.font.size = Pt(10); ch.legend.font.name = "Arial"
    for s, col in zip(ch.plots[0].series, (GREEN, STEEL)):
        s.format.fill.solid(); s.format.fill.fore_color.rgb = col
    ch.plots[0].gap_width = 55; ch.plots[0].overlap = -10
    ch.value_axis.has_major_gridlines = False; ch.value_axis.visible = False
    ch.category_axis.tick_labels.font.size = Pt(10); ch.category_axis.tick_labels.font.name = "Arial"
    ch.category_axis.format.line.color.rgb = HAIR
    ch.plots[0].has_data_labels = True
    dl = ch.plots[0].data_labels; dl.font.size = Pt(9); dl.font.name = "Arial"; dl.number_format = '$#,##0"K"'; dl.number_format_is_linked = False
    return ch

def strip_watermark(prs):
    """Light layouts carry the shield watermark as a background picture. Replace with plain white."""
    for i in (3, 4, 5, 6, 7, 8, 9, 10, 11, 13):
        layout = prs.slide_layouts[i]
        cSld = layout._element.find(qn("p:cSld"))
        bg = cSld.find(qn("p:bg"))
        if bg is not None:
            cSld.remove(bg)
        new_bg = etree.SubElement(cSld, qn("p:bg")); cSld.insert(0, new_bg)
        bgPr = etree.SubElement(new_bg, qn("p:bgPr"))
        sf = etree.SubElement(bgPr, qn("a:solidFill")); etree.SubElement(sf, qn("a:srgbClr"), val="FFFFFF")
        etree.SubElement(bgPr, qn("a:effectLst"))

def new_deck():
    prs = Presentation(SK)
    strip_watermark(prs)
    while len(prs.slides) > 0:
        rId = prs.slides._sldIdLst[0].rId
        prs.part.drop_rel(rId); del prs.slides._sldIdLst[0]
    return prs

EYEBROW = "TERRITORY REVIEW"
def std_header(slide, n, eyebrow, title, takeaway):
    fill(slide, 100, f"{EYEBROW}  ·  {n} OF 6  ·  {eyebrow}".upper())
    fill(slide, 101, title)
    fill(slide, 102, takeaway)



# ---------- shared ----------
def cover(prs, eyebrow, title, line, who, note):
    s = prs.slides.add_slide(prs.slide_layouts[L["COVER_MASCOT"]])
    fill(s, 102, eyebrow); fill(s, 104, title); fill(s, 105, line); fill(s, 106, who)
    notes(s, note)
    return s

def numbers_slide(prs, label, right_label, excalibur):
    s = prs.slides.add_slide(prs.slide_layouts[L["TITLE_CONTENT"]])
    std_header(s, 1, "By the numbers", f"{label} in numbers",
               "[One line on what the numbers say.]")
    drop(s, 103)
    heading(s, 0.80, 2.32, 5.60, "Bookings and quotes by quarter  ·  pre-filled")
    table(s, 0.80, 2.60, 5.60, [1.40, 1.40, 1.40, 1.40],
          ["Quarter", "Bookings 2025", "Bookings 2026", "Quotes 2026"],
          [["Q1", "", "", ""], ["Q2", "", "", ""], ["Q3", "", "", ""], ["Q4 to date", "", "", ""], ["YTD", "", "", ""]],
          row_h=0.29, size=10, header_h=0.30)
    heading(s, 0.80, 4.80, 5.60, "What the numbers don't show  ·  two sentences, yours")
    textbox(s, 0.80, 5.08, 5.60, 1.52, "[Two sentences.]", size=12, color=STEEL, fill_tint=True)
    heading(s, 7.13, 2.32, 5.40, f"{right_label}, YTD bookings  ·  pre-filled")
    rows = 5 if excalibur else 7
    table(s, 7.13, 2.60, 5.40, [3.40, 2.00], [right_label.split(",")[0].rstrip("s") if False else right_label, "YTD bookings"],
          [[f"{i}.  ", ""] for i in range(1, rows + 1)], row_h=0.29, size=10, header_h=0.30)
    if excalibur:
        heading(s, 7.13, 4.78, 5.40, "Excalibur dealers, YTD bookings  ·  pre-filled")
        table(s, 7.13, 5.06, 5.40, [3.40, 2.00], ["Excalibur dealer", "YTD bookings"],
              [["", ""], ["", ""], ["", ""]], row_h=0.29, size=10, header_h=0.30)
    fill(s, 104, "Source: [system], pulled [date]. Quotes are a count. Pre-filled by Nate.")
    notes(s, "Slide 1. Pre-filled. Add two sentences on what the numbers don't show.")
    return s

def buy_us_for_slide(prs, who, title, note):
    s = prs.slides.add_slide(prs.slide_layouts[L["TWO_CONTENT"]])
    std_header(s, 3, "What they buy us for", title, "[One line. The thing they count on most, and the gap that costs the most.]")
    fill(s, 103, f"What {who} count on us for")
    fill_lines(s, 105, ["1. [ ]", "2. [ ]", "3. [ ]"])
    fill(s, 106, "The one gap that costs me the most")
    fill_lines(s, 108, ["The gap: [product or resource]", "What it costs: [jobs, dealers, or $]", "What would fix it: [ ]"])
    drop(s, 109)
    notes(s, note)
    return s

def wins_loss_slide(prs):
    s = prs.slides.add_slide(prs.slide_layouts[L["TITLE_CONTENT"]])
    std_header(s, 5, "Three wins, one loss", "Three wins, one loss",
               "[One line. What the wins have in common.]")
    drop(s, 103); drop(s, 104)
    table(s, 0.80, 2.32, 11.73, [1.20, 3.00, 2.20, 2.70, 2.63],
          ["", "Job: dealer, account, job name", "Competitor", "What won it", "Repeatable? How"],
          [["Win 1", "", "", "", ""], ["Win 2", "", "", "", ""], ["Win 3", "", "", "", ""],
           ["Loss", "", "", "What decided it, and what else was true", "What we could have done differently"]],
          row_h=0.97, size=11, header_h=0.40)
    notes(s, "Slide 5. Three wins, one loss. For each win: the job, the competitor, and what won it. "
             "For the loss: what decided it, and what we could have done differently. Real jobs, real names. "
             "If the loss came down to the number, say what else was true.")
    return s

def q4_slide(prs, grow_label, grow_lines):
    s = prs.slides.add_slide(prs.slide_layouts[L["THREE_CARDS"]])
    std_header(s, 6, "Q4", "Q4 plan", "[One line. What Q4 looks like if these land.]")
    fill(s, 106, "Three projects I'm going after")
    fill_lines(s, 107, ["1. [Job, $, close date]", "2. [Job, $, close date]", "3. [Job, $, close date]"])
    fill(s, 111, grow_label)
    fill_lines(s, 112, grow_lines)
    fill(s, 116, "One thing I'd change")
    fill_lines(s, 117, ["What: [about how we run sales]", "Why: [ ]", "What it would take: [ ]"])
    drop(s, 118)
    notes(s, "Slide 6. Three projects you're going after, one you're going to grow, one thing you'd change about how we run sales.")
    return s

# ---------- RSM ----------
def rsm_winning(prs):
    s = prs.slides.add_slide(prs.slide_layouts[L["TITLE_CONTENT"]])
    std_header(s, 2, "Who's winning with what", "Who's winning with what",
               "[One line. Why they pick us, in their words if you have them.]")
    drop(s, 103); drop(s, 104)
    table(s, 0.80, 2.32, 11.73, [2.50, 1.10, 3.30, 4.83],
          ["Dealer", "Excalibur", "Products they perform with", "Why they pick Steel King over the other guys"],
          [[f"{i}.  ", "", "", ""] for i in range(1, 6)], row_h=0.776, size=11, header_h=0.40)
    notes(s, "Slide 2. Your top dealers, Excalibur and non-Excalibur, the products each one performs with, "
             "and why they pick Steel King over the other guys. In their words if you have them.")
    return s

def rsm_mindshare(prs):
    s = prs.slides.add_slide(prs.slide_layouts[L["THREE_CARDS"]])
    std_header(s, 4, "Mindshare", "How I get and keep mindshare", "[One line. The one thing that moves the needle with dealers.]")
    fill(s, 106, "How I get and keep mindshare")
    fill_lines(s, 107, ["[Visits, training, joint calls, quick ship, whatever it really is]"])
    fill(s, 111, "How new dealer reps get up to speed on us")
    fill_lines(s, 112, ["[What happens today when a dealer hires a new rep]"])
    fill(s, 116, "What's working, what isn't")
    fill_lines(s, 117, ["Working: [ ]", "Not working: [ ]"])
    drop(s, 118)
    notes(s, "Slide 4. How you get and keep mindshare with your dealers. How new dealer reps get up to speed on us. What's working, what isn't.")
    return s

# ---------- Chad ----------
def chad_products(prs):
    s = prs.slides.add_slide(prs.slide_layouts[L["TITLE_CONTENT"]])
    std_header(s, 2, "What's performing", "What's performing, and why Walmart picks us",
               "[One line. Why Walmart keeps choosing us, in their words if you have them.]")
    drop(s, 103); drop(s, 104)
    table(s, 0.80, 2.32, 11.73, [2.60, 3.60, 5.53],
          ["Product", "Where it's performing: site, program, format", "Why Walmart keeps choosing us"],
          [[f"{i}.  ", "", ""] for i in range(1, 6)], row_h=0.776, size=11, header_h=0.40)
    notes(s, "Slide 2 for Chad. Which products are performing across the Walmart sites and why Walmart keeps choosing us.")
    return s

def chad_mindshare(prs):
    s = prs.slides.add_slide(prs.slide_layouts[L["THREE_CARDS"]])
    std_header(s, 4, "Mindshare", "How I keep mindshare inside Walmart", "[One line. What keeps us in the room when engineering turns over.]")
    fill(s, 106, "Inside Walmart")
    fill_lines(s, 107, ["[How you stay in front of the decision makers as their engineering staff turns over]"])
    fill(s, 111, "With their integrators")
    fill_lines(s, 112, ["[Who they are and how you stay in front of them]"])
    fill(s, 116, "What's working, what isn't")
    fill_lines(s, 117, ["Working: [ ]", "Not working: [ ]"])
    drop(s, 118)
    notes(s, "Slide 4 for Chad. How you keep mindshare inside Walmart and with their integrators.")
    return s

# ---------- Mac ----------
def mac_products(prs):
    s = prs.slides.add_slide(prs.slide_layouts[L["TITLE_CONTENT"]])
    std_header(s, 2, "What's winning", "What's winning, and where growth is",
               "[One line. The product and account pattern that's working.]")
    drop(s, 103); drop(s, 104)
    heading(s, 0.80, 2.32, 5.60, "Products winning, by account")
    table(s, 0.80, 2.60, 5.60, [1.80, 1.80, 2.00], ["Product", "Account", "Why it's winning"],
          [["1.  ", "", ""], ["2.  ", "", ""], ["3.  ", "", ""]], row_h=1.10, size=11, header_h=0.40)
    heading(s, 7.13, 2.32, 5.40, "Biggest new customer opportunities")
    table(s, 7.13, 2.60, 5.40, [1.80, 1.80, 1.80], ["Account or segment", "Why", "What it takes"],
          [["1.  ", "", ""], ["2.  ", "", ""], ["3.  ", "", ""]], row_h=1.10, size=11, header_h=0.40)
    notes(s, "Slide 2 for Mac. Which products are winning with which accounts, and where the biggest new customer opportunities are.")
    return s

def mac_mindshare(prs):
    s = prs.slides.add_slide(prs.slide_layouts[L["THREE_CARDS"]])
    std_header(s, 4, "Getting in the door", "How I'm getting in the door", "[One line. The way in that's actually producing meetings.]")
    fill(s, 106, "How I'm getting in")
    fill_lines(s, 107, ["[Referrals, shows, integrators, outreach, whatever it really is]"])
    fill(s, 111, "What's working")
    fill_lines(s, 112, ["[ ]"])
    fill(s, 116, "What isn't")
    fill_lines(s, 117, ["[ ]"])
    drop(s, 118)
    notes(s, "Slide 4 for Mac. How you're getting in the door. What's working, what isn't.")
    return s

# ---------- build ----------
def build(kind, out):
    global EYEBROW
    EYEBROW = "TERRITORY REVIEW" if kind == "rsm" else "NATIONAL ACCOUNT REVIEW"
    prs = new_deck()
    if kind == "rsm":
        cover(prs, "SALES MEETING  ·  TERRITORY PRESENTATION", "[Territory name]",
              "Where we're winning, who's winning with what, and why they choose us.",
              "[Your name]  |  Sales Meeting  |  [Month] 2026",
              "Ten minutes. Six slides. Real dealer, account, and job names. No new data pulls. Pricing gets its own block on Day 2, leave it out of this one.")
        numbers_slide(prs, "[Territory]", "Top accounts", excalibur=True)
        rsm_winning(prs)
        buy_us_for_slide(prs, "dealers", "What they buy us for, and what's missing",
                         "Slide 3. The two or three things dealers count on us for. Then the one product or resource gap that costs you the most business.")
        rsm_mindshare(prs); wins_loss_slide(prs)
        q4_slide(prs, "One dealer I'm going to grow", ["Dealer: [name]", "From: [today]", "To: [target]", "How: [the specific move]"])
    elif kind == "walmart":
        cover(prs, "SALES MEETING  ·  NATIONAL ACCOUNT PRESENTATION", "Walmart, Sam's, and Ambient",
              "Where we're winning, what's performing, and why Walmart keeps choosing us.",
              "Chad  |  Sales Meeting  |  [Month] 2026",
              "Ten minutes. Six slides. Real sites, programs, and job names. No new data pulls. Pricing gets its own block on Day 2, leave it out of this one.")
        numbers_slide(prs, "Walmart", "Top programs and sites", excalibur=False)
        chad_products(prs)
        buy_us_for_slide(prs, "Walmart and their integrators", "What Walmart buys us for, and what's missing",
                         "Slide 3. The two or three things Walmart counts on us for. Then the one product or resource gap that costs you the most business.")
        chad_mindshare(prs); wins_loss_slide(prs)
        q4_slide(prs, "One program or site I'm going to grow", ["Program or site: [name]", "From: [today]", "To: [target]", "How: [the specific move]"])
    elif kind == "market":
        cover(prs, "SALES MEETING  ·  NATIONAL ACCOUNT PRESENTATION", "National accounts",
              "Where we're winning, what's winning with which accounts, and why they'd pick us.",
              "Mac  |  Sales Meeting  |  [Month] 2026",
              "Ten minutes. Six slides. Real accounts and job names. No new data pulls. Pricing gets its own block on Day 2, leave it out of this one.")
        numbers_slide(prs, "National accounts", "Top accounts", excalibur=False)
        mac_products(prs)
        buy_us_for_slide(prs, "national accounts", "Why a national account would pick us",
                         "Slide 3 for Mac. Why a national account would pick us, and what's missing to make that easier.")
        mac_mindshare(prs); wins_loss_slide(prs)
        q4_slide(prs, "One account I'm going to grow", ["Account: [name]", "From: [today]", "To: [target]", "How: [the specific move]"])
    prs.save(out)
    print("saved", out)

if __name__ == "__main__":
    outdir = sys.argv[1]
    build("rsm", f"{outdir}/Territory Presentation - RSM.pptx")
    build("walmart", f"{outdir}/National Account Presentation - Walmart (Chad).pptx")
    build("market", f"{outdir}/National Account Presentation - Mac.pptx")
