import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.oxml.ns import qn
from lxml import etree

SK = "/root/.claude/skills/synced/5dd94d0b-87b7-4407-be58-13dace0ecad4_8e7c9b18-ec18-4b4c-8b20-f72a48e99518/steel-king-ppt-template/assets/template.pptx"
GREEN = RGBColor(0x3C, 0x7D, 0x4D); STEEL = RGBColor(0x5B, 0x66, 0x70); INK = RGBColor(0x15, 0x18, 0x1A)
HAIR = RGBColor(0xE2, 0xE6, 0xE3); TINT = RGBColor(0xF2, 0xF5, 0xF2); WHITE = RGBColor(0xFF, 0xFF, 0xFF)
L = dict(COVER_MASCOT=15, TITLE_CONTENT=4, TWO_CONTENT=5, THREE_CARDS=6)
EYEBROW = "TERRITORY REVIEW"
ANSWER_PT = 12

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

def drop(slide, idx):
    sh = ph(slide, idx); sh._element.getparent().remove(sh._element)

def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text

def answer_format(paragraph, size=ANSWER_PT):
    """What the rep types here comes out bold, dark, and this size."""
    p = paragraph._p
    epr = p.find(qn("a:endParaRPr"))
    if epr is None:
        epr = etree.SubElement(p, qn("a:endParaRPr"))
    epr.set("lang", "en-US"); epr.set("sz", str(size * 100)); epr.set("b", "1")
    for old in epr.findall(qn("a:solidFill")) + epr.findall(qn("a:latin")): epr.remove(old)
    sf = etree.SubElement(epr, qn("a:solidFill")); etree.SubElement(sf, qn("a:srgbClr"), val="15181A")
    etree.SubElement(epr, qn("a:latin"), typeface="Arial")

def textbox(slide, x, y, w, h, text, size=11, bold=False, color=STEEL, italic=False):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True; tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = Inches(0); tf.margin_top = tf.margin_bottom = Inches(0)
    from pptx.enum.text import PP_ALIGN
    tf.paragraphs[0].alignment = PP_ALIGN.LEFT
    r = tf.paragraphs[0].add_run(); r.text = text
    r.font.name = "Arial"; r.font.size = Pt(size); r.font.bold = bold; r.font.italic = italic; r.font.color.rgb = color
    return tb

def heading(slide, x, y, w, text):
    return textbox(slide, x, y, w, 0.26, text.upper(), size=11, bold=True, color=GREEN)

def hint(slide, x, y, w, text):
    return textbox(slide, x, y, w, 0.26, text, size=10, color=STEEL)

def room(slide, text):
    return textbox(slide, 0.80, 6.34, 11.73, 0.26, "For the room:  " + text, size=11, bold=True, color=GREEN)

def answer_box(slide, x, y, w, h, size=ANSWER_PT):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True; tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = Inches(0.10); tf.margin_top = tf.margin_bottom = Inches(0.08)
    tb.fill.solid(); tb.fill.fore_color.rgb = WHITE
    tb.line.color.rgb = HAIR; tb.line.width = Pt(0.75)
    answer_format(tf.paragraphs[0], size)
    return tb

def _border(cell, color="E2E6E3", width=6350):
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        for old in tcPr.findall(qn(tag)): tcPr.remove(old)
        ln = etree.SubElement(tcPr, qn(tag), w=str(width), cap="flat", cmpd="sng", algn="ctr")
        sf = etree.SubElement(ln, qn("a:solidFill")); etree.SubElement(sf, qn("a:srgbClr"), val=color)
        etree.SubElement(ln, qn("a:prstDash"), val="solid")

def table(slide, x, y, w, col_w, header, rows, row_h=0.30, header_h=0.36, label_cols=1, row_heights=None):
    """Green header. Label columns in gray. Every other cell is an answer cell: empty, typed text comes out bold ink."""
    n = len(rows) + 1
    heights = row_heights or [row_h] * len(rows)
    gf = slide.shapes.add_table(n, len(header), Inches(x), Inches(y), Inches(w), Inches(header_h + sum(heights)))
    t = gf.table
    tblPr = t._tbl.tblPr
    tblPr.set("firstRow", "0"); tblPr.set("bandRow", "0")
    for st in tblPr.findall(qn("a:tableStyleId")): tblPr.remove(st)
    for i, cw in enumerate(col_w): t.columns[i].width = Inches(cw)
    t.rows[0].height = Inches(header_h)
    for r, hgt in enumerate(heights, start=1): t.rows[r].height = Inches(hgt)
    for ci, txt in enumerate(header):
        c = t.cell(0, ci); c.fill.solid(); c.fill.fore_color.rgb = GREEN
        c.margin_left = c.margin_right = Inches(0.08); c.margin_top = c.margin_bottom = Inches(0.03)
        c.vertical_anchor = MSO_ANCHOR.MIDDLE
        r = c.text_frame.paragraphs[0].add_run(); r.text = txt
        r.font.name = "Arial"; r.font.size = Pt(10); r.font.bold = True; r.font.color.rgb = WHITE
        _border(c, "3C7D4D")
    for ri, row in enumerate(rows, start=1):
        for ci, txt in enumerate(row):
            c = t.cell(ri, ci); c.fill.solid(); c.fill.fore_color.rgb = WHITE
            c.margin_left = c.margin_right = Inches(0.08); c.margin_top = c.margin_bottom = Inches(0.04)
            c.vertical_anchor = MSO_ANCHOR.TOP
            p = c.text_frame.paragraphs[0]
            if ci < label_cols and txt:
                c.fill.fore_color.rgb = TINT
                r = p.add_run(); r.text = txt
                r.font.name = "Arial"; r.font.size = Pt(10); r.font.bold = True; r.font.color.rgb = STEEL
            else:
                answer_format(p)
            _border(c)
    return t

def strip_watermark(prs):
    for i in (3, 4, 5, 6, 7, 8, 9, 10, 11, 13):
        layout = prs.slide_layouts[i]
        cSld = layout._element.find(qn("p:cSld"))
        bg = cSld.find(qn("p:bg"))
        if bg is not None: cSld.remove(bg)
        new_bg = etree.Element(qn("p:bg")); cSld.insert(0, new_bg)
        bgPr = etree.SubElement(new_bg, qn("p:bgPr"))
        sf = etree.SubElement(bgPr, qn("a:solidFill")); etree.SubElement(sf, qn("a:srgbClr"), val="FFFFFF")
        etree.SubElement(bgPr, qn("a:effectLst"))

def new_deck():
    prs = Presentation(SK)
    while len(prs.slides) > 0:
        rId = prs.slides._sldIdLst[0].rId
        prs.part.drop_rel(rId); del prs.slides._sldIdLst[0]
    strip_watermark(prs)
    return prs

def header(slide, n, tag, title, ask):
    fill(slide, 100, f"{EYEBROW}  ·  {n} OF 6  ·  {tag}".upper())
    fill(slide, 101, title)
    fill(slide, 102, ask)

def content_slide(prs, n, tag, title, ask, note):
    s = prs.slides.add_slide(prs.slide_layouts[L["TITLE_CONTENT"]])
    header(s, n, tag, title, ask); drop(s, 103); drop(s, 104); notes(s, note)
    return s

def cards_slide(prs, n, tag, title, ask, cards, note):
    """cards: list of (heading, hint). Answer box inside each card."""
    s = prs.slides.add_slide(prs.slide_layouts[L["THREE_CARDS"]])
    header(s, n, tag, title, ask)
    xs = (1.20, 5.22, 9.24)
    for (h_idx, b_idx), x, (head, hnt) in zip(((106, 107), (111, 112), (116, 117)), xs, cards):
        fill(s, h_idx, head); drop(s, b_idx)
        hint(s, x, 4.30, 2.91, hnt)
        answer_box(s, x, 4.60, 2.91, 1.44)
    drop(s, 118); notes(s, note)
    return s

# ---------- shared slides ----------
def cover(prs, eyebrow, title, line, who, note):
    s = prs.slides.add_slide(prs.slide_layouts[L["COVER_MASCOT"]])
    fill(s, 102, eyebrow); fill(s, 104, title); fill(s, 105, line); fill(s, 106, who); notes(s, note)
    return s

def numbers_slide(prs, label, right_label, excalibur):
    s = content_slide(prs, 1, "By the numbers", f"{label} in numbers",
                      "Numbers are pre-filled. Add two sentences on what they don't show.",
                      "Slide 1. Pre-filled by Nate. Add two sentences on what the numbers don't show. Two minutes, then move on.")
    heading(s, 0.80, 2.32, 5.60, "Bookings and quotes by quarter")
    table(s, 0.80, 2.60, 5.60, [1.40, 1.40, 1.40, 1.40], ["Quarter", "Bookings 2025", "Bookings 2026", "Quotes 2026"],
          [["Q1", "", "", ""], ["Q2", "", "", ""], ["Q3", "", "", ""], ["Q4 to date", "", "", ""], ["YTD", "", "", ""]],
          row_h=0.30, header_h=0.32)
    heading(s, 0.80, 4.78, 5.60, "What the numbers don't show")
    hint(s, 0.80, 5.02, 5.60, "Two sentences. Yours.")
    answer_box(s, 0.80, 5.28, 5.60, 1.32)
    heading(s, 7.13, 2.32, 5.40, f"{right_label}, YTD bookings")
    rows = 5 if excalibur else 7
    table(s, 7.13, 2.60, 5.40, [3.40, 2.00], [right_label, "YTD bookings"], [["", ""]] * rows, row_h=0.30, header_h=0.32, label_cols=0)
    if excalibur:
        heading(s, 7.13, 4.78, 5.40, "Excalibur dealers, YTD bookings")
        table(s, 7.13, 5.06, 5.40, [3.40, 2.00], ["Excalibur dealer", "YTD bookings"], [["", ""]] * 3, row_h=0.30, header_h=0.32, label_cols=0)
    textbox(s, 0.80, 6.96, 8.6, 0.26, "Source: [system], pulled [date]. Quotes are a count. Pre-filled by Nate.", size=9, color=STEEL)
    return s

def buy_us_for_slide(prs, title, ask, left_head, right_head, note, room_line):
    s = content_slide(prs, 3, "What they buy us for", title, ask, note)
    heading(s, 0.80, 2.32, 5.60, left_head)
    table(s, 0.80, 2.60, 5.60, [0.50, 5.10], ["", "In their words if you have them"],
          [["1", ""], ["2", ""], ["3", ""]], row_h=1.10, header_h=0.36)
    heading(s, 7.13, 2.32, 5.40, right_head)
    table(s, 7.13, 2.60, 5.40, [1.70, 3.70], ["", "One gap. The one that costs the most."],
          [["The gap", ""], ["What it costs us", ""], ["What would fix it", ""]], row_h=1.10, header_h=0.36)
    room(s, room_line)
    return s

def wins_loss_slide(prs, job_label):
    s = content_slide(prs, 5, "Three wins, one loss", "Three wins, one loss",
                      "Three wins: job, competitor, what won it. One loss: what decided it, what we'd do differently.",
                      "Slide 5. Three wins, one loss. Real jobs, real names. Spend most of the time on the wins: what won it and whether "
                      "the rest of the team can do the same thing. On the loss, what decided it and what we could have done differently. "
                      "If it came down to the number, say what else was true.")
    table(s, 0.80, 2.32, 11.73, [0.90, 2.90, 1.90, 3.20, 2.83],
          ["", job_label, "Competitor", "What won it  (loss: what decided it)", "Can we do it again?  (loss: what we'd change)"],
          [["Win 1", "", "", "", ""], ["Win 2", "", "", "", ""], ["Win 3", "", "", "", ""], ["Loss", "", "", "", ""]],
          header_h=0.40, row_heights=[0.92, 0.92, 0.92, 0.80])
    room(s, "What won it, and can the rest of us do the same thing?")
    return s

def q4_slide(prs, grow_head, grow_hint):
    return cards_slide(prs, 6, "Q4", "Q4 plan",
                       "Three projects, one you'll grow, and one thing you'd change about how we run sales.",
                       [("Three projects I'm going after", "Job, dollars, close date"),
                        (grow_head, grow_hint),
                        ("One thing I'd change", "About how we run sales. What, and why.")],
                       "Slide 6. Three projects you're going after, one you're going to grow, one thing you'd change about how we run sales.")

# ---------- RSM ----------
def rsm_winning(prs):
    s = content_slide(prs, 2, "Who's winning with what", "Who's winning with what",
                      "Your top dealers, the products they perform with, and why they pick Steel King, in their words.",
                      "Slide 2. Your top dealers, Excalibur and non-Excalibur, the products each one performs with, and why they pick "
                      "Steel King over the other guys. In their words if you have them. This is the slide the room should spend time on.")
    table(s, 0.80, 2.32, 11.73, [2.40, 1.00, 3.30, 5.03],
          ["Dealer", "Excalibur", "Products they perform with", "Why they pick Steel King, in their words"],
          [["", "", "", ""]] * 5, row_h=0.72, header_h=0.40, label_cols=0)
    room(s, "Are we hearing the same reasons across territories?")
    return s

def rsm_mindshare(prs):
    return cards_slide(prs, 4, "Mindshare", "Mindshare",
                       "How you get and keep mindshare, how new dealer reps learn us, and what's working.",
                       [("How I get and keep mindshare", "Visits, training, joint calls, quick ship"),
                        ("How new dealer reps get up to speed on us", "When a dealer hires someone new"),
                        ("What's working, what isn't", "One of each")],
                       "Slide 4. How you get and keep mindshare with your dealers. How new dealer reps get up to speed on us. What's working, what isn't.")

# ---------- Chad ----------
def chad_products(prs):
    s = content_slide(prs, 2, "What's performing", "What's performing, and why Walmart picks us",
                      "Which products perform across the Walmart sites, and why Walmart keeps choosing us.",
                      "Slide 2 for Chad. Which products are performing across the Walmart sites and why Walmart keeps choosing us.")
    table(s, 0.80, 2.32, 11.73, [2.40, 3.60, 5.73],
          ["Product", "Where it's performing: site, program, format", "Why Walmart keeps choosing us, in their words"],
          [["", "", ""]] * 5, row_h=0.72, header_h=0.40, label_cols=0)
    room(s, "Which of these reasons would hold at another national account?")
    return s

def chad_mindshare(prs):
    return cards_slide(prs, 4, "Mindshare", "Mindshare inside Walmart",
                       "How you keep mindshare inside Walmart and with their integrators, and what's working.",
                       [("Inside Walmart", "As their engineering staff turns over"),
                        ("With their integrators", "Who they are, how you stay in front"),
                        ("What's working, what isn't", "One of each")],
                       "Slide 4 for Chad. How you keep mindshare inside Walmart and with their integrators.")

# ---------- Mac ----------
def mac_products(prs):
    s = content_slide(prs, 2, "What's winning", "What's winning, and where growth is",
                      "Which products are winning with which accounts, and where the biggest new customer opportunities are.",
                      "Slide 2 for Mac. Which products are winning with which accounts, and where the biggest new customer opportunities are.")
    heading(s, 0.80, 2.32, 5.60, "Products winning, by account")
    table(s, 0.80, 2.60, 5.60, [1.70, 1.70, 2.20], ["Product", "Account", "Why it's winning"],
          [["", "", ""]] * 3, row_h=1.05, header_h=0.36, label_cols=0)
    heading(s, 7.13, 2.32, 5.40, "Biggest new customer opportunities")
    table(s, 7.13, 2.60, 5.40, [1.70, 1.85, 1.85], ["Account or segment", "Why", "What it takes"],
          [["", "", ""]] * 3, row_h=1.05, header_h=0.36, label_cols=0)
    room(s, "Which of these wins could a dealer territory copy?")
    return s

def mac_mindshare(prs):
    return cards_slide(prs, 4, "Getting in the door", "Getting in the door",
                       "How you're getting in the door with national accounts, and what's working.",
                       [("How I'm getting in", "Referrals, shows, integrators, outreach"),
                        ("What's working", "And why"),
                        ("What isn't", "And what would help")],
                       "Slide 4 for Mac. How you're getting in the door. What's working, what isn't.")

# ---------- build ----------
PREP = "Ten minutes. Six slides. Real names, real jobs. No new data pulls. Pricing gets its own block on Day 2, leave it out of this one."

def build(kind, out):
    global EYEBROW
    EYEBROW = "TERRITORY REVIEW" if kind == "rsm" else "NATIONAL ACCOUNT REVIEW"
    prs = new_deck()
    if kind == "rsm":
        cover(prs, "SALES MEETING  ·  TERRITORY PRESENTATION", "[Territory name]",
              "Where we're winning, who's winning with what, and why they choose us.", "[Your name]  |  Sales Meeting  |  [Month] 2026", PREP)
        numbers_slide(prs, "[Territory]", "Top accounts", True)
        rsm_winning(prs)
        buy_us_for_slide(prs, "What they buy us for, and what's missing",
                         "Two or three things dealers count on us for, then the one gap that costs you the most business.",
                         "What dealers count on us for", "The one gap that costs me the most",
                         "Slide 3. Lead with what dealers count on us for. Then the one gap that costs the most. One gap, not a list.",
                         "Which of these do we all hear?")
        rsm_mindshare(prs); wins_loss_slide(prs, "Job: dealer, account, job name")
        q4_slide(prs, "One dealer I'm going to grow", "Dealer, from, to, how")
    elif kind == "walmart":
        cover(prs, "SALES MEETING  ·  NATIONAL ACCOUNT PRESENTATION", "Walmart, Sam's, and Ambient",
              "Where we're winning, what's performing, and why Walmart keeps choosing us.", "Chad  |  Sales Meeting  |  [Month] 2026", PREP)
        numbers_slide(prs, "Walmart", "Top programs and sites", False)
        chad_products(prs)
        buy_us_for_slide(prs, "What Walmart buys us for, and what's missing",
                         "Two or three things Walmart counts on us for, then the one gap that costs you the most business.",
                         "What Walmart counts on us for", "The one gap that costs me the most",
                         "Slide 3 for Chad. Lead with what Walmart counts on us for. Then the one gap that costs the most.",
                         "Which of these would the dealers say too?")
        chad_mindshare(prs); wins_loss_slide(prs, "Job: site, program, job name")
        q4_slide(prs, "One program or site I'm going to grow", "Program or site, from, to, how")
    elif kind == "market":
        cover(prs, "SALES MEETING  ·  NATIONAL ACCOUNT PRESENTATION", "National accounts",
              "Where we're winning, what's winning with which accounts, and why they'd pick us.", "Mac  |  Sales Meeting  |  [Month] 2026", PREP)
        numbers_slide(prs, "National accounts", "Top accounts", False)
        mac_products(prs)
        buy_us_for_slide(prs, "Why a national account would pick us",
                         "Why a national account would pick us, and what's missing to make that easier.",
                         "Why a national account picks us", "What's missing to make that easier",
                         "Slide 3 for Mac. Why a national account would pick us, and what's missing to make that easier.",
                         "Is this the same reason the dealers give?")
        mac_mindshare(prs); wins_loss_slide(prs, "Job: account, job name")
        q4_slide(prs, "One account I'm going to grow", "Account, from, to, how")
    prs.save(out); print("saved", out)

if __name__ == "__main__":
    outdir = sys.argv[1]
    build("rsm", f"{outdir}/Territory Presentation - RSM.pptx")
    build("walmart", f"{outdir}/National Account Presentation - Walmart (Chad).pptx")
    build("market", f"{outdir}/National Account Presentation - Mac.pptx")
