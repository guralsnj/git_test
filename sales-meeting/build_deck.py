"""Build the Steel King sales meeting deck on speed to lead and customer service.

Run:  python3 build_deck.py
Output: Speed_and_Service_Win_the_Order.pptx (same folder)
"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.enum.text import PP_ALIGN

TEMPLATE = (
    "/root/.claude/skills/synced/5dd94d0b-87b7-4407-be58-13dace0ecad4_"
    "8e7c9b18-ec18-4b4c-8b20-f72a48e99518/steel-king-ppt-template/assets/template.pptx"
)
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                   "Speed_and_Service_Win_the_Order.pptx")

GREEN = RGBColor(0x3C, 0x7D, 0x4D)
STEEL = RGBColor(0x5B, 0x66, 0x70)
LIGHT = RGBColor(0x8C, 0xC7, 0xA2)
INK = RGBColor(0x15, 0x18, 0x1A)
HAIR = RGBColor(0xE2, 0xE6, 0xE3)
TINT = RGBColor(0xF2, 0xF5, 0xF2)

L = dict(COVER=1, SECTION=2, AGENDA=3, TITLE_CONTENT=4, TWO_CONTENT=5,
         THREE_CARDS=6, BIG_NUMBERS=7, CONTENT_VISUAL=8, FULL_VISUAL=9,
         COMPARISON=10, TIMELINE=11, KEY_MESSAGE=12, NEXT_STEPS=13,
         CLOSING=14, COVER_MASCOT=15)

prs = Presentation(TEMPLATE)
while len(prs.slides) > 0:
    rId = prs.slides._sldIdLst[0].rId
    prs.part.drop_rel(rId)
    del prs.slides._sldIdLst[0]


# ---------- helpers ----------
def ph(slide, idx):
    for sh in slide.placeholders:
        if sh.placeholder_format.idx == idx:
            return sh
    raise KeyError(idx)


def fill(slide, idx, text):
    sh = ph(slide, idx)
    p = sh.text_frame.paragraphs[0]
    if p.runs:
        p.runs[0].text = text
        for r in p.runs[1:]:
            r.text = ""
    else:
        p.add_run().text = text


def fill_lines(slide, idx, lines):
    sh = ph(slide, idx)
    tf = sh.text_frame
    tf.text = lines[0]
    for line in lines[1:]:
        p = tf.add_paragraph()
        p.text = line
    for p in tf.paragraphs:
        p.space_after = Pt(7)


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def add(layout):
    return prs.slides.add_slide(prs.slide_layouts[L[layout]])


def header(slide, eyebrow, title, takeaway, source=None, src_idx=None):
    fill(slide, 100, eyebrow)
    fill(slide, 101, title)
    fill(slide, 102, takeaway)
    if source is not None and src_idx is not None:
        fill(slide, src_idx, "Source: " + source)


def big(slide, eyebrow, title, takeaway, nums, sowhat, source):
    header(slide, eyebrow, title, takeaway)
    for (n, label, ctx), (a, b, c) in zip(nums, [(104, 105, 106), (108, 109, 110), (112, 113, 114)]):
        fill(slide, a, n)
        fill(slide, b, label)
        fill(slide, c, ctx)
    fill(slide, 115, sowhat)
    fill(slide, 116, "Source: " + source)


def bar_chart(slide, x, y, w, h, categories, values, series_name, number_format='0"%"',
              horizontal=True, font_pt=12):
    data = CategoryChartData()
    data.categories = categories
    data.add_series(series_name, values)
    ctype = XL_CHART_TYPE.BAR_CLUSTERED if horizontal else XL_CHART_TYPE.COLUMN_CLUSTERED
    gf = slide.shapes.add_chart(ctype, Inches(x), Inches(y), Inches(w), Inches(h), data)
    ch = gf.chart
    ch.has_title = False
    ch.has_legend = False
    ch.font.size = Pt(font_pt)
    ch.font.name = "Arial"
    ch.font.color.rgb = INK
    plot = ch.plots[0]
    plot.gap_width = 55
    s = plot.series[0]
    s.format.fill.solid()
    s.format.fill.fore_color.rgb = GREEN
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.number_format = number_format
    dl.number_format_is_linked = False
    dl.font.size = Pt(font_pt)
    dl.font.bold = True
    dl.font.color.rgb = INK
    dl.position = XL_LABEL_POSITION.OUTSIDE_END
    va = ch.value_axis
    va.has_major_gridlines = False
    va.visible = False
    ca = ch.category_axis
    ca.format.line.color.rgb = HAIR
    ca.tick_labels.font.size = Pt(font_pt)
    ca.tick_labels.font.color.rgb = INK
    if horizontal:
        ca.reverse_order = True
    return ch


def remove_placeholder(slide, idx):
    sh = ph(slide, idx)
    sh._element.getparent().remove(sh._element)


def table(slide, x, y, w, h, col_w, rows, header_fill=GREEN):
    shape = slide.shapes.add_table(len(rows), len(rows[0]), Inches(x), Inches(y), Inches(w), Inches(h))
    tbl = shape.table
    for i, cw in enumerate(col_w):
        tbl.columns[i].width = Inches(cw)
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            cell = tbl.cell(r, c)
            cell.text = val
            cell.margin_left = Inches(0.08)
            cell.margin_right = Inches(0.08)
            cell.margin_top = Inches(0.05)
            cell.margin_bottom = Inches(0.05)
            for p in cell.text_frame.paragraphs:
                for run in p.runs:
                    run.font.name = "Arial"
                    run.font.size = Pt(12 if r else 11)
                    run.font.bold = (r == 0)
                    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF) if r == 0 else INK
            cell.fill.solid()
            cell.fill.fore_color.rgb = header_fill if r == 0 else (TINT if r % 2 == 0 else RGBColor(0xFF, 0xFF, 0xFF))
    return tbl


# =====================================================================
# 1. COVER
# =====================================================================
s = add("COVER_MASCOT")
fill(s, 102, "SALES MEETING")
fill(s, 104, "Speed and service win the order")
fill(s, 105, "What the research says about first touch, follow-up, and support in industrial buying")
fill(s, 106, "Nate Guralski  |  Steel King Sales Team  |  September 2026")
notes(s, "Open with the point: buyers pick a favorite before they call. Our first touch and how we support "
         "them after decide whether that favorite is Steel King. Everything in this deck is sourced. Widely "
         "quoted numbers that could not be traced to a real study were left out on purpose.")

# =====================================================================
# 2. KEY MESSAGE
# =====================================================================
s = add("KEY_MESSAGE")
fill(s, 101, "THE POINT")
fill(s, 103, "By the time a buyer contacts us, they have already picked a favorite. The first touch decides whether it stays us.")
fill(s, 104, "Eight in ten B2B deals go to the first vendor the buyer contacts. Most vendors take more than a day to answer.")
notes(s, "6sense 2025 Buyer Experience Report (about 4,000 buyers): the first vendor contacted wins roughly 8 in 10 deals, "
         "and 94% of buying groups had ranked their preferred vendors before making first contact. "
         "HBR 2011 audit of 2,241 U.S. companies: average response to a web lead was 42 hours; 23% never responded.")

# =====================================================================
# 3. AGENDA
# =====================================================================
s = add("AGENDA")
header(s, "AGENDA", "Five things we cover", "Why it matters, what the market shows, and what we change.")
items = [("The market is showing up", "5 min"),
         ("The buyer decides early", "10 min"),
         ("Speed to lead", "10 min"),
         ("Service after the sale", "10 min"),
         ("What we change", "15 min")]
for (t, m), (a, b) in zip(items, [(104, 105), (108, 109), (112, 113), (116, 117), (120, 121)]):
    fill(s, a, t)
    fill(s, b, m)
notes(s, "50 minutes. Leave the last 15 for the standards discussion and owners.")

# =====================================================================
# SECTION 01: MARKET
# =====================================================================
s = add("SECTION")
fill(s, 101, "01")
fill(s, 103, "The market is showing up")
fill(s, 104, "Demand for rack is real. The question is who gets the order.")

s = add("BIG_NUMBERS")
big(s, "MARKET", "Racking is the top planned equipment buy",
    "Buyers have budget and intent. We are competing for share, not for demand.",
    [("45%", "PLAN TO BUY RACKS AND SHELVING", "Top equipment priority in the 2026 and 2025 MMH outlook surveys"),
     ("$542K", "AVG PLANNED MH EQUIPMENT SPEND", "Up from $402K planned for 2025, a 35% jump"),
     ("+18%", "U.S. INDUSTRIAL CONSTRUCTION PIPELINE", "305M sq ft under construction, Q2 2026, fourth straight quarterly rise")],
    "Every one of these buyers will shortlist fewer than five suppliers. Our job is to be on the list and stay on it.",
    "Modern Materials Handling / Peerless Research 2026 and 2025 Outlook Surveys (n=103, n=110); Cushman & Wakefield U.S. Industrial MarketBeat Q2 2026.")
notes(s, "MMH 2026 Outlook (Jan 2026, n=103): racks and shelving named by 45% as a planned purchase, #1 category. "
         "Avg planned spend $541,670, up from under $402,000. MMH 2025 Warehouse/DC Operations Survey: net 80% plan to expand, "
         "33% plan to add square footage (from 25%), avg capex $2.16M. "
         "Cushman & Wakefield Q2 2026: vacancy 6.9%, pipeline 305 msf up 18% YoY. CBRE Q2 2026 puts vacancy at 6.5%. "
         "MHI/Deloitte 2026: 56% of supply chain leaders increasing tech investment, 52% plan to spend over $1M. "
         "Growth pockets: cold storage 2026 pipeline only 5.9M sq ft (lowest since 2020) after record 2025 deliveries; "
         "data center component manufacturing drove a 28% YoY jump in large manufacturing leases in Q1 2026 (CBRE). "
         "MHEDA 2026 members name data centers and cold storage as top niche growth markets.")

# =====================================================================
# SECTION 02: THE BUYER DECIDES EARLY
# =====================================================================
s = add("SECTION")
fill(s, 101, "02")
fill(s, 103, "The buyer decides early")
fill(s, 104, "Most of the decision is made before anyone calls a rep.")

s = add("BIG_NUMBERS")
big(s, "BUYER BEHAVIOR", "The shortlist is built before we get the call",
    "Buyers research alone, keep the list short, and mostly buy from the first vendor they contact.",
    [("62%", "OF THE BUYING PROCESS DONE ONLINE", "Before a technical buyer contacts anyone at the vendor"),
     ("71%", "VET FEWER THAN FIVE SUPPLIERS", "Industrial buyers; only 3% look at more than ten"),
     ("8 in 10", "DEALS GO TO FIRST VENDOR CONTACTED", "Buyers rank favorites first, then reach out")],
    "If the first call, form, or dealer inquiry goes unanswered for a day, we hand the order to whoever answered.",
    "GlobalSpec / TREW Marketing, 2026 State of Marketing to Engineers (n=1,000+); Thomasnet Industrial Buyer Habits report; 6sense 2025 Buyer Experience Report (approx. 4,000 buyers).")
notes(s, "GlobalSpec/TREW 2026: 62% of buying process online before contacting sales (66% for buyers 35 and under). "
         "Only 6% prefer never to talk to a salesperson; they contact sales for technical complexity, pricing, and inventory. "
         "Thomasnet: 71% vet fewer than five suppliers, 25% vet 5 to 10, 3% more than ten. "
         "6sense 2025: buyers make first contact at 61% of the buying process (down from 69% in 2024), 94% of buying groups had ranked "
         "preferred vendors before first contact and bought that favorite 77% of the time. First vendor contacted wins about 8 in 10. "
         "Gartner: buyers spend 17% of the purchase process meeting with all suppliers combined; any one rep gets 5 to 6% of that time.")

# --- chart: how distributors win business
s = add("CONTENT_VISUAL")
header(s, "WHAT BUYERS WEIGH", "Lead time and support outrank price",
       "Industrial buyers rank lead time first when they shortlist. Price is fifth on why distributors win.",
       "Thomasnet survey of 400+ registered industrial buyers (2021); Industrial Distribution Survey of Distributor Operations.", 105)
fill_lines(s, 103, [
    "Thomasnet shortlist factors, in order: lead time and availability, unit price, quality certifications, verified company information.",
    "Buyers expect an answer inside 24 hours. 44% say so outright.",
    "One buyer, verbatim: \"If I have to follow up twice, I eliminate you from consideration.\"",
    "Rejection reasons: delayed or no RFQ response, having to follow up more than once, unclear pricing, stale company info.",
])
remove_placeholder(s, 104)
bar_chart(s, 6.50, 2.32, 6.03, 4.28,
          ["Relationships", "Product availability", "Technical support", "Delivery time", "Price"],
          [81, 76, 66, 62, 53], "Why distributors say they win")
notes(s, "Chart: Industrial Distribution annual Survey of Distributor Operations, reasons distributors say they earn business "
         "(relationships 81%, availability 76%, technical support 66%, delivery time 62%, price 53%). Self-reported by distributors. "
         "Thomasnet 2021 survey of 400+ buyers published rankings, not percentages, for the shortlist factors. "
         "The 44%-within-24-hours figure is from Thomasnet's Industrial Buyer Search Habits survey; year not disclosed on the page.")

# --- two content: what buyers want vs what they get
s = add("TWO_CONTENT")
header(s, "THE REP", "Buyers want a rep who knows their business",
       "Human help cuts buyer regret in half. Generic help gets us removed from the list.",
       "Salesforce State of the Connected Customer 2023 (3,300 business buyers); Gartner B2B Buying Report (2022 survey, n=441); LinkedIn State of Sales 2021; Gartner 2025 sales survey via Salesforce.", 109)
fill(s, 103, "WHAT BUYERS WANT")
fill_lines(s, 105, [
    "86% are more likely to buy when the company understands their goals.",
    "84% expect the rep to act as a trusted advisor.",
    "Rep-assisted purchases carry half the regret of self-service buys: 21% vs 43%.",
    "Only 6% of technical buyers want no salesperson at all.",
])
fill(s, 106, "WHAT BUYERS REPORT GETTING")
fill_lines(s, 108, [
    "59% say most reps don't take the time to understand them.",
    "73% say most sales interactions feel transactional.",
    "69% see conflicts between the website and what the rep says.",
    "Top deal killers: misleading price or product info (48%), not understanding the buyer's company (44%).",
])
notes(s, "Salesforce 2023 (6th ed., 3,300 business buyers): 86%, 84%, 59%, 73%. "
         "Gartner B2B Buying Report: high regret 43% self-service digital, 26% rep-led, 21% rep-assisted digital; "
         "self-service buyers 1.65x more likely to regret. "
         "Gartner 2025 sales survey (n=632): 69% report inconsistencies between website and rep. "
         "LinkedIn State of Sales 2021 (400+ buyers): deal killers 48% misleading info, 44% not understanding company, 43% not understanding own product. "
         "GlobalSpec 2025: only 6% prefer never to talk to a salesperson.")

# =====================================================================
# SECTION 03: SPEED TO LEAD
# =====================================================================
s = add("SECTION")
fill(s, 101, "03")
fill(s, 103, "Speed to lead")
fill(s, 104, "The first hour is worth more than the next week.")

s = add("BIG_NUMBERS")
big(s, "RESPONSE TIME", "An hour's delay cuts qualification odds 7x",
    "The value of a lead decays by the hour. Most companies let it decay for nearly two days.",
    [("7x", "MORE LIKELY TO QUALIFY", "Contact within one hour vs. waiting one more hour"),
     ("60x", "MORE LIKELY TO QUALIFY", "Contact within one hour vs. waiting 24 hours or more"),
     ("42 hrs", "AVERAGE RESPONSE TIME", "Across 2,241 U.S. companies audited; 23% never responded")],
    "A lead that sits overnight is not the same lead in the morning. Speed is the cheapest advantage we have.",
    "Oldroyd, McElheran, Elkington, \"The Short Life of Online Sales Leads,\" Harvard Business Review, March 2011 (1.25M leads, 42 companies; audit of 2,241 companies).")
notes(s, "HBR March 2011. Two datasets: (1) 1.25 million leads at 42 companies, 29 B2C and 13 B2B: firms that tried to contact "
         "within an hour were nearly 7x as likely to qualify the lead as those that waited even one more hour, and more than 60x "
         "as likely as those waiting 24 hours or longer. (2) Audit of 2,241 U.S. companies: 37% responded within an hour, 16% within "
         "1 to 24 hours, 24% took more than 24 hours, 23% never responded. Average 42 hours among those responding within 30 days. "
         "The older 2007 InsideSales/Oldroyd study (often mislabeled as an MIT study) found qualification odds drop 21x between "
         "a 5-minute and a 30-minute response. That was vendor-sponsored; HBR is the cleaner citation.")

# --- chart: how companies respond
s = add("CONTENT_VISUAL")
header(s, "THE GAP", "Most companies take a day or more to answer",
       "Slow response is the norm. That makes fast response a differentiator.",
       "HBR 2011 audit (2,241 companies); Workato lead response study, March 2026 (114 B2B companies); InsideSales 2021 Lead Response Management (5.7M leads).", 105)
fill_lines(s, 103, [
    "2026, 114 B2B companies tested: one sent a personalized reply within five minutes. None called within five minutes.",
    "Average personalized email took 12 hours. Average phone call took 14.5 hours. Only 31% called at all.",
    "Platform data, 5.7M leads: fewer than 1% of first attempts happen inside five minutes. 77% of leads got no response.",
    "Reps average 1.3 call attempts before giving up.",
])
remove_placeholder(s, 104)
bar_chart(s, 6.50, 2.32, 6.03, 4.28,
          ["Within 1 hour", "1 to 24 hours", "More than 24 hours", "Never responded"],
          [37, 16, 24, 23], "Share of companies, HBR 2011 audit")
notes(s, "Chart: HBR 2011 audit of 2,241 companies. "
         "Workato March 2026: 114 B2B companies received demo requests. Only 1 sent a personalized email within 5 minutes, none called "
         "within 5 minutes; avg personalized email 11h54m; avg phone 14h29m; 31% called at all; about 20% never emailed. "
         "InsideSales 2021 (400+ companies, 5.7M leads, 55M activities): under 1% of attempts within 5 minutes, under 15% within the first day, "
         "57% of attempts wait more than a week, 77% of leads not responded to at all. "
         "InsideSales 2013 audit of 14,061 companies: 47% did not respond; reps averaged 1.3 call attempts (7,960 companies, 2008 to 2012). "
         "Drift 2017 (433 SaaS companies): 7% responded within 5 minutes, 55% had not responded after 5 business days.")

# --- persistence
s = add("TITLE_CONTENT")
header(s, "PERSISTENCE", "Most reps quit after one attempt",
       "Most deals need several touches. The second through sixth is where reps stop and conversions live.",
       "Velocify / Leads360 Ultimate Contact Strategy 2012 (3.5M leads, consumer-heavy verticals); RAIN Group, Top Performance in Sales Prospecting, 2018 (488 buyers, 489 sellers); InsideSales 2013 Lead Response Report.", 104)
fill_lines(s, 103, [
    "Half of all leads are called only once. 93% of leads that convert are reached by the sixth call.",
    "A six-call cadence lifted conversion 49%. Adding a five-email cadence lifted it 128% combined.",
    "Average rep makes 1.3 call attempts before moving on.",
    "New accounts take 8 touches on average to get a meeting. Top performers get there in 5 because each touch carries something useful.",
    "Buyers told RAIN what they will tolerate: 55% accept 2 to 4 contacts, 23% accept 5 to 10. Persistence is expected, not resented.",
    "Caveat: the Velocify data is consumer phone sales. The direction holds in B2B; the exact percentages are not ours to claim.",
])
notes(s, "Velocify (then Leads360) Dec 2012, 400+ companies, ~3.5M leads, skewed to mortgage, insurance, education: "
         "50% of leads called once; 93% of converted leads reached by the 6th call; 6-call cadence +49%, 5-email +53%, combined +128%; "
         "calling within 1 minute +391% conversion. Consumer phone-heavy data, so quote the pattern, not the precision. "
         "RAIN Group 2018: average touches to get a meeting with a new account 8 for the field, 5 for top performers; top performers "
         "generate 52 conversions per 100 targets vs 19. Buyer tolerance: 2 to 4 contacts 55%, 5 to 10 23%, 11 to 15 12%. "
         "InsideSales: 1.3 call attempts average. XANT 2016 audit: average persistence 4.5 touches, only 9.4% got 12 touches. "
         "NOT USED: '80% of sales take 5 follow-ups, 44% of reps quit after one.' That traces to a 1942 survey of fewer than 40 "
         "salespeople on Long Island and has no modern source.")

# --- what we left out
s = add("TITLE_CONTENT")
header(s, "CREDIBILITY", "The popular follow-up stats don't hold up",
       "Every figure in this deck traces to a named study. These popular ones don't, so they're not here.",
       "Verification notes in the companion research brief. Origins traced by Sales & Marketing Executives International (2021), HBR 2011, and the original publishers.", 104)
fill_lines(s, 103, [
    "\"80% of sales need 5 follow-ups; 44% of reps give up after one.\" Traces to a 1942 survey of under 40 door-to-door salesmen. No modern source.",
    "\"78% of buyers buy from the first company to respond.\" Attributed to a \"Lead Connect survey\" that no one can produce.",
    "\"35 to 50% of sales go to the first responder.\" A vendor blog assertion with no study behind it.",
    "\"The MIT study says 21x.\" It was a 2007 vendor-sponsored analysis by an MIT fellow, not an MIT publication. HBR's 7x and 60x are the clean numbers.",
    "\"A dissatisfied customer tells 9 to 15 people.\" Attributed to a White House office that never published it; the research is from 1979.",
    "Why it matters: if we set standards on bad numbers, the first skeptic in the room wins. Use the sourced ones.",
])
notes(s, "This slide is optional in a short meeting, but it earns trust. If someone quotes one of these at us later, "
         "we know where it came from. The companion research brief has the full trace for each.")

# =====================================================================
# SECTION 04: SERVICE AFTER THE SALE
# =====================================================================
s = add("SECTION")
fill(s, 101, "04")
fill(s, 103, "Service after the sale")
fill(s, 104, "Service response is now a purchase criterion, not a cost center.")

s = add("BIG_NUMBERS")
big(s, "SERVICE", "Service response is now a buying criterion",
    "Buyers rank fast service response next to uptime. Poor service is the second reason business buyers leave.",
    [("95%", "RATE FAST SERVICE RESPONSE VERY IMPORTANT", "Warehouse automation buyers, 2026; up from 83% a year earlier"),
     ("41%", "OF BUSINESS BUYERS LEFT OVER POOR SERVICE", "Second only to price (65%) as a reason to stop buying"),
     ("63%", "WILL SWITCH AFTER ONE BAD EXPERIENCE", "Up nine points year over year")],
    "The install crew, the parts desk, and the engineer who calls back are part of the product. Buyers score them that way.",
    "Modern Materials Handling / Peerless 2026 Automation Study (120+ decision makers) and 2025 Automation Survey (n=139); Salesforce State of the AI Connected Customer 2024 (1,570 business buyers); Zendesk CX Trends 2025 (5,100 consumers).")
notes(s, "MMH/Peerless 2026 Automation Study: 'very important' when evaluating: fast service response 95% (from 83%), durability/uptime 92%, "
         "purchase price 78%, TCO/ROI 77%, parts availability 74%. 2025 Lift Truck User Survey: service response time is a top-5 purchase factor. "
         "Salesforce 7th ed. 2024: reasons customers stopped buying in the past year, poor customer service 43% overall, 41% among business buyers, "
         "behind high prices at 65%. "
         "Zendesk CX Trends 2025: 63% of consumers willing to switch after one bad experience, up 9% YoY. Consumer sample; use as direction. "
         "Microsoft 2019 State of Global Customer Service: 61% have stopped doing business with a brand over poor service; 95% call service important to brand choice. "
         "MHEDA Journal, Hytrol CRO: 'response time is within minutes, not hours' for uptime-critical end users.")

s = add("TWO_CONTENT")
header(s, "THE MATH", "Retention pays twice. Churn costs for years.",
       "Aftermarket carries double the margin of equipment, and a burned B2B customer stays gone.",
       "Deloitte 2026 Manufacturing Industry Outlook; Gartner CSO survey May 2025 (n=243); Accenture 2022 and 2019 B2B studies; Dimensional Research 2013 (n=1,046, B2B split); Gallup 2016.", 109)
fill(s, 103, "WHAT SERVICE EARNS")
fill_lines(s, 105, [
    "Aftermarket services deliver margins more than two times higher than equipment sales alone.",
    "73% of chief sales officers now prioritize growth from existing customers over new logos.",
    "Companies that run service as a value center grow revenue 3.5x faster than those that run it as a cost center.",
    "62% of B2B customers bought more after a good service experience.",
])
fill(s, 106, "WHAT POOR SERVICE COSTS")
fill_lines(s, 108, [
    "80% of frequent B2B buyers switched a supplier in 24 months because it could not meet their expectations.",
    "66% of B2B customers stopped buying after a bad service interaction.",
    "51% of them were still avoiding that vendor two or more years later.",
    "Only 29% of B2B customers are fully engaged. Gallup calls the other 71% ready to take their business elsewhere.",
])
notes(s, "Deloitte 2026 outlook (600 execs): aftermarket margins more than 2x equipment. "
         "Gartner May 2025: 73% of CSOs prioritizing existing-customer growth; 57% rank retention/growth top-3. "
         "Accenture 2022 (2,030 service leaders, 3,428 B2B customers): service-as-value-center companies 3.5x revenue growth. "
         "Accenture Interactive 2019 (748 B2B buyers): 80% switched suppliers at least once in 24 months. "
         "Dimensional Research 2013 (Zendesk sponsored, n=1,046): 62% B2B bought more after good service; 66% B2B stopped after bad; "
         "51% B2B avoid vendor 2+ years later. Dated, but the only clean B2B/B2C split. "
         "Gallup 2016: 29% fully engaged, 60% indifferent, 11% actively disengaged; fully engaged deliver 50% higher revenue, 34% higher profitability. "
         "Rule of thumb, not a study: probability of selling to an existing customer 60 to 70% vs 5 to 20% for a new prospect (Marketing Metrics, 2010).")

s = add("TITLE_CONTENT")
header(s, "THE CHANNEL", "Our response time becomes the dealer's",
       "Dealers are being asked for contractor-grade service. Our response time is theirs.",
       "MMH / Peerless 2025 Lift Truck User Survey (n=150); MHEDA Journal 2025 Industry Outlook and 2026 Business Trends; UPS Industrial Buying Dynamics 2017 (1,500 U.S. buyers) and 2019.", 104)
fill_lines(s, 103, [
    "86% of lift truck spend runs through dealers. Under 15% goes direct. Rack is not measured the same way, but the channel is the same.",
    "MHEDA 2025: storage and handling customers now expect \"contractor-like service capabilities\" from the dealer.",
    "MHEDA 2026 trends: rising customer expectations squeezing margins; manufacturer-distributor partnerships \"becoming essential.\"",
    "Half of industrial buyers would switch to a supplier offering returns help, training, and on-site support. 87% of millennial buyers would.",
    "Younger buyers research online and expect to buy online. 80% would move to a supplier with a more usable web presence.",
    "Every quote, drawing, and parts question a dealer sends us has an end customer waiting behind it.",
])
notes(s, "MMH 2025 Lift Truck User Survey: 86% of expenditures through dealer channel, under 15% direct. "
         "MHEDA Journal 2025 outlook, storage and handling segment: record order intake post-election, customers expecting contractor-like service, "
         "direct online purchasing increasing. MHEDA 2026 Business Trends: #3 rising customer expectations and margin pressure; #15 strategic "
         "manufacturer-distributor partnerships essential; #11 next-gen buyers shifting to online research and purchase. "
         "UPS 2017 (1,500 U.S. industrial buyers): 50% would switch for returns/training/on-site maintenance; 80% would shift to a supplier "
         "with a more user-friendly web presence. UPS 2019: 87% of millennial buyers likely to shift business for better post-sales support. "
         "Aleran 2025 (200 manufacturers/distributors, vendor-sponsored): 88% have lost deals to slow manual quoting; 71% take a day or more to quote.")

# =====================================================================
# SECTION 05: WHAT WE CHANGE
# =====================================================================
s = add("SECTION")
fill(s, 101, "05")
fill(s, 103, "What we change")
fill(s, 104, "Three standards, measured, owned, and on the scorecard.")

s = add("THREE_CARDS")
header(s, "STANDARDS", "Three standards we hold ourselves to",
       "Each one maps to a number in this deck. Each one gets measured.",
       "Proposed. Targets set from the research above; final numbers to be confirmed with the team.", 118)
fill(s, 106, "1. First touch inside one hour")
fill(s, 107, "Every inbound lead, web form, dealer inquiry, and RFQ gets a human acknowledgment within one business hour. "
             "A real answer or quote inside 24 hours. Nobody waits two days to hear from Steel King.")
fill(s, 111, "2. Six touches before we close a lead")
fill(s, 112, "No lead is marked dead after one call. Minimum six touches across phone and email over two weeks, "
             "each with something useful: a drawing, a lead time, a reference. Log every touch.")
fill(s, 116, "3. Same-day service response")
fill(s, 117, "Post-sale questions, parts, and install issues get a same-day reply and a named owner. Dealer questions answered "
             "within four business hours. Response time reported monthly next to bookings.")
notes(s, "These are proposals for discussion. The one-hour target comes from HBR's 7x finding. Six touches comes from Velocify's "
         "93%-by-the-sixth-call and RAIN's 5 to 8 touches. Same-day service comes from MMH's 95% and Thomasnet's 24-hour expectation. "
         "Ask the room: what gets in the way of each one today? That is the list we fix.")

s = add("TIMELINE")
header(s, "ROLLOUT", "Measure first, then set the bar",
       "We cannot manage response time until we can see it. Baseline comes before targets.",
       "Dates are proposed and can move with the team's input.", 124)
phases = [("OCT 2026", "Baseline", "Pull current response times for web leads, dealer RFQs, and service requests. Secret-shop our own inbound. Owner: Nate with inside sales."),
          ("NOV 2026", "Standards and cadence", "Agree the three standards. Build the six-touch cadence and templates. Assign lead ownership rules. Owner: regional managers."),
          ("DEC 2026", "Tools and dashboard", "Route leads with alerts. Response-time report by rep and by dealer. Owner: marketing and inside sales."),
          ("Q1 2027", "Scorecard and review", "Response time and follow-up rate on the sales scorecard. Review monthly in L10. Owner: Nate.")]
for (lab, head, body), (a, b, c) in zip(phases, [(106, 107, 108), (111, 112, 113), (116, 117, 118), (121, 122, 123)]):
    fill(s, a, lab)
    fill(s, b, head)
    fill(s, c, body)
notes(s, "Baseline first. HBR's audit method is simple: submit a lead to ourselves and time the reply. Do it for web, phone, and a dealer inquiry.")

s = add("KEY_MESSAGE")
fill(s, 101, "THE POINT")
fill(s, 103, "Nobody buys rack because we called back fast. They stop considering us because we didn't.")
fill(s, 104, "Speed and service don't win on their own. They keep us on the shortlist long enough for product and price to win.")
notes(s, "Close the argument here before the action list.")

# --- next steps table
s = add("NEXT_STEPS")
header(s, "ACTIONS", "Next steps", "What we agree to today, who owns it, and when.")
fill(s, 104, "Owners and dates are proposed for discussion in the meeting.")
remove_placeholder(s, 103)
rows = [
    ["Action", "Owner", "Date", "Status"],
    ["Secret-shop our own inbound (web form, phone, dealer RFQ) and report response times", "Nate", "Oct 15, 2026", "Proposed"],
    ["Pull 90-day CRM data: time to first touch and touches per lead, by rep", "Inside sales lead", "Oct 15, 2026", "Proposed"],
    ["Draft the six-touch cadence with templates and required content per touch", "Regional managers", "Nov 1, 2026", "Proposed"],
    ["Agree the three standards and the exceptions", "Sales team", "Nov sales meeting", "Proposed"],
    ["Lead routing alerts and a response-time dashboard by rep and dealer", "Marketing (John) with inside sales", "Dec 15, 2026", "Proposed"],
    ["Add response time and follow-up rate to the sales scorecard", "Nate", "Jan 2027 L10", "Proposed"],
]
table(s, 0.80, 2.32, 11.73, 4.10, [6.13, 2.20, 1.60, 1.80], rows)
notes(s, "Owners are placeholders based on role. Confirm names in the room.")

# --- closing
s = add("CLOSING")
fill(s, 102, "What we decided")
fill(s, 103, "Three standards: first touch inside an hour, six touches before a lead closes, same-day service response. Baseline in October, on the scorecard by January.")
fill(s, 104, "Nate Guralski  |  VP Sales & Marketing  |  Steel King Industries")
notes(s, "Restate the three standards and the first date. Then stop.")

prs.save(OUT)
print("saved", OUT, "slides:", len(prs.slides))
