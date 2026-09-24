"""Build the Steel King sales meeting deck on speed to lead and customer service.

Run:  python3 build_deck.py
Outputs (same folder):
  Speed_and_Service_Win_the_Order.pptx   the deck, talk track in speaker notes
  talk-track.md                          the same talk track as a printable document
"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "..", "templates", "steel-king-template.pptx")
OUT = os.path.join(HERE, "Speed_and_Service_Win_the_Order.pptx")
TALK = os.path.join(HERE, "talk-track.md")

GREEN = RGBColor(0x3C, 0x7D, 0x4D)
INK = RGBColor(0x15, 0x18, 0x1A)
HAIR = RGBColor(0xE2, 0xE6, 0xE3)
TINT = RGBColor(0xF2, 0xF5, 0xF2)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

L = dict(COVER=1, SECTION=2, AGENDA=3, TITLE_CONTENT=4, TWO_CONTENT=5,
         THREE_CARDS=6, BIG_NUMBERS=7, CONTENT_VISUAL=8, FULL_VISUAL=9,
         COMPARISON=10, TIMELINE=11, KEY_MESSAGE=12, NEXT_STEPS=13,
         CLOSING=14, COVER_MASCOT=15)

prs = Presentation(TEMPLATE)
while len(prs.slides) > 0:
    rId = prs.slides._sldIdLst[0].rId
    prs.part.drop_rel(rId)
    del prs.slides._sldIdLst[0]

TALK_TRACK = []  # (slide number, title, say, if_asked, source)


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


def talk(slide, title, say, if_asked="", source=""):
    """Write the talk track into the speaker notes and collect it for the document."""
    n = len(prs.slides)
    parts = ["SAY", say.strip()]
    if if_asked:
        parts += ["", "IF ASKED", if_asked.strip()]
    if source:
        parts += ["", "SOURCE", source.strip()]
    slide.notes_slide.notes_text_frame.text = "\n".join(parts)
    TALK_TRACK.append((n, title, say.strip(), if_asked.strip(), source.strip()))


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


def bar_chart(slide, x, y, w, h, categories, values, series_name, font_pt=12):
    data = CategoryChartData()
    data.categories = categories
    data.add_series(series_name, values)
    gf = slide.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(x), Inches(y), Inches(w), Inches(h), data)
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
    dl.number_format = '0"%"'
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
    ca.reverse_order = True
    return ch


def remove_placeholder(slide, idx):
    sh = ph(slide, idx)
    sh._element.getparent().remove(sh._element)


def table(slide, x, y, w, h, col_w, rows, body_pt=11):
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
            cell.margin_top = Inches(0.04)
            cell.margin_bottom = Inches(0.04)
            for p in cell.text_frame.paragraphs:
                for run in p.runs:
                    run.font.name = "Arial"
                    run.font.size = Pt(body_pt if r else 11)
                    run.font.bold = (r == 0)
                    run.font.color.rgb = WHITE if r == 0 else INK
            cell.fill.solid()
            cell.fill.fore_color.rgb = GREEN if r == 0 else (TINT if r % 2 == 0 else WHITE)
    return tbl


# =====================================================================
# 1. COVER
# =====================================================================
s = add("COVER_MASCOT")
fill(s, 102, "SALES MEETING")
fill(s, 104, "Speed and service win the order")
fill(s, 105, "What the research says about first touch, follow-up, and support in industrial buying")
fill(s, 106, "Nate Guralski  |  Steel King Sales Team  |  September 2026")
talk(s, "Cover",
     """This is not a pep talk about customer service. It is what the data says about how warehouse and
     industrial buyers pick a rack supplier, and what that means for how fast we answer and how long we stay
     on a lead. Every number in here comes from a named study. I will tell you where each one comes from,
     and I will tell you which popular numbers I left out because they don't hold up. Fifty minutes. The last
     fifteen are ours to decide what we change.""",
     "",
     "Companion research brief has every citation, URL, and a confidence grade.")

# =====================================================================
# 2. KEY MESSAGE
# =====================================================================
s = add("KEY_MESSAGE")
fill(s, 101, "THE POINT")
fill(s, 103, "By the time a buyer contacts us, they have already picked a favorite. The first touch decides whether it stays us.")
fill(s, 104, "Eight in ten B2B deals go to the first vendor the buyer contacts. Most vendors take more than a day to answer.")
talk(s, "The point",
     """Here is the whole deck in one line. Buyers do their homework before they call anyone. When they do
     call, they already have a favorite, and about eight times out of ten they buy from the first vendor they
     reach. So the first touch is not the start of the sale. It is the last chance to stay in it. And the
     average company takes almost two days to make that first touch. That gap is the opportunity.""",
     """If someone says our buyers aren't like software buyers: fair, and I'll show industrial data next to the
     cross-industry data all the way through. The pattern is the same in every sample we found.""",
     "6sense 2025 Buyer Experience Report, about 4,000 B2B buyers. HBR 2011 audit of 2,241 U.S. companies.")

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
talk(s, "Agenda",
     """Five parts. First, the demand is there, so this is about share. Second, how the buyer actually
     decides. Third, what speed does to a lead. Fourth, what service does after the PO. Fifth, three
     standards I want us to agree on and a plan to measure them. Hold questions on the standards until the
     end. Everything before that is evidence.""")

# =====================================================================
# SECTION 01: MARKET
# =====================================================================
s = add("SECTION")
fill(s, 101, "01")
fill(s, 103, "The market is showing up")
fill(s, 104, "Demand for rack is real. The question is who gets the order.")
talk(s, "Section 01",
     """Quick market check first, because none of this matters if nobody is buying rack. They are.""")

s = add("BIG_NUMBERS")
big(s, "MARKET", "Racking is the top planned equipment buy",
    "Buyers have budget and intent. We are competing for share, not for demand.",
    [("45%", "PLAN TO BUY RACKS AND SHELVING", "Top equipment category in the 2026 and 2025 MMH outlook surveys"),
     ("$542K", "AVG PLANNED MH EQUIPMENT SPEND", "Up from $402K planned for 2025, a 35% jump"),
     ("+18%", "U.S. INDUSTRIAL CONSTRUCTION PIPELINE", "305M sq ft under construction, Q2 2026, fourth straight quarterly rise")],
    "Every one of these buyers will shortlist fewer than five suppliers. Our job is to be on the list and stay on it.",
    "Modern Materials Handling / Peerless Research 2026 and 2025 Outlook Surveys (n=103, n=110); Cushman & Wakefield U.S. Industrial MarketBeat Q2 2026.")
talk(s, "Racking is the top planned equipment buy",
     """Modern Materials Handling surveys warehouse and DC operators every January on what they plan to
     buy. For the second year running, racks and shelving is the number one category. Forty-five percent
     say they plan to buy. Average planned spend on equipment and systems jumped from about four hundred
     thousand to five hundred forty-two thousand. That is a thirty-five percent increase in one year.
     On the real estate side, Cushman and Wakefield has three hundred five million square feet of
     industrial space under construction, up eighteen percent year over year and rising four quarters
     straight. Every one of those buildings needs storage. The point: demand is not our problem. Share is.""",
     """Sample sizes are small, about a hundred operators each year, but the survey has run for two decades
     and the direction is consistent. MHI's 2026 report backs it: 56% of supply chain leaders are increasing
     investment and 52% plan to spend over a million. Cold storage and data center adjacent manufacturing
     are the two hottest pockets right now.""",
     """MMH/Peerless 2026 Outlook Survey, Jan 2026, n=103. MMH/Peerless 2025 Outlook, n=110. Cushman &
     Wakefield U.S. Industrial MarketBeat, Q2 2026. MHI/Deloitte 2026 Annual Industry Report.""")

# =====================================================================
# SECTION 02: THE BUYER DECIDES EARLY
# =====================================================================
s = add("SECTION")
fill(s, 101, "02")
fill(s, 103, "The buyer decides early")
fill(s, 104, "Most of the decision is made before anyone calls a rep.")
talk(s, "Section 02",
     """Now, how the buyer decides. This is where the industrial data and the broad B2B data say the
     same thing.""")

s = add("BIG_NUMBERS")
big(s, "BUYER BEHAVIOR", "The shortlist is built before we get the call",
    "Buyers research alone, keep the list short, and mostly buy from the first vendor they contact.",
    [("62%", "OF THE BUYING PROCESS DONE ONLINE", "Engineers and technical buyers, before contacting anyone at the vendor"),
     ("71%", "VET FEWER THAN FIVE SUPPLIERS", "Industrial buyers on Thomasnet; only 3% look at more than ten"),
     ("8 in 10", "DEALS GO TO FIRST VENDOR CONTACTED", "B2B buyers across industries; favorites ranked before first contact")],
    "If the first call, form, or dealer inquiry goes unanswered for a day, we hand the order to whoever answered.",
    "GlobalSpec / TREW Marketing, 2026 State of Marketing to Engineers (n=1,000+); Thomasnet Industrial Buyer Habits report; 6sense 2025 Buyer Experience Report (approx. 4,000 B2B buyers).")
talk(s, "The shortlist is built before we get the call",
     """Three numbers. GlobalSpec surveys a thousand engineers and technical buyers every year. Sixty-two
     percent of the buying process is finished online before they contact anyone at the vendor. For buyers
     under thirty-five it is sixty-six percent. Thomasnet, the industrial sourcing platform, found seventy-one
     percent of industrial buyers look at fewer than five suppliers. Only three percent look at more than ten.
     And 6sense surveyed four thousand B2B buyers: ninety-four percent had ranked a favorite before they
     contacted anyone, and the first vendor they contacted won about eight times out of ten. Put those
     together. The buyer has done the research, the list is short, and the first company that picks up the
     phone usually gets the order. If our reply sits for a day, that first company is not us.""",
     """The 6sense sample skews to technology purchases. The GlobalSpec and Thomasnet samples are engineers
     and industrial buyers. I put them side by side on purpose: the industrial buyer is doing the same
     thing. Gartner adds that buyers spend only seventeen percent of the process talking to any supplier,
     and any one rep gets five or six percent of that time. Every minute of contact has to count.""",
     """GlobalSpec/TREW 2026 State of Marketing to Engineers. Thomasnet Industrial Buyer Habits. 6sense 2025
     Buyer Experience Report. Gartner B2B Buying Journey.""")

# --- chart: how distributors win business
s = add("CONTENT_VISUAL")
header(s, "WHAT BUYERS WEIGH", "Lead time and support outrank price",
       "Industrial buyers rank lead time first when they shortlist. Price is fifth on why distributors win.",
       "Thomasnet survey of 400+ registered industrial buyers (2021); Industrial Distribution annual Survey of Distributor Operations.", 105)
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
talk(s, "Lead time and support outrank price",
     """This is the industrial buyer specifically. Thomasnet asked four hundred registered industrial buyers
     what they weigh when they shortlist a new supplier. Number one was availability and lead time. Price
     was second. Quality certs third. Verified company information fourth. They also said they expect an
     answer in under twenty-four hours, and forty-four percent said that outright. One buyer put it this
     way, and I want you to remember the quote: if I have to follow up twice, I eliminate you from
     consideration. Their reasons for rejecting a supplier were slow or no RFQ response, having to chase,
     unclear pricing, and stale company information. On the right is the distributor side. Industrial
     Distribution asks distributors every year why they win business. Relationships, availability, technical
     support, and delivery all beat price. Price is fifth. Our dealers know this. The buyer is telling us
     the same thing from the other side.""",
     """The Thomasnet survey published rankings, not percentages, for the shortlist factors. The chart is
     self-reported by distributors, so it is what they believe wins, not what buyers said. The two
     sources agree, which is why both are here.""",
     """Thomasnet, What Industrial Buyers Care About Most When Shortlisting New Suppliers, 2021, 400+ buyers.
     Thomasnet Industrial Buyer Search Habits survey (44%). Industrial Distribution Survey of Distributor
     Operations.""")

# --- two content: what buyers want vs what they get
s = add("TWO_CONTENT")
header(s, "THE REP", "Buyers want a rep who knows their business",
       "Human help cuts buyer regret in half. Generic help gets us removed from the list.",
       "GlobalSpec / TREW 2025 and 2026 State of Marketing to Engineers; Salesforce State of the Connected Customer 2023 (3,300 business buyers); Gartner B2B Buying Report (2022 survey, n=441); LinkedIn State of Sales 2021.", 109)
fill(s, 103, "WHAT BUYERS WANT")
fill_lines(s, 105, [
    "Only 6% of technical buyers want no salesperson. They call for technical complexity (47%) and pricing or inventory (40%).",
    "66% of engineers rate the vendor's engineering experts as highly trustworthy.",
    "86% of business buyers are more likely to buy when the company understands their goals.",
    "Rep-assisted purchases carry half the regret of self-service buys: 21% vs 43%.",
])
fill(s, 106, "WHAT BUYERS REPORT GETTING")
fill_lines(s, 108, [
    "59% say most reps don't take the time to understand them.",
    "73% say most sales interactions feel transactional.",
    "69% see conflicts between the website and what the rep says.",
    "Top deal killers: misleading price or product info (48%), not understanding the buyer's company (44%).",
])
talk(s, "Buyers want a rep who knows their business",
     """The self-serve story gets overplayed. Engineers and technical buyers still want a person. Only six
     percent say they never want to talk to a salesperson. They call us for two reasons: the application is
     technically complex, or they need pricing and inventory, meaning lead time. Two thirds rate the vendor's
     engineers as highly trustworthy. That is our application engineers and our regional managers. Gartner
     found something I think about a lot: buyers who purchased self-serve with no rep regretted it forty-three
     percent of the time. Buyers who had a rep alongside digital tools regretted it twenty-one percent. A rep
     cuts regret in half. Now the right side. Fifty-nine percent of business buyers say most reps don't take
     time to understand them. Seventy-three percent say the interaction felt transactional. Sixty-nine percent
     found the website and the rep telling them different things. The two deal killers buyers name most:
     misleading information on price or product, and not understanding the buyer's company. So the buyer
     wants us in the deal. They just want us to show up prepared.""",
     """The left column mixes industrial and cross-industry sources. The 6%, 47%, 40%, and 66% are engineers
     and technical buyers. The 86% and the regret numbers are B2B across industries. The right column is
     all cross-industry B2B. Nobody has run the right-column questions on industrial buyers specifically.""",
     """GlobalSpec/TREW 2025 (6%, 47%, 40%) and 2026 (66%). Salesforce State of the Connected Customer 2023,
     3,300 business buyers (86%, 59%, 73%). Gartner B2B Buying Report, 2022 buyer survey (21% vs 43%).
     Gartner 2025 sales survey via Salesforce (69%). LinkedIn State of Sales 2021 (48%, 44%).""")

# =====================================================================
# SECTION 03: SPEED TO LEAD
# =====================================================================
s = add("SECTION")
fill(s, 101, "03")
fill(s, 103, "Speed to lead")
fill(s, 104, "The first hour is worth more than the next week.")
talk(s, "Section 03",
     """Now speed. This is the most studied question in sales research, and the answer has not changed in
     fifteen years.""")

s = add("BIG_NUMBERS")
big(s, "RESPONSE TIME", "An hour's delay cuts qualification odds 7x",
    "The value of a lead decays by the hour. Most companies let it decay for nearly two days.",
    [("7x", "MORE LIKELY TO QUALIFY", "Contact within one hour vs. waiting one more hour"),
     ("60x", "MORE LIKELY TO QUALIFY", "Contact within one hour vs. waiting 24 hours or more"),
     ("42 hrs", "AVERAGE RESPONSE TIME", "Across 2,241 U.S. companies audited; 23% never responded")],
    "A lead that sits overnight is not the same lead in the morning. Speed is the cheapest advantage we have.",
    "Oldroyd, McElheran, Elkington, \"The Short Life of Online Sales Leads,\" Harvard Business Review, March 2011 (1.25M leads, 42 companies; audit of 2,241 companies).")
talk(s, "An hour's delay cuts qualification odds 7x",
     """The cleanest study on this is Harvard Business Review, 2011. Two datasets. First, one and a quarter
     million leads at forty-two companies, both B2C and B2B. Companies that tried to reach a lead within an
     hour were nearly seven times as likely to qualify it as companies that waited even one more hour.
     Compared to companies that waited a day, sixty times. Second, they audited two thousand two hundred
     forty-one U.S. companies by submitting a lead and timing the response. Average response, among
     companies that responded at all, was forty-two hours. Twenty-three percent never responded. Think about
     what forty-two hours means for a rack inquiry. The buyer filled out our form Tuesday morning. We call
     back Thursday. In between, they heard from two competitors and maybe a dealer. The lead is not dead,
     but it is no longer ours to lose. It is ours to win back.""",
     """Why not the famous 21x number? That is from a 2007 study sponsored by InsideSales.com, done by an MIT
     fellow. It gets called the MIT study. It was never published by MIT. HBR is peer-edited and the
     sample is bigger, so I use HBR's 7x and 60x. If someone quotes 47 hours, HBR says 42. The 47 is a
     transcription error that has been copied for a decade.""",
     "Oldroyd, McElheran, Elkington, The Short Life of Online Sales Leads, Harvard Business Review, March 2011.")

# --- chart: how companies respond
s = add("CONTENT_VISUAL")
header(s, "THE GAP", "Most companies take a day or more to answer",
       "Slow response is the norm. That makes fast response a differentiator.",
       "Aleran survey of 200 U.S. manufacturers and distributors, July 2025 (vendor-sponsored); Workato lead response study, March 2026 (114 B2B companies); InsideSales 2021 Lead Response Management (5.7M leads); HBR 2011 audit.", 105)
fill_lines(s, 103, [
    "Manufacturers and distributors: 71% take a day or more to produce a quote. 88% say they have lost deals to slow, manual quoting.",
    "2026 test of 114 B2B companies: one sent a personalized reply within five minutes. None called within five minutes. Only 31% called at all.",
    "Platform data across 5.7M leads: 77% of leads never got a response. Reps average 1.3 call attempts before giving up.",
])
remove_placeholder(s, 104)
bar_chart(s, 6.50, 2.32, 6.03, 4.28,
          ["Within 1 hour", "1 to 24 hours", "More than 24 hours", "Never responded"],
          [37, 16, 24, 23], "Share of companies, HBR 2011 audit")
talk(s, "Most companies take a day or more to answer",
     """Has it gotten better since 2011? No. Start with our own industry. A 2025 survey of two hundred U.S.
     manufacturers and distributors found seventy-one percent take a day or more to produce a quote, and
     eighty-eight percent admit they have lost deals to slow, manual quoting. In March of this year, Workato
     sent demo requests to a hundred fourteen B2B companies. One of them sent a personalized reply within
     five minutes. Zero called within five minutes. Only thirty-one percent ever called. InsideSales looked
     at five point seven million leads across their platform: seventy-seven percent never got a response at
     all, and the average rep made one point three call attempts before moving on. The chart is the HBR
     audit: thirty-seven percent inside an hour, twenty-three percent never. Here is why this is good news.
     Fast response is not table stakes. Almost nobody does it. The bar is on the floor. If we answer a rack
     inquiry the same morning, we are already ahead of two thirds of the market.""",
     """The Aleran survey is sponsored by a quoting-software vendor, so treat the exact percentages with care.
     The Workato and InsideSales samples are heavy on software companies. HBR's sample was mixed B2C and
     B2B. Every audit, in every industry, for fifteen years, lands in the same place.""",
     """Aleran Built to Sell survey, July 2025, n=200. Workato, What We Learned from 114 Companies, March
     2026. InsideSales 2021 Lead Response Management, 5.7M leads. HBR 2011.""")

# --- persistence
s = add("TITLE_CONTENT")
header(s, "PERSISTENCE", "Most reps quit after one attempt",
       "Most deals need several touches. The second through sixth is where reps stop and conversions live.",
       "Velocify / Leads360 Ultimate Contact Strategy 2012 (3.5M leads, consumer-heavy verticals); RAIN Group, Top Performance in Sales Prospecting, 2018 (488 buyers, 489 sellers); InsideSales 2013 Lead Response Report; Thomasnet 2021.", 104)
fill_lines(s, 103, [
    "Half of all leads are called only once. 93% of leads that convert are reached by the sixth call.",
    "A six-call cadence lifted conversion 49%. Adding a five-email cadence lifted it 128% combined.",
    "Average rep makes 1.3 call attempts before moving on.",
    "New accounts take 8 touches on average to get a meeting. Top performers get there in 5 because each touch carries something useful.",
    "Buyers told RAIN what they will tolerate: 55% accept 2 to 4 contacts, 23% accept 5 to 10. Persistence is expected, not resented.",
    "Persistence runs one direction. We chase the lead. The buyer never chases us: \"If I have to follow up twice, I eliminate you.\"",
])
talk(s, "Most reps quit after one attempt",
     """Speed gets you the first conversation. Persistence gets you the order. Velocify studied three and a
     half million leads. Half were called exactly once. Ninety-three percent of the leads that eventually
     converted were reached by the sixth call. A six-call cadence lifted conversion forty-nine percent.
     Add a five-email sequence and it was one hundred twenty-eight percent combined. Average rep across the
     industry makes one point three attempts. RAIN Group asked both buyers and sellers. The average rep
     needs eight touches to get a meeting with a new account. Top performers need five, because each of
     their touches carries something the buyer can use: a drawing, a lead time, a reference, a code
     question answered. And buyers are fine with it. Fifty-five percent said two to four contacts is
     reasonable, another twenty-three percent said five to ten. Last line matters. Persistence runs one
     direction. We chase the lead. The buyer never has to chase us. That is the Thomasnet quote again.""",
     """The Velocify data is mortgage, insurance, and education phone sales, not industrial. I am quoting
     the pattern, not defending the exact percentages for rack. The RAIN data is B2B and self-reported.
     The InsideSales 1.3 attempts is platform data across 8,000 companies. Three different methods, same
     shape: most reps stop early, most conversions happen later.""",
     """Velocify/Leads360 Ultimate Contact Strategy, Dec 2012. RAIN Group Top Performance in Sales
     Prospecting, 2018. InsideSales 2014 Lead Response Report. Thomasnet 2021.""")

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
talk(s, "The popular follow-up stats don't hold up",
     """One minute on this because some of you have seen these numbers in a training deck or on LinkedIn.
     Eighty percent of sales need five follow-ups and forty-four percent of reps quit after one. That
     traces to a 1942 survey of fewer than forty door-to-door salesmen on Long Island. Seventy-eight
     percent buy from the first responder. Attributed to a survey nobody can find. Thirty-five to fifty
     percent go to the first responder. A vendor blog, no study. The MIT study. Vendor-sponsored, never
     published by MIT. The dissatisfied customer tells nine to fifteen people. Attributed to a White House
     office that never published it. I checked all of them and pulled them. Here is why I care. We are
     about to set standards. If the standard rests on a made-up number, the first skeptic who Googles it
     wins the argument. Everything on these slides survives that Google search.""",
     """If this slide feels like a detour, skip it and reference the brief. But if anyone quotes one of these
     numbers back at us later, this is the answer.""",
     """Conciergr and Sales & Marketing Executives International trace of the 80%/5-touches claim, 2021.
     Full trace for each item in the research brief, section 8.""")

# =====================================================================
# SECTION 04: SERVICE AFTER THE SALE
# =====================================================================
s = add("SECTION")
fill(s, 101, "04")
fill(s, 103, "Service after the sale")
fill(s, 104, "Service response is now a purchase criterion, not a cost center.")
talk(s, "Section 04",
     """The PO is not the finish line. Service after the sale is now part of how buyers pick the next
     supplier.""")

s = add("BIG_NUMBERS")
big(s, "SERVICE", "Service response is now a buying criterion",
    "Fast service response now ranks next to uptime. Poor service is the #2 reason buyers leave.",
    [("95%", "RATE FAST SERVICE RESPONSE VERY IMPORTANT", "Warehouse automation buyers, 2026; up from 83% a year earlier"),
     ("41%", "OF BUSINESS BUYERS LEFT OVER POOR SERVICE", "Second only to price (65%) as a reason to stop buying"),
     ("50%", "WOULD SWITCH FOR BETTER POST-SALE SUPPORT", "U.S. industrial buyers; 87% of millennial buyers say so")],
    "The install crew, the parts desk, and the engineer who calls back are part of the product. Buyers score them that way.",
    "Modern Materials Handling / Peerless 2026 Automation Study (120+ decision makers) and 2025 Automation Survey (n=139); Salesforce State of the AI Connected Customer 2024 (1,570 business buyers); UPS Industrial Buying Dynamics 2017 (1,500 U.S. buyers) and 2019.")
talk(s, "Service response is now a buying criterion",
     """Three numbers, and the first one is ours. Modern Materials Handling asks warehouse decision makers
     what is very important when they evaluate equipment. In 2026, ninety-five percent said fast service
     response times. A year earlier it was eighty-three percent. That is a twelve-point jump in one year,
     and it now sits right next to durability and uptime at ninety-two percent. Purchase price was
     seventy-eight. Service response outranks price. Second number, Salesforce asked business buyers why
     they stopped buying from a brand in the past year. Price was first at sixty-five percent. Poor customer
     service was second at forty-one percent. Third, UPS surveyed fifteen hundred U.S. industrial buyers.
     Half said they would switch to a supplier that offered better post-sale support: returns, training,
     on-site help. Among millennial buyers, who are now the majority of purchasing managers, eighty-seven
     percent. For a rack buyer, service is the install crew showing up when promised, the parts desk
     answering the same day, and the engineer calling back on a damaged upright. That is the product now.""",
     """The MMH 2026 study is automation buyers, not rack buyers specifically, but it is the same warehouse
     decision maker. The MMH 2025 lift truck survey puts service response time in the top five purchase
     factors, and the MHEDA Journal quotes Hytrol's CRO saying response for uptime-critical customers has
     to be within minutes, not hours.""",
     """MMH/Peerless 2026 Automation Study (95%, 92%, 78%) and 2025 Automation Survey (83%). Salesforce State
     of the AI Connected Customer, 7th ed., 2024, 1,570 business buyers (41%, 65%). UPS Industrial Buying
     Dynamics 2017 (50%) and 2019 (87%).""")

s = add("TWO_CONTENT")
header(s, "THE MATH", "Retention pays twice. Churn costs for years.",
       "Aftermarket carries double the margin of equipment, and a burned B2B customer stays gone.",
       "Deloitte 2026 Manufacturing Industry Outlook (n=600); Salesforce Trends in Manufacturing 2024 (n=830); Gartner CSO survey May 2025 (n=243); Accenture 2022 and 2019 B2B studies; Dimensional Research 2013 (n=1,046, B2B split); Gallup 2016.", 109)
fill(s, 103, "WHAT SERVICE EARNS")
fill_lines(s, 105, [
    "Aftermarket services deliver margins more than two times higher than equipment sales alone.",
    "97% of manufacturers are changing their service and aftermarket operations. Of those tracking satisfaction, only 44% are hitting their goals.",
    "73% of chief sales officers now prioritize growth from existing customers over new logos.",
    "Companies that run service as a value center grow revenue 3.5x faster than those that run it as a cost center.",
])
fill(s, 106, "WHAT POOR SERVICE COSTS")
fill_lines(s, 108, [
    "80% of frequent B2B buyers switched a supplier in 24 months because it could not meet their expectations.",
    "66% of B2B customers stopped buying after a bad service interaction.",
    "51% of them were still avoiding that vendor two or more years later.",
    "Only 29% of B2B customers are fully engaged. Gallup calls the other 71% ready to take their business elsewhere.",
])
talk(s, "Retention pays twice. Churn costs for years.",
     """Left side, what service earns. Deloitte's 2026 manufacturing outlook: aftermarket services carry
     margins more than twice what equipment sales carry. Salesforce surveyed eight hundred thirty
     manufacturers: ninety-seven percent are changing how they run service and aftermarket, and of the
     ones that track satisfaction, only forty-four percent are hitting their goals. So everyone knows it
     matters and most are not there yet. Gartner: seventy-three percent of chief sales officers now put
     existing-customer growth ahead of new logos. Accenture: companies that treat service as a value center
     grow revenue three and a half times faster than companies that treat it as a cost. Right side, what
     poor service costs. Eighty percent of frequent B2B buyers switched a supplier in the last two years
     because it could not meet expectations. Sixty-six percent of B2B customers stopped buying after one bad
     service interaction, and half of those were still avoiding that vendor two years later. Gallup: only
     twenty-nine percent of B2B customers are fully engaged. The other seventy-one percent are, in their
     words, ready to take their business elsewhere. For us: a rack customer who had a good install and a
     fast parts response buys the next building from us without a bid. One who didn't tells the dealer.""",
     """The 66% and 51% are from 2013, the only study with a clean B2B versus B2C split on service-driven
     buying. Gallup is 2016. The rest is 2022 to 2026. The old numbers are flagged as dated in the brief.
     The rule of thumb that selling to an existing customer is 60 to 70% likely versus 5 to 20% for a new
     prospect is a textbook figure, not a study, so it is not on the slide.""",
     """Deloitte 2026 Manufacturing Industry Outlook. Salesforce Trends in Manufacturing 2024. Gartner CSO
     survey, May 2025. Accenture End-to-Endless Customer Service 2022. Accenture Interactive 2019.
     Dimensional Research for Zendesk 2013. Gallup Guide to Customer Centricity 2016.""")

s = add("TITLE_CONTENT")
header(s, "THE CHANNEL", "Our response time becomes the dealer's",
       "Dealers are being asked for contractor-grade service. Our response time is theirs.",
       "MMH / Peerless 2025 Lift Truck User Survey (n=150); MHEDA Journal 2025 Industry Outlook, 2026 Business Trends, and \"Advancing from Distributor to Integrator\"; UPS Industrial Buying Dynamics 2017 (1,500 U.S. buyers).", 104)
fill_lines(s, 103, [
    "86% of lift truck spend runs through dealers. Under 15% goes direct. Rack is not measured the same way, but the channel is the same.",
    "MHEDA 2025: storage and handling customers now expect \"contractor-like service capabilities\" from the dealer.",
    "MHEDA 2026 trends: rising customer expectations squeezing margins; manufacturer-distributor partnerships \"becoming essential.\"",
    "Hytrol's CRO to MHEDA members: for uptime-critical end users, response time has to be \"within minutes, not hours.\"",
    "Younger buyers research online and expect to buy online. 80% of industrial buyers would move to a supplier with a more usable web presence.",
    "Every quote, drawing, and parts question a dealer sends us has an end customer waiting behind it.",
])
talk(s, "Our response time becomes the dealer's",
     """Most of what we sell goes through a dealer, so let's talk about the dealer's clock. The best channel
     number in material handling is lift trucks: eighty-six percent of spend runs through dealers, under
     fifteen percent direct. Nobody publishes the same split for rack, but we know our own mix. MHEDA's 2025
     outlook for the storage and handling segment says customers now expect contractor-like service
     capabilities from the dealer. Its 2026 trends list has rising customer expectations squeezing dealer
     margins, and manufacturer-distributor partnerships becoming essential. Hytrol's chief revenue officer
     told MHEDA members that for uptime-critical end users, response has to be within minutes, not hours.
     And eighty percent of industrial buyers say they would move to a supplier with a more usable web
     presence, which is a marketing item for John and me. Here is the connection. When a dealer sends us a
     quote request, a drawing question, or a parts inquiry, there is an end customer on the other side of
     that dealer holding the twenty-four-hour clock from slide eight. The dealer cannot be faster than we
     are. Our response time is their response time.""",
     """If someone asks for a rack-specific channel share: it does not exist publicly. MHEDA's distributor
     statistics are member-only. Use our own bookings mix.""",
     """MMH/Peerless 2025 Lift Truck User Survey. MHEDA Journal 2025 Industry Outlook. MHEDA 2026 Material
     Handling Business Trends. MHEDA Journal, Advancing from Distributor to Integrator. UPS 2017.""")

# =====================================================================
# SECTION 05: WHAT WE CHANGE
# =====================================================================
s = add("SECTION")
fill(s, 101, "05")
fill(s, 103, "What we change")
fill(s, 104, "Three standards, measured, owned, and on the scorecard.")
talk(s, "Section 05",
     """That is the evidence. Now what we do with it. First, where it lands in our own process.""")

# --- Steel King touchpoints
s = add("FULL_VISUAL")
header(s, "AT STEEL KING", "Where this shows up in our own process",
       "Six touchpoints we control. Each has a number behind it.",
       "Sources as cited on the slides above.", 104)
remove_placeholder(s, 103)
rows = [
    ["Touchpoint", "What the buyer expects", "The number behind it", "Source"],
    ["Web RFQ or contact form", "A human reply the same day", "7x more likely to qualify inside an hour. 44% expect an answer within 24 hours.", "HBR 2011; Thomasnet"],
    ["Dealer quote request", "Price and lead time fast enough to hold their customer", "Lead time is the #1 shortlist factor. 71% of manufacturers take a day or more to quote.", "Thomasnet 2021; Aleran 2025"],
    ["Application and drawings", "A rep and an engineer who know the job", "66% rate vendor engineers highly trustworthy. Rep-assisted buys carry half the regret.", "GlobalSpec 2026; Gartner"],
    ["Follow-up on an open quote", "We chase. They never have to.", "93% of conversions by the 6th call. \"Follow up twice and I eliminate you.\"", "Velocify 2012; Thomasnet 2021"],
    ["Install, parts, and service", "Same-day response with a named owner", "95% rate fast service response very important. Poor service is the #2 reason buyers leave.", "MMH 2026; Salesforce 2024"],
    ["Existing accounts", "Contact between POs, not just at PO time", "73% of CSOs prioritize existing-customer growth. Aftermarket margins 2x equipment.", "Gartner 2025; Deloitte 2026"],
]
table(s, 0.80, 2.32, 11.73, 4.20, [2.10, 3.05, 4.68, 1.90], rows, body_pt=11)
talk(s, "Where this shows up in our own process",
     """Six places this research lands in our building. Web RFQs and contact forms: the buyer expects a
     human the same day, and the HBR data says an hour. Dealer quote requests: lead time is the number one
     shortlist factor, and most manufacturers take a day or more to quote. If we quote same day, the dealer
     wins the shortlist. Application and drawings: buyers trust the vendor's engineers more than anyone, and
     a rep in the deal cuts regret in half. Follow-up on open quotes: we chase, they never chase. Install,
     parts, and service: ninety-five percent rate response time very important, and poor service is the
     second reason buyers leave. Existing accounts: aftermarket margins are double equipment, and most CSOs
     are now prioritizing the accounts they already have. I want to hear from you on this one. Which of these
     six is our weakest today? Where does a lead or a dealer question sit longest? That is what the baseline
     in October will tell us, but you already know.""",
     """This is a discussion slide. Let it run a few minutes. Write down what the room says the weakest
     touchpoint is. It goes into the November standards conversation.""",
     "All sources as cited on the preceding slides.")

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
talk(s, "Three standards we hold ourselves to",
     """Three standards. I am proposing them, not announcing them, and I want the room to push on the
     numbers. One: first touch inside one business hour. Every web form, every dealer inquiry, every RFQ
     gets a human acknowledgment within the hour and a real answer or quote within twenty-four. That comes
     straight from the seven-x finding and Thomasnet's twenty-four-hour expectation. Two: six touches before
     we close a lead. Nobody marks a lead dead after one call. Six touches across phone and email over two
     weeks, and every touch carries something useful. That is Velocify's ninety-three percent by the sixth
     call and RAIN's five to eight touches. Three: same-day service response. Parts, install, post-sale
     questions get a same-day reply and a named owner. Dealer questions inside four business hours. That is
     the ninety-five percent from MMH and the dealer clock. And all three get reported monthly next to
     bookings, because what we don't measure, we don't do. What gets in the way of each of these today?
     That is the list we fix.""",
     """Expected pushback: we don't have the people to answer in an hour. Answer: acknowledgment in an hour
     is a text or a two-line email that says who owns it and when they'll have an answer. The quote gets
     twenty-four hours. Second pushback: six touches will annoy dealers. Answer: RAIN's buyers said two to
     ten contacts is normal. The rule is every touch carries something new.""",
     """HBR 2011 (1 hour). Thomasnet 2021 (24 hours). Velocify 2012 and RAIN 2018 (six touches). MMH 2026
     and MHEDA (same-day service).""")

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
talk(s, "Measure first, then set the bar",
     """Four steps. October, baseline. We pull ninety days of CRM data on time to first touch and touches
     per lead, and we secret-shop ourselves the way HBR did: submit a web lead, call the main line, send a
     dealer-style RFQ, and time the response. No judgment, just the number. November, we agree the standards
     with the baseline in front of us and build the six-touch cadence and templates. December, lead routing
     with alerts and a response-time report by rep and by dealer. First quarter, it goes on the sales
     scorecard and we review it monthly in L10. I own the baseline and the scorecard. Regional managers own
     the cadence. Marketing and inside sales own the tooling.""",
     """If someone asks why not set the standard now: because we don't know if we are at four hours or forty.
     The standard should be a stretch from where we are, not a guess.""",
     "")

s = add("KEY_MESSAGE")
fill(s, 101, "THE POINT")
fill(s, 103, "Nobody buys rack because we called back fast. They stop considering us because we didn't.")
fill(s, 104, "Speed and service don't win on their own. They keep us on the shortlist long enough for product and price to win.")
talk(s, "The point, restated",
     """Last thought before the action list. Nobody has ever bought a rack system because we called back
     fast. Speed and service do not win the order by themselves. What they do is keep us on the shortlist
     long enough for the product, the engineering, and the price to win. And the research says the
     shortlist is short, it is built early, and it closes fast. We just have to be on it when it closes.""")

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
table(s, 0.80, 2.32, 11.73, 4.10, [6.13, 2.20, 1.60, 1.80], rows, body_pt=12)
talk(s, "Next steps",
     """Six actions. I take the secret shop and the scorecard. Inside sales pulls the CRM data. Regional
     managers draft the cadence. We agree the standards at the November meeting. John and inside sales build
     the dashboard by mid-December. Names and dates are proposals. Let's fix them now before we leave.""",
     """Confirm real names for inside sales lead and regional managers in the room. Adjust dates to the
     actual November meeting date.""")

# --- closing
s = add("CLOSING")
fill(s, 102, "What we decided")
fill(s, 103, "Three standards: first touch inside an hour, six touches before a lead closes, same-day service response. Baseline in October, on the scorecard by January.")
fill(s, 104, "Nate Guralski  |  VP Sales & Marketing  |  Steel King Industries")
talk(s, "Closing",
     """Three standards. First touch inside an hour. Six touches before a lead closes. Same-day service
     response. Baseline in October. On the scorecard by January. Thanks.""")

prs.save(OUT)

# ---------- talk track document ----------
with open(TALK, "w") as f:
    f.write("# Talk track: Speed and service win the order\n\n")
    f.write("Steel King sales meeting, September 2026. Spoken script per slide, with the pushback to expect "
            "and the source to cite aloud. The same text is in the speaker notes of the PPTX.\n\n")
    f.write("Timing: 50 minutes. Sections 01 to 04 are evidence, about 35 minutes. Section 05 is discussion, 15 minutes.\n\n")
    for n, title, say, if_asked, source in TALK_TRACK:
        f.write(f"---\n\n## Slide {n}: {title}\n\n")
        f.write("**Say**\n\n" + " ".join(say.split()) + "\n\n")
        if if_asked:
            f.write("**If asked**\n\n" + " ".join(if_asked.split()) + "\n\n")
        if source:
            f.write("**Source to cite**\n\n" + " ".join(source.split()) + "\n\n")
print("saved", OUT, "slides:", len(prs.slides))
print("saved", TALK)
