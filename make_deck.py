"""
Builds CARVERSE_Pitch.pptx — the hackathon submission deck.
Covers the 4 required deliverables (Game Design, Action Mapping, Anti-Gaming,
Demo) + team collaboration, all backed by the real event-log numbers.

    python make_deck.py  ->  CARVERSE_Pitch.pptx   (open & edit in PowerPoint)
"""
from pptx import Presentation
from pptx.util import Inches as In, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import scoring_engine as E

# ---- palette ----
INK   = RGBColor(0x16, 0x22, 0x2E)   # near-black navy
RED   = RGBColor(0xE6, 0x39, 0x46)   # racing red
EMER  = RGBColor(0x0A, 0x9D, 0x8E)   # emerald
AMBER = RGBColor(0xF4, 0xA2, 0x61)
GREY  = RGBColor(0x5C, 0x66, 0x70)
LGREY = RGBColor(0xED, 0xF1, 0xF4)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
MINT  = RGBColor(0xE4, 0xF3, 0xF0)
FONT  = "Segoe UI"

# ---- data ----
ev, emp, loc, loc_name = E.load_data()
se, state = E.compute_scores(ev)
ind, dept, branch = E.leaderboards(se, loc_name)
summ = E.summary(ev, se, state)
state["branch"] = state["location_code"].map(loc_name).fillna(state["location_code"])
NOISE = round(100 * summ["noise_events"] / summ["total_events"])
CLEAN = int((state["delivered"] & (state["rework"] == 0)).sum())
TOP_IND = [(r["username"], r["department"], int(r["realized"])) for _, r in ind.head(5).iterrows()]
DEPTL = [(r["department"], int(r["realized"])) for _, r in dept.iterrows()]
prog = (state.groupby("branch")["delivered"].mean() * 100).round(1)
cars = state.groupby("branch").size()
BR = sorted([(b, prog[b]) for b in prog.index if cars[b] >= 20], key=lambda x: -x[1])[:5]

prs = Presentation()
prs.slide_width, prs.slide_height = In(13.333), In(7.5)
BLANK = prs.slide_layouts[6]
W = 13.333


def slide(bg=WHITE):
    s = prs.slides.add_slide(BLANK)
    r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    r.fill.solid(); r.fill.fore_color.rgb = bg; r.line.fill.background()
    r.shadow.inherit = False
    s.shapes._spTree.remove(r._element); s.shapes._spTree.insert(2, r._element)
    return s


def box(s, l, t, w, h, fill=None, line=None):
    shp = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, In(l), In(t), In(w), In(h))
    if fill is None: shp.fill.background()
    else: shp.fill.solid(); shp.fill.fore_color.rgb = fill
    if line is None: shp.line.fill.background()
    else: shp.line.color.rgb = line; shp.line.width = Pt(1)
    shp.shadow.inherit = False
    return shp


def txt(s, l, t, w, h, text, size=18, color=INK, bold=False, align=PP_ALIGN.LEFT,
        font=FONT, anchor=MSO_ANCHOR.TOP, italic=False, spacing=1.0):
    tb = s.shapes.add_textbox(In(l), In(t), In(w), In(h)); tf = tb.text_frame
    tf.word_wrap = True; tf.vertical_anchor = anchor
    lines = text.split("\n")
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align; p.line_spacing = spacing
        r = p.add_run(); r.text = ln
        f = r.font; f.size = Pt(size); f.bold = bold; f.italic = italic
        f.name = font; f.color.rgb = color
    return tb


def bullets(s, l, t, w, h, items, size=17, color=INK, gap=8):
    tb = s.shapes.add_textbox(In(l), In(t), In(w), In(h)); tf = tb.text_frame
    tf.word_wrap = True
    for i, (head, body) in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(gap); p.line_spacing = 1.05
        r = p.add_run(); r.text = "▸ " + head
        r.font.size = Pt(size); r.font.bold = True; r.font.name = FONT; r.font.color.rgb = color
        if body:
            r2 = p.add_run(); r2.text = "  " + body
            r2.font.size = Pt(size); r2.font.name = FONT; r2.font.color.rgb = GREY
    return tb


def header(s, kicker, title, tcolor=INK):
    box(s, 0, 0, W, 0.28, fill=RED)  # top race stripe
    txt(s, 0.7, 0.55, 11.9, 0.4, kicker, 13, EMER, bold=True)
    txt(s, 0.7, 0.95, 11.9, 1.0, title, 33, tcolor, bold=True)


def footer(s, n):
    txt(s, 0.7, 7.05, 6, 0.3, "CARVERSE GRAND PRIX", 10, GREY, bold=True)
    txt(s, 11.6, 7.05, 1.1, 0.3, str(n), 10, GREY, align=PP_ALIGN.RIGHT)


def stat(s, l, t, w, big, cap, color=RED):
    txt(s, l, t, w, 0.9, big, 46, color, bold=True)
    txt(s, l, t + 0.85, w, 0.5, cap, 12, GREY, bold=True)


# ============================ SLIDE 1 — TITLE ============================
s = slide(INK)
box(s, 0, 0, W, 0.35, fill=RED)
box(s, 0, 7.15, W, 0.35, fill=EMER)
txt(s, 0.9, 2.0, 11.5, 0.5, "CARVERSE MOBILITY · DEALERSHIP GAMIFICATION HACKATHON", 15, EMER, bold=True)
txt(s, 0.9, 2.6, 11.8, 1.8, "CARVERSE\nGRAND PRIX", 66, WHITE, bold=True, spacing=0.95)
txt(s, 0.9, 4.9, 11.5, 0.9,
    "Turn the dealership into a team racing game — every booking is a race car,\n"
    "every department is the pit crew, and delivery is the finish line.",
    19, RGBColor(0xC8, 0xD2, 0xDC))
txt(s, 0.9, 6.4, 11.5, 0.4, "A game people actually want to play — that rewards real progress, not clicks.",
    14, AMBER, italic=True, bold=True)

# ============================ SLIDE 2 — PROBLEM ============================
s = slide()
header(s, "THE PROBLEM", "Dealership work is invisible, siloed, and unmotivating")
bullets(s, 0.7, 2.1, 7.3, 4.5, [
    ("A booking crosses 7 departments", "Sales → Finance → Accounts → EDP → PDI → CCM → RTO. Handoffs stall; nobody sees the whole race."),
    ("Effort ≠ recognition", "The people who quietly unblock a booking get no credit; loudest ≠ most valuable."),
    ("“Employee of the Month” rewards volume", "Not clean, fast, collaborative delivery — the things that actually move the business."),
    ("The data already exists", "170k+ real events are logged every 2 months — but none of it motivates anyone."),
], size=17)
box(s, 8.4, 2.1, 4.2, 4.3, fill=MINT)
txt(s, 8.7, 2.4, 3.6, 0.5, "IN THE RAW DATA", 12, EMER, bold=True)
stat(s, 8.7, 2.9, 3.6, f"{summ['total_events']:,}", "EVENTS IN 2 MONTHS", RED)
stat(s, 8.7, 4.2, 3.6, f"{summ['distinct_actions']:,}", "DISTINCT ACTION TYPES", INK)
stat(s, 8.7, 5.5, 3.6, "9", "DEPARTMENTS INVOLVED", EMER)
footer(s, 2)

# ============================ SLIDE 3 — INSIGHT ============================
s = slide()
header(s, "KEY INSIGHT", "Not all actions are equal — most are just noise")
txt(s, 0.7, 2.0, 11.9, 0.6,
    "If you reward every click, the game is instantly broken. We separated signal from noise.", 17, GREY)
# noise column
box(s, 0.7, 2.8, 5.9, 3.6, fill=LGREY)
txt(s, 1.0, 3.0, 5.4, 0.5, "❌  NOISE — scores 0", 16, GREY, bold=True)
bullets(s, 1.0, 3.6, 5.4, 2.6, [
    ("USER_LOGGED_IN", "27,764× — trivial"),
    ("DISCOUNT_UPDATED", "9,188× — just an edit"),
    ("PDI_INFO_UPDATED", "8,428× — internal"),
    ("FINANCE_INFO_UPDATED", "5,588× — effort, not outcome"),
], size=14)
# milestone column
box(s, 6.7, 2.8, 5.9, 3.6, fill=MINT)
txt(s, 7.0, 3.0, 5.4, 0.5, "✅  MILESTONES — score", 16, EMER, bold=True)
bullets(s, 7.0, 3.6, 5.4, 2.6, [
    ("BOOKING_CREATED", "a real business event"),
    ("DISCOUNT_APPROVED", "pricing finalized"),
    ("INVOICE_APPROVED_EDP", "billing complete"),
    ("GATEPASS_ISSUED", "vehicle delivered 🏁"),
], size=14)
txt(s, 0.7, 6.55, 11.9, 0.6,
    f"Result: {NOISE}% of all {summ['total_events']:,} events earn ZERO points. We only pay for crossings.",
    17, RED, bold=True)
footer(s, 3)

# ============================ SLIDE 4 — BIG IDEA ============================
s = slide(INK)
box(s, 0, 0, W, 0.28, fill=RED)
txt(s, 0.7, 0.5, 11.9, 0.4, "THE BIG IDEA", 13, EMER, bold=True)
txt(s, 0.7, 0.95, 11.9, 1.6, "Every booking is a race car.\nEvery department is the pit crew.", 34, WHITE, bold=True, spacing=1.0)
cards = [("🏁 BOOKING", "The car enters the track"), ("🔧 EACH DEPT", "does its milestone → car moves a segment"),
         ("⚡ CLEAN + FAST", "= speed & clean multipliers"), ("🏆 DELIVERED", "car crosses the line → team payout")]
x = 0.7
for i, (h, b) in enumerate(cards):
    box(s, x, 3.2, 2.85, 2.6, fill=RGBColor(0x1F, 0x2E, 0x3D), line=EMER)
    txt(s, x + 0.25, 3.55, 2.4, 0.8, h, 18, AMBER, bold=True)
    txt(s, x + 0.25, 4.45, 2.4, 1.2, b, 14, RGBColor(0xC8, 0xD2, 0xDC))
    x += 3.05
txt(s, 0.7, 6.3, 11.9, 0.8,
    "Nobody wins the finish line alone — the delivery bonus is split across every department that touched the car.",
    16, EMER, italic=True, bold=True)

# ============================ SLIDE 5 — GAMEPLAY LOOP ============================
s = slide()
header(s, "GAME DESIGN · THE CORE LOOP", "What a player does, earns, and comes back for")
steps = [("1  ACT", "Complete a real milestone on a booking", EMER),
         ("2  SCORE", "You + your department + your branch all earn points", RED),
         ("3  BOOST", "Speed × Clean multipliers reward fast, rework-free work", AMBER),
         ("4  DELIVER", "Escrow releases + co-op bonus splits across the pit crew", EMER),
         ("5  COMPETE", "Weekly branch league · quests · boss battles · seasons", RED)]
y = 2.1
for h, b, c in steps:
    box(s, 0.7, y, 0.16, 0.72, fill=c)
    txt(s, 1.05, y, 3.0, 0.72, h, 19, c, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, 4.2, y, 8.3, 0.72, b, 16, INK, anchor=MSO_ANCHOR.MIDDLE)
    y += 0.9
txt(s, 0.7, 6.75, 11.9, 0.4,
    "Progression is by clean, collaborative behaviour — not raw revenue — so the game stays fair and fun for months.",
    13, GREY, italic=True)
footer(s, 5)

# ============================ SLIDE 6 — IT IS A GAME ============================
s = slide()
header(s, "GAME DESIGN · MECHANICS", "It's a real game — not a dashboard")
feats = [("🧑 Racers & Levels", "Every employee is a player who earns XP and levels up a reputation."),
         ("👥 Teams & Pit Crews", "Departments and branches form squads with shared goals."),
         ("🎯 Quests", "“Deliver 800 cars with zero rework” — live, trackable missions."),
         ("👹 Boss Battles", "Weekly raids vs “The Bottleneck” — damage = clean deliveries."),
         ("🏢 Branch League", "24 branches race head-to-head every week; the table resets each season."),
         ("🏅 Rewards & Podium", "Medals, streaks, and real perks tied to the leaderboard.")]
x, y = 0.7, 2.1
for i, (h, b) in enumerate(feats):
    col = i % 3; row = i // 3
    bx = 0.7 + col * 4.05; by = 2.1 + row * 2.35
    box(s, bx, by, 3.8, 2.1, fill=MINT)
    txt(s, bx + 0.25, by + 0.2, 3.3, 0.6, h, 16, EMER, bold=True)
    txt(s, bx + 0.25, by + 0.85, 3.3, 1.1, b, 13, INK)
footer(s, 6)

# ============================ SLIDE 7 — ACTION MAPPING ============================
s = slide()
header(s, "DELIVERABLE · ACTION MAPPING", "The 16 scored actions + 2 derived (under the 20 cap)")
rows = []
for act, (pts, tier, d, label) in E.SCORING_ACTIONS.items():
    rows.append((tier, act, label, d, pts))
for name, (pts, tier, d, label) in E.DERIVED.items():
    rows.append((tier, name + " *", label, d, pts))
rows.sort(key=lambda r: -r[4])
n = len(rows) + 1
tbl = s.shapes.add_table(n, 5, In(0.7), In(1.95), In(11.9), In(4.9)).table
tbl.columns[0].width = In(0.9); tbl.columns[1].width = In(4.0)
tbl.columns[2].width = In(4.2); tbl.columns[3].width = In(2.0); tbl.columns[4].width = In(0.8)
hdr = ["TIER", "ACTION CODE", "MEANING", "DEPT", "PTS"]
for c, h in enumerate(hdr):
    cell = tbl.cell(0, c); cell.text = h
    cell.fill.solid(); cell.fill.fore_color.rgb = INK
    p = cell.text_frame.paragraphs[0]; p.runs[0].font.size = Pt(11)
    p.runs[0].font.bold = True; p.runs[0].font.color.rgb = WHITE; p.runs[0].font.name = FONT
tiercol = {"S": RED, "A": AMBER, "B": EMER}
for i, (tier, act, mean, d, pts) in enumerate(rows, 1):
    vals = [tier, act, mean, d, str(pts)]
    for c, v in enumerate(vals):
        cell = tbl.cell(i, c); cell.text = v
        cell.fill.solid(); cell.fill.fore_color.rgb = WHITE if i % 2 else LGREY
        p = cell.text_frame.paragraphs[0]; r = p.runs[0]
        r.font.size = Pt(10.5); r.font.name = FONT; r.font.color.rgb = INK
        if c == 0: r.font.bold = True; r.font.color.rgb = tiercol.get(tier, INK)
        if c == 4: r.font.bold = True; r.font.color.rgb = EMER
txt(s, 0.7, 6.95, 11.9, 0.3, "* derived milestone (e.g. full document set completed) — gamify the outcome, not each upload.", 10, GREY, italic=True)
footer(s, 7)

# ============================ SLIDE 8 — ANTI-GAMING ============================
s = slide()
header(s, "DELIVERABLE · ANTI-GAMING STRATEGY", "Six guardrails so the system can't be farmed")
guards = [("1 · Milestones only", "Updates, logins, notes = 0 points."),
          ("2 · Once per booking", "Each milestone scores a single time per car."),
          ("3 · Delivery escrow", "40% of points are held until the car actually delivers."),
          ("4 · Net progress", "Cancelled (−50) & dormant bookings reverse points."),
          ("5 · Rework penalty", "Escalation loops shrink the clean-multiplier to 0.6×."),
          ("6 · Co-op payout", "Delivery bonus splits across depts — no solo farming.")]
for i, (h, b) in enumerate(guards):
    col = i % 2; row = i // 2
    bx = 0.7 + col * 6.05; by = 2.1 + row * 1.5
    box(s, bx, by, 5.8, 1.35, fill=WHITE, line=EMER)
    txt(s, bx + 0.3, by + 0.15, 5.2, 0.5, h, 16, RED, bold=True)
    txt(s, bx + 0.3, by + 0.65, 5.2, 0.6, b, 13.5, GREY)
txt(s, 0.7, 6.8, 11.9, 0.5,
    "“If a user repeats an action 50× a day, do they get 50× points?”  →  No. That's the test we built for.",
    15, INK, bold=True, italic=True)
footer(s, 8)

# ============================ SLIDE 9 — COLLABORATION ============================
s = slide()
header(s, "TEAM COLLABORATION BY DESIGN", "You can't win the finish line alone")
bullets(s, 0.7, 2.1, 7.3, 4.2, [
    ("Co-op delivery payout", "Every delivered car pays a bonus split across ALL departments that cleanly touched it — weighted by contribution."),
    ("Relay bonuses", "Fast, clean handoffs between departments (e.g. Finance → Accounts) earn extra multipliers."),
    ("Branch as a squad", "24 branches compete as teams in a weekly league — one slow department drags the whole branch down."),
    ("Boss battles are raids", "The weekly boss can only be beaten by multiple departments coordinating deliveries."),
], size=16)
box(s, 8.4, 2.1, 4.2, 4.3, fill=INK)
txt(s, 8.7, 2.4, 3.6, 0.5, "WHY IT WORKS", 12, EMER, bold=True)
txt(s, 8.7, 3.0, 3.6, 3.2,
    "A sales rep is now incentivised to help Finance unblock a loan — because the delivery "
    "bonus only lands when the WHOLE car crosses the line.\n\nSelfish behaviour literally scores less.",
    15, WHITE)
footer(s, 9)

# ============================ SLIDE 10 — PROOF ============================
s = slide()
header(s, "PROVEN ON REAL DATA", "Not a mockup — scored on the actual May–Jun 2026 log")
row = [(f"{summ['bookings']:,}", "BOOKINGS RACED", INK), (f"{summ['delivered']:,}", "DELIVERED 🏁", EMER),
       (f"{NOISE}%", "NOISE IGNORED", RED), (f"{summ['users']}", "PLAYERS", AMBER)]
x = 0.7
for big, cap, c in row:
    box(s, x, 2.0, 2.85, 1.5, fill=LGREY)
    txt(s, x + 0.25, 2.15, 2.4, 0.9, big, 40, c, bold=True)
    txt(s, x + 0.25, 3.0, 2.4, 0.4, cap, 11, GREY, bold=True)
    x += 3.05
txt(s, 0.7, 3.9, 5.8, 0.4, "TOP RACERS (individual points)", 13, EMER, bold=True)
bullets(s, 0.7, 4.3, 5.8, 2.4,
        [(f"{i+1}. {n}", f"{d} · {p:,} pts") for i, (n, d, p) in enumerate(TOP_IND)], size=13, gap=4)
txt(s, 6.9, 3.9, 5.7, 0.4, "BRANCH GRAND PRIX (% delivered)", 13, EMER, bold=True)
bullets(s, 6.9, 4.3, 5.7, 2.4,
        [(f"{i+1}. {b}", f"{pr}% delivered") for i, (b, pr) in enumerate(BR)], size=13, gap=4)
footer(s, 10)

# ============================ SLIDE 11 — SUSTAINABLE ============================
s = slide()
header(s, "BUILT TO LAST", "Sustainable for months of live play")
bullets(s, 0.7, 2.2, 11.9, 4.0, [
    ("Seasons reset the race", "Weekly & monthly leagues mean no single star runs away with it — everyone starts fresh."),
    ("Rewards clean behaviour, not burnout", "Speed is bonus, not the base; quality and collaboration are what compound."),
    ("Self-balancing against gaming", "Escrow + net-progress + once-per-booking mean exploits simply don't pay."),
    ("Plugs into existing data", "Runs directly on the event log the dealership already produces — zero new tracking."),
    ("Fully tunable", "All 20 actions, weights, and multipliers live in one config — easy to rebalance."),
], size=17)
footer(s, 11)

# ============================ SLIDE 12 — DEMO ============================
s = slide(INK)
box(s, 0, 0, W, 0.28, fill=RED)
txt(s, 0.7, 0.5, 11.9, 0.4, "LIVE DEMO", 13, EMER, bold=True)
txt(s, 0.7, 0.95, 11.9, 1.0, "See it running on the real data", 33, WHITE, bold=True)
bullets(s, 0.7, 2.4, 11.9, 3.2, [
    ("Interactive dashboard", "Branch race track, live leaderboards, quests, boss battle & anti-gaming proof."),
    ("Every number is real", "Computed by our scoring engine from all 170,162 logged events."),
    ("Two artifacts", "A one-file web app (carverse.html) + this deck — both from one Python pipeline."),
], size=17, color=WHITE)
for i, (h, _) in enumerate([("scoring_engine.py", ""), ("build_site.py → carverse.html", ""), ("make_deck.py → this deck", "")]):
    txt(s, 0.9, 5.4 + i*0.45, 11, 0.4, "› " + h, 14, AMBER, bold=True, font="Consolas")

# ============================ SLIDE 13 — CLOSE ============================
s = slide()
box(s, 0, 0, W, 0.28, fill=RED); box(s, 0, 7.22, W, 0.28, fill=EMER)
txt(s, 0.7, 2.6, 11.9, 1.4, "Make the dealership\na game worth winning.", 44, INK, bold=True, spacing=1.0)
txt(s, 0.7, 4.6, 11.9, 0.6, "CARVERSE GRAND PRIX — rewards real progress, forces collaboration, can't be farmed.", 18, GREY)
txt(s, 0.7, 5.4, 11.9, 0.5, "Thank you.", 20, RED, bold=True)

prs.save("CARVERSE_Pitch.pptx")
print("saved CARVERSE_Pitch.pptx —", len(prs.slides._sldIdLst), "slides")
