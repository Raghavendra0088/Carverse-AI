# -*- coding: utf-8 -*-
"""
DealerVerse AI — complete project deck.
Modern 16:9 dark theme matching the app. Every detail, real data.
"""
import os
LIMIT = int(os.environ.get("SLIDE_LIMIT", "0"))  # 0 = all
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ---------- palette ----------
BG      = RGBColor(0x0B, 0x0E, 0x1A)
BG2     = RGBColor(0x10, 0x14, 0x24)
CARD    = RGBColor(0x17, 0x1C, 0x30)
CARD2   = RGBColor(0x1E, 0x24, 0x3D)
LINE    = RGBColor(0x2A, 0x31, 0x4E)
VIOLET  = RGBColor(0x7C, 0x5C, 0xFF)
VIOLET2 = RGBColor(0xA7, 0x8B, 0xFA)
TXT     = RGBColor(0xE8, 0xEA, 0xF2)
MUTED   = RGBColor(0x96, 0x9C, 0xB6)
WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
GREEN   = RGBColor(0x22, 0xC5, 0x5E)
AMBER   = RGBColor(0xF5, 0x9E, 0x0B)
CYAN    = RGBColor(0x38, 0xBD, 0xF8)
TEAL    = RGBColor(0x2E, 0xE6, 0xA6)
PINK    = RGBColor(0xFB, 0x71, 0x85)
GOLD    = RGBColor(0xFF, 0xC5, 0x3D)
RED     = RGBColor(0xF4, 0x3F, 0x5E)

DEPT_COLOR = {
    "SALES": VIOLET, "FINANCE": GREEN, "ACCOUNTS": AMBER, "EDP": CYAN,
    "CUSTOMER CARE": TEAL, "PDI": PINK, "RTO / REGN TEAM": VIOLET2,
}

FONT = "Segoe UI"
FONT_L = "Segoe UI Light"
FONT_SB = "Segoe UI Semibold"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]

# ---------- helpers ----------
import atexit, sys
def _save_now():
    out = r"C:\Users\reddy\Desktop\trash\CARVERSE\DealerVerse_AI_Complete.pptx"
    prs.save(out)
    print("Saved:", out, "| Slides:", len(prs.slides._sldIdLst))
atexit.register(_save_now)
def slide():
    if LIMIT and len(prs.slides._sldIdLst) >= LIMIT:
        sys.exit(0)
    s = prs.slides.add_slide(BLANK)
    r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, SH)
    r.fill.solid(); r.fill.fore_color.rgb = BG
    r.line.fill.background()
    r.shadow.inherit = False
    return s

def _noline(sp): sp.line.fill.background()
def _noshadow(sp): sp.shadow.inherit = False

def rect(s, x, y, w, h, fill=CARD, line=None, line_w=1.0, radius=True):
    shp_type = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    sp = s.shapes.add_shape(shp_type, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid(); sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line; sp.line.width = Pt(line_w)
    _noshadow(sp)
    if radius:
        try:
            sp.adjustments[0] = 0.06
        except Exception:
            pass
    return sp

def txt(s, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
        space_after=4, line_spacing=1.0):
    """runs: list of paragraphs; each paragraph is list of (text,size,color,bold,font) tuples."""
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
    for i, para in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(space_after); p.space_before = Pt(0)
        p.line_spacing = line_spacing
        for (t, sz, col, bold, fnt) in para:
            r = p.add_run(); r.text = t
            r.font.size = Pt(sz); r.font.color.rgb = col; r.font.bold = bold
            r.font.name = fnt
    return tb

def R(t, sz, col=TXT, bold=False, fnt=FONT):
    return (t, sz, col, bold, fnt)

def bullets(s, x, y, w, h, items, size=15, gap=8, color=TXT, marker="—",
            mcolor=VIOLET, lh=1.05):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(gap); p.space_before = Pt(0); p.line_spacing = lh
        if isinstance(it, tuple):
            head, body = it
            rm = p.add_run(); rm.text = marker + "  "
            rm.font.size = Pt(size); rm.font.color.rgb = mcolor; rm.font.bold = True; rm.font.name = FONT
            rh = p.add_run(); rh.text = head + "  "
            rh.font.size = Pt(size); rh.font.color.rgb = WHITE; rh.font.bold = True; rh.font.name = FONT
            rb = p.add_run(); rb.text = body
            rb.font.size = Pt(size); rb.font.color.rgb = color; rb.font.name = FONT
        else:
            rm = p.add_run(); rm.text = marker + "  "
            rm.font.size = Pt(size); rm.font.color.rgb = mcolor; rm.font.bold = True; rm.font.name = FONT
            rb = p.add_run(); rb.text = it
            rb.font.size = Pt(size); rb.font.color.rgb = color; rb.font.name = FONT
    return tb

def kicker(s, text, x=0.7, y=0.55, color=VIOLET2):
    txt(s, x, y, 9, 0.4, [[R(text.upper(), 12.5, color, True)]])

def heading(s, text, x=0.7, y=0.9, size=30, w=11.9, color=WHITE):
    txt(s, x, y, w, 0.9, [[R(text, size, color, True, FONT_SB)]])

def accent_bar(s, x=0.7, y=1.62, w=0.9):
    b = rect(s, x, y, w, 0.06, fill=VIOLET, radius=False)
    return b

def pagenum(s, n):
    txt(s, 12.2, 7.02, 0.9, 0.35, [[R(f"{n:02d}", 11, MUTED, True)]], align=PP_ALIGN.RIGHT)
    txt(s, 0.7, 7.02, 6, 0.35, [[R("DealerVerse AI", 10.5, MUTED, False)]])

_pageno = [0]
def footer(s):
    _pageno[0] += 1
    pagenum(s, _pageno[0])

def stat_tile(s, x, y, w, h, value, label, vcolor=VIOLET, sub=None):
    rect(s, x, y, w, h, fill=CARD, line=LINE, line_w=1.0)
    rect(s, x, y, 0.09, h, fill=vcolor, radius=False)
    paras = [[R(value, 27, vcolor, True, FONT_SB)],
             [R(label, 11.5, TXT, True)]]
    if sub:
        paras.append([R(sub, 9.5, MUTED)])
    txt(s, x+0.28, y+0.16, w-0.4, h-0.3, paras, anchor=MSO_ANCHOR.TOP, space_after=2)

def chip(s, x, y, w, text, color):
    h = 0.34
    c = rect(s, x, y, w, h, fill=None, line=color, line_w=1.25)
    txt(s, x, y+0.02, w, h, [[R(text, 10.5, color, True)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# ---------- table builder ----------
def table(s, x, y, w, col_w, headers, rows, header_fill=VIOLET, row_h=0.34,
          header_h=0.42, fs=11.5, hfs=11.5, zebra=(CARD, BG2), col_colors=None,
          align=None, header_color=WHITE):
    ncol = len(headers)
    total = sum(col_w)
    col_w = [cw * w / total for cw in col_w]
    align = align or [PP_ALIGN.LEFT]*ncol
    # header
    cx = x
    for j, hd in enumerate(headers):
        rect(s, cx, y, col_w[j], header_h, fill=header_fill, radius=False)
        txt(s, cx+0.1, y, col_w[j]-0.18, header_h, [[R(hd, hfs, header_color, True)]],
            align=align[j], anchor=MSO_ANCHOR.MIDDLE)
        cx += col_w[j]
    # rows
    ry = y + header_h
    for i, row in enumerate(rows):
        cx = x
        fill = zebra[i % 2]
        for j, cell in enumerate(row):
            rect(s, cx, ry, col_w[j], row_h, fill=fill, radius=False)
            col = TXT
            bold = False
            if col_colors and col_colors[j] is not None:
                col = col_colors[j]; bold = (j == 0)
            if j == 0:
                bold = True
            txt(s, cx+0.1, ry, col_w[j]-0.18, row_h, [[R(str(cell), fs, col, bold)]],
                align=align[j], anchor=MSO_ANCHOR.MIDDLE)
            cx += col_w[j]
        ry += row_h
    return ry

SHOTS = r"C:\Users\reddy\AppData\Local\Temp\claude\C--Users-reddy-Desktop-trash-CARVERSE-dealerverse-ai\e163ebe1-a966-4b1b-9534-08d92cfd15f0\scratchpad\shots"
def pic(s, name, x, y, w, border=VIOLET, bw=1.25):
    p = s.shapes.add_picture(os.path.join(SHOTS, name + ".png"), Inches(x), Inches(y), width=Inches(w))
    p.line.color.rgb = border; p.line.width = Pt(bw)
    p.shadow.inherit = False
    return p

def shot_caption(s, x, y, w, title, sub, color=VIOLET2):
    txt(s, x, y, w, 0.7, [[R(title, 13.5, color, True)], [R(sub, 11.5, MUTED)]], space_after=2, line_spacing=1.05)

# ==================================================================
# SLIDE 1 — TITLE
# ==================================================================
s = slide()
# big gradient-ish panels
rect(s, -1, -1, 8, 9.5, fill=BG2, radius=False)
# accent glow blocks
rect(s, 8.4, 0.9, 4.3, 1.5, fill=CARD, line=VIOLET, line_w=1.25)
txt(s, 8.7, 1.0, 3.9, 1.3,
    [[R("CarVerse Mobility Technologies", 11, VIOLET2, True)],
     [R("Dealership Gamification Hackathon", 12.5, TXT, True)],
     [R("Build a Game for a Real Dealership", 10.5, MUTED)]], space_after=3)

txt(s, 0.9, 2.35, 8.0, 0.5, [[R("PRESENTING", 14, VIOLET2, True)]])
txt(s, 0.85, 2.75, 11.6, 1.9,
    [[R("DealerVerse ", 62, WHITE, True, FONT_SB), R("AI", 62, VIOLET, True, FONT_SB)]])
txt(s, 0.9, 4.05, 11.6, 0.8, [[R("Level Up Your Dealership", 26, VIOLET2, True, FONT_L)]])
txt(s, 0.92, 4.95, 11.4, 1.2,
    [[R("A gamification platform that turns real dealership work into a game — employees earn XP, "
        "ranks, missions and team bonuses just by doing their jobs. Built end-to-end on the real "
        "170,162-event dataset. No fake numbers.", 15, MUTED)]], line_spacing=1.15)

for i,(t,c) in enumerate([("7 Departments gamified",VIOLET),("≤20 scoring actions",GREEN),
                          ("Anti-gaming by design",AMBER),("Full C2B2C buy/sell loop",CYAN)]):
    chip(s, 0.9+i*3.05, 6.3, 2.9, t, c)
txt(s, 0.9, 6.95, 11, 0.4, [[R("React · Vite · Tailwind · Framer Motion · Recharts · Zustand   ·   16 live screens", 11.5, MUTED)]])

# ==================================================================
# SLIDE 2 — THE CHALLENGE
# ==================================================================
s = slide(); kicker(s, "The Brief"); heading(s, "The Challenge"); accent_bar(s); footer(s)
txt(s, 0.7, 1.85, 12, 0.9,
    [[R("“Design a game that motivates people at a car dealership to perform better at their jobs, "
        "using real activity data — a game people will ", 15, TXT),
      R("actually want to play", 15, VIOLET2, True),
      R(" while doing their job.”", 15, TXT)]], line_spacing=1.15)

rect(s, 0.7, 2.9, 5.85, 3.9, fill=CARD, line=LINE)
txt(s, 1.0, 3.1, 5.3, 0.4, [[R("WHAT WE MUST DO", 12, VIOLET2, True)]])
bullets(s, 1.0, 3.6, 5.3, 3.0, [
    ("Motivate individuals & teams", "to move enquiries → bookings → delivery faster."),
    ("Separate signal from noise", "reward real progress, not data edits."),
    ("Encourage collaboration", "across departments through the lifecycle."),
    ("Be un-gameable", "repeating an action 50× must not pay 50×."),
], size=13.5, gap=9)

rect(s, 6.75, 2.9, 5.85, 3.9, fill=CARD, line=LINE)
txt(s, 7.05, 3.1, 5.3, 0.4, [[R("HARD CONSTRAINTS", 12, AMBER, True)]])
bullets(s, 7.05, 3.6, 5.3, 3.0, [
    ("Max 20 actions", "may be used for scoring — force real prioritisation."),
    ("Map to real workflow", "the actual booking → delivery pipeline."),
    ("Sustainable for months", "must survive a live dealership, not just a demo."),
    ("Key message", "“if your game rewards the wrong behaviour, it will fail.”"),
], size=13.5, gap=9, mcolor=AMBER)

# ==================================================================
# SLIDE 3 — JUDGING CRITERIA -> HOW WE WIN
# ==================================================================
s = slide(); kicker(s, "Scored Against"); heading(s, "Judging Criteria → How DealerVerse Answers"); accent_bar(s); footer(s)
rows = [
    ["Game Design Quality", "30%", "Full loop: XP → levels → ranks → missions → badges → leaderboards, per department."],
    ["Real-world relevance", "25%", "Every point is computed from the real 170k-event log; 7 real departments."],
    ["Anti-gaming robustness", "20%", "Once-per-booking scoring; 77.6% of events earn 0; outcomes >> updates."],
    ["Team collaboration", "15%", "Team-bonus mechanic pays every dept that moved a delivered car (+25 XP each)."],
    ["Creativity", "10%", "A playable app + a full C2B2C buy/sell marketplace layered on the workflow."],
]
table(s, 0.7, 2.0, 11.95, [3.0, 1.1, 7.4],
      ["Criterion", "Weight", "Our Answer"], rows,
      align=[PP_ALIGN.LEFT, PP_ALIGN.CENTER, PP_ALIGN.LEFT], row_h=0.72, fs=12.5, hfs=12.5,
      col_colors=[WHITE, GOLD, TXT])
txt(s, 0.7, 6.6, 12, 0.5,
    [[R("Every weighted criterion maps to a concrete, shipped mechanic — nothing is aspirational.", 13, VIOLET2, True)]])

# ==================================================================
# SLIDE 4 — THE SOLUTION (overview)
# ==================================================================
s = slide(); kicker(s, "Our Solution"); heading(s, "One Platform, Real Work Becomes a Game"); accent_bar(s); footer(s)
txt(s, 0.7, 1.8, 12, 0.8,
    [[R("Employees log in and do their normal jobs. DealerVerse reads the dealership event stream, "
        "scores only the milestones that move a car forward, and instantly turns that into XP, ranks, "
        "missions and team bonuses.", 15, MUTED)]], line_spacing=1.15)

cards = [
    ("Do the work", "A booking is created, a loan cleared, a car inspected, a gatepass issued.", VIOLET, "1"),
    ("Earn instantly", "Meaningful milestones award XP to the real person who did them.", GREEN, "2"),
    ("Level & rank up", "XP fills your level; per-department ladders promote you Rookie → Legend.", AMBER, "3"),
    ("Compete & collaborate", "Leaderboards, daily missions, streaks, and cross-dept team bonuses.", CYAN, "4"),
]
cw = 2.92; gap = 0.13; x0 = 0.7
for i,(t,d,c,n) in enumerate(cards):
    x = x0 + i*(cw+gap)
    rect(s, x, 3.0, cw, 3.4, fill=CARD, line=LINE)
    circ = rect(s, x+0.28, 3.28, 0.62, 0.62, fill=c, radius=True)
    try: circ.adjustments[0] = 0.5
    except: pass
    txt(s, x+0.28, 3.30, 0.62, 0.6, [[R(n, 22, BG, True, FONT_SB)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, x+0.28, 4.1, cw-0.5, 0.6, [[R(t, 15.5, WHITE, True)]])
    txt(s, x+0.28, 4.7, cw-0.5, 1.6, [[R(d, 12.5, MUTED)]], line_spacing=1.12)
txt(s, 0.7, 6.65, 12, 0.5,
    [[R("The gameplay loop is short, honest, and repeats every working day — that is what keeps it fun.", 13, VIOLET2, True)]])

# ==================================================================
# SLIDE 5 — THE DATA
# ==================================================================
s = slide(); kicker(s, "Grounded in Reality"); heading(s, "Built on the Real Dataset"); accent_bar(s); footer(s)
tiles = [
    ("170,162", "Real events", VIOLET, "May–Jun 2026"),
    ("3,800", "Bookings tracked", GREEN, "location + enquiry"),
    ("305", "Active players", CYAN, "with live XP"),
    ("41", "Locations", AMBER, "branch network"),
]
tw = 2.92
for i,(v,l,c,sub) in enumerate(tiles):
    stat_tile(s, 0.7+i*(tw+0.13), 2.0, tw, 1.5, v, l, vcolor=c, sub=sub)

rect(s, 0.7, 3.85, 5.85, 2.9, fill=CARD, line=LINE)
txt(s, 1.0, 4.05, 5.3, 0.4, [[R("SIX CORE FIELDS PER EVENT", 12, VIOLET2, True)]])
bullets(s, 1.0, 4.55, 5.3, 2.1, [
    "USER  ·  who performed the action",
    "DEPARTMENT  ·  which of the real teams",
    "ACTION  ·  3,347 distinct action types",
    "BOOKING  ·  location_code + enquiry_no",
    "TIME  ·  timestamp for streaks & weekly series",
    "LOCATION  ·  branch, for league standings",
], size=12.5, gap=6.5)

rect(s, 6.75, 3.85, 5.85, 2.9, fill=CARD, line=LINE)
txt(s, 7.05, 4.05, 5.3, 0.4, [[R("THREE REAL SOURCE FILES", 12, GREEN, True)]])
bullets(s, 7.05, 4.55, 5.3, 2.1, [
    ("z_event_log_may_june_2026.csv", "— 170,162 rows of raw activity."),
    ("z_employees.csv", "— names, departments, designations, managers, branch."),
    ("z_locations.csv", "— location codes → branch names."),
], size=12.5, gap=9, mcolor=GREEN)
txt(s, 7.05, 6.15, 5.3, 0.55,
    [[R("Processed once by ", 12, MUTED), R("scripts/build-data.mjs", 12, TXT, True),
      R(" into a compact JSON the app loads instantly.", 12, MUTED)]], line_spacing=1.1)

# ==================================================================
# SLIDE 6 — SIGNAL VS NOISE
# ==================================================================
s = slide(); kicker(s, "Anti-Gaming · Principle 1"); heading(s, "Signal vs Noise: 77.6% of Events Earn Zero"); accent_bar(s); footer(s)
txt(s, 0.7, 1.8, 12, 0.7,
    [[R("The dataset is dominated by trivial, repeatable events. Rewarding them would break the game. "
        "We reward only what moves a booking forward.", 14.5, MUTED)]], line_spacing=1.15)

rect(s, 0.7, 2.85, 6.0, 3.9, fill=CARD, line=RED, line_w=1.25)
txt(s, 1.0, 3.05, 5.4, 0.4, [[R("NOISE — SCORED AT 0 XP", 12.5, RED, True)]])
noise = [
    ["USER_LOGGED_IN", "27,764"],
    ["DISCOUNT_UPDATED", "9,188"],
    ["PDI_INFO_UPDATED", "8,428"],
    ["PAYMENTS_UPDATED", "7,816"],
    ["FINANCE_INFO_UPDATED", "5,588"],
]
table(s, 1.0, 3.55, 5.4, [3.4, 1.4], ["High-frequency action", "Count"], noise,
      header_fill=RED, align=[PP_ALIGN.LEFT, PP_ALIGN.RIGHT], row_h=0.42, fs=12.5, hfs=11.5,
      col_colors=[TXT, MUTED])
txt(s, 1.0, 6.05, 5.4, 0.6, [[R("Logins, edits, notes and repeat uploads → ", 11.5, MUTED), R("0 XP.", 11.5, RED, True)]])

rect(s, 6.9, 2.85, 5.7, 3.9, fill=CARD, line=GREEN, line_w=1.25)
txt(s, 7.2, 3.05, 5.1, 0.4, [[R("SIGNAL — REWARDED MILESTONES", 12.5, GREEN, True)]])
bullets(s, 7.2, 3.6, 5.1, 3.0, [
    ("Update vs Progress", "FINANCE_INFO_UPDATED = 0; FINANCE_CLEARED = 60."),
    ("Repeated vs One-time", "notes = 0; INVOICE_APPROVED scores once."),
    ("Granular vs Meaningful", "single doc uploads = 0; full DOCUMENT_SET = 30."),
    ("Effort vs Outcome", "commitments = 0; delivery-driving steps score high."),
], size=12.5, gap=10, mcolor=GREEN)

# ==================================================================
# SLIDE 7 — THE XP ENGINE (scoring table)
# ==================================================================
s = slide(); kicker(s, "The XP Engine"); heading(s, "17 Milestones + Team Bonus — Under the 20 Cap"); accent_bar(s); footer(s)
score_rows = [
    ["Booking Created", "SALES", "20"],
    ["Discount Approved", "SALES", "30"],
    ["Invoice Signed (Sales)", "SALES", "25"],
    ["Registration Approved", "SALES", "25"],
    ["Loan File Opened *", "FINANCE", "10"],
    ["Finance Cleared *", "FINANCE", "60"],
    ["Payment Credited", "ACCOUNTS", "45"],
    ["Gatepass Issued", "ACCOUNTS", "100"],
    ["Invoice Finalized", "EDP", "45"],
]
score_rows2 = [
    ["Registration Processed", "EDP", "25"],
    ["Registration Filed", "RTO / Regn", "40"],
    ["Insurance Approved", "CUST. CARE", "40"],
    ["Delivery Coordinated", "CUST. CARE", "20"],
    ["Gatepass Cleared (CCM)", "CUST. CARE", "20"],
    ["Inspected & Dispatched", "PDI", "20"],
    ["RTO Registration", "PDI", "20"],
    ["Document Set Completed *", "PDI", "30"],
    ["Team Bonus (co-op)", "ANY DEPT", "+25"],
]
table(s, 0.7, 1.95, 6.0, [3.5, 1.7, 0.9], ["Milestone", "Dept", "XP"], score_rows,
      align=[PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.RIGHT], row_h=0.40, fs=11.5, hfs=11,
      col_colors=[TXT, MUTED, GOLD])
table(s, 6.95, 1.95, 6.0, [3.5, 1.7, 0.9], ["Milestone", "Dept", "XP"], score_rows2,
      align=[PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.RIGHT], row_h=0.40, fs=11.5, hfs=11,
      col_colors=[TXT, MUTED, GOLD])
txt(s, 0.7, 6.5, 12, 0.7,
    [[R("*  ", 12, VIOLET2, True), R("Derived milestones", 12, WHITE, True),
      R(" — not a single log row but computed from the booking's history. ", 12, MUTED),
      R("XP weight = business impact: a Gatepass (100) and Finance Cleared (60) dwarf a status update (0).", 12, MUTED)]],
    line_spacing=1.15)

# ==================================================================
# SLIDE 8 — DERIVED MILESTONES
# ==================================================================
s = slide(); kicker(s, "Beyond Raw Logs"); heading(s, "Derived Milestones: Scoring What Isn't Logged"); accent_bar(s); footer(s)
txt(s, 0.7, 1.8, 12, 0.6,
    [[R("The real signal often isn't a single event. We reconstruct true milestones from each "
        "booking's event history, then credit the person who completed them.", 14, MUTED)]], line_spacing=1.15)
der = [
    ("Loan File Opened  ·  10 XP", "First finance touch on a booking — the loan work begins.", GREEN),
    ("Finance Cleared  ·  60 XP", "Booking had finance work AND advanced past the finance stage (ABNR+).", GREEN),
    ("Document Set Completed  ·  30 XP", "All 7 required docs present (PAN, Aadhaar, VIN, 4 photos) → credit the finisher.", PINK),
    ("Team Bonus  ·  +25 XP each", "A delivered car that ≥2 departments moved pays every contributor.", GOLD),
]
for i,(t,d,c) in enumerate(der):
    y = 2.75 + i*1.02
    rect(s, 0.7, y, 11.9, 0.9, fill=CARD, line=LINE)
    rect(s, 0.7, y, 0.09, 0.9, fill=c, radius=False)
    txt(s, 1.0, y+0.13, 4.4, 0.7, [[R(t, 14.5, c, True)]], anchor=MSO_ANCHOR.MIDDLE)
    txt(s, 5.4, y+0.13, 7.0, 0.7, [[R(d, 12.5, TXT)]], anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.1)
txt(s, 0.7, 6.95, 12, 0.4,
    [[R("This is how we “gamify the completion, not the pieces” — exactly as the brief asks.", 12.5, VIOLET2, True)]])

# ==================================================================
# SLIDE 9 — ANTI-GAMING STRATEGY
# ==================================================================
s = slide(); kicker(s, "Anti-Gaming · The Full Strategy"); heading(s, "Six Rules That Make It Un-Gameable"); accent_bar(s); footer(s)
rules = [
    ("Once per booking", "Every milestone scores exactly once per car. Repeating an action does nothing.", VIOLET),
    ("Cancelled = zero", "Cancelled bookings never score, so churning fake deals is worthless.", RED),
    ("Outcomes over effort", "Delivery-driving steps carry 3–10× the weight of any update.", GREEN),
    ("Noise floored to 0", "The five most frequent actions — 58k+ events — all pay nothing.", AMBER),
    ("Credit the real actor", "XP lands on whoever actually performed the step, in their real dept.", CYAN),
    ("≤20 scoring actions", "A hard whitelist; everything outside it is invisible to the score.", VIOLET2),
]
cw = 3.83; ch = 2.0; gx = 0.7; gy = 2.05
for i,(t,d,c) in enumerate(rules):
    col = i % 3; row = i // 3
    x = gx + col*(cw+0.18); y = gy + row*(ch+0.2)
    rect(s, x, y, cw, ch, fill=CARD, line=LINE)
    rect(s, x, y, cw, 0.09, fill=c, radius=False)
    txt(s, x+0.25, y+0.24, cw-0.5, 0.5, [[R(t, 15, c, True)]])
    txt(s, x+0.25, y+0.82, cw-0.5, 1.1, [[R(d, 12.5, TXT)]], line_spacing=1.12)
txt(s, 0.7, 6.75, 12, 0.4,
    [[R("“If users repeat an action 50 times, they get 50× points?”  →  No. In DealerVerse, they get 1×, once.", 13, VIOLET2, True)]])

# ==================================================================
# SLIDE 10 — COLLABORATION
# ==================================================================
s = slide(); kicker(s, "Team Collaboration Design"); heading(s, "Cars Get Delivered by Teams — So Teams Score"); accent_bar(s); footer(s)
txt(s, 0.7, 1.8, 12, 0.7,
    [[R("A delivered car is a relay. When two or more departments moved the same booking to delivery, "
        "every contributing department earns a ", 14.5, MUTED), R("+25 XP Team Bonus", 14.5, GOLD, True),
      R(" — collaboration literally pays.", 14.5, MUTED)]], line_spacing=1.15)
stat_tile(s, 0.7, 2.9, 3.75, 1.5, "1,508", "Team bonuses awarded", vcolor=GOLD, sub="on delivered cars")
stat_tile(s, 4.6, 2.9, 3.75, 1.5, "+25", "XP per dept, per car", vcolor=VIOLET, sub="every contributor")
stat_tile(s, 8.5, 2.9, 4.1, 1.5, "7 / 7", "Departments collaborating", vcolor=GREEN, sub="Sales → Customer Care")
rect(s, 0.7, 4.7, 11.9, 2.1, fill=CARD, line=LINE)
txt(s, 1.0, 4.9, 11.4, 0.4, [[R("THE RELAY — ONE CAR, MANY HANDS", 12, VIOLET2, True)]])
relay = ["Sales", "Finance", "Accounts", "EDP", "PDI", "Cust. Care", "RTO / Regn"]
rc = [VIOLET, GREEN, AMBER, CYAN, PINK, TEAL, VIOLET2]
xx = 1.0; bw = 1.44
for i,(name,c) in enumerate(zip(relay, rc)):
    rect(s, xx, 5.55, bw, 0.8, fill=BG2, line=c, line_w=1.5)
    txt(s, xx, 5.55, bw, 0.8, [[R(name, 11.5, c, True)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    if i < 6:
        txt(s, xx+bw-0.06, 5.55, 0.30, 0.8, [[R("›", 20, MUTED, True)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    xx += bw + 0.24
txt(s, 1.0, 6.45, 11.4, 0.3, [[R("Each hand-off that helps deliver the car is rewarded — no one can hoard the win.", 12, MUTED)]])

# ==================================================================
# SLIDE 11 — 7 DEPARTMENTS (live results)
# ==================================================================
s = slide(); kicker(s, "Every Team, Not Just Sales"); heading(s, "All 7 Departments Gamified — Live Standings"); accent_bar(s); footer(s)
dept_rows = [
    ["ACCOUNTS", "13", "256,675", "2,979"],
    ["SALES", "241", "199,015", "6,747"],
    ["EDP", "16", "178,565", "3,999"],
    ["CUSTOMER CARE", "17", "160,960", "4,608"],
    ["PDI", "3", "130,445", "4,143"],
    ["FINANCE", "14", "129,220", "3,647"],
    ["RTO / REGN TEAM", "1", "51,965", "841"],
]
dcol = [DEPT_COLOR[r[0]] for r in dept_rows]
# custom table with dept colour on name
ncol=4; col_w=[3.4,1.6,2.2,2.0]; x=0.7; w=11.95; y=2.0
tot=sum(col_w); col_w=[c*w/tot for c in col_w]
heads=["Department","Players","Total XP","Milestones"]
al=[PP_ALIGN.LEFT,PP_ALIGN.CENTER,PP_ALIGN.RIGHT,PP_ALIGN.RIGHT]
cx=x
for j,hd in enumerate(heads):
    rect(s,cx,y,col_w[j],0.44,fill=VIOLET,radius=False)
    txt(s,cx+0.1,y,col_w[j]-0.18,0.44,[[R(hd,12,WHITE,True)]],align=al[j],anchor=MSO_ANCHOR.MIDDLE)
    cx+=col_w[j]
ry=y+0.44
for i,row in enumerate(dept_rows):
    cx=x; fill=CARD if i%2==0 else BG2
    for j,cell in enumerate(row):
        rect(s,cx,ry,col_w[j],0.47,fill=fill,radius=False)
        col = dcol[i] if j==0 else (GOLD if j==2 else TXT)
        txt(s,cx+0.1,ry,col_w[j]-0.18,0.47,[[R(str(cell),12.5,col,j==0)]],align=al[j],anchor=MSO_ANCHOR.MIDDLE)
        cx+=col_w[j]
    ry+=0.47
txt(s, 0.7, ry+0.12, 12, 0.5,
    [[R("Note the fairness win: ", 12.5, VIOLET2, True),
      R("PDI has 3 people yet out-scores 241 in Sales — per-department ranks let small teams compete on their own scale.", 12.5, MUTED)]],
    line_spacing=1.12)

# ==================================================================
# SLIDE 12 — RANK LADDERS
# ==================================================================
s = slide(); kicker(s, "Progression"); heading(s, "Per-Department Rank Ladders — Rookie → Legend"); accent_bar(s); footer(s)
ladders = [
    ("SALES", ["Rookie","Lead Hunter","Negotiator","Deal Closer","Sales Champion","Dealership Legend"]),
    ("FINANCE", ["Finance Trainee","Loan Expert","Approval Master","Finance Hero","Finance Ace","Finance Legend"]),
    ("ACCOUNTS", ["Ledger Rookie","Payment Verifier","Credit Controller","Accounts Ace","Revenue Guardian","Accounts Legend"]),
    ("EDP", ["Data Rookie","System Operator","Invoice Master","Backend Ace","Ops Architect","EDP Legend"]),
    ("CUSTOMER CARE", ["Care Trainee","Support Specialist","Insurance Pro","Care Champion","Customer Hero","Care Legend"]),
    ("PDI", ["Inspection Rookie","Quality Checker","Dispatch Pro","PDI Ace","Delivery Master","PDI Legend"]),
    ("RTO / REGN TEAM", ["Filing Rookie","Registration Clerk","RTO Specialist","Compliance Pro","Registration Master","RTO Legend"]),
]
y=1.95; rh=0.66
for i,(d,ranks) in enumerate(ladders):
    c = DEPT_COLOR[d]
    yy = y + i*rh
    rect(s, 0.7, yy, 2.35, rh-0.1, fill=CARD, line=c, line_w=1.25)
    txt(s, 0.8, yy, 2.2, rh-0.1, [[R(d, 10.5, c, True)]], anchor=MSO_ANCHOR.MIDDLE)
    xx = 3.2; bw = 1.55
    for k,rk in enumerate(ranks):
        alpha = c
        rect(s, xx, yy, bw, rh-0.1, fill=BG2, line=LINE, line_w=0.75)
        rc_col = c if k>=4 else (TXT if k>=1 else MUTED)
        txt(s, xx+0.05, yy, bw-0.1, rh-0.1, [[R(rk, 8.8, rc_col, k>=4)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, line_spacing=0.95)
        xx += bw + 0.05
txt(s, 3.2, y-0.02+7*rh-0.02, 9, 0.3, [])
txt(s, 0.7, 6.62, 12, 0.5,
    [[R("6 tiers × 7 departments = 42 named ranks. ", 12.5, VIOLET2, True),
      R("Thresholds are computed from each department's real XP distribution, so every ladder is earned.", 12.5, MUTED)]], line_spacing=1.1)

# ==================================================================
# SLIDE 13 — PROGRESSION SYSTEMS
# ==================================================================
s = slide(); kicker(s, "The Full Game Loop"); heading(s, "Levels · Missions · Badges · Streaks"); accent_bar(s); footer(s)
blocks = [
    ("Ranks & Levels", VIOLET, [
        "6-tier ladder per department",
        "Rank ring shows % to next rank",
        "Top earner always reaches Legend",
    ]),
    ("Daily Missions", GREEN, [
        "2 department-specific quests/day",
        "Progress from real 'today' counts",
        "+ Team-bonus & streak missions",
    ]),
    ("Achievements", AMBER, [
        "First Booking, Centurion, Team Player",
        "Sharpshooter, Fast Approver, Gatekeeper",
        "Veteran, Legend, 7-Day Streak 🔥",
    ]),
    ("Streaks & Growth", CYAN, [
        "Consecutive active-day streaks",
        "Weekly XP delta (up/down %)",
        "Most Improved & Fastest Growing",
    ]),
]
cw=2.92
for i,(t,c,items) in enumerate(blocks):
    x=0.7+i*(cw+0.13)
    rect(s,x,2.0,cw,4.4,fill=CARD,line=LINE)
    rect(s,x,2.0,cw,0.09,fill=c,radius=False)
    txt(s,x+0.25,2.28,cw-0.5,0.5,[[R(t,15,c,True)]])
    bullets(s,x+0.25,2.95,cw-0.45,3.2,items,size=11.8,gap=9,mcolor=c)
txt(s, 0.7, 6.65, 12, 0.4,
    [[R("Short-term (missions) + mid-term (ranks) + long-term (season leaderboard) goals keep it fun for months.", 12.5, VIOLET2, True)]])

# ==================================================================
# SLIDE 14 — MISSIONS & AI COACH
# ==================================================================
s = slide(); kicker(s, "Engagement Engine"); heading(s, "Daily Missions & the AI Coach"); accent_bar(s); footer(s)
rect(s, 0.7, 1.95, 5.85, 4.6, fill=CARD, line=LINE)
txt(s, 1.0, 2.15, 5.3, 0.4, [[R("DAILY MISSIONS  (examples)", 12, GREEN, True)]])
mis = [
    ("🚗 Create 5 bookings", "Sales · 100 XP"),
    ("✅ Clear 3 loans", "Finance · Loan Hero badge"),
    ("🛡️ Issue 4 gatepasses", "Accounts · 150 XP"),
    ("☂️ Approve 4 insurances", "Cust. Care · 120 XP"),
    ("📝 File 5 registrations", "RTO · 150 XP"),
    ("🤝 Earn a Team Bonus", "Any · 25 XP + badge"),
    ("🔥 Keep your streak alive", "Any · streak multiplier"),
]
yy=2.65
for t,r in mis:
    rect(s, 1.0, yy, 5.25, 0.48, fill=BG2, line=LINE, line_w=0.75)
    txt(s, 1.15, yy, 3.6, 0.48, [[R(t, 12, TXT, True)]], anchor=MSO_ANCHOR.MIDDLE)
    txt(s, 4.7, yy, 1.5, 0.48, [[R(r, 10, GREEN, True)]], align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    yy+=0.53

rect(s, 6.75, 1.95, 5.85, 4.6, fill=CARD, line=VIOLET, line_w=1.25)
txt(s, 7.05, 2.15, 5.3, 0.4, [[R("🧠  AI COACH  (deterministic insights)", 12, VIOLET2, True)]])
coach = [
    "“You're only 340 XP from Deal Closer — keep clearing key milestones.”",
    "“Your XP is up 24% vs last week (1,120 XP in your last 7 active days).”",
    "“2 to go on 'Clear 3 loans' — finish it to grab the Loan Hero badge.”",
    "“Strong teamwork — 18 collaboration bonuses. You're carrying your squad.”",
    "“5-day streak going. One more active day keeps your multiplier climbing.”",
]
yy=2.7
for c in coach:
    txt(s, 7.05, yy, 5.3, 0.8, [[R(c, 12.5, TXT)]], line_spacing=1.08)
    yy+=0.78
txt(s, 7.05, 6.35, 5.3, 0.3, [[R("100% computed from the player's real data — no LLM, no guesswork.", 10.5, MUTED)]])

# ==================================================================
# SLIDE 15 — C2B2C LOOP
# ==================================================================
s = slide(); kicker(s, "The Big Differentiator"); heading(s, "A Full C2B2C Buy / Sell Loop, Gamified"); accent_bar(s); footer(s)
txt(s, 0.7, 1.75, 12, 0.6,
    [[R("Beyond scoring history, DealerVerse runs a live end-to-end used-car flow: a customer sells a car "
        "to the dealership, which reconditions, prices and re-sells it to a buyer — every step scored.", 13.5, MUTED)]], line_spacing=1.15)
# acquisition track
txt(s, 0.7, 2.6, 12, 0.35, [[R("① ACQUISITION  —  customer sells to dealership", 12.5, TEAL, True)]])
acq = [("Seller posts car","Customer"),("PDI inspects","+20 XP"),("Sales offers","+40 XP"),
       ("Seller accepts","—"),("Accounts pays out","+50 XP"),("Listed to inventory","auto-priced")]
xx=0.7; bw=1.92
for i,(t,x2) in enumerate(acq):
    rect(s, xx, 2.98, bw, 0.72, fill=CARD, line=TEAL, line_w=1.0)
    txt(s, xx+0.06, 3.02, bw-0.12, 0.44, [[R(t, 10.3, TXT, True)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, line_spacing=0.9)
    txt(s, xx+0.06, 3.42, bw-0.12, 0.24, [[R(x2, 9, TEAL, True)]], align=PP_ALIGN.CENTER)
    if i<5: txt(s, xx+bw-0.05, 2.98, 0.22, 0.72, [[R("›", 16, MUTED, True)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    xx+=bw+0.055
# sale track
txt(s, 0.7, 3.95, 12, 0.35, [[R("② SALE FULFILMENT  —  buyer purchases (each stage owned by a department)", 12.5, VIOLET2, True)]])
sale = [("Reserved","Sales · 20"),("Loan Approved","Finance · 60"),("Payment Cleared","Accounts · 50"),
        ("Invoice Generated","EDP · 25"),("Ownership Transferred","RTO · 30"),("Delivered","Cust.Care · 60")]
scol = [VIOLET, GREEN, AMBER, CYAN, VIOLET2, TEAL]
xx=0.7
for i,((t,x2),c) in enumerate(zip(sale,scol)):
    rect(s, xx, 4.33, bw, 0.72, fill=CARD, line=c, line_w=1.0)
    txt(s, xx+0.06, 4.37, bw-0.12, 0.44, [[R(t, 10.3, TXT, True)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, line_spacing=0.9)
    txt(s, xx+0.06, 4.77, bw-0.12, 0.24, [[R(x2, 9, c, True)]], align=PP_ALIGN.CENTER)
    if i<5: txt(s, xx+bw-0.05, 4.33, 0.22, 0.72, [[R("›", 16, MUTED, True)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    xx+=bw+0.055
# economics
rect(s, 0.7, 5.45, 11.9, 1.35, fill=BG2, line=LINE)
txt(s, 1.0, 5.6, 11.4, 0.35, [[R("BUILT-IN ECONOMICS", 11.5, AMBER, True)]])
eco = [
    ("Reconditioning", "4% of acquisition price", AMBER),
    ("Retail margin", "12% dealership markup", GREEN),
    ("Auto-retail", "(acq + recon) × 1.12, rounded", CYAN),
    ("Override", "Sales/Supervisor can re-price", VIOLET2),
]
xx=1.0; ew=2.85
for t,d,c in eco:
    txt(s, xx, 5.95, ew, 0.7, [[R(t, 11.5, c, True)], [R(d, 10.5, TXT)]], space_after=2, line_spacing=1.0)
    xx+=ew

# ==================================================================
# SLIDE 16 — ROLES & PERMISSIONS
# ==================================================================
s = slide(); kicker(s, "Access Model"); heading(s, "Four Roles, Department-Gated Actions"); accent_bar(s); footer(s)
roles = [
    ("Employee", VIOLET, "Logs in with a 4-digit PIN + department. Sees their dashboard, missions, rank, leaderboard. Can advance sale stages owned by their department."),
    ("Supervisor", AMBER, "An employee flag that unlocks every workflow gate — can inspect, offer, pay out and advance any stage. For managers running the floor."),
    ("Seller (Customer)", TEAL, "Posts a car to sell, tracks the acquisition offer, and accepts or declines the dealership's price."),
    ("Buyer", CYAN, "Browses the marketplace of certified inventory, reserves a car, and watches its sale progress to delivery."),
]
for i,(t,c,d) in enumerate(roles):
    y = 2.0 + i*1.18
    rect(s, 0.7, y, 11.9, 1.05, fill=CARD, line=LINE)
    rect(s, 0.7, y, 0.09, 1.05, fill=c, radius=False)
    txt(s, 1.0, y, 2.7, 1.05, [[R(t, 15, c, True)]], anchor=MSO_ANCHOR.MIDDLE)
    txt(s, 3.7, y+0.13, 8.7, 0.85, [[R(d, 12.5, TXT)]], anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.1)
txt(s, 0.7, 6.85, 12, 0.4,
    [[R("PINs are derived deterministically from the employee ID (client-side, no backend) — every one of 6,037 staff has a unique login.", 12, MUTED)]])

# ==================================================================
# SLIDE 17 — PRODUCT TOUR (16 screens)
# ==================================================================
s = slide(); kicker(s, "The Product"); heading(s, "16 Live Screens — A Complete Playable App"); accent_bar(s); footer(s)
screens = [
    ("Login", "PIN + dept, or Buyer/Seller", VIOLET),
    ("Dashboard", "XP, rank ring, level, missions", VIOLET),
    ("Activity Feed", "Live milestone stream", CYAN),
    ("Pipeline", "Booking-race lifecycle board", VIOLET2),
    ("Evaluate", "PDI inspects trade-ins", PINK),
    ("Customer Portal", "Seller posts a car", TEAL),
    ("Marketplace", "Buyer browses & reserves", CYAN),
    ("My Purchases", "Buyer tracks delivery", CYAN),
    ("Missions", "Daily quests + progress", GREEN),
    ("Career", "Rank ladder progression", VIOLET),
    ("Achievements", "Badge collection", AMBER),
    ("Leaderboard", "Overall & per-dept ranks", GOLD),
    ("Department", "Team standings & stats", AMBER),
    ("Manager", "Weekly charts, most improved", GREEN),
    ("AI Coach", "Personalised insights", VIOLET2),
    ("Profile", "Player card & history", TEAL),
]
cw=2.92; ch=1.12; x0=0.7; y0=2.0
for i,(t,d,c) in enumerate(screens):
    col=i%4; row=i//4
    x=x0+col*(cw+0.13); y=y0+row*(ch+0.14)
    rect(s,x,y,cw,ch,fill=CARD,line=LINE)
    rect(s,x,y,0.08,ch,fill=c,radius=False)
    txt(s,x+0.25,y+0.16,cw-0.4,0.4,[[R(t,13,c,True)]])
    txt(s,x+0.25,y+0.58,cw-0.4,0.45,[[R(d,10.8,MUTED)]],line_spacing=0.95)

# ==================================================================
# SLIDE 17b — SCREENSHOTS: PLAYER EXPERIENCE
# ==================================================================
s = slide(); kicker(s, "Live App · Player Experience"); heading(s, "What a Player Sees Every Day"); accent_bar(s); footer(s)
IW = 5.95  # 1.6:1 screenshots
pic(s, "dashboard", 0.55, 2.05, IW, border=VIOLET)
pic(s, "career", 6.83, 2.05, IW, border=CYAN)
shot_caption(s, 0.55, 5.95, IW, "Personal Dashboard", "Rank ring, total XP, streak, XP-growth chart & AI-Coach teaser — real numbers.", VIOLET2)
shot_caption(s, 6.83, 5.95, IW, "Career Progression", "The full department rank ladder with 'You are here' and per-rank headcount.", CYAN)

# ==================================================================
# SLIDE 17c — SCREENSHOTS: MISSIONS & COMPETITION
# ==================================================================
s = slide(); kicker(s, "Live App · Engagement"); heading(s, "Missions, Coaching & Competition"); accent_bar(s); footer(s)
pic(s, "missions", 0.55, 2.05, IW, border=GREEN)
pic(s, "leaderboard", 6.83, 2.05, IW, border=GOLD)
shot_caption(s, 0.55, 5.95, IW, "Daily Missions", "Department quests with live progress bars from real 'today' activity.", GREEN)
shot_caption(s, 6.83, 5.95, IW, "Leaderboard", "“Ranked on real, milestone-only XP — activity can't be farmed.”", GOLD)

# ==================================================================
# SLIDE 17d — SCREENSHOTS: MANAGER & C2B2C
# ==================================================================
s = slide(); kicker(s, "Live App · Ops & Marketplace"); heading(s, "Manager Analytics & the Live Marketplace"); accent_bar(s); footer(s)
pic(s, "manager", 0.55, 2.05, IW, border=AMBER)
pic(s, "market", 6.83, 2.05, IW, border=CYAN)
shot_caption(s, 0.55, 5.95, IW, "Manager Dashboard", "1.1M XP across 305 players; weekly momentum line per department.", AMBER)
shot_caption(s, 6.83, 5.95, IW, "Buyer Marketplace", "Certified inventory with N-pt check, warranty, price & est. EMI — reserve to buy.", CYAN)

# ==================================================================
# SLIDE 18 — LEADERBOARD / RESULTS
# ==================================================================
s = slide(); kicker(s, "It Actually Runs"); heading(s, "Live Leaderboard — Real People, Real XP"); accent_bar(s); footer(s)
lb = [
    ["1", "MANOJ KUMAR B R", "PDI", "97,030", "PDI Legend"],
    ["2", "S HEMAVATHI", "RTO / Regn", "51,965", "RTO Legend"],
    ["3", "JYOTHILAKSHMI C", "Accounts", "36,670", "Accounts Legend"],
    ["4", "VIGNESH", "EDP", "35,100", "EDP Legend"],
    ["5", "ASHWINI RAI MITCHELL", "Customer Care", "35,065", "Care Legend"],
]
ncol=5; col_w=[0.7,3.6,2.3,1.8,2.6]; x=0.7; w=11.95; y=2.0
tot=sum(col_w); col_w=[c*w/tot for c in col_w]
heads=["#","Player","Department","XP","Rank"]
al=[PP_ALIGN.CENTER,PP_ALIGN.LEFT,PP_ALIGN.LEFT,PP_ALIGN.RIGHT,PP_ALIGN.LEFT]
cx=x
for j,hd in enumerate(heads):
    rect(s,cx,y,col_w[j],0.46,fill=VIOLET,radius=False)
    txt(s,cx+0.1,y,col_w[j]-0.18,0.46,[[R(hd,12.5,WHITE,True)]],align=al[j],anchor=MSO_ANCHOR.MIDDLE)
    cx+=col_w[j]
ry=y+0.46
medal=[GOLD, RGBColor(0xC0,0xC6,0xD8), RGBColor(0xCD,0x8B,0x5C), MUTED, MUTED]
for i,row in enumerate(lb):
    cx=x; fill=CARD if i%2==0 else BG2
    for j,cell in enumerate(row):
        rect(s,cx,ry,col_w[j],0.56,fill=fill,radius=False)
        if j==0: col=medal[i]
        elif j==3: col=GOLD
        elif j==4: col=VIOLET2
        else: col=TXT
        txt(s,cx+0.1,ry,col_w[j]-0.18,0.56,[[R(str(cell),12.5,col,j in (0,1,3))]],align=al[j],anchor=MSO_ANCHOR.MIDDLE)
        cx+=col_w[j]
    ry+=0.56
txt(s, 0.7, ry+0.15, 12, 0.9,
    [[R("Every name, department and XP total is computed from the real event log — the app is fully seeded and playable. ", 12.5, MUTED)],
     [R("305 players ranked · 1,508 team bonuses · 3,800 bookings · live activity feed of 800 recent milestones.", 12.5, VIOLET2, True)]],
    space_after=6, line_spacing=1.1)

# ==================================================================
# SLIDE 19 — ARCHITECTURE
# ==================================================================
s = slide(); kicker(s, "Under the Hood"); heading(s, "Architecture & Tech Stack"); accent_bar(s); footer(s)
# pipeline flow
flow = [("3 real CSVs","170k events", MUTED),("build-data.mjs","score + derive", GREEN),
        ("dealerverse.json","compact precompute", CYAN),("React app","instant load", VIOLET)]
xx=0.7; bw=2.85
for i,(t,d,c) in enumerate(flow):
    rect(s,xx,2.0,bw,1.0,fill=CARD,line=c,line_w=1.25)
    txt(s,xx+0.15,2.15,bw-0.3,0.5,[[R(t,13,c,True)]],align=PP_ALIGN.CENTER)
    txt(s,xx+0.15,2.58,bw-0.3,0.35,[[R(d,10.5,MUTED)]],align=PP_ALIGN.CENTER)
    if i<3: txt(s,xx+bw-0.06,2.0,0.28,1.0,[[R("›",22,MUTED,True)]],align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
    xx+=bw+0.10

rect(s, 0.7, 3.35, 5.85, 3.15, fill=CARD, line=LINE)
txt(s, 1.0, 3.55, 5.3, 0.4, [[R("FRONTEND STACK", 12, VIOLET2, True)]])
bullets(s, 1.0, 4.05, 5.3, 2.4, [
    ("React 19 + Vite", "SPA, hash-routed, 16 screens"),
    ("Tailwind CSS", "dark design system"),
    ("Framer Motion", "animations & confetti"),
    ("Recharts", "manager & department charts"),
    ("Zustand", "identity + workflow state (persisted)"),
    ("PapaParse", "CSV parsing in the pipeline"),
], size=12, gap=6.5)

rect(s, 6.75, 3.35, 5.85, 3.15, fill=CARD, line=LINE)
txt(s, 7.05, 3.55, 5.3, 0.4, [[R("DESIGN DECISIONS", 12, GREEN, True)]])
bullets(s, 7.05, 4.05, 5.3, 2.4, [
    ("Precompute, don't recompute", "170k rows scored once, offline."),
    ("Two zustand stores", "identity.v1 + workflow.v2, localStorage-persisted."),
    ("No backend needed", "fully client-side; deploys as static files."),
    ("Deterministic everything", "PINs, coach insights, thresholds — reproducible."),
    ("Real actor crediting", "XP maps to the true person & department."),
], size=12, gap=6.5, mcolor=GREEN)

# ==================================================================
# SLIDE 20 — IMPACT / WHY IT WINS
# ==================================================================
s = slide(); kicker(s, "Why It Wins"); heading(s, "From Activity Log to Behaviour Change"); accent_bar(s); footer(s)
wins = [
    ("Rewards the right behaviour", "Outcomes and hand-offs score; noise and edits don't. The game can't push people to game it.", GREEN),
    ("Fair to every department", "Per-department ladders mean a 3-person PDI team competes as hard as 241 in Sales.", VIOLET),
    ("Collaboration is structural", "Team bonuses make helping another department the fastest way to earn — 1,508 already paid.", GOLD),
    ("Actually built & playable", "Not slides — a 16-screen app running on real data, plus a full buy/sell marketplace.", CYAN),
    ("Sustainable for months", "Once-per-booking + whitelist keeps the economy stable; missions & seasons refresh engagement.", TEAL),
]
for i,(t,d,c) in enumerate(wins):
    y=2.0+i*0.94
    rect(s,0.7,y,11.9,0.82,fill=CARD,line=LINE)
    rect(s,0.7,y,0.09,0.82,fill=c,radius=False)
    txt(s,1.0,y,3.9,0.82,[[R(t,14,c,True)]],anchor=MSO_ANCHOR.MIDDLE)
    txt(s,4.95,y+0.1,7.5,0.65,[[R(d,12,TXT)]],anchor=MSO_ANCHOR.MIDDLE,line_spacing=1.05)

# ==================================================================
# SLIDE 21 — ROADMAP
# ==================================================================
s = slide(); kicker(s, "What's Next"); heading(s, "Roadmap"); accent_bar(s); footer(s)
road = [
    ("Now — Shipped", GREEN, ["7 depts gamified", "≤20-action XP engine", "Ranks, missions, badges", "Full C2B2C loop", "PIN login, 16 screens"]),
    ("Next — Live pilot", AMBER, ["Stream events in real time", "Branch-vs-branch leagues", "Weekly 'boss battle' challenges", "Manager reward controls"]),
    ("Later — Scale", CYAN, ["Multi-dealer tenancy", "Mobile app + push streaks", "Configurable scoring per group", "Analytics & churn insight"]),
]
cw=3.83
for i,(t,c,items) in enumerate(road):
    x=0.7+i*(cw+0.18)
    rect(s,x,2.0,cw,4.4,fill=CARD,line=LINE)
    rect(s,x,2.0,cw,0.6,fill=c,radius=False)
    txt(s,x+0.2,2.0,cw-0.4,0.6,[[R(t,14.5,BG,True)]],anchor=MSO_ANCHOR.MIDDLE)
    bullets(s,x+0.25,2.85,cw-0.45,3.4,items,size=12.5,gap=11,mcolor=c)
txt(s, 0.7, 6.6, 12, 0.4,
    [[R("The scoring engine is a whitelist + weights table — new milestones or seasons are a config change, not a rebuild.", 12.5, VIOLET2, True)]])

# ==================================================================
# SLIDE 22 — THANK YOU / DEMO
# ==================================================================
s = slide()
rect(s, -1, -1, 15, 9.5, fill=BG2, radius=False)
rect(s, 0, 0, 0.25, 7.5, fill=VIOLET, radius=False)
txt(s, 0.9, 1.4, 11.6, 0.5, [[R("DEMO IT YOURSELF", 14, VIOLET2, True)]])
txt(s, 0.85, 1.85, 11.8, 1.1, [[R("Thank You", 54, WHITE, True, FONT_SB)]])
txt(s, 0.9, 3.0, 11.6, 0.6, [[R("DealerVerse AI — real dealership work, turned into a game people want to play.", 16, MUTED)]])

rect(s, 0.9, 3.9, 5.8, 2.5, fill=CARD, line=LINE)
txt(s, 1.2, 4.1, 5.2, 0.4, [[R("RUN IT", 12, GREEN, True)]])
txt(s, 1.2, 4.6, 5.2, 1.7,
    [[R("cd dealerverse-ai", 13, TXT, True, "Consolas")],
     [R("npm install", 13, TXT, True, "Consolas")],
     [R("npm run data     # build JSON (once)", 12, MUTED, False, "Consolas")],
     [R("npm run dev      # open the app", 13, TXT, True, "Consolas")]], space_after=8)

rect(s, 6.95, 3.9, 5.75, 2.5, fill=CARD, line=LINE)
txt(s, 7.25, 4.1, 5.1, 0.4, [[R("SAMPLE LOGINS  (dept · PIN)", 12, CYAN, True)]])
logins = [
    ("Sales · AJEYA GOWDA", "9183"),
    ("Finance · SANTOSH K SHETTY", "3596"),
    ("PDI · MANOJ KUMAR B R", "9774"),
    ("Accounts · JYOTHILAKSHMI C", "6064"),
]
yy=4.62
for t,p in logins:
    txt(s, 7.25, yy, 4.0, 0.4, [[R(t, 12, TXT)]])
    txt(s, 11.3, yy, 1.15, 0.4, [[R(p, 13, GOLD, True, "Consolas")]], align=PP_ALIGN.RIGHT)
    yy+=0.42
txt(s, 7.25, 6.28, 5.1, 0.3, [[R("Or log in as a Buyer / Seller to try the marketplace.", 10.5, MUTED)]])
txt(s, 0.9, 6.75, 11.6, 0.4, [[R("CarVerse Dealership Gamification Hackathon  ·  Built on 170,162 real events", 11.5, MUTED)]])

# ---------- save ----------
_END = True
out = r"C:\Users\reddy\Desktop\trash\CARVERSE\DealerVerse_AI_Complete.pptx"
prs.save(out)
print("Saved:", out)
print("Slides:", len(prs.slides._sldIdLst))
