"""
CARVERSE Grand Prix - Dashboard
Run:  streamlit run app.py
"""
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

import scoring_engine as E

st.set_page_config(page_title="CARVERSE Grand Prix", page_icon="🏁", layout="wide")

# --------------------------------------------------------------------------- #
#  THEME / CSS
# --------------------------------------------------------------------------- #
CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Orbitron:wght@600;700;800&family=Rajdhani:wght@500;600;700&display=swap');

html, body, [class*="css"], .stMarkdown, p, span, div { font-family: 'Rajdhani', sans-serif; }
.stApp { background: radial-gradient(circle at 15% -10%, #17213d 0%, #0b0f19 45%, #06090f 100%); }
.block-container { padding-top: 1.5rem; max-width: 1300px; }

/* checkered top strip */
.flag-strip { height: 8px; border-radius: 6px; margin-bottom: 4px;
  background-image: repeating-linear-gradient(45deg,#fff 0 10px,#111 10px 20px);
  opacity:.85; }

h1 { font-family:'Orbitron',sans-serif!important; font-weight:800!important; letter-spacing:2px;
  background:linear-gradient(90deg,#ff2e2e,#ff9d3d); -webkit-background-clip:text;
  -webkit-text-fill-color:transparent; margin-bottom:0!important; }
h2,h3 { font-family:'Orbitron',sans-serif!important; letter-spacing:.5px; color:#f0f4ff!important; }
.subtitle { color:#8b97b4; margin-top:2px; font-size:1.02rem; }

/* stat cards */
.stat-card { background:linear-gradient(160deg,#161d2e,#0e1320); border:1px solid #26304a;
  border-radius:14px; padding:14px 16px; box-shadow:0 8px 22px rgba(0,0,0,.45); transition:.2s; height:100%; }
.stat-card:hover { border-color:#ff2e2e; box-shadow:0 0 20px rgba(255,46,46,.28); transform:translateY(-2px); }
.stat-label { font-size:.68rem; letter-spacing:1.6px; text-transform:uppercase; color:#8b97b4; }
.stat-value { font-family:'Orbitron'; font-size:1.7rem; font-weight:800; color:#fff; line-height:1.25; }
.stat-sub { font-size:.72rem; color:#5f6b88; }
.accent .stat-value { color:#ff5b4d; }

/* leaderboard rows */
.lb-row { display:flex; align-items:center; gap:14px; background:linear-gradient(160deg,#141b2b,#0f1420);
  border:1px solid #222c44; border-left:4px solid #2a9d8f; border-radius:12px; padding:10px 16px; margin-bottom:8px; }
.lb-row:hover { border-left-color:#ff2e2e; }
.lb-rank { font-family:'Orbitron'; font-size:1.15rem; font-weight:800; width:44px; color:#8b97b4; }
.lb-name { flex:1; font-size:1.05rem; font-weight:700; color:#eef2ff; }
.lb-dept { font-size:.78rem; color:#8b97b4; text-transform:uppercase; letter-spacing:1px; }
.lb-pts { font-family:'Orbitron'; font-weight:800; font-size:1.15rem; color:#ffb703; }
.gold{border-left-color:#ffd700!important;} .silver{border-left-color:#c0c0c0!important;} .bronze{border-left-color:#cd7f32!important;}

/* quest cards */
.quest { background:linear-gradient(160deg,#141b2b,#0f1420); border:1px solid #222c44; border-radius:14px;
  padding:16px 18px; margin-bottom:12px; }
.quest-title { font-family:'Orbitron'; font-weight:700; font-size:1.05rem; color:#fff; }
.quest-desc { color:#8b97b4; font-size:.9rem; margin:2px 0 8px; }
.quest-num { font-family:'Orbitron'; color:#2ee6a6; font-weight:700; }

/* tabs */
.stTabs [data-baseweb="tab-list"] { gap:6px; }
.stTabs [data-baseweb="tab"] { background:#121829; border:1px solid #222c44; border-radius:10px 10px 0 0;
  padding:8px 16px; font-weight:600; }
.stTabs [aria-selected="true"] { background:#1c2540; border-color:#ff2e2e; color:#fff!important; }

/* progress bars */
.stProgress > div > div > div { background:linear-gradient(90deg,#ff2e2e,#ff9d3d)!important; }

.pill { display:inline-block; background:#1c2540; border:1px solid #2a3556; border-radius:20px;
  padding:2px 12px; font-size:.8rem; color:#a9b6d6; margin:2px 4px 2px 0; }
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)

DEPT_COLORS = {"SALES":"#4cc9f0","FINANCE":"#f72585","ACCOUNTS":"#ffb703",
               "EDP":"#8338ec","CUSTOMER CARE":"#2ee6a6","PDI":"#ff8fab","RTO / REGN TEAM":"#adb5bd"}
STAGE_COLOR = {"PRE-BOOKING":"#457b9d","PENDING-BOOKING":"#4cc9f0","IBNA":"#ffb703",
               "ABNR":"#ff9d3d","RBND":"#2ee6a6","DELIVERED":"#2a9d8f","DORMANT":"#6c757d","CANCELLED":"#e63946"}

def dark(fig, h=380):
    fig.update_layout(height=h, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                      font=dict(color="#c9d4ee", family="Rajdhani"), margin=dict(t=20,b=10,l=10,r=10),
                      xaxis=dict(gridcolor="#1b2338"), yaxis=dict(gridcolor="#1b2338"))
    return fig

# --------------------------------------------------------------------------- #
#  DATA (cached)
# --------------------------------------------------------------------------- #
@st.cache_data(show_spinner="🏁 Warming up the grid — crunching 170k dealership events...")
def build():
    ev, emp, loc, loc_name = E.load_data()
    se, state = E.compute_scores(ev)
    ind, dept, branch = E.leaderboards(se, loc_name)
    audit = E.anti_gaming_audit(ev)
    summ = E.summary(ev, se, state)
    state["branch"] = state["location_code"].map(loc_name).fillna(state["location_code"])
    se["branch"] = se["location_code"].map(loc_name).fillna(se["location_code"])
    return dict(ev=ev, se=se, state=state, ind=ind, dept=dept, branch=branch,
                audit=audit, summ=summ, loc_name=loc_name)

D = build()

# --------------------------------------------------------------------------- #
#  HEADER
# --------------------------------------------------------------------------- #
st.markdown("<div class='flag-strip'></div>", unsafe_allow_html=True)
st.markdown("<h1>🏁 CARVERSE GRAND PRIX</h1>"
            "<p class='subtitle'>Every booking is a race car. Move it from enquiry to delivery — "
            "cleanly and fast — to score. Departments are the pit crew; nobody wins the finish line alone.</p>",
            unsafe_allow_html=True)
st.write("")

def card(label, value, sub="", accent=False):
    cls = "stat-card accent" if accent else "stat-card"
    st.markdown(f"<div class='{cls}'><div class='stat-label'>{label}</div>"
                f"<div class='stat-value'>{value}</div><div class='stat-sub'>{sub}</div></div>",
                unsafe_allow_html=True)

s = D["summ"]
noise_pct = 100 * s["noise_events"] / s["total_events"]
cols = st.columns(6)
data = [("Events processed", f"{s['total_events']:,}", "raw log rows", False),
        ("Bookings on track", f"{s['bookings']:,}", "unique race cars", False),
        ("Delivered", f"{s['delivered']:,}", "crossed the line 🏁", True),
        ("Distinct actions", f"{s['distinct_actions']:,}", "in the raw data", False),
        ("Scored actions", f"{s['scored_action_types']} / 20", "under the cap", False),
        ("Noise ignored", f"{noise_pct:.0f}%", "earns ZERO points", True)]
for c, (l, v, sub, a) in zip(cols, data):
    with c: card(l, v, sub, a)
st.write("")

tab_track, tab_lead, tab_quest, tab_anti, tab_rules = st.tabs(
    ["🏁 Race Track", "🏆 Leaderboards", "🎯 Quests & Boss Battles",
     "🛡️ Anti-Gaming", "📖 How Scoring Works"])

# --------------------------------------------------------------------------- #
#  RACE TRACK
# --------------------------------------------------------------------------- #
with tab_track:
    st.subheader("🏎️ Branch Grand Prix — race to the finish line")
    st.caption("Each branch is a car. Position = % of its bookings that reached delivery. First to 100% wins.")

    prog = (D["state"].groupby("branch")["delivered"].mean() * 100)
    total = D["state"].groupby("branch").size()
    race = pd.DataFrame({"prog": prog, "cars": total}).dropna()
    race = race[race["cars"] >= 20].sort_values("prog").tail(10)

    fig = go.Figure()
    for br, row in race.iterrows():
        fig.add_trace(go.Scatter(x=[0, row["prog"]], y=[br, br], mode="lines",
                                 line=dict(color="#222c44", width=8), showlegend=False, hoverinfo="skip"))
        fig.add_trace(go.Scatter(x=[row["prog"]], y=[br], mode="text", text="🏎️",
                                 textfont=dict(size=24), showlegend=False,
                                 hovertemplate=f"{br}<br>{row['prog']:.0f}% delivered<br>{int(row['cars'])} cars<extra></extra>"))
    fig.add_vline(x=100, line=dict(color="#ffffff", width=2, dash="dot"))
    fig.add_annotation(x=100, y=len(race)-0.4, text="🏁 FINISH", showarrow=False,
                       font=dict(color="#fff", size=14, family="Orbitron"), xanchor="right")
    fig.update_xaxes(range=[0, 108], title="% delivered")
    st.plotly_chart(dark(fig, 430), width="stretch")

    st.divider()
    c1, c2 = st.columns([1.3, 1])
    with c1:
        st.markdown("**Where every car sits on the track right now**")
        order = ["PRE-BOOKING","PENDING-BOOKING","IBNA","ABNR","RBND","DELIVERED","DORMANT","CANCELLED"]
        counts = D["state"]["current_stage"].value_counts().reindex(order).fillna(0)
        f2 = go.Figure(go.Bar(x=[E.STAGE_LABEL[o] for o in order], y=counts.values,
                              marker_color=[STAGE_COLOR[o] for o in order],
                              text=[int(v) for v in counts.values], textposition="outside"))
        f2.update_layout(xaxis_tickangle=-25)
        st.plotly_chart(dark(f2, 360), width="stretch")
    with c2:
        st.markdown("**⚡ Fastest deliveries (cycle time)**")
        fast = D["state"][D["state"]["delivered"]].dropna(subset=["cycle_days"])
        fast = fast[fast["cycle_days"] >= 0].sort_values("cycle_days").head(9)
        for _, r in fast.iterrows():
            st.markdown(f"<div class='lb-row' style='padding:6px 14px'>"
                        f"<span class='lb-rank' style='font-size:.9rem'>{r['cycle_days']:.1f}d</span>"
                        f"<span class='lb-name' style='font-size:.9rem'>🏎️ {r['booking']}</span>"
                        f"<span class='lb-dept'>{r['branch']}</span></div>", unsafe_allow_html=True)

# --------------------------------------------------------------------------- #
#  LEADERBOARDS
# --------------------------------------------------------------------------- #
with tab_lead:
    view = st.radio("League", ["🏢 Branch vs Branch", "👥 Department", "🧑 Individual"],
                    horizontal=True, label_visibility="collapsed")
    if view.startswith("🏢"):
        b = D["branch"].head(12)
        fig = px.bar(b, x="realized", y="branch", orientation="h", text="realized")
        fig.update_traces(marker_color="#ff5b4d")
        fig.update_layout(yaxis={"categoryorder":"total ascending"}, xaxis_title="points")
        st.plotly_chart(dark(fig, 520), width="stretch")
    elif view.startswith("👥"):
        dpt = D["dept"]
        fig = px.bar(dpt, x="department", y="realized", text="realized", color="department",
                     color_discrete_map=DEPT_COLORS)
        fig.update_layout(showlegend=False, xaxis_title="", yaxis_title="points")
        st.plotly_chart(dark(fig, 460), width="stretch")
        st.info("💡 Finance scores lowest **by design**: the raw data has no clean `FINANCE_APPROVED` "
                "action — only `FINANCE_INFO_UPDATED` (effort). So finance is rewarded only when a "
                "booking actually clears the finance stage, never for editing data.")
    else:
        dept_f = st.selectbox("Filter by department", ["ALL"] + sorted(D["ind"]["department"].unique()))
        ind = D["ind"] if dept_f == "ALL" else D["ind"][D["ind"]["department"] == dept_f]
        medals = ["🥇","🥈","🥉"]
        cls = ["gold","silver","bronze"]
        for i, (_, r) in enumerate(ind.head(20).iterrows()):
            badge = medals[i] if i < 3 else f"{i+1}"
            rowcls = f"lb-row {cls[i]}" if i < 3 else "lb-row"
            st.markdown(f"<div class='{rowcls}'><span class='lb-rank'>{badge}</span>"
                        f"<span class='lb-name'>{r['username']}</span>"
                        f"<span class='lb-dept'>{r['department']}</span>"
                        f"<span class='lb-pts'>{int(r['realized']):,}</span></div>", unsafe_allow_html=True)

# --------------------------------------------------------------------------- #
#  QUESTS & BOSS BATTLES
# --------------------------------------------------------------------------- #
with tab_quest:
    delivered = int(D["state"]["delivered"].sum())
    clean_deliv = int((D["state"]["delivered"] & (D["state"]["rework"] == 0)).sum())
    fast_deliv = int((D["state"]["delivered"] & (D["state"]["cycle_days"] <= 14)).sum())
    doc_sets = int((D["se"]["action_code"] == "DOCUMENT_SET_COMPLETED").sum())

    st.subheader("🎯 Active Quests")
    st.caption("Team goals derived from real activity this period. Progress is live.")
    quests = [
        ("🏁 Checkered Flag", "Deliver 1,500 vehicles", delivered, 1500),
        ("✨ Clean Run", "Deliver 800 cars with ZERO rework / escalation loops", clean_deliv, 800),
        ("⚡ Speed Demon", "Deliver 600 cars within 14 days of booking", fast_deliv, 600),
        ("📑 Paperwork Pro", "Complete 1,000 full document sets", doc_sets, 1000),
    ]
    qc = st.columns(2)
    for i, (title, desc, val, goal) in enumerate(quests):
        pct = min(val / goal, 1.0)
        status = "✅ CLEARED!" if pct >= 1 else ""
        with qc[i % 2]:
            st.markdown(f"<div class='quest'><div class='quest-title'>{title} {status}</div>"
                        f"<div class='quest-desc'>{desc}</div></div>", unsafe_allow_html=True)
            st.progress(pct, text=f"{val:,} / {goal:,}  ({pct*100:.0f}%)")

    st.divider()
    st.subheader("👹 Weekly Boss Battle — “The Bottleneck”")
    st.markdown("Branches battle the boss by pushing cars **invoice → delivery** with no escalations. "
                "Damage dealt = clean deliveries.")
    boss_hp, dmg = 2000, clean_deliv
    st.progress(min(dmg / boss_hp, 1.0), text=f"Boss HP {max(boss_hp-dmg,0):,} / {boss_hp:,}")
    st.markdown(f"<span class='pill'>💥 Damage dealt: {dmg:,}</span>"
                f"<span class='pill'>🏆 Top striker: {D['branch'].iloc[0]['branch']}</span>"
                f"<span class='pill'>⚔️ {len(D['branch'])} branches in the raid</span>",
                unsafe_allow_html=True)

# --------------------------------------------------------------------------- #
#  ANTI-GAMING
# --------------------------------------------------------------------------- #
with tab_anti:
    st.subheader("🛡️ Why this system can't be farmed")
    a = D["audit"].head(25)
    fig = px.bar(a, x="event_count", y="action_code", orientation="h", color="role",
                 color_discrete_map={"SCORED milestone":"#2ee6a6","doc component (rolls up)":"#ffb703",
                                     "NOISE - 0 pts":"#48546f"}, text="event_count")
    fig.update_layout(yaxis={"categoryorder":"total ascending"}, legend_title="", xaxis_title="event count")
    st.plotly_chart(dark(fig, 640), width="stretch")
    st.info(f"🚫 The 4 most frequent actions in the whole dataset — **USER_LOGGED_IN, DISCOUNT_UPDATED, "
            f"PDI_INFO_UPDATED, FINANCE_INFO_UPDATED** — are pure effort/noise and score **0**. "
            f"Overall **{noise_pct:.0f}%** of all events earn nothing.")

    st.markdown("### The six guardrails")
    guards = [
        ("Milestones only", "Updates, logins, notes = 0 pts. Only outcomes that move a booking forward score."),
        ("Once per booking", "Each milestone scores a single time per car — repeating it 50× does nothing."),
        ("Delivery escrow", f"{int(E.ESCROW_FRACTION*100)}% of points are held until the car actually delivers."),
        ("Net progress", f"Cancelled (-{int(-E.CANCEL_PENALTY)} pts) and dormant bookings reverse points."),
        ("Rework penalty", "Escalation loops shrink the clean-multiplier down to 0.6×."),
        ("Co-op payout", "Delivery bonus is split across depts — you can't win solo, forcing collaboration."),
    ]
    g = st.columns(3)
    for i, (t, d) in enumerate(guards):
        with g[i % 3]:
            st.markdown(f"<div class='quest'><div class='quest-title' style='font-size:.95rem'>"
                        f"{i+1}. {t}</div><div class='quest-desc'>{d}</div></div>", unsafe_allow_html=True)

# --------------------------------------------------------------------------- #
#  RULES
# --------------------------------------------------------------------------- #
with tab_rules:
    st.subheader("📖 The scoring actions — 16 real + 2 derived, under the 20 cap")
    rows = []
    for act, (pts, tier, dept, label) in E.SCORING_ACTIONS.items():
        rows.append({"Tier": tier, "Action": act, "Means": label, "Dept": dept, "Points": pts})
    for name, (pts, tier, dept, label) in E.DERIVED.items():
        rows.append({"Tier": tier, "Action": name + " (derived)", "Means": label, "Dept": dept, "Points": pts})
    tbl = pd.DataFrame(rows).sort_values("Points", ascending=False)
    st.dataframe(tbl, hide_index=True, width="stretch", height=680)
    st.markdown("""
### The gameplay loop
1. **A milestone fires** → the player, their department **and** their branch all score.
2. **Speed × Clean multipliers** reward fast, rework-free movement.
3. When a car is **DELIVERED**, held escrow is released and a **co-op bonus** is split across every
   department that touched it — nobody wins the finish line alone.
4. Weekly **Branch-vs-Branch league** + **Boss battles** reset the competition so it stays fun
   long-term and no single star runs away with it.
""")

st.caption("CARVERSE Grand Prix · data: May–Jun 2026 event log · scoring_engine.py is the single source of truth")
