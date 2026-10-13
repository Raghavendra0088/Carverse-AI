"""
CARVERSE Grand Prix - Scoring Engine
=====================================
Turns raw dealership event logs into a gamified score that rewards REAL
business progress (moving a booking toward delivery) and actively resists
gaming (repeated updates, logins, notes earn nothing).

Design principles (from the hackathon brief):
  * Not all actions are equal        -> only OUTCOME milestones score.
  * Repeated actions can be gamed     -> every milestone scores ONCE per booking.
  * Effort < Outcome                  -> delivery-tier weights dominate.
  * Reward forward progress           -> cancelled/dormant bookings lose points,
                                         escrow is released only on delivery.
  * Encourage collaboration           -> delivered cars pay a co-op bonus that is
                                         split across every dept that touched them.

Booking identity = (location_code, enquiry_no)  [unique, per the dataset notes].
"""

from __future__ import annotations
import os
import pandas as pd

# --------------------------------------------------------------------------- #
#  CONFIG
# --------------------------------------------------------------------------- #
DATA_DIR = os.path.join(os.path.dirname(__file__), "carverse files")
EVENTS_CSV    = os.path.join(DATA_DIR, "z_event_log_may_june_2026.csv")
EMPLOYEES_CSV = os.path.join(DATA_DIR, "z_employees.csv")
LOCATIONS_CSV = os.path.join(DATA_DIR, "z_locations.csv")

# Lifecycle track (the "race track"). Higher index = closer to the finish line.
STAGE_ORDER = {
    "PRE-BOOKING": 0,
    "PENDING-BOOKING": 1,
    "IBNA": 2,
    "ABNR": 3,
    "RBND": 4,
    "DELIVERED": 5,
}
OFF_TRACK = {"DORMANT": -1, "CANCELLED": -2}
STAGE_LABEL = {
    "PRE-BOOKING": "Enquiry / Pre-Booking",
    "PENDING-BOOKING": "Booking Confirmed",
    "IBNA": "Invoicing & Finance",
    "ABNR": "Approvals & Registration",
    "RBND": "Ready for Delivery",
    "DELIVERED": "Delivered 🏁",
    "DORMANT": "Dormant ⚠️",
    "CANCELLED": "Cancelled ❌",
}

# ---------------------------------------------------------------------------
#  THE <=20 SCORING ACTIONS  (hard constraint: max 20)
#  Each is a genuine milestone. tier drives weight. Everything else = 0 points.
# ---------------------------------------------------------------------------
#   tier S : customer handover (the win)
#   tier A : hard money / legal gates
#   tier B : forward-progress gates
SCORING_ACTIONS = {
    # action_code                    : (points, tier, department, label)
    "GATEPASS_ISSUED_BY_ACCOUNTS"    : (100, "S", "ACCOUNTS",      "Vehicle cleared for delivery"),
    "CREDIT_APPROVED"                : ( 45, "A", "ACCOUNTS",      "Payment verified & credited"),
    "INVOICE_APPROVED_EDP"           : ( 45, "A", "EDP",           "Invoice billing finalized"),
    "INSURANCE_APPROVED_BY_CCM"      : ( 40, "A", "CUSTOMER CARE", "Insurance approved"),
    "REGISTRATION_ADDED"             : ( 40, "A", "EDP",           "Registration filed"),
    "BOOKING_CREATED"                : ( 30, "B", "SALES",         "Booking created"),
    "DISCOUNT_APPROVED"              : ( 25, "B", "SALES",         "Pricing finalized"),
    "INVOICE_APPROVED_SALES"         : ( 25, "B", "SALES",         "Sales invoice signed off"),
    "RGISTRATION_APPROVED_SALES"     : ( 25, "B", "SALES",         "Registration approved (sales)"),
    "RGISTRATION_CREATED_EDP"        : ( 25, "B", "EDP",           "Registration processed (EDP)"),
    "RTO_REQUEST"                    : ( 20, "B", "PDI",           "RTO registration initiated"),
    "DISPATCH_BY_PDI"                : ( 20, "B", "PDI",           "Vehicle inspected & dispatched"),
    "DISPATCH_BY_CCM"                : ( 20, "B", "CUSTOMER CARE", "Delivery dispatch coordinated"),
    "GATEPASS_BY_CCM"                : ( 20, "B", "CUSTOMER CARE", "Gatepass cleared (CCM)"),
}

# Two DERIVED milestones (computed, not raw action codes) -> solve two brief traps.
DERIVED = {
    # name                 : (points, tier, department, label)
    "DOCUMENT_SET_COMPLETED": (30, "B", "PDI",     "Full document set completed"),
    "FINANCE_STAGE_CLEARED" : (30, "B", "FINANCE", "Finance cleared, booking advanced"),
}

# Required document components for DOCUMENT_SET_COMPLETED (gamify completion, not pieces).
REQUIRED_DOCS = {
    "DOCUMENTS_PAN", "DOCUMENTS_AADHAAR FRONT", "DOCUMENTS_VIN",
    "DOCUMENTS_FRONT", "DOCUMENTS_REAR", "DOCUMENTS_LEFT", "DOCUMENTS_RIGHT",
}

ESCALATION_ACTIONS = {"DISCOUNT_ESCALATED", "INVOICE_ESCALATED",
                      "REGISTARTION_ESCALATED", "INVOICE_CANCEL_ESCALATED"}

# Anti-gaming knobs -------------------------------------------------------- #
ESCROW_FRACTION       = 0.40   # % of milestone points HELD until the car delivers
DELIVERY_COOP_BONUS   = 150    # split across every dept that cleanly moved a delivered car
CLEAN_FULL            = 1.00   # multiplier when a booking has no rework
CLEAN_FLOOR           = 0.60   # worst-case multiplier for rework-heavy bookings
REWORK_PENALTY_STEP   = 0.10   # each extra escalation (beyond 1/type) costs this
CANCEL_PENALTY        = -50    # a cancelled booking reverses points
DORMANT_PENALTY       = -15


# --------------------------------------------------------------------------- #
#  LOAD
# --------------------------------------------------------------------------- #
def load_data():
    ev = pd.read_csv(EVENTS_CSV, dtype=str, keep_default_na=False)
    emp = pd.read_csv(EMPLOYEES_CSV, dtype=str, keep_default_na=False)
    loc = pd.read_csv(LOCATIONS_CSV, dtype=str, keep_default_na=False)

    ev["created_date"] = pd.to_datetime(ev["created_date"], errors="coerce")
    ev["booking"] = ev["location_code"] + "|" + ev["enquiry_no"]
    # keep only rows tied to a real booking for booking-level logic
    ev["has_booking"] = (ev["enquiry_no"] != "-") & (ev["enquiry_no"] != "")

    loc_name = dict(zip(loc["location_code"], loc["location_name"]))
    return ev, emp, loc, loc_name


# --------------------------------------------------------------------------- #
#  CORE SCORING
# --------------------------------------------------------------------------- #
def compute_scores(ev: pd.DataFrame):
    """Returns (scoring_events_df, booking_state_df)."""
    booked = ev[ev["has_booking"]].copy()

    # ---- booking state: furthest stage reached, delivered?, timings, rework --
    booked["stage_rank"] = booked["stage"].map(STAGE_ORDER)
    grp = booked.groupby("booking")

    state = pd.DataFrame(index=grp.size().index)
    state["location_code"] = grp["location_code"].first()
    state["max_stage_rank"] = grp["stage_rank"].max()
    state["delivered"] = grp["stage"].apply(lambda s: (s == "DELIVERED").any())
    state["cancelled"] = grp["stage"].apply(lambda s: (s == "CANCELLED").any())
    state["dormant"]   = grp["stage"].apply(lambda s: (s == "DORMANT").any())

    # current stage label = furthest track stage reached (delivered/cancelled win)
    def _cur(s):
        if (s == "CANCELLED").any(): return "CANCELLED"
        if (s == "DELIVERED").any(): return "DELIVERED"
        ontrack = [x for x in s if x in STAGE_ORDER]
        if ontrack:
            return max(ontrack, key=lambda x: STAGE_ORDER[x])
        if (s == "DORMANT").any(): return "DORMANT"
        return "PRE-BOOKING"
    state["current_stage"] = grp["stage"].apply(_cur)

    # cycle time (booking_created -> delivered) in days, for delivered cars
    created = booked[booked["action_code"] == "BOOKING_CREATED"].groupby("booking")["created_date"].min()
    delivered_at = (booked[booked["action_code"] == "GATEPASS_ISSUED_BY_ACCOUNTS"]
                    .groupby("booking")["created_date"].min())
    state["created_at"] = created
    state["delivered_at"] = delivered_at
    state["cycle_days"] = (state["delivered_at"] - state["created_at"]).dt.total_seconds() / 86400

    # rework = escalations beyond the first per escalation type
    esc = booked[booked["action_code"].isin(ESCALATION_ACTIONS)]
    esc_counts = esc.groupby(["booking", "action_code"]).size()
    extra = (esc_counts - 1).clip(lower=0).groupby("booking").sum()
    state["rework"] = extra.reindex(state.index).fillna(0).astype(int)

    # clean multiplier per booking
    state["clean_mult"] = (CLEAN_FULL - state["rework"] * REWORK_PENALTY_STEP).clip(lower=CLEAN_FLOOR)

    # ---- build scoring events -------------------------------------------- --
    rows = []

    # (1) raw milestone actions, ONCE per (booking, action) -> credit first actor
    mil = booked[booked["action_code"].isin(SCORING_ACTIONS)].sort_values("created_date")
    first = mil.groupby(["booking", "action_code"], as_index=False).first()
    for r in first.itertuples(index=False):
        pts, tier, dept, label = SCORING_ACTIONS[r.action_code]
        rows.append(_event(r.booking, r.action_code, label, tier, pts,
                           r.user_id, r.username, dept, r.location_code, r.created_date))

    # (2) DERIVED: DOCUMENT_SET_COMPLETED -> when all REQUIRED_DOCS present
    docs = booked[booked["action_code"].isin(REQUIRED_DOCS)]
    if len(docs):
        have = docs.groupby("booking")["action_code"].nunique()
        complete = have[have >= len(REQUIRED_DOCS)].index
        # credit the actor of the LAST required doc (the one who finished the set)
        lastdoc = (docs[docs["booking"].isin(complete)]
                   .sort_values("created_date").groupby("booking").last())
        pts, tier, dept, label = DERIVED["DOCUMENT_SET_COMPLETED"]
        for bk, r in lastdoc.iterrows():
            rows.append(_event(bk, "DOCUMENT_SET_COMPLETED", label, tier, pts,
                               r["user_id"], r["username"], dept, r["location_code"], r["created_date"]))

    # (3) DERIVED: FINANCE_STAGE_CLEARED -> booking had finance work AND advanced past IBNA
    fin = booked[booked["department"] == "FINANCE"]
    if len(fin):
        advanced = state.index[state["max_stage_rank"] >= STAGE_ORDER["ABNR"]]
        fin_ok = fin[fin["booking"].isin(advanced)].sort_values("created_date").groupby("booking").last()
        pts, tier, dept, label = DERIVED["FINANCE_STAGE_CLEARED"]
        for bk, r in fin_ok.iterrows():
            rows.append(_event(bk, "FINANCE_STAGE_CLEARED", label, tier, pts,
                               r["user_id"], r["username"], dept, r["location_code"], r["created_date"]))

    se = pd.DataFrame(rows)

    # ---- apply anti-gaming economics ------------------------------------- --
    se = se.merge(state[["clean_mult", "delivered", "cancelled", "dormant"]],
                  left_on="booking", right_index=True, how="left")

    # base after clean multiplier
    se["adj_points"] = se["base_points"] * se["clean_mult"]
    # escrow: keep (1-frac) now; release the rest ONLY if delivered
    se["escrow"] = se["adj_points"] * ESCROW_FRACTION
    se["realized"] = se["adj_points"] - se["escrow"] + se["escrow"] * se["delivered"].astype(int)
    # penalties on off-track bookings
    se.loc[se["cancelled"], "realized"] += CANCEL_PENALTY / se.groupby("booking")["booking"].transform("count")
    se.loc[se["dormant"] & ~se["delivered"], "realized"] += DORMANT_PENALTY / se.groupby("booking")["booking"].transform("count")
    se["realized"] = se["realized"].round(1)

    # ---- delivery co-op bonus (collaboration) ---------------------------- --
    delivered_bookings = state.index[state["delivered"]]
    coop_rows = []
    for bk in delivered_bookings:
        contrib = se[se["booking"] == bk]
        if contrib.empty:
            continue
        # speed bonus multiplier for the whole car
        cd = state.at[bk, "cycle_days"]
        speed = _speed_mult(cd)
        pool = DELIVERY_COOP_BONUS * speed
        # split proportional to each actor's base contribution on this car
        by_user = contrib.groupby(["user_id", "username", "department", "location_code"])["base_points"].sum()
        total = by_user.sum()
        for (uid, uname, dept, loc), base in by_user.items():
            share = round(pool * base / total, 1)
            coop_rows.append(_event(bk, "DELIVERY_COOP_BONUS", "Co-op delivery payout",
                                    "S", share, uid, uname, dept, loc,
                                    state.at[bk, "delivered_at"], realized=share))
    if coop_rows:
        se = pd.concat([se, pd.DataFrame(coop_rows)], ignore_index=True)

    return se, state.reset_index()


def _event(booking, action, label, tier, base_points, uid, uname, dept, loc, ts, realized=None):
    return {
        "booking": booking, "action_code": action, "label": label, "tier": tier,
        "base_points": float(base_points), "user_id": uid, "username": uname,
        "department": dept, "location_code": loc, "created_date": ts,
        "realized": realized if realized is not None else float(base_points),
    }


def _speed_mult(cycle_days):
    if pd.isna(cycle_days):    return 1.0
    if cycle_days <= 7:        return 1.5
    if cycle_days <= 14:       return 1.25
    if cycle_days <= 30:       return 1.0
    return 0.8


# --------------------------------------------------------------------------- #
#  LEADERBOARDS
# --------------------------------------------------------------------------- #
def leaderboards(se: pd.DataFrame, loc_name: dict):
    ind = (se.groupby(["user_id", "username", "department"])["realized"]
           .sum().round(0).reset_index().sort_values("realized", ascending=False))
    ind["rank"] = range(1, len(ind) + 1)

    dept = (se.groupby("department")["realized"].sum().round(0)
            .reset_index().sort_values("realized", ascending=False))

    branch = (se.groupby("location_code")["realized"].sum().round(0).reset_index())
    branch["branch"] = branch["location_code"].map(loc_name).fillna(branch["location_code"])
    branch = branch.sort_values("realized", ascending=False)
    branch["rank"] = range(1, len(branch) + 1)
    return ind, dept, branch


# --------------------------------------------------------------------------- #
#  ANTI-GAMING AUDIT  (proves noise -> 0 points)
# --------------------------------------------------------------------------- #
def anti_gaming_audit(ev: pd.DataFrame):
    counts = ev["action_code"].value_counts()
    scored = set(SCORING_ACTIONS) | REQUIRED_DOCS
    audit = pd.DataFrame({"action_code": counts.index, "event_count": counts.values})
    audit["is_scored"] = audit["action_code"].isin(SCORING_ACTIONS)
    audit["role"] = audit["action_code"].apply(
        lambda a: "SCORED milestone" if a in SCORING_ACTIONS
        else ("doc component (rolls up)" if a in REQUIRED_DOCS
              else "NOISE - 0 pts"))
    return audit


def summary(ev, se, state):
    return {
        "total_events": len(ev),
        "distinct_actions": ev["action_code"].nunique(),
        "scored_action_types": len(SCORING_ACTIONS) + len(DERIVED),
        "bookings": int(state.shape[0]),
        "delivered": int(state["delivered"].sum()),
        "cancelled": int(state["cancelled"].sum()),
        "users": ev["user_id"].nunique(),
        "branches": ev["location_code"].nunique(),
        "scoring_events": len(se),
        "noise_events": int(len(ev) - ev["action_code"].isin(set(SCORING_ACTIONS) | REQUIRED_DOCS).sum()),
    }


if __name__ == "__main__":
    ev, emp, loc, loc_name = load_data()
    se, state = compute_scores(ev)
    ind, dept, branch = leaderboards(se, loc_name)
    s = summary(ev, se, state)
    print("=== SUMMARY ==="); [print(f"  {k:22}: {v}") for k, v in s.items()]
    print("\n=== TOP 10 INDIVIDUALS ==="); print(ind.head(10).to_string(index=False))
    print("\n=== DEPARTMENT LEAGUE ==="); print(dept.to_string(index=False))
    print("\n=== TOP 10 BRANCHES ==="); print(branch.head(10)[["rank","branch","realized"]].to_string(index=False))
    noise = int(len(ev) - ev["action_code"].isin(set(SCORING_ACTIONS) | REQUIRED_DOCS).sum())
    print(f"\nAnti-gaming: {noise:,}/{len(ev):,} events ({100*noise/len(ev):.1f}%) earn ZERO points.")
