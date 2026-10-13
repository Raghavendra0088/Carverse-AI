"""
Builds carverse.html — a self-contained, animated single-page dashboard.
Pulls every number from scoring_engine.py (the real event log), embeds it as
JSON inline, and writes one portable HTML file (opens by double-click, offline).

    python build_site.py   ->   carverse.html
"""
import json, os
import pandas as pd
import scoring_engine as E

DEPT_COLORS = {
    "SALES": "#2a78d6", "CUSTOMER CARE": "#1baf7a", "ACCOUNTS": "#eda100",
    "EDP": "#4a3aa7", "FINANCE": "#e34948", "PDI": "#e87ba4", "RTO / REGN TEAM": "#eb6834",
    "ACCESSORIES": "#6d7a74", "TRUEVALUE": "#9aa6b2",
}
STAGE_META = [
    ("PRE-BOOKING", "Enquiry", "#86b6ef"),
    ("PENDING-BOOKING", "Booked", "#5598e7"),
    ("IBNA", "Invoice & Finance", "#2a78d6"),
    ("ABNR", "Approvals & Regn", "#eda100"),
    ("RBND", "Ready", "#1baf7a"),
    ("DELIVERED", "Delivered", "#0a9d8e"),
    ("DORMANT", "Dormant", "#9aa6b2"),
    ("CANCELLED", "Cancelled", "#e34948"),
]


def build_payload():
    ev, emp, loc, loc_name = E.load_data()
    se, state = E.compute_scores(ev)
    ind, dept, branch = E.leaderboards(se, loc_name)
    audit = E.anti_gaming_audit(ev)
    summ = E.summary(ev, se, state)
    state["branch"] = state["location_code"].map(loc_name).fillna(state["location_code"])

    noise_pct = round(100 * summ["noise_events"] / summ["total_events"])

    # branch race
    prog = state.groupby("branch")["delivered"].mean() * 100
    cars = state.groupby("branch").size()
    bpts = branch.set_index("branch")["realized"]
    race = pd.DataFrame({"prog": prog, "cars": cars, "points": bpts}).dropna(subset=["prog"])
    race = race[race["cars"] >= 20].sort_values("prog", ascending=False).head(12)
    branch_race = [{"name": i, "prog": round(r["prog"], 1), "cars": int(r["cars"]),
                    "points": int(r["points"]) if pd.notna(r["points"]) else 0}
                   for i, r in race.iterrows()]

    # stage distribution
    counts = state["current_stage"].value_counts()
    stages = [{"key": k, "label": lbl, "count": int(counts.get(k, 0)), "color": c}
              for k, lbl, c in STAGE_META]

    # individuals (top 60)
    individuals = [{"name": r["username"], "dept": r["department"], "points": int(r["realized"])}
                   for _, r in ind.head(60).iterrows()]

    # departments
    departments = [{"name": r["department"], "points": int(r["realized"]),
                    "color": DEPT_COLORS.get(r["department"], "#7c8aa5")}
                   for _, r in dept.iterrows()]

    # quests
    delivered = int(state["delivered"].sum())
    clean = int((state["delivered"] & (state["rework"] == 0)).sum())
    fast = int((state["delivered"] & (state["cycle_days"] <= 14)).sum())
    docsets = int((se["action_code"] == "DOCUMENT_SET_COMPLETED").sum())
    quests = {"delivered": delivered, "clean": clean, "fast": fast, "docsets": docsets,
              "top_branch": branch_race[0]["name"] if branch_race else "-",
              "n_branches": int(summ["branches"])}

    # anti-gaming audit (top 22)
    audit_rows = [{"action": r["action_code"], "count": int(r["event_count"]), "role": r["role"]}
                  for _, r in audit.head(22).iterrows()]

    # fastest deliveries
    fastdf = state[state["delivered"]].dropna(subset=["cycle_days"])
    fastdf = fastdf[fastdf["cycle_days"] >= 0].sort_values("cycle_days").head(10)
    fastest = [{"booking": i if isinstance(i, str) else r.get("booking", ""),
                "branch": r["branch"], "days": round(r["cycle_days"], 1)}
               for i, r in fastdf.iterrows()]
    # booking id is the index? state was reset -> 'booking' col exists
    if "booking" in fastdf.columns:
        fastest = [{"booking": r["booking"], "branch": r["branch"], "days": round(r["cycle_days"], 1)}
                   for _, r in fastdf.iterrows()]

    # scoring actions table
    actions = []
    for act, (pts, tier, d, label) in E.SCORING_ACTIONS.items():
        actions.append({"tier": tier, "action": act, "means": label, "dept": d, "points": pts})
    for name, (pts, tier, d, label) in E.DERIVED.items():
        actions.append({"tier": tier, "action": name + " *", "means": label, "dept": d, "points": pts})
    actions.sort(key=lambda x: -x["points"])

    return {
        "summary": {"events": summ["total_events"], "bookings": summ["bookings"],
                    "delivered": delivered, "distinct": summ["distinct_actions"],
                    "scored": summ["scored_action_types"], "noise_pct": noise_pct,
                    "users": summ["users"], "branches": summ["branches"],
                    "cancelled": summ["cancelled"]},
        "branch_race": branch_race, "stages": stages, "individuals": individuals,
        "departments": departments, "quests": quests, "audit": audit_rows,
        "fastest": fastest, "actions": actions,
        "dept_colors": DEPT_COLORS,
    }


def main():
    payload = build_payload()
    data_json = json.dumps(payload, ensure_ascii=False)
    here = os.path.dirname(__file__)
    with open(os.path.join(here, "template.html"), encoding="utf-8") as f:
        html = f.read()
    html = html.replace("/*__DATA__*/", data_json)
    out = os.path.join(here, "carverse.html")
    with open(out, "w", encoding="utf-8") as f:
        f.write(html)
    print("wrote", out, f"({len(html)//1024} KB)")


if __name__ == "__main__":
    main()
