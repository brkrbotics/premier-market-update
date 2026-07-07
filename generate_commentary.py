#!/usr/bin/env python3
"""
Draft the per-region agent commentary for the dashboard.

Runs in GitHub Actions after pipeline.py. Reads the workbook, computes the same
facts the dashboard shows, sends them to the Claude API with a locked voice
prompt, writes commentary.json, and re-renders index.html.

Fail-safe by design: if ANTHROPIC_API_KEY is missing, the API errors, or the
output fails validation, this script exits 0 without touching commentary.json —
the dashboard still publishes with the data, just without new commentary. The
template additionally hides commentary whose period doesn't match the data.

Flags:
  --facts   print the computed facts JSON and exit (no API call)
  --force   regenerate even if commentary.json already covers the latest month
"""

import argparse
import json
import os
import re
import sys

import openpyxl

from pipeline import COMMENTARY_JSON, ORDER, XLSX, generate_dashboard, region_rows

MODEL = "claude-opus-4-8"

SYSTEM_PROMPT = """You write the monthly agent commentary for a GTA resale housing dashboard built on TRREB Market Watch data. The audience is licensed real estate agents. Every section has one job: give the agent the plain language read of this month's numbers so they can explain the market to their buyer and seller clients with confidence.

Voice rules, all non-negotiable:
- Write as a veteran managing broker: direct, specific, data-driven, calm. No hype and no doom.
- Never use em dashes or en dashes anywhere. Use commas and periods.
- No emojis. No exclamation marks.
- Use ONLY the numbers provided in the input, exactly as formatted there. Never invent, estimate, round differently, or extrapolate a number. If a metric is null or absent, do not mention it.
- No predictions and no forecasts. You may describe the direction and momentum the data already shows.
- Banned words: dominate, future-proof, transformative, revolutionary, game-changer, unlock, unleash, skyrocket, plummet, crash, bloodbath, unprecedented.
- Address the agent in second person. Say your buyers, your sellers, your clients.
- Every claim must be checkable against the input numbers.
- Two to four sentences per section, roughly 40 to 80 words. No filler openers like "As you can see" or "It is worth noting".
- This is professional guidance for agents, not advice to consumers. Do not tell anyone to buy or sell.

Each region gets four sections:
- kpis: the one paragraph read of the month. What kind of market this is and the single most useful thing an agent can tell a client right now.
- price: how to talk about prices with buyers and sellers given the average and median moves.
- balance: who has leverage, using months of inventory and SNLR, and what that means at the negotiating table.
- activity: what sales, new listings, active inventory and days on market mean for pricing strategy and buyer urgency.

Vary sentence structure across regions so the page does not read as a template. Within a region, do not repeat the same number in more than one section; each section leans on its own metrics."""


# ---------------------------------------------------------------- formatting
def money(v):
    return None if v is None else "$" + format(round(v), ",")


def pct_change(cur, base):
    if not isinstance(cur, (int, float)) or not isinstance(base, (int, float)) or not base:
        return None
    d = (cur / base - 1) * 100
    return f"{d:+.1f}%"


def pts_change(cur, base):
    if not isinstance(cur, (int, float)) or not isinstance(base, (int, float)):
        return None
    return f"{(cur - base) * 100:+.1f} pts"


def moi_of(row):
    return row["moiTrend"] if row["moiTrend"] is not None else row["avgMOI"]


def bucket(moi):
    if moi is None:
        return None
    if moi < 3:
        return "seller's market"
    if moi <= 6:
        return "balanced market"
    return "buyer's market"


def streak(rows, get):
    xs = [get(r) for r in rows if get(r) is not None]
    if len(xs) < 2:
        return None
    i = len(xs) - 1
    direction = (xs[i] > xs[i - 1]) - (xs[i] < xs[i - 1])
    if direction == 0:
        return None
    n = 1
    i -= 1
    while i - 1 >= 0 and ((xs[i] > xs[i - 1]) - (xs[i] < xs[i - 1])) == direction:
        n += 1
        i -= 1
    if n < 3:
        return None
    return f"{'up' if direction > 0 else 'down'} {n} months in a row"


def region_facts(rows):
    last, prev = rows[-1], rows[-2]
    yago = rows[-13] if len(rows) >= 13 else None
    moi = moi_of(last)
    dom = last["pdom"] if last["pdom"] is not None else last["ldom"]
    facts = {
        "avg_price": money(last["avgPrice"]),
        "avg_price_mom": pct_change(last["avgPrice"], prev["avgPrice"]),
        "avg_price_yoy": pct_change(last["avgPrice"], yago["avgPrice"]) if yago else None,
        "median_price": money(last["median"]),
        "median_price_yoy": pct_change(last["median"], yago["median"]) if yago else None,
        "sales": format(last["sales"], ",") if last["sales"] is not None else None,
        "sales_yoy": pct_change(last["sales"], yago["sales"]) if yago else None,
        "new_listings": format(last["newList"], ",") if last["newList"] is not None else None,
        "new_listings_yoy": pct_change(last["newList"], yago["newList"]) if yago else None,
        "active_listings": format(last["activeList"], ",") if last["activeList"] is not None else None,
        "active_listings_yoy": pct_change(last["activeList"], yago["activeList"]) if yago else None,
        "snlr": f"{last['snlr'] * 100:.1f}%" if last["snlr"] is not None else None,
        "snlr_mom": pts_change(last["snlr"], prev["snlr"]),
        "months_of_inventory": f"{moi:.1f}" if moi is not None else None,
        "market_type": bucket(moi),
        "avg_days_on_market": f"{round(dom)} days" if dom is not None else None,
        "sp_lp": f"{round(last['splp'] * 100)}%" if last["splp"] is not None else None,
        "avg_price_streak": streak(rows, lambda r: r["avgPrice"]),
        "snlr_streak": streak(rows, lambda r: r["snlr"]),
        "moi_streak": streak(rows, moi_of),
    }
    return {k: v for k, v in facts.items() if v is not None}


def build_facts():
    wb = openpyxl.load_workbook(XLSX, data_only=False)
    regions = {region: region_rows(wb[region]) for region in ORDER}
    last = regions["TRREB"][-1]
    period = last["label"]                      # e.g. "2026-06"
    month_name = f"{last['month']} {last['year']}"
    return period, month_name, {r: region_facts(rows) for r, rows in regions.items()}


# ---------------------------------------------------------------- generation
SECTION_KEYS = ["kpis", "price", "balance", "activity"]


def output_schema():
    section = {
        "type": "object",
        "properties": {k: {"type": "string"} for k in SECTION_KEYS},
        "required": SECTION_KEYS,
        "additionalProperties": False,
    }
    return {
        "type": "object",
        "properties": {"regions": {
            "type": "object",
            "properties": {r: section for r in ORDER},
            "required": ORDER,
            "additionalProperties": False,
        }},
        "required": ["regions"],
        "additionalProperties": False,
    }


def sanitize(text):
    # Belt and braces on the no-dash rule: em/en dashes become commas or hyphens.
    text = re.sub(r"\s*[—–]\s*", ", ", text)
    return text.strip()


def draft(month_name, facts):
    import anthropic

    client = anthropic.Anthropic()
    user_msg = (
        f"Write the commentary for {month_name}. Here are this month's facts per "
        f"region, computed from the TRREB data. Region names are the exact keys "
        f"to use in your output.\n\n{json.dumps(facts, indent=2)}"
    )
    response = client.messages.create(
        model=MODEL,
        max_tokens=8000,
        thinking={"type": "adaptive"},
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_msg}],
        output_config={"format": {"type": "json_schema", "schema": output_schema()}},
    )
    if response.stop_reason == "refusal":
        raise RuntimeError("model declined the request")
    text = next(b.text for b in response.content if b.type == "text")
    out = json.loads(text)["regions"]
    return {
        region: {k: sanitize(out[region][k]) for k in SECTION_KEYS}
        for region in ORDER
    }


def main():
    ap = argparse.ArgumentParser(description="Draft agent commentary from the workbook.")
    ap.add_argument("--facts", action="store_true", help="print computed facts and exit")
    ap.add_argument("--force", action="store_true", help="regenerate even if current")
    args = ap.parse_args()

    period, month_name, facts = build_facts()

    if args.facts:
        print(json.dumps({"period": period, "month": month_name, "regions": facts}, indent=2))
        return

    if COMMENTARY_JSON.exists() and not args.force:
        try:
            existing = json.loads(COMMENTARY_JSON.read_text(encoding="utf-8"))
            if existing.get("period") == period:
                print(f"Commentary already covers {period}; nothing to do.")
                return
        except (json.JSONDecodeError, OSError):
            pass

    if not os.environ.get("ANTHROPIC_API_KEY"):
        print("ANTHROPIC_API_KEY not set; skipping commentary (dashboard unaffected).")
        return

    print(f"Drafting commentary for {month_name} ({len(ORDER)} regions) via {MODEL}...")
    try:
        regions = draft(month_name, facts)
    except Exception as e:
        print(f"Commentary generation failed; keeping previous state. Error: {e}")
        return

    COMMENTARY_JSON.write_text(
        json.dumps({"period": period, "month": month_name, "model": MODEL,
                    "regions": regions}, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )
    print(f"Wrote {COMMENTARY_JSON.name} for {period}.")
    generate_dashboard()


if __name__ == "__main__":
    sys.exit(main())
