"""
retailer_rates.py
------------------
Given a live-looked-up Network Tariff Code (NTC) plus the state/distribution
and charge lines already extracted from a bill, this module finds the
retailer + rate-card combination in Master_Retailers_Comparison.xlsx
("Comparison" tab) that produces the LOWEST total cost for THIS customer's
actual usage, and returns a `proposed_charges` list shaped exactly like
`bill_data["charges"]` (same order, same length) so it can be passed
straight into:

    single_excel_filler.fill_quote(bill_data, out_path,
                                    tariff_override=ntc,
                                    proposed_charges=proposed_charges)

...to populate the "New Proposed Offer" (J/K) columns without touching any
of the row-insertion / merge / formula logic that already works for the
"Current Energy Offer" (E/F) columns.

HOW THE COMPARISON SHEET IS LAID OUT
("Comparison" tab of Master_Retailers_Comparison.xlsx):

    Col A-F  : Type, Tariff Code, CL Component(s), State, Distribution,
               # Retailers Offering
    Col G-T  : ORIGIN      block (14 cols)
    Col U-AH : MOMENTUM    block
    Col AI-AV: 1ST ENERGY  block
    Col AW-BJ: NBE         block
    Col BK-BX: ALINTA      block
    Col BY-CL: EA          block

Each 14-column block, in order:
    Daily supply charge, Peak, Peak 2, Shoulder, Off Peak, CL1, CL2, CL3,
    Capacity Charges, Demand 1, Demand 2, Solar, Discount, Sign-up Credit

A blank cell means that retailer doesn't publish a rate for that field on
that tariff row - this is NOT the same as a rate of $0, and is treated as
"can't fully price this retailer" below (see `_missing`).

ASSUMPTIONS - FLAGGED HERE BECAUSE THEY MATERIALLY CHANGE THE NUMBERS AND
ARE EASY TO GET WRONG SILENTLY. Worth spot-checking against 2-3 known real
offers before trusting this for a live quote:

  1. UNITS: the Comparison sheet's rate fields (everything except Discount
     and Sign-up Credit) are stored in cents and converted to dollars at
     load time (see CENTS_TO_DOLLARS_FIELDS) to match the bill's own
     extracted rates and the template's Current Offer column. Discount is
     read as a plain percentage (20 means 20%) and Sign-up Credit is
     assumed to already be a dollar amount - if either of those turns out
     to be stored differently for some rows, this will misprice just that
     row.
  2. DISCOUNT SCOPE: the same discount fraction is applied to every matched
     charge line, including the Daily Supply Charge, mirroring how the
     template already treats discount per-row (E/F, J/K). If a real offer's
     discount is usage-only, this slightly overstates that retailer's
     saving.
  3. SIGN-UP CREDIT: this is a one-off bill credit, not a per-usage rate, so
     it doesn't fit the line-item J/K structure. It is NOT applied to the
     ranking or written anywhere on the quote - it's returned in
     `meta["signup_credit"]` for Raj to surface manually. Two retailers
     that are ~equal on running cost won't be re-ranked just because one
     has a bigger sign-up credit.
  4. CONTROLLED LOAD: the "CL Component(s)" column (e.g. "060, 070") hints
     that a Base Tariff row is meant to be paired with a specific CL add-on
     tariff row for full accuracy. This module reads CL1/CL2/CL3 straight
     off the SAME row as the base tariff (whatever each retailer publishes
     there) and does NOT separately look up a "Controlled Load Add-on" type
     row. Flag any CL-heavy bill for a manual check.
  5. MATCHING: Tariff Code + State + Distribution (all three, case/space
     insensitive). If nothing matches on all three, falls back to Tariff
     Code + State, then Tariff Code alone - `meta["match_level"]` tells you
     which one actually fired so a loose match is never silent.
  6. UNMATCHED CHARGE LINES: a bill charge whose description can't be
     mapped to one of the 12 rate fields (see `classify_charge`) is left
     out of both the ranking and the proposed_charges output for that row
     (rate=None) - it's flagged in `meta["unmatched_descriptions"]` rather
     than guessed at.
"""
from __future__ import annotations

import os
import re
from dataclasses import dataclass, field

import openpyxl

COMPARISON_SHEET = "Comparison"

# 14 fields per retailer block, in on-sheet column order.
FIELD_ORDER = [
    "daily_supply", "peak", "peak2", "shoulder", "offpeak",
    "cl1", "cl2", "cl3", "capacity", "demand1", "demand2",
    "solar", "discount", "signup_credit",
]

# 1-indexed starting column of each retailer's 14-col block (col G = 7).
RETAILER_START_COL = {
    "ORIGIN": 7,
    "MOMENTUM": 21,
    "1ST ENERGY": 35,
    "NBE": 49,
    "ALINTA": 63,
    "EA": 77,
}

# Fields that represent a credit paid TO the customer rather than a charge.
CREDIT_FIELDS = {"solar"}

# Fields the ranking/costing logic actually uses per bill line (discount and
# signup_credit are handled separately, not matched against a charge line).
CHARGE_FIELDS = [f for f in FIELD_ORDER if f not in ("discount", "signup_credit")]

# The Comparison sheet stores every rate field in CENTS (e.g. 50.00 = 50c),
# same as most AU retail rate cards, but the bill's own extracted rates -
# and the template's Current Offer column - are in DOLLARS (e.g. 0.50).
# Every field in CHARGE_FIELDS is divided by 100 at load time below so the
# New Proposed Offer lines up with the Current Offer's units. Discount (a
# percentage) and Sign-up Credit (already a dollar amount) are left alone.
CENTS_TO_DOLLARS_FIELDS = set(CHARGE_FIELDS)

# Brand colors for the divider column (I30, merged I30:I{total_row-1}) so
# the quote visually flags which retailer the New Proposed Offer is from.
# ARGB hex, as required by openpyxl's PatternFill.
RETAILER_COLORS = {
    "MOMENTUM": "FFADD8E6",     # Light Blue
    "ALINTA": "FFFFA500",       # Orange
    "NBE": "FF90EE90",          # Light Green (Next Business Energy)
    "ORIGIN": "FFFF0000",       # Red
    "1ST ENERGY": "FF00008B",   # Dark Blue
    "EA": "FF808080",           # Grey (Energy Australia)
}


def retailer_color_hex(retailer: str | None) -> str | None:
    if not retailer:
        return None
    return RETAILER_COLORS.get(retailer.strip().upper())


def _norm(s) -> str:
    return re.sub(r"\s+", " ", str(s or "")).strip().lower()


@dataclass
class TariffRow:
    row_num: int
    type_: str
    tariff_code: str
    cl_components: str
    state: str
    distribution: str
    retailers: dict = field(default_factory=dict)  # retailer -> {field: value|None}


def load_comparison(path: str) -> list[TariffRow]:
    """Loads every tariff row from the Comparison sheet into memory."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[COMPARISON_SHEET]
    rows = []
    for r in range(3, ws.max_row + 1):
        tariff_code = ws.cell(row=r, column=2).value
        if tariff_code in (None, ""):
            continue
        retailers = {}
        for name, start_col in RETAILER_START_COL.items():
            vals = {}
            for i, fname in enumerate(FIELD_ORDER):
                raw = ws.cell(row=r, column=start_col + i).value
                if fname in CENTS_TO_DOLLARS_FIELDS and isinstance(raw, (int, float)):
                    raw = raw / 100.0
                vals[fname] = raw
            retailers[name] = vals
        rows.append(
            TariffRow(
                row_num=r,
                type_=ws.cell(row=r, column=1).value,
                tariff_code=str(tariff_code).strip(),
                cl_components=ws.cell(row=r, column=3).value,
                state=ws.cell(row=r, column=4).value,
                distribution=ws.cell(row=r, column=5).value,
                retailers=retailers,
            )
        )
    return rows


# ---------------------------------------------------------------------------
# Charge line -> rate field classification
# ---------------------------------------------------------------------------
def classify_charges(charges: list[dict]) -> list[str | None]:
    """
    Maps each bill charge line to one of CHARGE_FIELDS by keyword, in order.
    Handles the "second occurrence" ambiguity for peak/peak2 and
    demand1/demand2 by counting - the first "peak"-ish line found is `peak`,
    the next is `peak2`, etc. Returns a list the same length as `charges`,
    with None for any line that couldn't be confidently classified.
    """
    seen_peak = 0
    seen_demand = 0
    out: list[str | None] = []
    for c in charges:
        d = _norm(c.get("description"))
        u = _norm(c.get("unit"))
        text = f"{d} {u}"

        if "solar" in text or "feed" in text or "feed-in" in text or "fit" in d.split():
            out.append("solar")
        elif "daily" in text or ("supply" in text and "charge" in text):
            out.append("daily_supply")
        elif "shoulder" in text:
            out.append("shoulder")
        elif "off peak" in text or "offpeak" in text or "off-peak" in text:
            out.append("offpeak")
        elif "controlled load" in text or re.search(r"\bcl\s*1\b", text):
            out.append("cl1")
        elif re.search(r"\bcl\s*2\b", text):
            out.append("cl2")
        elif re.search(r"\bcl\s*3\b", text):
            out.append("cl3")
        elif "capacity" in text:
            out.append("capacity")
        elif "demand" in text:
            seen_demand += 1
            out.append("demand1" if seen_demand == 1 else "demand2")
        elif "peak" in text:  # catches plain "Peak Usage" after off-peak/shoulder are ruled out above
            seen_peak += 1
            out.append("peak" if seen_peak == 1 else "peak2")
        else:
            out.append(None)
    return out


# ---------------------------------------------------------------------------
# Matching + costing
# ---------------------------------------------------------------------------
def _find_candidates(rows: list[TariffRow], ntc: str, state: str | None, distribution: str | None):
    ntc_n = _norm(ntc)
    state_n = _norm(state)
    dist_n = _norm(distribution)

    by_code = [r for r in rows if _norm(r.tariff_code) == ntc_n]
    if not by_code:
        return [], "no_tariff_code_match"

    if state_n:
        by_code_state = [r for r in by_code if _norm(r.state) == state_n]
    else:
        by_code_state = by_code

    if state_n and dist_n:
        by_all = [r for r in by_code_state if _norm(r.distribution) == dist_n]
        if by_all:
            return by_all, "tariff_state_distribution"

    if by_code_state:
        return by_code_state, "tariff_state"

    return by_code, "tariff_code_only"


def _missing(value) -> bool:
    return value is None or value == ""


def _cost_for_retailer(charges: list[dict], field_map: list[str | None], rates: dict, billing_days: float):
    """
    Returns (total_cost, per_line_rates, unavailable) for one retailer's
    rate card against this bill's actual quantities.
    `unavailable` is True if any matched line has no published rate for that
    retailer on this tariff row (can't be fully priced -> disqualified).
    """
    discount_raw = rates.get("discount")
    discount_frac = (discount_raw / 100.0) if isinstance(discount_raw, (int, float)) else 0.0

    total = 0.0
    per_line = []
    unavailable = False
    for charge, fname in zip(charges, field_map):
        if fname is None:
            per_line.append(None)
            continue
        rate = rates.get(fname)
        if _missing(rate):
            unavailable = True
            per_line.append(None)
            continue
        qty = charge.get("quantity") or 0
        signed_rate = -abs(rate) if fname in CREDIT_FIELDS else rate
        line_cost = qty * signed_rate * (1 - discount_frac)
        total += line_cost
        per_line.append({"rate": signed_rate, "discount_pct": discount_frac})
    return total, per_line, unavailable


def find_best_offer(
    rows: list[TariffRow],
    charges: list[dict],
    ntc: str,
    state: str | None,
    distribution: str | None,
    billing_days: float = 30,
):
    """
    Returns a dict:
        {
          "retailer": "NBE" | None,
          "tariff_row": TariffRow | None,
          "total_cost": float | None,
          "per_line_rates": [ {"rate":.., "discount_pct":..} | None, ... ],
          "match_level": "...",
          "field_map": [...],
          "candidates_considered": int,
          "eligible_retailers": [...],
        }
    If nothing prices cleanly, retailer/tariff_row/total_cost are None and
    the caller should fall back to leaving the New Proposed Offer columns
    blank rather than guessing.
    """
    field_map = classify_charges(charges)
    candidates, match_level = _find_candidates(rows, ntc, state, distribution)

    best = None
    eligible = []
    for row in candidates:
        for retailer, rates in row.retailers.items():
            # Skip retailers with no data at all on this row.
            if all(_missing(v) for v in rates.values()):
                continue
            total, per_line, unavailable = _cost_for_retailer(charges, field_map, rates, billing_days)
            if unavailable:
                continue  # can't fully price this retailer for this bill's line items
            eligible.append((retailer, row, total))
            if best is None or total < best[2]:
                best = (retailer, row, total, per_line)

    if best is None:
        return {
            "retailer": None,
            "tariff_row": None,
            "total_cost": None,
            "per_line_rates": [None] * len(charges),
            "match_level": match_level,
            "field_map": field_map,
            "candidates_considered": len(candidates),
            "eligible_retailers": [],
        }

    retailer, row, total, per_line = best
    return {
        "retailer": retailer,
        "tariff_row": row,
        "total_cost": total,
        "per_line_rates": per_line,
        "match_level": match_level,
        "field_map": field_map,
        "candidates_considered": len(candidates),
        "eligible_retailers": sorted({e[0] for e in eligible}),
        "signup_credit": row.retailers[retailer].get("signup_credit"),
    }


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------
def build_proposed_charges(bill_data: dict, ntc: str, comparison_path: str) -> tuple[list[dict], dict]:
    """
    High-level helper for app.py:

        proposed_charges, meta = build_proposed_charges(
            extracted["data"], live_ntc, "Master_Retailers_Comparison.xlsx"
        )
        fill_quote(extracted["data"], out_path,
                   tariff_override=live_ntc,
                   proposed_charges=proposed_charges)

    `proposed_charges` is the SAME length/order as bill_data["charges"], each
    entry either None (leave that row's J/K blank - couldn't price it) or
    {"rate_before_discount": .., "conditional_discount_pct": .., "is_credit": ..}
    ready to be written by single_excel_filler in place of the bill's own
    rate_before_discount/conditional_discount_pct for that row.
    """
    charges = bill_data.get("charges") or []
    rows = load_comparison(comparison_path)
    result = find_best_offer(
        rows,
        charges,
        ntc=ntc,
        state=bill_data.get("state"),
        distribution=bill_data.get("distribution_region"),
        billing_days=bill_data.get("billing_period_days") or 30,
    )

    proposed_charges = []
    for charge, line in zip(charges, result["per_line_rates"]):
        if line is None:
            proposed_charges.append(None)
        else:
            proposed_charges.append(
                {
                    "rate_before_discount": abs(line["rate"]),
                    "conditional_discount_pct": line["discount_pct"],
                    "is_credit": line["rate"] < 0,
                }
            )

    unmatched = [
        c.get("description")
        for c, f in zip(charges, result["field_map"])
        if f is None
    ]

    current_total = None
    try:
        current_total = sum(
            (c.get("quantity") or 0)
            * (c.get("rate_before_discount") or 0)
            * (1 - (c.get("conditional_discount_pct") or 0))
            * (-1 if c.get("is_credit") else 1)
            for c in charges
        )
    except Exception:
        pass

    meta = {
        "retailer": result["retailer"],
        "tariff_code": result["tariff_row"].tariff_code if result["tariff_row"] else None,
        "match_level": result["match_level"],
        "candidates_considered": result["candidates_considered"],
        "eligible_retailers": result["eligible_retailers"],
        "proposed_total": result["total_cost"],
        "current_total": current_total,
        "estimated_saving": (
            current_total - result["total_cost"]
            if current_total is not None and result["total_cost"] is not None
            else None
        ),
        "signup_credit": result.get("signup_credit"),
        "unmatched_descriptions": unmatched,
    }
    return proposed_charges, meta
