"""
Sends a SINGLE PDF that contains bills for MULTIPLE sites (e.g. a Strata /
Owners Corporation portfolio statement, or several bills merged into one PDF
for a consolidated quote) to Gemini and gets back a list of per-site
extractions matching ConsolidatedBillExtraction in schema.py.

This mirrors gemini_client.py's dual-model cross-check strategy (Flash as
primary with thinking enabled, Flash-Lite as a cheap second opinion, retry
with backoff on transient 429/503 errors, degrade to single-model if one
side fails outright) but adapted for a *list* of sites instead of one bill:

- The two models are not guaranteed to segment the PDF into sites in the same
  order, or even find the same number of sites (one might merge two meters
  read on the same page, the other might not). Rather than compare list
  index-to-index like the single-bill client does, sites from the two models
  are matched by nmi_or_mirn first, falling back to site_address, before
  being cross-checked field by field. Anything that can't be matched at all
  is flagged rather than silently dropped.
- The primary (Flash) model's site list is always what actually gets
  returned/used; the secondary model exists purely to flag discrepancies for
  human review, same as the single-bill client.

Usage:
    from consolidated_gemini_client import extract_consolidated_bills
    result = extract_consolidated_bills("portfolio_bill.pdf")
"""
import json
import math
import os
import random
import time
from typing import Optional

from google import genai
from google.genai import types

from schema import ConsolidatedBillExtraction, get_consolidated_json_schema

PRIMARY_MODEL = "gemini-2.5-flash"
SECONDARY_MODEL = "gemini-2.5-flash-lite"

# Same retry tuning as gemini_client.py — see that file's module docstring
# for the reasoning (a full run of retries on one model call takes roughly
# 4 + 8 + 16 + 32 = 60s plus jitter before giving up).
MAX_RETRIES = 5
BASE_DELAY_SECONDS = 4
MAX_DELAY_SECONDS = 30
RETRYABLE_ERROR_MARKERS = ("503", "UNAVAILABLE", "429", "RESOURCE_EXHAUSTED", "500", "INTERNAL")

SYSTEM_PROMPT = """You are a deterministic Australian electricity/gas bill extraction engine.
Read the attached PDF, which contains bills/statements for MULTIPLE separate sites or meters
(for example: a Strata / Owners Corporation portfolio statement covering several properties, or
several individual bills that have been merged into one PDF). Return JSON matching the supplied
schema exactly — one entry in "bills" per distinct site/meter/NMI, in the order they appear.

Rules for splitting sites:
- Treat each distinct NMI/MIRN or site address as its own separate entry in "bills", even if
  several appear on the same page or share the same customer_name/oc_number.
- Do not combine charges from two different NMIs/sites into a single entry.
- If the same site appears more than once (e.g. duplicated summary + detail pages), only include
  it once, using the most recent billing perioed.

Rules for each site's extraction (same as single-bill extraction):
- Preserve numbers, names, and addresses exactly as they appear on the bill.
- Convert cent-based rates (c/kWh, c/day) to dollars (divide by 100).
- Do not invent, infer, or calculate values that are not clearly present. If a field cannot be
  found confidently, return null for it rather than guessing.
- distribution_region must be the network DISTRIBUTOR (poles-and-wires company), not the retailer.
  Only fill it if it's explicitly printed on the bill (e.g. next to faults/emergencies contact info).
- Never swap peak, off-peak, shoulder etc with different charge. (e.g. Peak charges must go with Peak & off-peak charges must go with off-peak and so on).
- Include every charge line used to reach that site's total: daily/supply charge, and every energy
  charge band (peak, off-peak, shoulder, controlled load, anytime, demand), plus solar feed-in as a
  credit line if present. Do not include Rounding Adjustment. Do not include GST as its own line, and do not include the total itself
  as a line.
- If the bill shows a conditional / guaranteed / pay-on-time discount, capture it as a decimal
  fraction on the charge line(s) it applies to.
- oc_number vs customer_name: if the customer-name position on the bill shows an Owners
  Corporation / Strata Plan number (e.g. "OC 123456", "SP 12345") followed by a mailing/postal
  address rather than a person's name, put that number in oc_number and leave customer_name null.
  Otherwise fill customer_name normally and leave oc_number null.

  
Rules for is_usage:
- Include every charge line used to reach the total.
- is_usage: Set this to TRUE only if the charge is directly related to energy consumption 
  (e.g., Peak, Off-Peak, Shoulder, Controlled Load, Usage, kWh, MJ).
- is_usage: Set this to FALSE if the charge is a fixed fee (e.g., Daily Supply Charge, 
  Service Fee), a credit (e.g., Solar Feed-in Tariff), or a cost adjustment (e.g., Demand Charge).
"""


def _client(api_key: str) -> genai.Client:
    return genai.Client(api_key=api_key)


def _is_retryable(error: Exception) -> bool:
    msg = str(error)
    return any(marker in msg for marker in RETRYABLE_ERROR_MARKERS)


def _is_rate_limit(error: Exception) -> bool:
    msg = str(error)
    return "429" in msg or "RESOURCE_EXHAUSTED" in msg


def _call_with_retry(fn, keys: list[str]):
    """Same key-failover retry as single_gemini_client._call_with_retry: backs
    off on the same key for transient 503/500 errors, but jumps straight to
    the next configured key on a 429 rate-limit instead of waiting it out."""
    last_error = None
    key_idx = 0
    for attempt in range(MAX_RETRIES):
        key = keys[key_idx % len(keys)]
        client = _client(key)
        try:
            return fn(client)
        except Exception as e:  # noqa: BLE001
            last_error = e
            is_last_attempt = attempt == MAX_RETRIES - 1
            if not _is_retryable(e) or is_last_attempt:
                raise
            if _is_rate_limit(e) and len(keys) > 1:
                key_idx += 1
                continue
            delay = min(BASE_DELAY_SECONDS * (2 ** attempt), MAX_DELAY_SECONDS)
            delay += random.uniform(0, delay * 0.25)
            time.sleep(delay)
    raise last_error  # pragma: no cover


def _extract_with_model(keys: list[str], pdf_bytes: bytes, model: str, config_type: str) -> dict:
    config_kwargs = dict(
        system_instruction=SYSTEM_PROMPT,
        temperature=0,
        response_mime_type="application/json",
        response_json_schema=get_consolidated_json_schema(),
    )

    if "gemini-3" in model:
        if config_type == "primary":
            config_kwargs["thinking_config"] = types.ThinkingConfig(thinking_level="high")
        elif config_type == "secondary":
            config_kwargs["thinking_config"] = types.ThinkingConfig(thinking_level="medium")
    else:
        if config_type == "primary":
            # Multi-site PDFs need more room to think than a single bill —
            # double the single-bill budget.
            config_kwargs["thinking_config"] = types.ThinkingConfig(thinking_budget=4096)
        elif config_type == "secondary":
            config_kwargs["thinking_config"] = types.ThinkingConfig(thinking_budget=0)

    def _do_call(client: genai.Client):
        response = client.models.generate_content(
            model=model,
            contents=[
                types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf"),
                "Extract every site/bill in this PDF into the schema.",
            ],
            config=types.GenerateContentConfig(**config_kwargs),
        )
        return json.loads(response.text)

    return _call_with_retry(_do_call, keys)


def _floats_close(a: Optional[float], b: Optional[float], tol: float = 0.01) -> bool:
    if a is None or b is None:
        return a == b
    return math.isclose(a, b, abs_tol=tol)


def _norm_str(s: Optional[str]) -> Optional[str]:
    if s is None:
        return None
    return " ".join(s.split()).strip().lower()


def _strings_match(a: Optional[str], b: Optional[str]) -> bool:
    na, nb = _norm_str(a), _norm_str(b)
    if na is None or nb is None:
        return na == nb
    return na == nb


def _site_key(bill: dict) -> Optional[str]:
    """Best available identifier to match the same site across both models' outputs."""
    nmi = _norm_str(bill.get("nmi_or_mirn"))
    if nmi:
        return f"nmi:{nmi}"
    addr = _norm_str(bill.get("site_address"))
    if addr:
        return f"addr:{addr}"
    return None


def _compare_bill(primary: dict, secondary: dict) -> list[str]:
    """Same field-by-field comparison as gemini_client._compare, for one matched site pair."""
    mismatches = []

    scalar_str_fields = [
        "customer_name", "oc_number", "site_address", "nmi_or_mirn", "distribution_region",
        "tariff_classification", "current_energy_retailer", "state",
    ]
    for f in scalar_str_fields:
        if not _strings_match(primary.get(f), secondary.get(f)):
            mismatches.append(f)

    scalar_num_fields = ["billing_period_days", "total_due"]
    for f in scalar_num_fields:
        if not _floats_close(primary.get(f), secondary.get(f)):
            mismatches.append(f)

    p_charges = primary.get("charges") or []
    s_charges = secondary.get("charges") or []
    if len(p_charges) != len(s_charges):
        mismatches.append("charges (different number of line items found)")
    else:
        for i, (pc, sc) in enumerate(zip(p_charges, s_charges)):
            line_mismatches = []
            if not _strings_match(pc.get("description"), sc.get("description")):
                line_mismatches.append("description")
            if not _floats_close(pc.get("quantity"), sc.get("quantity")):
                line_mismatches.append("quantity")
            if not _floats_close(pc.get("rate_before_discount"), sc.get("rate_before_discount")):
                line_mismatches.append("rate_before_discount")
            if not _floats_close(pc.get("conditional_discount_pct"), sc.get("conditional_discount_pct")):
                line_mismatches.append("conditional_discount_pct")
            if line_mismatches:
                mismatches.append(f"charges[{i}] ({pc.get('description', '?')}): {', '.join(line_mismatches)}")

    return mismatches


def _compare_site_lists(primary_bills: list, secondary_bills: list) -> list[str]:
    """
    Matches sites between the two models by NMI/address and cross-checks each
    matched pair. Returns a flat list of human-readable review notes, one
    (or more) per site that has discrepancies or couldn't be matched at all.
    """
    review_notes = []

    secondary_by_key = {}
    secondary_unkeyed = []
    for sb in secondary_bills:
        key = _site_key(sb)
        if key:
            secondary_by_key.setdefault(key, []).append(sb)
        else:
            secondary_unkeyed.append(sb)

    matched_secondary_ids = set()

    for idx, pb in enumerate(primary_bills):
        label = pb.get("site_address") or pb.get("nmi_or_mirn") or f"bills[{idx}]"
        key = _site_key(pb)
        candidates = secondary_by_key.get(key, []) if key else []

        if not candidates:
            review_notes.append(
                f"{label}: no matching site found by the second-opinion model "
                f"(could be a real discrepancy, or just how it segmented the PDF) — recommend manual check."
            )
            continue

        sb = candidates.pop(0)
        matched_secondary_ids.add(id(sb))
        mismatches = _compare_bill(pb, sb)
        if mismatches:
            review_notes.append(f"{label}: {', '.join(mismatches)}")

    # Anything left in the secondary model's output that primary never matched
    # (either genuinely extra sites it found, or unkeyed leftovers) is worth a
    # heads-up too, since it may mean primary silently missed a site.
    leftover = [sb for group in secondary_by_key.values() for sb in group if id(sb) not in matched_secondary_ids]
    leftover += secondary_unkeyed
    for sb in leftover:
        label = sb.get("site_address") or sb.get("nmi_or_mirn") or "(unidentified site)"
        review_notes.append(
            f"{label}: only found by the second-opinion model, not the primary model — recommend manual check."
        )

    return review_notes


def extract_consolidated_bills(pdf_path: str, api_keys: Optional[list[str]] = None) -> dict:
    """
    Runs the two-model cross-check on a multi-site PDF and returns:
        {
            "success": bool,
            "data": {"bills": [dict, ...]}   # validated against ConsolidatedBillExtraction
            "review_notes": [str, ...]       # per-site discrepancies/unmatched sites
            "model_errors": {"primary": str|None, "secondary": str|None}
            "degraded_mode": bool            # True if only one model succeeded
        }

    api_keys: ordered list of keys to try, failing over to the next one on a
    429 rate-limit. Defaults to a single-item list from GEMINI_API_KEY.
    """
    keys = api_keys or ([os.environ["GEMINI_API_KEY"]] if os.environ.get("GEMINI_API_KEY") else None)
    if not keys:
        raise RuntimeError("GEMINI_API_KEY environment variable is not set.")

    with open(pdf_path, "rb") as f:
        pdf_bytes = f.read()

    primary_raw, secondary_raw = None, None
    model_errors = {"primary": None, "secondary": None}

    try:
        primary_raw = _extract_with_model(keys, pdf_bytes, PRIMARY_MODEL, config_type="primary")
    except Exception as e:  # noqa: BLE001
        model_errors["primary"] = str(e)

    try:
        secondary_raw = _extract_with_model(keys, pdf_bytes, SECONDARY_MODEL, config_type="secondary")
    except Exception as e:  # noqa: BLE001
        model_errors["secondary"] = str(e)

    if primary_raw is None and secondary_raw is None:
        return {
            "success": False,
            "data": None,
            "review_notes": [],
            "model_errors": model_errors,
            "degraded_mode": False,
        }

    degraded_mode = primary_raw is None
    source_raw = secondary_raw if degraded_mode else primary_raw

    try:
        validated = ConsolidatedBillExtraction.model_validate(source_raw)
    except Exception as e:  # noqa: BLE001
        return {
            "success": False,
            "data": source_raw,
            "review_notes": [],
            "model_errors": {**model_errors, "validation": str(e)},
            "degraded_mode": degraded_mode,
        }

    review_notes = []
    if not degraded_mode and secondary_raw is not None:
        review_notes = _compare_site_lists(
            source_raw.get("bills") or [], secondary_raw.get("bills") or []
        )

    return {
        "success": True,
        "data": validated.model_dump(),
        "review_notes": review_notes,
        "model_errors": model_errors,
        "degraded_mode": degraded_mode,
    }
