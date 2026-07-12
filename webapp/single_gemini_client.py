"""
Sends the bill PDF directly to Gemini (native multimodal document understanding
— no OCR/text-extraction step, which was the biggest source of scrambled-table
errors in the old Groq pipeline) and gets back JSON matching schema.py.

Accuracy strategy: every bill is run through TWO models independently —
Gemini 2.5 Flash (primary, thinking enabled) and Gemini 2.5 Flash-Lite (cheap
second opinion). Fields where they agree are returned with high confidence.
Fields where they disagree keep Flash's value but are listed in
`review_fields` so the review screen / console can flag them for a human
to double check, rather than silently trusting a single model's guess.

RELIABILITY NOTES (free tier):
- 503 UNAVAILABLE means Google's servers are momentarily overloaded, not that
  anything is wrong with the request. Their own guidance is to retry with
  exponential backoff, which is what `_call_with_retry` below does — it will
  usually clear up within 2-4 attempts. 429 RESOURCE_EXHAUSTED (rate limit)
  gets the same treatment since a short wait is often enough for the window
  to reset, though hitting your daily quota repeatedly won't be fixed by
  retrying and will still surface after MAX_RETRIES.
- If Flash keeps failing after all retries but Flash-Lite succeeds, the
  Flash-Lite result is used as a degraded single-model fallback instead of
  failing the whole extraction outright (flagged via `degraded_mode` in the
  result so you can tell it apart from a normal cross-checked extraction).
"""
import json
import math
import os
import random
import time
from typing import Optional

from google import genai
from google.genai import types

from schema import BillExtraction, get_json_schema

PRIMARY_MODEL = "gemini-2.5-flash"
SECONDARY_MODEL = "gemini-2.5-flash-lite"

# Retry tuning: with these defaults, a run that hits 503s on every attempt
# waits roughly 4 + 8 + 16 + 32 = 60s (plus jitter) before giving up on a
# single model call. Lower MAX_RETRIES if you'd rather fail fast while
# testing; raise BASE_DELAY_SECONDS if you're still seeing frequent failures.
MAX_RETRIES = 5
BASE_DELAY_SECONDS = 4
MAX_DELAY_SECONDS = 30
RETRYABLE_ERROR_MARKERS = ("503", "UNAVAILABLE", "429", "RESOURCE_EXHAUSTED", "500", "INTERNAL")

SYSTEM_PROMPT = """You are a deterministic Australian electricity/gas bill extraction engine.
Read the attached bill PDF and return JSON matching the supplied schema exactly.

Rules:
- Preserve numbers, names, and addresses exactly as they appear on the bill.
- Convert cent-based rates (c/kWh, c/day) to dollars (divide by 100).
- Do not invent, infer, or calculate values that are not clearly present. If a field cannot be
  found confidently, return null for it rather than guessing.
- Never swap peak, off-peak, shoulder etc with different charge. (e.g. Peak charges must go with Peak & off-peak charges must go with off-peak and so on).
- distribution_region must be the network DISTRIBUTOR (poles-and-wires company), not the retailer.
  Only fill it if it's explicitly printed on the bill (e.g. next to faults/emergencies contact info).
- Include every charge line used to reach the total: daily/supply charge, and every energy charge
  band (peak, off-peak, shoulder, controlled load, anytime, demand), plus solar feed-in as a credit
  line if present. Do not include GST as its own line, and do not include the total itself as a line.
- If the bill shows a conditional / guaranteed / pay-on-time discount, capture it as a decimal
  fraction on the charge line(s) it applies to.
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
    """
    Runs fn(client), retrying with exponential backoff + jitter on transient
    server errors (503/500). On a rate-limit (429) error, instead of backing
    off on the same key, immediately fails over to the next key in `keys` (if
    more than one is configured) -- no extra backoff wait, no extra tokens,
    just the same single call retried against a key that isn't rate-limited.
    Anything non-retryable (bad key, invalid schema, etc.) raises immediately.
    """
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
                key_idx += 1  # move straight to the next key, no backoff needed
                continue
            delay = min(BASE_DELAY_SECONDS * (2 ** attempt), MAX_DELAY_SECONDS)
            delay += random.uniform(0, delay * 0.25)  # jitter avoids retry storms
            time.sleep(delay)
    raise last_error  # pragma: no cover — loop always returns or raises above



def _extract_with_model(keys: list[str], pdf_bytes: bytes, model: str, config_type: str) -> dict:
    config_kwargs = dict(
        system_instruction=SYSTEM_PROMPT,
        temperature=0,
        response_mime_type="application/json",
        response_json_schema=get_json_schema(),
    )

    # Handle thinking config safely based on model generation limits
    if "gemini-3" in model:
        # Gemini 3 series uses structural levels ('high', 'medium', etc.)
        if config_type == "primary":
            config_kwargs["thinking_config"] = types.ThinkingConfig(thinking_level="high")
        elif config_type == "secondary":
            config_kwargs["thinking_config"] = types.ThinkingConfig(thinking_level="medium")
    else:
        # Gemini 2.5 series uses standard token budgets
        if config_type == "primary":
            # Give primary flash up to 2048 tokens to think through complex nested bill tables
            config_kwargs["thinking_config"] = types.ThinkingConfig(thinking_budget=2048)
        elif config_type == "secondary":
            # Turn thinking entirely off (0 budget) for the fast/cheap second opinion check
            config_kwargs["thinking_config"] = types.ThinkingConfig(thinking_budget=0)

    def _do_call(client: genai.Client):
        response = client.models.generate_content(
            model=model,
            contents=[
                types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf"),
                "Extract this bill into the schema.",
            ],
            config=types.GenerateContentConfig(**config_kwargs),
        )
        return json.loads(response.text)

    return _call_with_retry(_do_call, keys)


def _floats_close(a: Optional[float], b: Optional[float], tol: float = 0.01) -> bool:
    if a is None or b is None:
        return a == b
    return math.isclose(a, b, abs_tol=tol)


def _strings_match(a: Optional[str], b: Optional[str]) -> bool:
    if a is None or b is None:
        return a == b
    return " ".join(a.split()).strip().lower() == " ".join(b.split()).strip().lower()


def _compare(primary: dict, secondary: dict) -> list[str]:
    """Returns a list of top-level field names where the two models disagree."""
    review_fields = []

    scalar_str_fields = [
        "customer_name", "site_address", "nmi_or_mirn", "distribution_region",
        "tariff_classification", "current_energy_retailer", "state",
    ]
    for f in scalar_str_fields:
        if not _strings_match(primary.get(f), secondary.get(f)):
            review_fields.append(f)

    scalar_num_fields = ["billing_period_days", "total_due"]
    for f in scalar_num_fields:
        if not _floats_close(primary.get(f), secondary.get(f)):
            review_fields.append(f)

    p_charges = primary.get("charges") or []
    s_charges = secondary.get("charges") or []
    if len(p_charges) != len(s_charges):
        review_fields.append("charges (different number of line items found)")
    else:
        for i, (pc, sc) in enumerate(zip(p_charges, s_charges)):
            mismatches = []
            if not _strings_match(pc.get("description"), sc.get("description")):
                mismatches.append("description")
            if not _floats_close(pc.get("quantity"), sc.get("quantity")):
                mismatches.append("quantity")
            if not _floats_close(pc.get("rate_before_discount"), sc.get("rate_before_discount")):
                mismatches.append("rate_before_discount")
            if not _floats_close(pc.get("conditional_discount_pct"), sc.get("conditional_discount_pct")):
                mismatches.append("conditional_discount_pct")
            if mismatches:
                review_fields.append(f"charges[{i}] ({pc.get('description', '?')}): {', '.join(mismatches)}")

    return review_fields


def extract_bill(pdf_path: str, api_keys: Optional[list[str]] = None) -> dict:
    """
    Runs the two-model cross-check and returns:
        {
            "success": bool,
            "data": dict (validated against BillExtraction),
            "review_fields": [str, ...]      # fields Flash and Flash-Lite disagreed on
            "model_errors": {"primary": str|None, "secondary": str|None}
            "degraded_mode": bool            # True if only one model succeeded
        }

    api_keys: ordered list of keys to try. If a call hits a rate limit (429)
    on the current key, it fails over to the next key in the list -- no
    splitting, no parallel calls, just the same single request retried on a
    key that isn't rate-limited. Defaults to a single-item list from the
    GEMINI_API_KEY env var if not provided.
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

    # Both models are down (or both hit a non-retryable error) -> nothing to return.
    if primary_raw is None and secondary_raw is None:
        return {
            "success": False,
            "data": None,
            "review_fields": [],
            "model_errors": model_errors,
            "degraded_mode": False,
        }

    # Primary is down but the second opinion made it through: use it rather
    # than failing the whole extraction. There's nothing to cross-check
    # against, so this is flagged as degraded_mode instead of silently
    # pretending it's a normal two-model result.
    degraded_mode = primary_raw is None
    source_raw = secondary_raw if degraded_mode else primary_raw

    try:
        validated = BillExtraction.model_validate(source_raw)
    except Exception as e:  # noqa: BLE001
        return {
            "success": False,
            "data": source_raw,
            "review_fields": [],
            "model_errors": {**model_errors, "validation": str(e)},
            "degraded_mode": degraded_mode,
        }

    review_fields = []
    if not degraded_mode and secondary_raw is not None:
        review_fields = _compare(primary_raw, secondary_raw)

    return {
        "success": True,
        "data": validated.model_dump(),
        "review_fields": review_fields,
        "model_errors": model_errors,
        "degraded_mode": degraded_mode,
    }
