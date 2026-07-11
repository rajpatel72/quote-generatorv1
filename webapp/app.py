"""
Energy Bill -> Excel Quote Web App
-----------------------------------
Thin web UI around the existing extraction pipeline:
    single_gemini_client.extract_bill()          + single_excel_filler.fill_quote()
    consolidated_gemini_client.extract_consolidated_bills() + consolidated_excel_filler.fill_consolidated_quote()

UI notes:
- Results (including the generated Excel bytes) are cached in st.session_state,
  keyed by mode. Clicking "Download" -- or a failed/retried download -- never
  forces a re-run of the Gemini extraction. The user only pays the processing
  cost once per bill, and can re-download as many times as needed.
- A "Start over" button explicitly clears state when ready for the next bill.
- Extracted data renders as proper tables (metadata table + charges table),
  not raw JSON.
- Download filenames:
    Single:       "Quote - {Street Address} - {Mon YYYY}.xlsx"
    Consolidated: "Consolidated - {Client Name} - {Mon YYYY}.xlsx"
  The month/year used is the date the quote was generated (bills don't carry
  a reliable "issue date" field in the schema) -- flag if you'd rather this
  used something else, e.g. a billing period end date.

Rate Comparison Table (new)
----------------------------
Each site (single or consolidated) now gets an editable "Rate Comparison —
New Offer" table, seeded from the extracted bill charges. Users can:
  - edit the "New Rate" / "New Disc %" columns per line
  - add or remove custom line items (num_rows="dynamic")
  - click "Auto-fill from tariff" to pull suggested rates automatically

FUTURE INTEGRATION NOTE (read this before wiring the New Offer template):
This table is UI + data-model only today. Edited rows are captured into the
generated result dict under `rate_comparison` (single) / `rate_comparisons`
(consolidated, keyed by site index) but are NOT YET passed into
fill_quote() / fill_consolidated_quote(). To finish the loop later:
  1. Add a `new_offer_rates` (or similar) parameter to fill_quote() and
     fill_consolidated_quote() in single_excel_filler.py /
     consolidated_excel_filler.py that writes these rows into the
     template's "New Offer" section.
  2. Pass `res["rate_comparison"]` / `res["rate_comparisons"]` straight
     through when calling those functions -- the row shape
     (description / quantity / unit / current_rate / new_rate / ...)
     is already template-ready, see `comparison_df_to_records()`.
  3. For automatic (non-manual) rate suggestions, add a function such as
     `get_rate_card(tariff_code) -> dict[str, float]` to tariff_bridge.py.
     `_auto_fill_new_rates()` below calls this via getattr() already, so
     it will start working the moment tariff_bridge exposes it -- no
     app.py changes required.
"""
import os
import re
import tempfile
from datetime import datetime

import pandas as pd
import streamlit as st
from pypdf import PdfWriter

import tariff_bridge as tariff_bridge_module
from single_gemini_client import extract_bill
from single_excel_filler import fill_quote
from consolidated_gemini_client import extract_consolidated_bills
from consolidated_excel_filler import fill_consolidated_quote
from api_key_pool import load_api_keys
from tariff_bridge import render_auto_lookup, get_cached_tariff, get_cached_tariff_debug, clear_cached_tariff, render_bulk_lookup_button

st.set_page_config(page_title="Bills \u2192 Quote", page_icon="\u26A1", layout="wide")

# ---------------------------------------------------------------------------
# Look & feel -- premium, quiet, brand-consistent
# ---------------------------------------------------------------------------
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=Inter:wght@400;500;600&display=swap');

    :root {
        --ink: #10192E;
        --ink-soft: #4A5573;
        --line: rgba(16, 25, 46, 0.08);
        --navy: #16294D;
        --navy-2: #23407A;
        --accent: #C4123A;
        --accent-soft: rgba(196, 18, 58, 0.08);
        --good: #1C7A43;
        --good-soft: rgba(46, 204, 113, 0.10);
        --warn: #8A6100;
        --warn-soft: rgba(255, 193, 7, 0.12);
        --card: #FFFFFF;
        --canvas: #F4F6FB;
    }

    html, body { font-family: 'Inter', 'Space Grotesk', sans-serif; }
    p, span, label, li, div { font-family: 'Inter', 'Space Grotesk', sans-serif; }

    .stApp { background: var(--canvas); }
    .block-container { padding-top: 4.2rem; padding-bottom: 3rem; max-width: 1280px; margin: 0 auto; }

    /* ---- Header ---- */
    .hero-badge {
        display: inline-block; padding: 0.22rem 0.8rem; border-radius: 999px;
        background: var(--accent-soft); border: 1px solid rgba(196, 18, 58, 0.28);
        color: var(--accent) !important; font-size: 0.7rem; font-weight: 600; letter-spacing: 0.08em;
        text-transform: uppercase; margin-bottom: 0.65rem;
    }
    .hero-title {
        font-family: 'Space Grotesk', sans-serif;
        font-size: 2.05rem; font-weight: 700; line-height: 1.15; margin: 0.1rem 0 0 0;
        color: var(--navy) !important;
        letter-spacing: -0.01em;
    }
    .hero-sub { color: var(--ink-soft) !important; font-size: 0.95rem; margin: 0.6rem 0 1.5rem 0; line-height: 1.6; }
    .hero-sub ul { margin: 0; padding-left: 1.1rem; }
    .hero-sub li { margin-bottom: 0.2rem; color: var(--ink-soft) !important; }

    /* ---- Section labels ---- */
    .section-label {
        display: flex; align-items: center; gap: 0.5rem;
        font-size: 0.72rem; font-weight: 700; letter-spacing: 0.09em; text-transform: uppercase;
        color: var(--navy-2) !important; margin: 1.4rem 0 0.6rem 0;
    }
    .section-label::before {
        content: ""; display: inline-block; width: 6px; height: 6px; border-radius: 50%;
        background: var(--accent); flex-shrink: 0;
    }
    .meta-caption { color: #7A84A0 !important; font-size: 0.8rem; }

    /* ---- Status banners ---- */
    .field-flag {
        background: var(--warn-soft); border: 1px solid rgba(255, 193, 7, 0.4);
        color: var(--warn) !important; border-radius: 10px; padding: 0.55rem 0.9rem; margin-bottom: 0.4rem;
        font-size: 0.88rem;
    }
    .ok-banner {
        background: var(--good-soft); border: 1px solid rgba(46, 204, 113, 0.35);
        color: var(--good) !important; border-radius: 10px; padding: 0.65rem 1rem; font-size: 0.9rem;
        margin-bottom: 0.7rem;
    }
    .info-banner {
        background: rgba(35, 64, 122, 0.06); border: 1px solid rgba(35, 64, 122, 0.18);
        color: var(--navy-2) !important; border-radius: 10px; padding: 0.6rem 0.9rem; font-size: 0.88rem;
        margin-bottom: 0.6rem;
    }
    .field-flag *, .ok-banner *, .info-banner * { color: inherit !important; }

    /* ---- Cards (top-level bordered containers only) ---- */
    [data-testid="stVerticalBlockBorderWrapper"] {
        background: var(--card);
        border: 1px solid var(--line) !important;
        border-radius: 16px !important;
        box-shadow: 0 1px 2px rgba(16, 25, 46, 0.03), 0 8px 24px rgba(16, 25, 46, 0.035);
    }
    /* prevent a bordered container from ever nesting its own shadow/border again */
    [data-testid="stVerticalBlockBorderWrapper"] [data-testid="stVerticalBlockBorderWrapper"] {
        box-shadow: none;
    }

    /* ---- Tabs as a quiet segmented control ---- */
    .stTabs [data-baseweb="tab-list"] {
        gap: 0.3rem; background: rgba(16, 25, 46, 0.04); padding: 0.3rem;
        border-radius: 12px; border: 1px solid var(--line);
    }
    .stTabs [data-baseweb="tab"] {
        height: 2.5rem; border-radius: 9px; font-weight: 600; font-size: 0.9rem;
        color: var(--ink-soft) !important; background: transparent;
    }
    .stTabs [data-baseweb="tab"] p { color: inherit !important; }
    .stTabs [aria-selected="true"] {
        background: var(--card) !important;
        box-shadow: 0 1px 3px rgba(16, 25, 46, 0.08);
        padding: 10px;
    }
    .stTabs [aria-selected="true"] p { color: var(--navy) !important; }

    /* ---- Buttons ---- */
    .stButton > button, .stDownloadButton > button {
        border-radius: 10px; font-weight: 600; letter-spacing: 0.01em;
        border: 1px solid var(--line);
        color: var(--ink) !important;
    }
    .stButton > button p, .stDownloadButton > button p { color: inherit !important; }
    .stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] {
        background: linear-gradient(135deg, var(--navy), var(--navy-2));
        border: none;
        color: #FFFFFF !important;
    }
    .stButton > button[kind="primary"]:hover, .stDownloadButton > button[kind="primary"]:hover {
        filter: brightness(1.12);
        color: #FFFFFF !important;
    }
    .stButton > button[kind="secondary"]:hover, .stDownloadButton > button[kind="secondary"]:hover {
        border-color: var(--navy-2); color: var(--navy-2) !important;
    }

    .st-emotion-cache-ds661v, .st-emotion-cache-rzg2vg {   
    border: solid 1px grey;
    }

    /* ---- Metrics ---- */
    [data-testid="stMetric"] {
        background: rgba(16, 25, 46, 0.025); border: 1px solid var(--line);
        border-radius: 12px; padding: 0.7rem 0.9rem 0.5rem 0.9rem;
    }
    [data-testid="stMetricLabel"] { color: var(--ink-soft) !important; font-size: 0.78rem; }
    [data-testid="stMetricValue"] { color: var(--ink) !important; }

    /* ---- Rate comparison: Current / New offer panels ----
       Scoped via st.container(key=...), which Streamlit >=1.32 renders as
       a stable `st-key-<key>` class -- far more reliable than guessing at
       internal testids, and immune to nested-card doubling since these
       panels never use container(border=True) themselves. */
    div[class*="st-key-cmp-current-"] {
        border: 1px solid #E0A46B; border-radius: 12px; padding: 0.75rem 0.85rem 0.5rem 0.85rem;
        background: #FFFDFB;
    }
    div[class*="st-key-cmp-new-"] {
        border: 1px solid #8FC46B; border-radius: 12px; padding: 0.75rem 0.85rem 0.5rem 0.85rem;
        background: #FBFEFA;
    }
    .offer-header {
        text-align: center; font-weight: 700; font-size: 0.92rem;
        padding: 0.55rem 0.5rem; border-radius: 8px; margin-bottom: 0.5rem;
    }
    .offer-header.current { background: #F6C69A; color: #7A3D00 !important; }
    .offer-header.new { background: #B8E0A0; color: #1F5C0B !important; }
    .offer-header * { color: inherit !important; }
    .offer-gst-note {
        text-align: center; font-style: italic; font-size: 0.78rem;
        color: var(--ink-soft) !important; margin: -0.15rem 0 0.5rem 0;
    }
    .offer-total-bar {
        display: flex; justify-content: space-between; align-items: center;
        font-weight: 700; font-size: 0.92rem; padding: 0.55rem 0.9rem;
        border-radius: 8px; margin-top: 0.6rem;
    }
    .offer-total-bar.current { background: #F6C69A; color: #7A3D00 !important; }
    .offer-total-bar.new { background: #B8E0A0; color: #1F5C0B !important; }
    .offer-total-bar * { color: inherit !important; }
    .offer-row-controls { display: flex; gap: 0.5rem; margin-top: 0.6rem; }
    .offer-row-controls .stButton > button { width: 100%; font-size: 0.82rem; padding: 0.3rem 0.5rem; }

    /* ---- Misc ---- */
    div[data-testid="stStatusWidget"] { border-radius: 12px; }
    hr { border-color: var(--line) !important; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Config / secrets
# ---------------------------------------------------------------------------
API_KEYS = load_api_keys(st.secrets, os.environ)
GEMINI_API_KEY = API_KEYS[0] if API_KEYS else None
APP_PASSWORD = st.secrets.get("APP_PASSWORD")  # optional shared team password

if GEMINI_API_KEY:
    os.environ["GEMINI_API_KEY"] = GEMINI_API_KEY


def _require_password():
    if not APP_PASSWORD:
        return True
    if st.session_state.get("authed"):
        return True
    st.markdown('<div class="hero-badge">Team Access</div>', unsafe_allow_html=True)
    st.markdown('<div class="hero-title">Locked</div>', unsafe_allow_html=True)
    pw = st.text_input("Password", type="password")
    if st.button("Enter", type="primary"):
        if pw == APP_PASSWORD:
            st.session_state["authed"] = True
            st.rerun()
        else:
            st.error("Wrong password.")
    return False


if not _require_password():
    st.stop()

if not GEMINI_API_KEY:
    st.error(
        "API_KEY is not configured. Add it under this app"
    )
    st.stop()

# ---------------------------------------------------------------------------
# Session state
# ---------------------------------------------------------------------------
st.session_state.setdefault("single_extracted", None)
st.session_state.setdefault("single_result", None)
st.session_state.setdefault("consolidated_result", None)
st.session_state.setdefault("consolidated_extracted", None)
st.session_state.setdefault("consolidated_result", None)
st.session_state.setdefault("consolidated_client_name", "")

EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


# ---------------------------------------------------------------------------
# Filename helpers
# ---------------------------------------------------------------------------
def _sanitize_filename_part(s: str, max_len: int = 60) -> str:
    """Strip characters that break filenames and cap the length."""
    if not s:
        return "Unknown"
    s = re.sub(r'[\\/:*?"<>|]', "-", str(s))
    s = re.sub(r"\s+", " ", s).strip(" -")
    if len(s) > max_len:
        s = s[:max_len].rstrip() + "..."
    return s or "Unknown"


def _month_year() -> str:
    """Generation date, formatted like 'Jul 2026'. Bills don't carry a
    reliable issue-date field in the schema, so this uses today's date."""
    return datetime.now().strftime("%b %Y")


def build_single_filename(site_address: str) -> str:
    addr = _sanitize_filename_part(site_address or "Unknown Site", max_len=60)
    return f"Quote - {addr} - {_month_year()}.xlsx"


def build_consolidated_filename(client_name: str, bills: list) -> str:
    if not client_name:
        first = bills[0] if bills else {}
        client_name = first.get("customer_name") or first.get("oc_number") or "Multiple Sites"
    name = _sanitize_filename_part(client_name, max_len=40)
    return f"Consolidated - {name} - {_month_year()}.xlsx"


# ---------------------------------------------------------------------------
# Error helpers
# ---------------------------------------------------------------------------
def _friendly_error(e: Exception) -> str:
    msg = str(e)
    if "GEMINI_API_KEY" in msg:
        return "The API key isn't configured correctly."
    if "429" in msg or "RESOURCE_EXHAUSTED" in msg:
        return "Rate limit / Try again in a few minutes."
    if "503" in msg or "UNAVAILABLE" in msg:
        return "Servers were temporarily overloaded and retries didn't clear it. Try again shortly."
    if "PdfReadError" in msg or "pypdf" in msg.lower():
        return "The PDF couldn't be read \u2014 it may be corrupted, password-protected, or not a real PDF."
    return "Something went wrong while processing this bill."


def _show_model_errors(model_errors: dict):
    labels = {"primary": "Primary model (Flash)", "secondary": "Second-opinion model (Flash-Lite)"}
    any_shown = False
    for key, label in labels.items():
        err = model_errors.get(key)
        if err:
            any_shown = True
            st.markdown(f"**{label} failed:**")
            with st.expander("Show technical details", expanded=False):
                st.code(err)
    if not any_shown and model_errors.get("validation"):
        st.markdown("**The extracted data didn't match the expected format:**")
        with st.expander("Show technical details", expanded=False):
            st.code(model_errors["validation"])


# ---------------------------------------------------------------------------
# Tabular rendering of extracted bill data (instead of raw JSON)
# ---------------------------------------------------------------------------
def _money(v):
    return f"${v:,.2f}" if isinstance(v, (int, float)) else "\u2014"


def render_bill_tables(data: dict, heading: str | None = None):
    if heading:
        st.markdown(f'<div class="section-label">{heading}</div>', unsafe_allow_html=True)

    meta_rows = [
        ("Customer", data.get("customer_name") or data.get("oc_number") or "\u2014"),
        ("Site Address", data.get("site_address") or "\u2014"),
        ("NMI / MIRN", data.get("nmi_or_mirn") or "\u2014"),
        ("Distributor", data.get("distribution_region") or "\u2014"),
        ("Tariff", data.get("tariff_classification") or "\u2014"),
        ("Retailer", data.get("current_energy_retailer") or "\u2014"),
        ("State", data.get("state") or "\u2014"),
        ("Billing Period", f"{data['billing_period_days']} days" if data.get("billing_period_days") else "\u2014"),
        ("Total Due", _money(data.get("total_due"))),
    ]
    meta_df = pd.DataFrame(meta_rows, columns=["Field", "Value"])
    st.dataframe(meta_df, hide_index=True, width='stretch')

    charges = data.get("charges") or []
    if not charges:
        st.info("No charge line items were extracted.")
        return

    rows = []
    for c in charges:
        rows.append(
            {
                "Description": ("\u21A9 " if c.get("is_credit") else "") + (c.get("description") or ""),
                "Quantity": c.get("quantity"),
                "Unit": c.get("unit") or "",
                "Rate": c.get("rate_before_discount"),
                "Discount": c.get("conditional_discount_pct"),
            }
        )
    charges_df = pd.DataFrame(rows)
    st.dataframe(
        charges_df,
        hide_index=True,
        width='stretch',
        column_config={
            "Quantity": st.column_config.NumberColumn(format="%.2f"),
            "Rate": st.column_config.NumberColumn(format="$%.4f"),
            "Discount": st.column_config.NumberColumn(format="percent"),
        },
    )


# ---------------------------------------------------------------------------
# Rate Comparison Table -- side-by-side "Current Energy Offer" vs.
# "New Proposed Offer", modelled directly on the existing Excel layout:
#   Before Discount | Conditional Discount | After Discount | Total  (x2)
# ---------------------------------------------------------------------------
def _init_comparison_master(charges: list | None) -> pd.DataFrame:
    """One row per bill charge. Quantity + Description are kept as hidden
    bookkeeping columns (not displayed) so 'Total' can still be computed and
    'Auto-fill from tariff' can still match rows by description."""
    rows = []
    for c in charges or []:
        rows.append(
            {
                "Description": ("\u21A9 " if c.get("is_credit") else "") + (c.get("description") or ""),
                "Quantity": c.get("quantity") if c.get("quantity") is not None else 1,
                "Current Before": c.get("rate_before_discount"),
                "Current Cond %": c.get("conditional_discount_pct"),
                "New Before": c.get("rate_before_discount"),
                "New Cond %": c.get("conditional_discount_pct"),
            }
        )
    if not rows:
        rows.append(
            {
                "Description": "", "Quantity": 1,
                "Current Before": None, "Current Cond %": None,
                "New Before": None, "New Cond %": None,
            }
        )
    return pd.DataFrame(rows)


def _reconcile_master_length(master: pd.DataFrame, new_len: int) -> pd.DataFrame:
    """Keeps the hidden master table in sync when the user adds/removes rows
    in the New Proposed Offer editor (num_rows='dynamic'). Rows added beyond
    the original bill charges have no 'Current' counterpart -- exactly right
    for a brand-new charge that only exists in the new offer -- and default
    to quantity 1 (a flat new rate) since there's no bill quantity for them."""
    cur_len = len(master)
    if new_len == cur_len:
        return master
    if new_len > cur_len:
        extra = pd.DataFrame(
            {
                "Description": [""] * (new_len - cur_len),
                "Quantity": [1] * (new_len - cur_len),
                "Current Before": [None] * (new_len - cur_len),
                "Current Cond %": [None] * (new_len - cur_len),
                "New Before": [None] * (new_len - cur_len),
                "New Cond %": [None] * (new_len - cur_len),
            }
        )
        return pd.concat([master, extra], ignore_index=True)
    return master.iloc[:new_len].reset_index(drop=True)


def comparison_df_to_records(master: pd.DataFrame | None) -> list[dict]:
    """Template-ready row shape for the future New Offer section."""
    if master is None or master.empty:
        return []
    return master.to_dict("records")


def _auto_fill_new_rates(master: pd.DataFrame, tariff_code: str | None) -> tuple[pd.DataFrame, bool]:
    """Best-effort auto-fill of 'New Before' from a (future) tariff_bridge
    rate card. Looks for an optional `get_rate_card(tariff_code)` function on
    the tariff_bridge module so this starts working the moment that lands
    there, with no changes needed in app.py. Returns (master, filled_any)."""
    lookup_fn = getattr(tariff_bridge_module, "get_rate_card", None)
    if not lookup_fn or not tariff_code:
        return master, False
    try:
        rate_card = lookup_fn(tariff_code) or {}
    except Exception:
        return master, False
    if not rate_card:
        return master, False

    updated = master.copy()
    filled_any = False
    for idx, row in updated.iterrows():
        match = rate_card.get(row.get("Description"))
        if match is not None:
            updated.at[idx, "New Before"] = match
            filled_any = True
    return updated, filled_any


def _after_discount(before, cond_pct):
    if before is None or (isinstance(before, float) and pd.isna(before)):
        return None
    cond = cond_pct if cond_pct is not None and not (isinstance(cond_pct, float) and pd.isna(cond_pct)) else 0
    return before * (1 - cond / 100)


def render_comparison_table(charges: list | None, state_key: str, tariff_code: str | None = None) -> pd.DataFrame:
    """Renders the Current Energy Offer / New Proposed Offer side-by-side
    tables for one site and returns the underlying master DataFrame
    (Description, Quantity, Current Before/Cond %, New Before/Cond %).
    Persisted in st.session_state under `state_key` so edits survive
    reruns (tariff lookups, tab switches, etc.)."""
    st.markdown('<div class="section-label">Rate Comparison \u2014 New Offer</div>', unsafe_allow_html=True)
    st.caption(
        "Current Energy Offer is read from the bill. Edit New Proposed Offer rates on the right, or add a "
        "row for a brand-new charge that isn't on the current bill at all."
    )

    if state_key not in st.session_state:
        st.session_state[state_key] = _init_comparison_master(charges)
    master = st.session_state[state_key]

    _, auto_col = st.columns([3, 1.2])
    with auto_col:
        if st.button("\u2728 Auto-fill from tariff", key=f"autofill_{state_key}"):
            master, filled = _auto_fill_new_rates(master, tariff_code)
            st.session_state[state_key] = master
            if filled:
                st.toast("Auto-filled new rates from the tariff rate card.")
            else:
                st.markdown(
                    '<div class="info-banner">\u2139\uFE0F Automatic rate lookup isn\'t available for this '
                    "tariff yet \u2014 enter the new rates manually below.</div>",
                    unsafe_allow_html=True,
                )

    left, right = st.columns(2, gap="small")

    # ---- Current Energy Offer (read-only, from the bill) ----
    with left:
        with st.container(border=True):
            st.markdown('<div class="offer-header current">Current Energy Offer</div>', unsafe_allow_html=True)
            st.markdown('<div class="offer-gst-note">Rates are Inclusive of GST</div>', unsafe_allow_html=True)
            current_view = pd.DataFrame(
                {
                    "Before Discount": master["Current Before"],
                    "Conditional Discount": master["Current Cond %"],
                }
            )
            current_view["After Discount"] = [
                _after_discount(b, c) for b, c in zip(current_view["Before Discount"], current_view["Conditional Discount"])
            ]
            current_view["Total"] = [
                (a * q if a is not None and pd.notna(q) else None)
                for a, q in zip(current_view["After Discount"], master["Quantity"])
            ]
            st.dataframe(
                current_view,
                hide_index=True,
                width='stretch',
                column_config={
                    "Before Discount": st.column_config.NumberColumn(format="$%.4f"),
                    "Conditional Discount": st.column_config.NumberColumn(format="%.1f%%"),
                    "After Discount": st.column_config.NumberColumn(format="$%.4f"),
                    "Total": st.column_config.NumberColumn(format="$%.2f"),
                },
            )
            current_total = current_view["Total"].dropna().sum() if current_view["Total"].notna().any() else 0.0
            st.markdown(
                f'<div class="offer-total-bar current"><span>Total (GST Incl.)</span><span>{_money(current_total)}</span></div>',
                unsafe_allow_html=True,
            )

    # ---- New Proposed Offer (editable) ----
    with right:
        with st.container(border=True):
            st.markdown('<div class="offer-header new">New Proposed Offer</div>', unsafe_allow_html=True)
            st.markdown('<div class="offer-gst-note">Rates are Inclusive of GST</div>', unsafe_allow_html=True)
            new_view = pd.DataFrame(
                {
                    "Before Discount": master["New Before"],
                    "Conditional Discount": master["New Cond %"],
                }
            )
            new_view["After Discount"] = [
                _after_discount(b, c) for b, c in zip(new_view["Before Discount"], new_view["Conditional Discount"])
            ]
            new_view["Total"] = [
                (a * q if a is not None and pd.notna(q) else None)
                for a, q in zip(new_view["After Discount"], master["Quantity"])
            ]
            edited_new = st.data_editor(
                new_view,
                key=f"{state_key}_new_editor",
                num_rows="dynamic",
                hide_index=True,
                width='stretch',
                column_config={
                    "Before Discount": st.column_config.NumberColumn(format="$%.4f"),
                    "Conditional Discount": st.column_config.NumberColumn(format="%.1f%%"),
                    "After Discount": st.column_config.NumberColumn(format="$%.4f", disabled=True),
                    "Total": st.column_config.NumberColumn(format="$%.2f", disabled=True),
                },
            )
            master = _reconcile_master_length(master, len(edited_new))
            master["New Before"] = edited_new["Before Discount"].values
            master["New Cond %"] = edited_new["Conditional Discount"].values
            st.session_state[state_key] = master

            new_total = edited_new["Total"].dropna().sum() if edited_new["Total"].notna().any() else 0.0
            st.markdown(
                f'<div class="offer-total-bar new"><span>Total (GST Incl.)</span><span>{_money(new_total)}</span></div>',
                unsafe_allow_html=True,
            )

    if current_total or new_total:
        delta = new_total - current_total
        pct = (delta / current_total * 100) if current_total else 0.0
        sign = "\U0001F53B" if delta > 0 else ("\U0001F53C" if delta < 0 else "\u2796")
        st.caption(f"{sign} Estimated change: **{_money(delta)}** ({pct:+.1f}%) vs. the current bill.")

    return master


def _clear_comparison_state(prefix: str):
    for k in list(st.session_state.keys()):
        if k.startswith(prefix):
            del st.session_state[k]


# ---------------------------------------------------------------------------
# Single-site pipeline
# ---------------------------------------------------------------------------
def _clean_nmi(data: dict) -> str | None:
    nmi = data.get("nmi_or_mirn")
    nmi = str(nmi).strip() if nmi else ""
    return nmi or None


def extract_single(pdf_bytes: bytes, filename: str):
    """Phase 1: Gemini extraction only. Splitting this out from Excel
    generation is what makes the NMI tariff lookup possible - the NMI has
    to be known (and confirmable/overridable) before the quote is built."""
    with tempfile.TemporaryDirectory() as tmp:
        pdf_path = os.path.join(tmp, filename)
        with open(pdf_path, "wb") as f:
            f.write(pdf_bytes)

        status = st.status("Reading bill\u2026", expanded=True)
        try:
            status.write("\U0001F4E4 Analyzing bill for data extraction (this can take up to a minute)\u2026")
            result = extract_bill(pdf_path, api_keys=API_KEYS)

            if not result["success"]:
                status.update(label="Extraction failed", state="error")
                st.error("Couldn't extract data from this bill.")
                _show_model_errors(result["model_errors"])
                return None

            status.update(
                label="\u2705 Extraction complete \u2014 cross-checked the two models' answers",
                state="complete",
                expanded=False,
            )
            return {
                "data": result["data"],
                "review_fields": result["review_fields"],
                "degraded_mode": result.get("degraded_mode", False),
                "source_filename": filename,
            }
        except Exception as e:  # noqa: BLE001
            status.update(label="Something went wrong", state="error")
            st.error(_friendly_error(e))
            with st.expander("Show technical details"):
                st.exception(e)
            return None


def generate_single_excel(extracted: dict, tariff_override: str | None, comparison_df: pd.DataFrame | None = None):
    """Phase 2: build the Excel quote from already-extracted data, using the
    live-looked-up NTC (if any) instead of whatever tariff Gemini read off
    the bill text.

    `comparison_df` (the edited rate-comparison table) is captured onto the
    result as `rate_comparison` so it travels with the quote. It is not yet
    passed into fill_quote() -- see the module docstring for the wiring
    plan once the template's New Offer section is ready."""
    with tempfile.TemporaryDirectory() as tmp:
        out_path = os.path.join(tmp, "generated_quote.xlsx")
        tariff_debug: dict = {}
        try:
            with st.spinner("\U0001F4CA Building the Excel quote\u2026"):
                fill_quote(
                    extracted["data"],
                    out_path,
                    tariff_debug=tariff_debug,
                    tariff_override=tariff_override,
                )
                with open(out_path, "rb") as f:
                    excel_bytes = f.read()
        except Exception as e:  # noqa: BLE001
            st.error(_friendly_error(e))
            with st.expander("Show technical details"):
                st.exception(e)
            return None

        return {
            "success": True,
            "excel_bytes": excel_bytes,
            "filename": build_single_filename(extracted["data"].get("site_address")),
            "source_filename": extracted["source_filename"],
            "data": extracted["data"],
            "review_fields": extracted["review_fields"],
            "degraded_mode": extracted["degraded_mode"],
            "tariff_debug": tariff_debug,
            "rate_comparison": comparison_df_to_records(comparison_df),
            "generated_at": datetime.now().strftime("%H:%M:%S"),
        }


def reset_single():
    extracted = st.session_state.get("single_extracted")
    result = st.session_state.get("single_result")
    nmi = None
    if extracted:
        nmi = _clean_nmi(extracted["data"])
    elif result:
        nmi = _clean_nmi(result["data"])
    clear_cached_tariff(nmi)
    _clear_comparison_state("single_comparison_df")
    st.session_state.update(single_extracted=None, single_result=None)


def render_single_result(res: dict):
    if res.get("degraded_mode"):
        st.warning(
            "\u26A0\uFE0F Only one model returned a result for this bill \u2014 it wasn't cross-checked "
            "against a second opinion. Please review the figures carefully before sending this quote out."
        )

    if res["review_fields"]:
        st.markdown(f"**\u26A0\uFE0F {len(res['review_fields'])} field(s) need a manual double-check:**")
        for f in res["review_fields"]:
            st.markdown(f'<div class="field-flag">{f}</div>', unsafe_allow_html=True)
    else:
        st.markdown(
            '<div class="ok-banner">\u2705 Both models agreed on every field \u2014 high confidence extraction.</div>',
            unsafe_allow_html=True,
        )

    render_bill_tables(res["data"], heading="Extracted Bill Data")

    tariff_debug = res.get("tariff_debug") or {}
    if tariff_debug.get("source") == "browser_popup_override":
        st.caption(f"\u26A1 Tariff on this quote (B16): **{tariff_debug.get('tariff')}** \u2014 live lookup from the portal.")
    else:
        st.caption("\u2139\uFE0F Tariff on this quote (B16) came from the bill text \u2014 no live lookup was used, so double-check it.")

    if res.get("rate_comparison"):
        st.markdown('<div class="section-label">Rate Comparison Saved</div>', unsafe_allow_html=True)
        st.caption(
            f"{len(res['rate_comparison'])} line item(s) saved with this quote for the New Offer section "
            "(not yet written into the Excel file \u2014 coming soon)."
        )

    st.markdown('<div class="section-label">Download</div>', unsafe_allow_html=True)
    st.download_button(
        "\u2B07\uFE0F Download Excel Quote",
        data=res["excel_bytes"],
        file_name=res["filename"],
        mime=EXCEL_MIME,
        type="primary",
        key="download_single",
    )
    st.markdown(
        f'<span class="meta-caption">File: {res["filename"]} \u00B7 generated at {res["generated_at"]} '
        f'from {res["source_filename"]}. Download it to add New Offer.</span>',
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------------------
# Consolidated pipeline
# ---------------------------------------------------------------------------
def extract_consolidated(pdf_bytes_list, filenames):
    """Phase 1: Extract data from all PDFs, but hold off on Excel generation."""
    with tempfile.TemporaryDirectory() as tmp:
        pdf_paths = []
        for data, name in zip(pdf_bytes_list, filenames):
            p = os.path.join(tmp, name)
            with open(p, "wb") as f:
                f.write(data)
            pdf_paths.append(p)

        status = st.status("Processing bill(s)\u2026", expanded=True)
        try:
            if len(pdf_paths) > 1:
                status.write(f"\U0001F4CE Merging {len(pdf_paths)} PDFs into one file\u2026")
                merged_path = os.path.join(tmp, "merged.pdf")
                writer = PdfWriter()
                for p in pdf_paths:
                    writer.append(p)
                writer.write(merged_path)
            else:
                merged_path = pdf_paths[0]

            status.write("\U0001F4E4 Sending to analyzer for multi-site extraction (can take a few minutes)\u2026")
            result = extract_consolidated_bills(merged_path, api_keys=API_KEYS)

            if not result["success"]:
                status.update(label="Extraction failed", state="error")
                st.error("Couldn't extract data from these bill(s).")
                _show_model_errors(result["model_errors"])
                return None

            bills = result["data"]["bills"]
            if not bills:
                status.update(label="No sites found", state="error")
                st.error("No sites were found in the uploaded PDF(s). Double-check the file(s) and try again.")
                return None

            status.update(label=f"\u2705 Extracted {len(bills)} site(s)", state="complete", expanded=False)
            return {
                "bills": bills,
                "review_notes": result.get("review_notes", []),
                "n_sites": len(bills),
            }
        except Exception as e:
            status.update(label="Something went wrong", state="error")
            st.error(_friendly_error(e))
            with st.expander("Show technical details"):
                st.exception(e)
            return None


def generate_consolidated_excel(extracted_data, client_name, comparison_dfs: dict | None = None):
    """Phase 2: Build the Excel quote using cached NTC lookups.

    `comparison_dfs` (per-site edited rate tables, keyed by site index) are
    captured onto the result as `rate_comparisons` -- see the module
    docstring for the plan to wire these into fill_consolidated_quote()."""
    bills = extracted_data["bills"]
    comparison_dfs = comparison_dfs or {}
    with tempfile.TemporaryDirectory() as tmp:
        out_path = os.path.join(tmp, "consolidated_quote.xlsx")

        # Gather any live tariffs fetched by the user
        tariff_overrides = {}
        for b in bills:
            nmi = _clean_nmi(b)
            if nmi:
                cached = get_cached_tariff(nmi)
                if cached:
                    tariff_overrides[nmi] = cached

        try:
            fill_consolidated_quote(
                bills,
                out_path,
                client_name=client_name or None,
                tariff_overrides=tariff_overrides
            )
            with open(out_path, "rb") as f:
                excel_bytes = f.read()
        except Exception as e:
            st.error(_friendly_error(e))
            with st.expander("Show technical details"):
                st.exception(e)
            return None

        rate_comparisons = {
            str(i): comparison_df_to_records(df) for i, df in comparison_dfs.items()
        }

        return {
            "success": True,
            "excel_bytes": excel_bytes,
            "filename": build_consolidated_filename(client_name, bills),
            "bills": bills,
            "review_notes": extracted_data.get("review_notes", []),
            "rate_comparisons": rate_comparisons,
            "generated_at": datetime.now().strftime("%H:%M:%S"),
            "n_sites": len(bills),
        }


def reset_consolidated():
    extracted = st.session_state.get("consolidated_extracted")
    result = st.session_state.get("consolidated_result")

    # Clear out cached tariffs for all NMIs involved in this batch
    if extracted:
        for b in extracted["bills"]:
            clear_cached_tariff(_clean_nmi(b))
    elif result:
        for b in result["bills"]:
            clear_cached_tariff(_clean_nmi(b))

    _clear_comparison_state("cons_comparison_df")
    st.session_state.update(
        consolidated_extracted=None,
        consolidated_result=None,
        consolidated_client_name=""
    )


def render_consolidated_result(res: dict):
    st.markdown(
        f'<div class="ok-banner">\u2705 Found {res["n_sites"]} site(s) and built the consolidated quote.</div>',
        unsafe_allow_html=True,
    )

    if res["review_notes"]:
        st.markdown(f"**\u26A0\uFE0F {len(res['review_notes'])} discrepancy(ies) between models \u2014 please review:**")
        for note in res["review_notes"]:
            st.markdown(f'<div class="field-flag">{note}</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-label">Sites Overview</div>', unsafe_allow_html=True)
    overview_df = pd.DataFrame(
        [
            {
                "Site Address": b.get("site_address") or "\u2014",
                "Customer": b.get("customer_name") or b.get("oc_number") or "\u2014",
                "Total Due": b.get("total_due"),
            }
            for b in res["bills"]
        ]
    )
    st.dataframe(
        overview_df,
        hide_index=True,
        width='stretch',
        column_config={"Total Due": st.column_config.NumberColumn(format="dollar")},
    )

    st.markdown('<div class="section-label">Per-Site Detail</div>', unsafe_allow_html=True)
    rate_comparisons = res.get("rate_comparisons") or {}
    for i, bill in enumerate(res["bills"], start=1):
        label = bill.get("site_address") or bill.get("customer_name") or f"Site {i}"
        with st.expander(f"{i}. {label}"):
            render_bill_tables(bill)
            saved_rows = rate_comparisons.get(str(i))
            if saved_rows:
                st.markdown('<div class="section-label">Rate Comparison Saved</div>', unsafe_allow_html=True)
                st.caption(f"{len(saved_rows)} line item(s) saved for the New Offer section (not yet written into the Excel file).")

    st.markdown('<div class="section-label">Download</div>', unsafe_allow_html=True)
    st.download_button(
        "\u2B07\uFE0F Download Consolidated Excel Quote",
        data=res["excel_bytes"],
        file_name=res["filename"],
        mime=EXCEL_MIME,
        type="primary",
        key="download_consolidated",
    )
    st.markdown(
        f'<span class="meta-caption">File: {res["filename"]} \u00B7 generated at {res["generated_at"]}. '
        f"Download it and fill new offer</span>",
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------
_LOGO_PATH = os.path.join(os.path.dirname(__file__), "assets", "logo.png")
header_l, header_r = st.columns([1, 6])
with header_l:
    if os.path.exists(_LOGO_PATH):
        st.image(_LOGO_PATH, width=64)
with header_r:
    st.markdown('<div class="hero-badge">\u26A1 Quoting Tool</div>', unsafe_allow_html=True)
    st.markdown('<div class="hero-title">Bills \u2192 Quote</div>', unsafe_allow_html=True)

st.markdown(
    '<div class="hero-sub"><ul>'
    "<li>Upload an electricity/gas bill PDF and get a quote back to fill New Offer.</li>"
    '<li>For a consolidated quote, switch to \u201cConsolidated / Multi-site\u201d.</li>'
    "<li>Always double-check the extracted figures and new rates before sending a quote out.</li>"
    "</ul></div>",
    unsafe_allow_html=True,
)

tab_single, tab_consolidated = st.tabs(["\U0001F4C4 Single Site", "\U0001F3E2 Consolidated / Multi-site"])

with tab_single:
    if st.session_state.single_result is not None:
        with st.container(border=True):
            render_single_result(st.session_state.single_result)
        st.button("\U0001F504 Process another bill", key="reset_single", on_click=reset_single)

    elif st.session_state.single_extracted is not None:
        extracted = st.session_state.single_extracted
        nmi = _clean_nmi(extracted["data"])

        with st.container(border=True):
            if extracted.get("degraded_mode"):
                st.warning(
                    "\u26A0\uFE0F Only one model returned a result for this bill \u2014 it wasn't cross-checked "
                    "against a second opinion. Please review the figures carefully."
                )
            if extracted["review_fields"]:
                st.markdown(f"**\u26A0\uFE0F {len(extracted['review_fields'])} field(s) need a manual double-check:**")
                for f in extracted["review_fields"]:
                    st.markdown(f'<div class="field-flag">{f}</div>', unsafe_allow_html=True)
            else:
                st.markdown(
                    '<div class="ok-banner">\u2705 Both models agreed on every field \u2014 high confidence extraction.</div>',
                    unsafe_allow_html=True,
                )

            render_bill_tables(extracted["data"], heading="Extracted Bill Data")

            st.markdown('<div class="section-label">Network Tariff (NTC) Lookup</div>', unsafe_allow_html=True)
            if not nmi:
                st.caption(
                    "No NMI/MIRN was extracted from this bill, so a live tariff lookup isn't possible \u2014 "
                    "the quote will use the tariff printed on the bill."
                )
            else:
                cached_tariff = get_cached_tariff(nmi)
                if cached_tariff:
                    st.markdown(
                        f'<div class="ok-banner">\u2705 Live NTC fetched for NMI {nmi}: <b>{cached_tariff}</b> '
                        f'\u2014 this will be used on the quote instead of the bill\'s printed tariff.</div>',
                        unsafe_allow_html=True,
                    )
                    st.button(
                        "\u21BA Re-fetch tariff",
                        key="refetch_tariff_single",
                        on_click=lambda: (clear_cached_tariff(nmi), st.rerun()),
                    )
                else:
                    st.caption(
                        f"Bill shows tariff **{extracted['data'].get('tariff_classification') or '\u2014'}** for NMI **{nmi}**. "
                        "Click below to fetch the live network tariff from the sales portal (opens in a new tab \u2014 "
                        "requires the GloBird NTC Fetcher Tampermonkey script and an active portal login)."
                    )
                    render_auto_lookup(nmi, key="single")

            comparison_df = render_comparison_table(
                extracted["data"].get("charges"),
                state_key="single_comparison_df",
                tariff_code=get_cached_tariff(nmi) or extracted["data"].get("tariff_classification"),
            )

            st.markdown('<div class="section-label">Generate Quote</div>', unsafe_allow_html=True)
            gcol1, gcol2 = st.columns([1, 1])
            with gcol1:
                if st.button("\U0001F4CA Generate Excel Quote", type="primary", key="run_generate_single"):
                    tariff_override = get_cached_tariff(nmi)
                    res = generate_single_excel(extracted, tariff_override, comparison_df)
                    if res:
                        st.session_state.single_result = res
                        st.rerun()
            with gcol2:
                st.button("\U0001F504 Start over", key="reset_single_extracted", on_click=reset_single)

    else:
        with st.container(border=True):
            uploaded = st.file_uploader("Upload one bill PDF", type=["pdf"], key="single_uploader")
            run = st.button("Extract Bill Data", type="primary", disabled=uploaded is None, key="run_single")
        if run and uploaded:
            extracted = extract_single(uploaded.getvalue(), uploaded.name)
            if extracted:
                st.session_state.single_extracted = extracted
                st.rerun()

with tab_consolidated:
    if st.session_state.consolidated_result is not None:
        with st.container(border=True):
            render_consolidated_result(st.session_state.consolidated_result)
        st.button("\U0001F504 Process another batch", key="reset_consolidated", on_click=reset_consolidated)

    elif st.session_state.consolidated_extracted is not None:
        extracted = st.session_state.consolidated_extracted
        client_name = st.session_state.consolidated_client_name

        # Calculate exactly which NMIs still need lookup
        all_nmis = []
        for b in extracted["bills"]:
            nmi_val = _clean_nmi(b)
            if nmi_val:
                all_nmis.append(nmi_val)

        # Deduplicate and find missing ones
        unique_nmis = list(dict.fromkeys(all_nmis))
        missing_nmis = [n for n in unique_nmis if not get_cached_tariff(n)]

        with st.container(border=True):
            st.markdown(f'<div class="ok-banner">\u2705 Found {extracted["n_sites"]} site(s). Review data, fetch live tariffs, and set new offer rates below before generating the quote.</div>', unsafe_allow_html=True)

            # --- Bulk tariff lookup ---
            st.markdown('<div class="section-label">Network Tariff Lookup (Bulk)</div>', unsafe_allow_html=True)
            if missing_nmis:
                st.caption(f"There are **{len(missing_nmis)}** site(s) missing live tariffs. Click below to fetch them all at once.")
                render_bulk_lookup_button(missing_nmis, key="run_bulk_lookup")
            elif unique_nmis:
                st.markdown('<div class="ok-banner">\u2705 All live tariffs successfully fetched for this portfolio!</div>', unsafe_allow_html=True)
            else:
                st.caption("No NMIs were found in this portfolio to look up.")

            st.markdown('<div class="section-label">Sites, Tariffs \u0026 Rate Comparison</div>', unsafe_allow_html=True)

            comparison_dfs = {}
            for i, bill in enumerate(extracted["bills"], start=1):
                label = bill.get("site_address") or bill.get("customer_name") or f"Site {i}"
                nmi = _clean_nmi(bill)

                with st.expander(f"{i}. {label} (NMI: {nmi or 'None'})", expanded=False):
                    render_bill_tables(bill)

                    if not nmi:
                        st.caption("No NMI extracted \u2014 automated lookup suspended for this site.")
                    else:
                        cached = get_cached_tariff(nmi)
                        if cached:
                            st.markdown(f'<div class="ok-banner">\u2705 Live NTC fetched: <b>{cached}</b></div>', unsafe_allow_html=True)
                            st.button(
                                "\u21BA Reset this tariff",
                                key=f"refetch_cons_{nmi}_{i}",
                                on_click=lambda target=nmi: (clear_cached_tariff(target), st.rerun()),
                            )
                        else:
                            st.caption(f"Bill shows tariff **{bill.get('tariff_classification') or '\u2014'}**. Awaiting bulk lookup above.")

                    comparison_dfs[i] = render_comparison_table(
                        bill.get("charges"),
                        state_key=f"cons_comparison_df_{i}",
                        tariff_code=get_cached_tariff(nmi) if nmi else bill.get("tariff_classification"),
                    )

            st.markdown('<div class="section-label">Generate Quote</div>', unsafe_allow_html=True)
            gcol1, gcol2 = st.columns([1, 1])
            with gcol1:
                if st.button("\U0001F4CA Generate Consolidated Quote", type="primary", key="run_gen_consolidated"):
                    with st.spinner("Building the Excel quote\u2026"):
                        res = generate_consolidated_excel(extracted, client_name, comparison_dfs)
                        if res:
                            st.session_state.consolidated_result = res
                            st.rerun()
            with gcol2:
                st.button("\U0001F504 Start over", key="reset_cons_extracted", on_click=reset_consolidated)

    else:
        with st.container(border=True):
            uploaded_files = st.file_uploader(
                "Upload one or more bill PDFs (multi-site portfolio, or several separate bills)",
                type=["pdf"],
                accept_multiple_files=True,
                key="consolidated_uploader",
            )
            client_name = st.text_input(
                "Client name to show as 'Quote Prepared For' (optional, e.g. the Strata management company)",
                key="client_name_input",
            )
            run = st.button(
                "Extract Bill Data", type="primary", disabled=not uploaded_files, key="run_consolidated"
            )

        if run and uploaded_files:
            data_list = [f.getvalue() for f in uploaded_files]
            names = [f.name for f in uploaded_files]
            extracted = extract_consolidated(data_list, names)
            if extracted:
                st.session_state.consolidated_extracted = extracted
                st.session_state.consolidated_client_name = client_name
                st.rerun()

st.divider()
st.markdown('<span class="meta-caption">Internal tool \u2014 OZ Admin Team.</span>', unsafe_allow_html=True)
