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
"""
import os
import re
import tempfile
from datetime import datetime

import pandas as pd
import streamlit as st
from pypdf import PdfWriter

from single_gemini_client import extract_bill
from single_excel_filler import fill_quote
from consolidated_gemini_client import extract_consolidated_bills
from consolidated_excel_filler import fill_consolidated_quote
from api_key_pool import load_api_keys

st.set_page_config(page_title="Bills \u2192 Quote", page_icon="\u26A1", layout="centered")

# ---------------------------------------------------------------------------
# Look & feel
# ---------------------------------------------------------------------------
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&display=swap');
 
    html, body, [class*="css"] { font-family: 'Space Grotesk', sans-serif; }
    .block-container { padding-top: 3.2rem; max-width: 820px; }
 
    .hero-badge {
        display: inline-block; padding: 0.2rem 0.75rem; border-radius: 999px;
        background: rgba(236, 27, 64, 0.08); border: 1px solid rgba(236, 27, 64, 0.35);
        color: #C4123A; font-size: 0.72rem; font-weight: 600; letter-spacing: 0.06em;
        text-transform: uppercase; margin-bottom: 0.7rem;
    }
    .hero-title {
        font-size: 2.3rem; font-weight: 700; line-height: 1.15; margin-bottom: 0.15rem;
        background: linear-gradient(90deg, #1D3A66, #25477C 45%, #EC1B40);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
    }
    .hero-sub { color: #55607A; font-size: 1rem; margin-bottom: 1.6rem; }
 
    .field-flag {
        background: rgba(255, 193, 7, 0.12); border: 1px solid rgba(255, 193, 7, 0.45);
        color: #8A6100; border-radius: 8px; padding: 0.5rem 0.9rem; margin-bottom: 0.4rem;
        font-size: 0.9rem;
    }
    .ok-banner {
        background: rgba(46, 204, 113, 0.12); border: 1px solid rgba(46, 204, 113, 0.45);
        color: #1C7A43; border-radius: 8px; padding: 0.65rem 1rem; font-size: 0.92rem;
        margin-bottom: 0.7rem;
    }
    .section-label {
        font-size: 0.78rem; font-weight: 700; letter-spacing: 0.08em; text-transform: uppercase;
        color: #3A5686; margin: 1.1rem 0 0.4rem 0;
    }
    .meta-caption { color: #6B7690; font-size: 0.82rem; }
 
    .stButton > button, .stDownloadButton > button {
        border-radius: 10px; font-weight: 600; letter-spacing: 0.01em;
    }
    div[data-testid="stStatusWidget"] { border-radius: 12px; }
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
st.session_state.setdefault("single_result", None)
st.session_state.setdefault("consolidated_result", None)

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
    st.dataframe(meta_df, hide_index=True, use_container_width=True)

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
        use_container_width=True,
        column_config={
            "Quantity": st.column_config.NumberColumn(format="%.2f"),
            "Rate": st.column_config.NumberColumn(format="$%.4f"),
            "Discount": st.column_config.NumberColumn(format="percent"),
        },
    )


# ---------------------------------------------------------------------------
# Single-site pipeline
# ---------------------------------------------------------------------------
def process_single(pdf_bytes: bytes, filename: str):
    with tempfile.TemporaryDirectory() as tmp:
        pdf_path = os.path.join(tmp, filename)
        with open(pdf_path, "wb") as f:
            f.write(pdf_bytes)

        status = st.status("Processing bill\u2026", expanded=True)
        try:
            status.write("\U0001F4E4 Analyzing bill for data extraction (this can take up to a minute)\u2026")
            result = extract_bill(pdf_path, api_keys=API_KEYS)

            if not result["success"]:
                status.update(label="Extraction failed", state="error")
                st.error("Couldn't extract data from this bill.")
                _show_model_errors(result["model_errors"])
                return None

            status.write("\u2705 Extraction complete \u2014 cross-checking the two models' answers\u2026")
            status.write("\U0001F4CA Building the Excel quote\u2026")
            out_path = os.path.join(tmp, "generated_quote.xlsx")
            fill_quote(result["data"], out_path)
            with open(out_path, "rb") as f:
                excel_bytes = f.read()

            status.update(label="Done", state="complete", expanded=False)
            return {
                "success": True,
                "excel_bytes": excel_bytes,
                "filename": build_single_filename(result["data"].get("site_address")),
                "source_filename": filename,
                "data": result["data"],
                "review_fields": result["review_fields"],
                "degraded_mode": result.get("degraded_mode", False),
                "generated_at": datetime.now().strftime("%H:%M:%S"),
            }
        except Exception as e:  # noqa: BLE001
            status.update(label="Something went wrong", state="error")
            st.error(_friendly_error(e))
            with st.expander("Show technical details"):
                st.exception(e)
            return None


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
def process_consolidated(pdf_bytes_list, filenames, client_name):
    with tempfile.TemporaryDirectory() as tmp:
        pdf_paths = []
        for data, name in zip(pdf_bytes_list, filenames):
            p = os.path.join(tmp, name)
            with open(p, "wb") as f:
                f.write(data)
            pdf_paths.append(p)

        status = st.status("Processing bill(s)…", expanded=True)
        try:
            if len(pdf_paths) > 1:
                status.write(f"📎 Merging {len(pdf_paths)} PDFs into one file…")
                merged_path = os.path.join(tmp, "merged.pdf")
                writer = PdfWriter()
                for p in pdf_paths:
                    writer.append(p)
                writer.write(merged_path)
            else:
                merged_path = pdf_paths[0]

            status.write("📤 Sending to analyzer for multi-site extraction (can take a few minutes)…")
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

            status.write(f"\u2705 Found {len(bills)} site(s) \u2014 building the consolidated Excel quote\u2026")
            out_path = os.path.join(tmp, "consolidated_quote.xlsx")
            fill_consolidated_quote(bills, out_path, client_name=client_name or None)
            with open(out_path, "rb") as f:
                excel_bytes = f.read()

            status.update(label="Done", state="complete", expanded=False)
            return {
                "success": True,
                "excel_bytes": excel_bytes,
                "filename": build_consolidated_filename(client_name, bills),
                "bills": bills,
                "review_notes": result.get("review_notes", []),
                "generated_at": datetime.now().strftime("%H:%M:%S"),
                "n_sites": len(bills),
            }
        except Exception as e:  # noqa: BLE001
            status.update(label="Something went wrong", state="error")
            st.error(_friendly_error(e))
            with st.expander("Show technical details"):
                st.exception(e)
            return None


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
        use_container_width=True,
        column_config={"Total Due": st.column_config.NumberColumn(format="dollar")},
    )

    st.markdown('<div class="section-label">Per-Site Detail</div>', unsafe_allow_html=True)
    for i, bill in enumerate(res["bills"], start=1):
        label = bill.get("site_address") or bill.get("customer_name") or f"Site {i}"
        with st.expander(f"{i}. {label}"):
            render_bill_tables(bill)

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
if os.path.exists(_LOGO_PATH):
    st.image(_LOGO_PATH, width=260)
st.markdown('<div class="hero-badge">\u26A1 Quoting Tool</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="hero-sub"><li>Upload an electricity/gas bill PDF and get a quote back to fill New Offer.</li><li> For Consolidated Quote swith to "Consolidated / Multi-Site".</li><li>Must check it thoroughly to insure no error.</li></div>',
    unsafe_allow_html=True,
)

tab_single, tab_consolidated = st.tabs(["\U0001F4C4 Single Site", "\U0001F3E2 Consolidated / Multi-site"])

with tab_single:
    if st.session_state.single_result is None:
        with st.container(border=True):
            uploaded = st.file_uploader("Upload one bill PDF", type=["pdf"], key="single_uploader")
            run = st.button("Generate Quote", type="primary", disabled=uploaded is None, key="run_single")
        if run and uploaded:
            res = process_single(uploaded.getvalue(), uploaded.name)
            if res:
                st.session_state.single_result = res
                st.rerun()
    else:
        with st.container(border=True):
            render_single_result(st.session_state.single_result)
        st.button(
            "\U0001F504 Process another bill",
            key="reset_single",
            on_click=lambda: st.session_state.update(single_result=None),
        )

with tab_consolidated:
    if st.session_state.consolidated_result is None:
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
                "Generate Consolidated Quote", type="primary", disabled=not uploaded_files, key="run_consolidated"
            )
        if run and uploaded_files:
            data_list = [f.getvalue() for f in uploaded_files]
            names = [f.name for f in uploaded_files]
            res = process_consolidated(data_list, names, client_name)
            if res:
                st.session_state.consolidated_result = res
                st.rerun()
    else:
        with st.container(border=True):
            render_consolidated_result(st.session_state.consolidated_result)
        st.button(
            "\U0001F504 Process another batch",
            key="reset_consolidated",
            on_click=lambda: st.session_state.update(consolidated_result=None),
        )

st.divider()
st.markdown('<span class="meta-caption">Internal tool \u2014 OZ Admin Team.</span>', unsafe_allow_html=True)
