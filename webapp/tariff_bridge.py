"""
Bridges the "GloBird NMI NTC Fetcher" Tampermonkey userscript back into this Streamlit app.
"""
import os
import streamlit as st
import streamlit.components.v1 as components

# Any page on this origin will do - the userscript's @match takes care of
# running there, reads `ntc_lookup_nmi` off the URL, and does the rest.
LOOKUP_URL_TEMPLATE = "https://globird-salesportal.azurewebsites.net/?ntc_lookup_nmi={nmi}"
BULK_LOOKUP_URL_TEMPLATE = "https://globird-salesportal.azurewebsites.net/?ntc_lookup_bulk={nmis}"

_MESSAGE_TYPE = "globird-ntc-lookup-result"
_BULK_MESSAGE_TYPE = "globird-ntc-bulk-result"

_SESSION_KEY = "auto_tariff_by_nmi"
_DEBUG_SESSION_KEY = "auto_tariff_debug_by_nmi"

# ---------------------------------------------------------------------------
# DYNAMIC NATIVE CUSTOM COMPONENT SETUP
# ---------------------------------------------------------------------------
COMPONENT_DIR = os.path.join(os.path.dirname(__file__), "ntc_lookup_component")
if not os.path.exists(COMPONENT_DIR):
    os.makedirs(COMPONENT_DIR)

INDEX_HTML_CONTENT = """<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    body { margin: 0; padding: 0; font-family: sans-serif; overflow: hidden; background: transparent; }
    .status-box {
      color: #EC1B40; font-size: 0.88rem; font-weight: 600; padding: 0.6rem; 
      background: rgba(236, 27, 64, 0.04); border: 1px dashed rgba(236, 27, 64, 0.3); 
      border-radius: 8px; box-sizing: border-box; width: 100%; margin-bottom: 6px;
    }
    button {
      cursor:pointer; width:100%; padding: 8px; border-radius: 4px; background: #EC1B40; color: white; border: none; font-weight: bold;
    }
  </style>
</head>
<body>
  <div id="status" class="status-box">⚡ Initializing...</div>
  <button id="lookup-btn">Open Portal to Fetch NTC</button>

  <script>
    let lookupTarget = ""; 
    let isBulk = false;
    let lookupUrlTemplate = "";
    let messageType = "";
    let popupRef = null;
    let popupCheckInterval = null;

    window.addEventListener("message", function(event) {
      if (event.data.type === "streamlit:render") {
        const args = event.data.args;
        lookupUrlTemplate = args.lookup_url_template;
        messageType = args.message_type;
        
        if (args.nmis) {
            isBulk = true;
            lookupTarget = args.nmis.join(",");
            document.getElementById("status").textContent = `⚡ Ready to fetch ${args.nmis.length} tariffs at once...`;
            document.getElementById("lookup-btn").textContent = "Fetch All Missing Tariffs";
        } else {
            isBulk = false;
            lookupTarget = args.nmi;
            document.getElementById("status").textContent = `⚡ Click button to lookup NMI...`;
            document.getElementById("lookup-btn").textContent = "Open Portal to Fetch NTC";
        }
      }
    });

    function resetToRetry(message) {
      if (popupCheckInterval) clearInterval(popupCheckInterval);
      document.getElementById("status").textContent = message;
      document.getElementById("lookup-btn").style.display = "block";
    }

    document.getElementById("lookup-btn").onclick = function() {
      const popupUrl = isBulk ? lookupUrlTemplate.replace("{nmis}", lookupTarget) : lookupUrlTemplate.replace("{nmi}", lookupTarget);
      const popup = window.open(popupUrl, "_blank", "width=1000,height=750");

      if (!popup) {
        resetToRetry("⚠️ Popup blocked! Please allow popups for this site and try again.");
        return;
      }

      popupRef = popup;
      document.getElementById("status").textContent = isBulk ? "Waiting for bulk data from portal tab..." : "Waiting for data from portal tab...";
      document.getElementById("lookup-btn").style.display = "none";

      popupCheckInterval = setInterval(function() {
        if (popup.closed) {
          resetToRetry("⚠️ The portal tab was closed before it finished. Click to try again.");
        }
      }, 800);
    };

    // Listen for data from the popup
    window.addEventListener("message", function(event) {
      if (event.data) {
        if (isBulk && event.data.type === messageType) {
            if (popupCheckInterval) clearInterval(popupCheckInterval);
            document.getElementById("status").textContent = "✅ Bulk NTC data received - updating all sites...";
            window.parent.postMessage({
                isStreamlitMessage: true,
                type: "streamlit:setComponentValue",
                value: { bulk_results: event.data.results }
            }, "*");
        } 
        else if (!isBulk && event.data.type === messageType && String(event.data.nmi) === String(lookupTarget)) {
            if (popupCheckInterval) clearInterval(popupCheckInterval);
            document.getElementById("status").textContent = "✅ NTC received - updating...";
            window.parent.postMessage({
                isStreamlitMessage: true,
                type: "streamlit:setComponentValue",
                value: { ntc: event.data.ntc, nmi: lookupTarget }
            }, "*");
        }
      }
    });

    window.parent.postMessage({ isStreamlitMessage: true, type: "streamlit:componentReady", apiVersion: 1 }, "*");
    window.parent.postMessage({ isStreamlitMessage: true, type: "streamlit:setFrameHeight", height: 110 }, "*");
  </script>
</body>
</html>
"""

# Write component HTML dynamically at runtime
with open(os.path.join(COMPONENT_DIR, "index.html"), "w", encoding="utf-8") as f:
    f.write(INDEX_HTML_CONTENT)

# Declare component inside Streamlit architecture
_ntc_lookup_component = components.declare_component("ntc_lookup_component", path=COMPONENT_DIR)


# ---------------------------------------------------------------------------
# BACKWARD COMPATIBLE INTERFACES FOR APP.PY
# ---------------------------------------------------------------------------
def consume_incoming_tariff() -> None:
    """Kept as a functional stub so app.py doesn't crash if called."""
    pass


def get_cached_tariff(nmi: str | None) -> str | None:
    """Returns the most recent auto-looked-up NTC for this NMI, if any."""
    if not nmi:
        return None
    return st.session_state.get(_SESSION_KEY, {}).get(str(nmi))


def get_cached_tariff_debug(nmi: str | None) -> dict | None:
    """Returns the diagnostic info recorded alongside the cached NTC, if any."""
    if not nmi:
        return None
    return st.session_state.get(_DEBUG_SESSION_KEY, {}).get(str(nmi))


def clear_cached_tariff(nmi: str | None) -> None:
    """Drops any cached auto-lookup value for this NMI."""
    if not nmi:
        return
    st.session_state.get(_SESSION_KEY, {}).pop(str(nmi), None)
    st.session_state.get(_DEBUG_SESSION_KEY, {}).pop(str(nmi), None)


def _save_tariff_to_state(nmi: str, ntc: str):
    """Internal helper to save a fetched tariff securely to session state."""
    st.session_state.setdefault(_SESSION_KEY, {})
    st.session_state[_SESSION_KEY][str(nmi)] = ntc
    st.session_state.setdefault(_DEBUG_SESSION_KEY, {})
    st.session_state[_DEBUG_SESSION_KEY][str(nmi)] = {
        "matched": True,
        "source": "globird_ntc_lookup",
        "tariff": ntc,
    }


def render_lookup_button(nmi: str | None, key: str = "auto") -> None:
    """Original single-site lookup."""
    if not nmi:
        st.caption("No NMI extracted — automated lookup suspended.")
        return

    if get_cached_tariff(nmi):
        return

    component_result = _ntc_lookup_component(
        nmi=nmi,
        lookup_url_template=LOOKUP_URL_TEMPLATE,
        message_type=_MESSAGE_TYPE,
        key=f"ntc_comp_{key}_{nmi}"
    )

    if component_result and isinstance(component_result, dict):
        ntc = component_result.get("ntc")
        returned_nmi = component_result.get("nmi")

        if ntc and returned_nmi:
            _save_tariff_to_state(returned_nmi, ntc)
            st.rerun()


def render_bulk_lookup_button(nmis: list, key: str = "bulk_auto") -> None:
    """New multi-site bulk lookup for Consolidated Quotes."""
    if not nmis:
        return

    component_result = _ntc_lookup_component(
        nmis=nmis,
        lookup_url_template=BULK_LOOKUP_URL_TEMPLATE,
        message_type=_BULK_MESSAGE_TYPE,
        key=f"bulk_comp_{key}"
    )

    if component_result and isinstance(component_result, dict):
        results = component_result.get("bulk_results", [])
        if results:
            for res in results:
                if res.get("ntc") and res.get("nmi"):
                    _save_tariff_to_state(res.get("nmi"), res.get("ntc"))
            st.rerun()


def persist_single_bill(nmi: str | None, result: dict) -> None:
    """Stub to prevent app.py from crashing if called."""
    pass

def persist_consolidated_bill(result: dict) -> None:
    """Stub to prevent app.py from crashing if called."""
    pass


# Add this at the end of tariff_bridge.py
def render_auto_lookup(nmi: str | None, key: str = "auto") -> None:
    """Alias for render_lookup_button to maintain compatibility with app.py."""
    render_lookup_button(nmi, key=key)
