"""
Bridges the "GloBird NMI NTC Fetcher" Tampermonkey userscript back into this Streamlit app.
"""
import os
import streamlit as st
import streamlit.components.v1 as components

# Any page on this origin will do - the userscript's @match takes care of
# running there, reads `ntc_lookup_nmi` off the URL, and does the rest.
LOOKUP_URL_TEMPLATE = "https://globird-salesportal.azurewebsites.net/?ntc_lookup_nmi={nmi}"

_MESSAGE_TYPE = "globird-ntc-lookup-result"

_SESSION_KEY = "auto_tariff_by_nmi"
_DEBUG_SESSION_KEY = "auto_tariff_debug_by_nmi"

# ---------------------------------------------------------------------------
# DYNAMIC NATIVE CUSTOM COMPONENT SETUP
# This safely escapes Streamlit's iframe sandbox restrictions by communicating
# directly over Streamlit's internal WebSocket channel rather than altering URLs.
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
      border-radius: 8px; box-sizing: border-box; width: 100%;
    }
  </style>
</head>
<body>
  <div id="status" class="status-box">⚡ Click button to lookup NMI...</div>
  <button id="lookup-btn" style="cursor:pointer; width:100%; padding: 8px; border-radius: 4px; background: #EC1B40; color: white; border: none;">
    Open Portal to Fetch NTC
  </button>

  <script>
    let nmi = "";
    let lookupUrlTemplate = "";
    let messageType = "";
    let popupRef = null;
    let popupCheckInterval = null;

    window.addEventListener("message", function(event) {
      if (event.data.type === "streamlit:render") {
        nmi = event.data.args.nmi;
        lookupUrlTemplate = event.data.args.lookup_url_template;
        messageType = event.data.args.message_type;
      }
    });

    function resetToRetry(message) {
      if (popupCheckInterval) {
        clearInterval(popupCheckInterval);
        popupCheckInterval = null;
      }
      document.getElementById("status").textContent = message;
      document.getElementById("lookup-btn").style.display = "block";
      document.getElementById("lookup-btn").textContent = "Retry: Open Portal to Fetch NTC";
    }

    document.getElementById("lookup-btn").onclick = function() {
      const popupUrl = lookupUrlTemplate.replace("{nmi}", nmi);

      // IMPORTANT: do NOT pass "noopener" here. The userscript running in
      // the popup sends its result back via `window.opener.postMessage(...)`.
      // "noopener" sets window.opener to null in the popup, which silently
      // breaks that entirely (the userscript just gives up with no error
      // visible in this tab) - this was the root cause of lookups never
      // coming back.
      const popup = window.open(popupUrl, "_blank", "width=1000,height=750");

      if (!popup) {
        document.getElementById("status").textContent = "⚠️ Popup blocked! Please allow popups for this site and try again.";
        return;
      }

      popupRef = popup;
      document.getElementById("status").textContent = "Waiting for data from portal (a new tab just opened)...";
      document.getElementById("lookup-btn").style.display = "none";

      // Safety net: if the person closes the portal tab (or it fails to
      // load / the userscript isn't installed) before it ever sends data
      // back, don't leave this stuck on "Waiting..." forever - flip back
      // to a retry state.
      popupCheckInterval = setInterval(function() {
        if (popup.closed) {
          resetToRetry("⚠️ The portal tab was closed before it finished. Click below to try again.");
        }
      }, 800);
    };

    // Listen for data from the popup
    window.addEventListener("message", function(event) {
      if (event.data && event.data.type === messageType && String(event.data.nmi) === String(nmi)) {
        if (popupCheckInterval) {
          clearInterval(popupCheckInterval);
          popupCheckInterval = null;
        }
        document.getElementById("status").textContent = "✅ NTC received - updating...";
        window.parent.postMessage({
          isStreamlitMessage: true,
          type: "streamlit:setComponentValue",
          value: { ntc: event.data.ntc, nmi: nmi }
        }, "*");
      }
    });

    window.parent.postMessage({ isStreamlitMessage: true, type: "streamlit:componentReady", apiVersion: 1 }, "*");
    window.parent.postMessage({ isStreamlitMessage: true, type: "streamlit:setFrameHeight", height: 90 }, "*");
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


def render_lookup_button(nmi: str | None, key: str = "auto") -> None:
    """
    Renders the custom lookup automation interface. Receives the data values 
    from JavaScript execution natively and assigns them directly to session state.
    """
    if not nmi:
        st.caption("No NMI extracted — automated lookup suspended.")
        return

    # Check if we already have a cached tariff for this NMI to prevent reloading loops
    if get_cached_tariff(nmi):
        return

    # Run the native component and capture any values returned by JS window.parent.postMessage
    component_result = _ntc_lookup_component(
        nmi=nmi,
        lookup_url_template=LOOKUP_URL_TEMPLATE,
        message_type=_MESSAGE_TYPE,
        key=f"ntc_comp_{key}_{nmi}"
    )

    # Process data instantly when streamed back from the JavaScript frame
    if component_result and isinstance(component_result, dict):
        ntc = component_result.get("ntc")
        returned_nmi = component_result.get("nmi")

        if ntc and returned_nmi:
            st.session_state.setdefault(_SESSION_KEY, {})
            st.session_state[_SESSION_KEY][str(returned_nmi)] = ntc
            st.session_state.setdefault(_DEBUG_SESSION_KEY, {})
            st.session_state[_DEBUG_SESSION_KEY][str(returned_nmi)] = {
                "matched": True,
                "source": "globird_ntc_lookup",
                "tariff": ntc,
            }
            # Issue a clean Python rerun to refresh state variables and process your tool pipeline
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
