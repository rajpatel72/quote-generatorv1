"""
Bridges the UConX Tampermonkey userscript (see uconx-tariff-grabber.user.js)
back into this Streamlit app.

Why this exists: a popup window is a different origin than the Streamlit
app, so plain JS in the app can never read the popup's content directly
(Same-Origin Policy). The userscript solves that by running *inside* the
UConX page and using window.postMessage to hand the value back across the
origin boundary - postMessage is explicitly designed to be safe for that.

The remaining wrinkle: Streamlit itself has no live JS -> Python channel
outside of a full custom-component build. So once our JS receives the
postMessage, it navigates the top-level page with the result as a URL
query param, which triggers a normal Streamlit rerun - and on that rerun,
consume_incoming_tariff() (called from app.py) picks the value up from
st.query_params and stores it in session_state, keyed by NMI, then strips
it from the URL so it can't be reapplied on later reruns.

REQUIRES: the companion Tampermonkey (or similar) userscript installed in
the browser doing the lookup. Without it, the popup opens normally but
nothing gets sent back - the manual "look at the page and paste the value"
fallback is preserved for anyone who hasn't installed it.
"""
import streamlit as st
import streamlit.components.v1 as components

LOOKUP_URL_TEMPLATE = "https://firstenergy.uconx.com.au/agent/tools/get-address-and-meter-data?nmi={nmi}"

_QP_TARIFF = "auto_tariff"
_QP_NMI = "auto_tariff_nmi"
_SESSION_KEY = "auto_tariff_by_nmi"


def consume_incoming_tariff() -> None:
    """
    Call once near the top of the script, before rendering any UI, on
    every rerun. If the browser just navigated back from the popup flow
    with a tariff result in the URL, stash it in session_state (keyed by
    NMI) and strip those two params from the URL so a later rerun (e.g.
    clicking "process another bill") doesn't reapply a stale value.
    """
    params = st.query_params
    tariff = params.get(_QP_TARIFF)
    nmi = params.get(_QP_NMI)
    if tariff and nmi:
        st.session_state.setdefault(_SESSION_KEY, {})
        st.session_state[_SESSION_KEY][nmi] = tariff
        del params[_QP_TARIFF]
        del params[_QP_NMI]


def get_cached_tariff(nmi: str | None) -> str | None:
    """Returns the most recent auto-looked-up tariff for this NMI, if any."""
    if not nmi:
        return None
    return st.session_state.get(_SESSION_KEY, {}).get(str(nmi))


def render_lookup_button(nmi: str | None, key: str) -> None:
    """
    Renders a button that opens the UConX lookup page for `nmi` in a
    popup. If the userscript is installed, the result comes back
    automatically (the page will visibly reload once it arrives - that's
    the query-param rerun, not an error). If it isn't installed, the
    popup just opens normally and the person can read the value off the
    page and enter it manually elsewhere in the UI.
    """
    if not nmi:
        st.caption("No NMI was extracted from this bill, so auto-lookup isn't available.")
        return

    url = LOOKUP_URL_TEMPLATE.format(nmi=nmi)
    html = f"""
    <div style="font-family: 'Space Grotesk', sans-serif;">
      <button id="lookup-btn-{key}" style="
          padding: 0.5rem 1rem; border-radius: 8px; border: 1px solid #ccc;
          background: #fff; cursor: pointer; font-size: 0.9rem;">
        \U0001F50D Auto-lookup Network Tariff
      </button>
      <span id="lookup-status-{key}" style="margin-left: 0.6rem; color: #666; font-size: 0.85rem;"></span>
    </div>
    <script>
      (function() {{
        const btn = document.getElementById("lookup-btn-{key}");
        const status = document.getElementById("lookup-status-{key}");

        btn.addEventListener("click", function() {{
          status.textContent = "Opening UConX in a popup\u2026";

          // This iframe is sandboxed by Streamlit and can't reliably
          // navigate window.top itself once we're outside the original
          // click's brief activation window (that's why the old version
          // silently failed here). Instead, the userscript running in the
          // popup does the navigation - it's a normal, unsandboxed tab, so
          // it can navigate its opener's top window directly, the same
          // way OAuth popups redirect their opener on completion.
          // It can't read this tab's URL itself (cross-origin), so we hand
          // it over now, while we're still same-origin and mid-click.
          const returnTo = window.top.location.origin + window.top.location.pathname;
          const popupUrl = "{url}" + "&return_to=" + encodeURIComponent(returnTo);
          const popup = window.open(popupUrl, "uconxLookup_{key}", "width=900,height=700");

          if (!popup) {{
            status.textContent = "Popup was blocked \u2014 allow popups for this site and try again.";
            return;
          }}

          let settled = false;
          status.textContent = "Waiting for the userscript to find the tariff \u2014 " +
                                "this page will refresh automatically once it does\u2026";

          function handleMessage(event) {{
            if (!event.data || event.data.type !== "uconx-nmi-tariff-result") return;
            if (String(event.data.nmi) !== "{nmi}") return;
            settled = true;
            window.removeEventListener("message", handleMessage);

            if (!event.data.tariff) {{
              status.textContent = "Userscript ran but found no tariff on the page \u2014 check it manually.";
            }}
            // On success the popup navigates this tab and closes itself
            // directly - there's nothing left to do here.
          }}
          window.addEventListener("message", handleMessage);

          // If nothing comes back in 15s, the userscript probably isn't
          // installed - let the person know instead of waiting forever.
          setTimeout(function() {{
            if (!settled) {{
              status.textContent = "No response yet \u2014 is the Tampermonkey script installed and enabled? " +
                                    "You can still read the value off the popup and enter it manually.";
            }}
          }}, 15000);
        }});
      }})();
    </script>
    """
    components.html(html, height=50)
