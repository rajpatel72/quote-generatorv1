"""
Reads however many GEMINI_API_KEYs are configured in Streamlit secrets / env,
so the extraction clients can fail over from key 1 -> key 2 -> key 3 ... when
one hits a rate limit. This does NOT run anything in parallel and does NOT
split bills into chunks -- it's the same single call as before, just with a
list of keys to try in order instead of one fixed key.

Supported ways to supply multiple keys (checked in this order, first match wins):
    1. GEMINI_API_KEYS = "key1, key2, key3"   (comma or newline separated)
    2. Numbered slots: GEMINI_API_KEY (or GEMINI_API_KEY_1), GEMINI_API_KEY_2,
       GEMINI_API_KEY_3, ...
"""
from typing import Optional


def load_api_keys(secrets, env) -> list[str]:
    keys: list[str] = []

    bundle = _get(secrets, env, "GEMINI_API_KEYS")
    if bundle:
        for part in bundle.replace("\n", ",").split(","):
            if part.strip():
                keys.append(part.strip())

    first = _get(secrets, env, "GEMINI_API_KEY") or _get(secrets, env, "GEMINI_API_KEY_1")
    if first:
        keys.append(first)
    i = 2
    while True:
        k = _get(secrets, env, f"GEMINI_API_KEY_{i}")
        if not k:
            break
        keys.append(k)
        i += 1

    seen = set()
    deduped = []
    for k in keys:
        if k not in seen:
            seen.add(k)
            deduped.append(k)
    return deduped


def _get(secrets, env, name: str) -> Optional[str]:
    val = None
    try:
        val = secrets.get(name)
    except Exception:  # noqa: BLE001
        val = None
    if not val:
        val = env.get(name)
    return val
