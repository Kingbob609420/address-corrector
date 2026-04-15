import io
import os
import importlib
import pandas as pd
import streamlit as st

# Force fresh import every deploy — prevents Streamlit Cloud from serving stale module
import address_corrector as _ac
importlib.reload(_ac)
import re as _re
from address_corrector import (
    CORRECTORS, DISPLAY_LABELS, _detect_columns, _write_excel, apply_autofix,
    correct_address_line, correct_city, correct_state, correct_country, correct_postal_code,
    detect_country_from_postal, infer_province_from_canadian_postal, infer_us_state_from_zip,
    _US_STATE_CODES, _NULL_PLACEHOLDERS, _STATE_CODE_TO_COUNTRY,
    _COUNTRY_NAME_INDEX, _STATE_FUZZY_INDEX,
)


# ── AI street-name spell-checker ──────────────────────────────────────────────
def _ai_correct_street(line: str) -> str:
    """
    Use Claude claude-haiku-3-5 to fix misspelled words in a street address line.
    Falls back to the original value silently if no API key or any error.
    """
    if not line or not str(line).strip():
        return line
    try:
        import anthropic  # type: ignore

        # Resolve API key: Streamlit Cloud secrets → env var
        api_key = ""
        try:
            api_key = st.secrets.get("ANTHROPIC_API_KEY", "")
        except Exception:
            pass
        if not api_key:
            api_key = os.environ.get("ANTHROPIC_API_KEY", "")
        if not api_key:
            return line  # No key available — skip silently

        client = anthropic.Anthropic(api_key=api_key)
        msg = client.messages.create(
            model="claude-haiku-3-5",
            max_tokens=150,
            messages=[{
                "role": "user",
                "content": (
                    "Fix ONLY misspelled street name words in this address line. "
                    "Do NOT change the building/house number at the start. "
                    "Do NOT alter standard road-type abbreviations "
                    "(DR, ST, AVE, BLVD, RD, LN, CT, PL, WAY, HWY, CIR, TER, etc.). "
                    "If nothing looks misspelled, return it exactly as given. "
                    "Return ONLY the corrected address line — no explanation, no punctuation changes.\n\n"
                    f"Address: {line}"
                ),
            }],
        )
        result = msg.content[0].text.strip()
        # Safety guard: if AI returns something wildly different in length, keep original
        if result and 0.5 < len(result) / max(len(str(line)), 1) < 2.5:
            return result
        return line
    except Exception:
        return line  # Any failure → return unchanged


st.set_page_config(
    page_title="Address Corrector",
    page_icon="📍",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');

*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

html, body,
[data-testid="stAppViewContainer"],
[data-testid="stMain"],
[data-testid="stMainBlockContainer"] {
    font-family: 'Inter', sans-serif !important;
    background: #f5f4f0 !important;
    color: #111118;
}

/* Kill all Streamlit chrome */
#MainMenu, footer, header,
[data-testid="stToolbar"],
[data-testid="stDecoration"],
[data-testid="stStatusWidget"],
[data-testid="stSidebarCollapsedControl"] { display:none !important; }

.block-container {
    padding: 0 !important;
    max-width: 100% !important;
}

/* ══ NAV ══════════════════════════════════════════════ */
.nav {
    position: sticky;
    top: 0;
    z-index: 100;
    background: rgba(245,244,240,.85);
    backdrop-filter: blur(12px);
    -webkit-backdrop-filter: blur(12px);
    border-bottom: 1px solid rgba(0,0,0,.07);
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 0 3.5rem;
    height: 58px;
}
.nav-logo {
    font-weight: 800;
    font-size: 1.05rem;
    color: #111118;
    letter-spacing: -.4px;
}
.nav-cta {
    background: #6c47ff;
    color: #fff !important;
    border: none;
    border-radius: 100px;
    padding: .45rem 1.3rem;
    font-size: .82rem;
    font-weight: 700;
    cursor: pointer;
    text-decoration: none;
    transition: background .15s, transform .1s;
    display: inline-block;
}
.nav-cta:hover { background: #5a35ed; transform: translateY(-1px); }

/* ══ HERO ══════════════════════════════════════════════ */
.hero {
    display: flex;
    flex-direction: column;
    align-items: center;
    text-align: center;
    padding: 7rem 2rem 5rem;
}
.hero-eyebrow {
    font-size: .72rem;
    font-weight: 700;
    letter-spacing: 1.8px;
    text-transform: uppercase;
    color: #6c47ff;
    margin-bottom: 1.4rem;
}
.hero-title {
    font-size: clamp(2.6rem, 5vw, 4rem);
    font-weight: 900;
    line-height: 1.12;
    letter-spacing: -2px;
    color: #111118;
    max-width: 780px;
    margin-bottom: 1.5rem;
}
.hl {
    background: #6c47ff;
    color: #fff;
    border-radius: 6px;
    padding: .04em .18em;
    display: inline-block;
    transform: rotate(-.5deg);
}
.hero-sub {
    font-size: 1.05rem;
    color: #71717a;
    line-height: 1.75;
    max-width: 460px;
    margin-bottom: 2.4rem;
}
.hero-btn {
    background: #6c47ff;
    color: #fff;
    border: none;
    border-radius: 100px;
    padding: .85rem 2.4rem;
    font-size: .95rem;
    font-weight: 700;
    cursor: pointer;
    transition: background .15s, box-shadow .15s, transform .1s;
    text-decoration: none;
    box-shadow: 0 4px 24px rgba(108,71,255,.35);
    display: inline-block;
}
.hero-btn:hover {
    background: #5a35ed;
    box-shadow: 0 8px 32px rgba(108,71,255,.45);
    transform: translateY(-2px);
}

/* ══ DIVIDER ══════════════════════════════════════════ */
.divider {
    width: 100%;
    height: 1px;
    background: linear-gradient(to right, transparent, rgba(0,0,0,.1) 30%, rgba(0,0,0,.1) 70%, transparent);
    margin: 0;
}

/* ══ TOOL SECTION ══════════════════════════════════════ */
.tool-wrap {
    max-width: 1060px;
    margin: 0 auto;
    padding: 4rem 2rem 6rem;
}
.tool-label {
    text-align: center;
    font-size: .7rem;
    font-weight: 700;
    letter-spacing: 1.6px;
    text-transform: uppercase;
    color: #a1a1aa;
    margin-bottom: .9rem;
}
.tool-heading {
    text-align: center;
    font-size: 1.7rem;
    font-weight: 800;
    letter-spacing: -.8px;
    color: #111118;
    margin-bottom: .6rem;
}
.tool-sub {
    text-align: center;
    font-size: .9rem;
    color: #71717a;
    margin-bottom: 2.8rem;
}

/* ══ INPUT PANELS ══════════════════════════════════════ */
.panels-grid {
    display: grid;
    grid-template-columns: 1fr auto 1fr;
    gap: 0;
    align-items: stretch;
    margin-bottom: 2.8rem;
}
.panel {
    background: #ffffff;
    border: 1px solid rgba(0,0,0,.08);
    border-radius: 18px;
    padding: 1.8rem 2rem;
    box-shadow: 0 2px 16px rgba(0,0,0,.05);
}
.panel-icon  { font-size: 1.6rem; margin-bottom: .7rem; }
.panel-title { font-size: .95rem; font-weight: 700; color: #111118; margin-bottom: .3rem; }
.panel-desc  { font-size: .78rem; color: #a1a1aa; line-height: 1.6; margin-bottom: 1rem; }

.or-col {
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    padding: 0 1.2rem;
    gap: .4rem;
}
.or-line-seg {
    flex: 1;
    width: 1px;
    background: rgba(0,0,0,.08);
}
.or-circle {
    width: 30px; height: 30px;
    border-radius: 50%;
    border: 1px solid rgba(0,0,0,.1);
    background: #f5f4f0;
    display: flex; align-items: center; justify-content: center;
    font-size: .65rem; font-weight: 700; color: #a1a1aa;
    flex-shrink: 0;
}

/* ══ TEXT INPUTS ═══════════════════════════════════════ */
[data-testid="stTextInput"] input,
[data-testid="stTextInput"] input:not([type]),
div[data-testid="stTextInput"] > div > div > input {
    background: #ffffff !important;
    border: 1.5px solid rgba(0,0,0,.10) !important;
    border-radius: 10px !important;
    color: #111118 !important;
    font-family: 'Inter', sans-serif !important;
    font-size: .88rem !important;
    padding: .55rem .85rem !important;
    transition: border-color .15s, box-shadow .15s !important;
    box-shadow: 0 1px 3px rgba(0,0,0,.04) !important;
}
[data-testid="stTextInput"] input:focus,
div[data-testid="stTextInput"] > div > div > input:focus {
    border-color: #6c47ff !important;
    box-shadow: 0 0 0 3px rgba(108,71,255,.12) !important;
    outline: none !important;
    background: #ffffff !important;
}
[data-testid="stTextInput"] input::placeholder { color: #c4c4cc !important; }
[data-testid="stTextInput"] label {
    font-size: .73rem !important;
    font-weight: 600 !important;
    color: #71717a !important;
    letter-spacing: .02em !important;
    text-transform: uppercase !important;
}
/* Force white background on ALL inputs globally */
input[type="text"], input:not([type]) {
    background-color: #ffffff !important;
    color: #111118 !important;
}

/* ══ PRIMARY BUTTON (purple, not red) ═════════════════ */
[data-testid="stButton"] > button[kind="primary"],
[data-testid="stButton"] > button {
    background: #6c47ff !important;
    color: #ffffff !important;
    border: none !important;
    border-radius: 100px !important;
    font-weight: 700 !important;
    font-size: .88rem !important;
    height: 46px !important;
    box-shadow: 0 4px 16px rgba(108,71,255,.30) !important;
    transition: background .15s, box-shadow .15s, transform .1s !important;
}
[data-testid="stButton"] > button:hover {
    background: #5a35ed !important;
    box-shadow: 0 6px 22px rgba(108,71,255,.42) !important;
    transform: translateY(-1px) !important;
}

/* ══ TABS ══════════════════════════════════════════════ */
[data-testid="stTabs"] [data-testid="stTab"] {
    font-weight: 600 !important;
    font-size: .85rem !important;
    color: #71717a !important;
    border-radius: 8px 8px 0 0 !important;
    padding: .55rem 1.4rem !important;
    transition: color .15s !important;
}
[data-testid="stTabs"] [data-testid="stTab"][aria-selected="true"] {
    color: #6c47ff !important;
    border-bottom: 2px solid #6c47ff !important;
}
[data-testid="stTabs"] [role="tablist"] {
    border-bottom: 1.5px solid rgba(0,0,0,.08) !important;
    margin-bottom: 1.8rem !important;
}

/* Textarea */
textarea {
    background: #fafafa !important;
    border: 1px solid rgba(0,0,0,.09) !important;
    border-radius: 10px !important;
    color: #111118 !important;
    font-family: 'SF Mono','Fira Code',monospace !important;
    font-size: .77rem !important;
    transition: border-color .15s !important;
}
textarea:focus {
    border-color: #6c47ff !important;
    box-shadow: 0 0 0 3px rgba(108,71,255,.1) !important;
    outline: none !important;
}
textarea::placeholder { color: #d4d4d8 !important; }

/* File uploader */
[data-testid="stFileUploader"] section {
    background: #fafafa !important;
    border: 1.5px dashed rgba(0,0,0,.1) !important;
    border-radius: 12px !important;
    transition: border-color .15s !important;
}
[data-testid="stFileUploader"] section:hover {
    border-color: #6c47ff !important;
}
[data-testid="stFileUploader"] button {
    background: #f5f4f0 !important;
    border: 1px solid rgba(0,0,0,.1) !important;
    color: #52525b !important;
    border-radius: 8px !important;
    font-size: .8rem !important;
}
[data-testid="stFileUploader"] small,
[data-testid="stFileUploader"] span { color: #a1a1aa !important; font-size: .78rem !important; }

/* ══ RESULTS SECTION ═══════════════════════════════════ */
.results-card {
    background: #ffffff;
    border: 1px solid rgba(0,0,0,.08);
    border-radius: 18px;
    padding: 2rem;
    box-shadow: 0 2px 16px rgba(0,0,0,.05);
}

/* Metric tiles */
.metrics-row { display: grid; grid-template-columns: repeat(4,1fr); gap: .75rem; margin-bottom: 1.6rem; }
.mtile {
    background: #f5f4f0;
    border-radius: 12px;
    padding: 1rem 1.2rem;
    border: 1px solid rgba(0,0,0,.06);
}
.mtile-val { font-size: 1.8rem; font-weight: 800; color: #111118; letter-spacing: -1px; }
.mtile-lbl { font-size: .68rem; font-weight: 600; text-transform: uppercase; letter-spacing: .7px; color: #a1a1aa; margin-top: .2rem; }

/* Column badges */
.badge-row { display: flex; flex-wrap: wrap; gap: .45rem; margin-bottom: 1.4rem; }
.bdg {
    display: inline-flex; align-items: center; gap: .35rem;
    border-radius: 100px;
    padding: .3rem .8rem;
    font-size: .73rem;
    font-weight: 600;
}
.bdg-ok  { background: #f0fdf4; border: 1px solid #bbf7d0; color: #16a34a; }
.bdg-no  { background: #fefce8; border: 1px solid #fde68a; color: #ca8a04; }
.bdg-src { font-weight: 400; opacity: .7; }

/* Radio pills */
[data-testid="stRadio"] > div { display:flex; gap:.4rem; flex-wrap:wrap; }
[data-testid="stRadio"] label {
    background: #f5f4f0 !important;
    border: 1px solid rgba(0,0,0,.09) !important;
    border-radius: 100px !important;
    padding: .28rem .85rem !important;
    font-size: .76rem !important;
    font-weight: 500 !important;
    color: #71717a !important;
    cursor: pointer;
    transition: all .15s !important;
}
[data-testid="stRadio"] label:has(input:checked) {
    background: #6c47ff !important;
    border-color: #6c47ff !important;
    color: #ffffff !important;
}

/* Expander */
[data-testid="stExpander"] {
    background: #fafafa !important;
    border: 1px solid rgba(0,0,0,.07) !important;
    border-radius: 10px !important;
    margin-bottom: 1rem !important;
}

/* Download buttons */
[data-testid="stDownloadButton"] button {
    border-radius: 100px !important;
    font-weight: 700 !important;
    font-size: .82rem !important;
    height: 40px !important;
    transition: all .15s !important;
}
[data-testid="stDownloadButton"]:first-child > button {
    background: #6c47ff !important;
    color: #fff !important;
    border: none !important;
    box-shadow: 0 4px 14px rgba(108,71,255,.3) !important;
}
[data-testid="stDownloadButton"]:first-child > button:hover {
    background: #5a35ed !important;
    box-shadow: 0 6px 20px rgba(108,71,255,.4) !important;
    transform: translateY(-1px) !important;
}

/* Success toast */
.toast-ok {
    display: flex; align-items: center; gap: .5rem;
    background: #f0fdf4; border: 1px solid #bbf7d0;
    border-radius: 8px; padding: .5rem .85rem;
    font-size: .76rem; color: #16a34a; font-weight: 600;
    margin-top: .5rem;
}

/* Legend */
.legend { display:flex; align-items:center; gap:1rem; margin-top:.7rem; font-size:.73rem; color:#a1a1aa; }
.ldot { width:10px; height:10px; border-radius:3px; display:inline-block; margin-right:.3rem; }

/* Scrollbar */
::-webkit-scrollbar { width:5px; height:5px; }
::-webkit-scrollbar-track { background:#f5f4f0; }
::-webkit-scrollbar-thumb { background:#d4d4d8; border-radius:3px; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# NAV
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="nav">
    <span class="nav-logo">📍 Address Corrector</span>
    <a class="nav-cta" href="#load">Get Started</a>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# HERO
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="hero">
    <div class="hero-eyebrow">Address standardisation tool</div>
    <h1 class="hero-title">
        The address data you<br>need to <span class="hl">trust</span>
    </h1>
    <p class="hero-sub">
        Paste from Excel or upload a file. Columns are auto-detected,
        corrections applied instantly — across all 249 countries.
    </p>
    <a class="hero-btn" href="#load">Correct my addresses</a>
</div>
<div class="divider"></div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TOOL  (anchored via id="load" on the wrapper)
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="tool-wrap" id="load">
    <div class="tool-label">Step 1</div>
    <div class="tool-heading">Load your addresses</div>
    <div class="tool-sub">Correct a single address instantly, or process hundreds at once.</div>
</div>
""", unsafe_allow_html=True)

def _parse_full_address(text):
    """
    Parse a free-text address into (addr1, addr2, addr3, city, state, country, postal).
    Strategy: scan ALL parts for postal/country/state using patterns + lookups,
    then assign what remains to address/city (working from end = city, front = street).
    """
    parts = [p.strip() for p in _re.split(r"[,\n;|]+", text) if p.strip()]
    if not parts:
        return "", "", "", "", "", "", ""

    used   = [False] * len(parts)
    postal = state = country = ""

    # ── Postal patterns (tried against every part, highest specificity first) ───
    _postal_pats = [
        r"^[A-Za-z]\d[A-Za-z]\d[A-Za-z]\d$",          # Canadian A1A1A1
        r"^[A-Za-z]{1,2}\d[A-Za-z\d]?\d[A-Za-z]{2}$", # UK SW1A2AA
        r"^\d{5}-\d{4}$",                               # US ZIP+4
        r"^\d{5}-\d{5,}$",                              # malformed 5+N ZIP
        r"^\d{4}[A-Za-z]{2}$",                          # Dutch 1234AB
        r"^\d{4,6}$",                                   # generic numeric
    ]

    # ── 1. Find postal (last matching part wins so street numbers don't steal) ──
    for i in reversed(range(len(parts))):
        s = parts[i].replace(" ", "")
        for pat in _postal_pats:
            if _re.match(pat, s, _re.I):
                postal = parts[i]
                used[i] = True
                break
        if postal:
            break

    # ── 2. Find country (working from right, skip used) ─────────────────────────
    for i in reversed(range(len(parts))):
        if used[i]:
            continue
        part = parts[i]
        test = part.strip().lower()
        corrected_c = correct_country(part)
        # Match if: (a) lookup returns a different (shorter) code, or (b) exact key
        if (corrected_c and corrected_c != part.strip().upper()) or test in _COUNTRY_NAME_INDEX:
            country = part
            used[i] = True
            break

    # ── 3. Find state (working from right, skip used) ────────────────────────────
    for i in reversed(range(len(parts))):
        if used[i]:
            continue
        part  = parts[i]
        test  = part.strip().lower()
        short = part.strip()
        if (len(short) <= 3 and short.isalpha()) or test in _STATE_FUZZY_INDEX:
            state = part
            used[i] = True
            break

    # ── 4. Remaining → last unused = city, earlier unused = address lines ────────
    remaining = [parts[i] for i in range(len(parts)) if not used[i]]
    city      = remaining[-1]        if remaining          else ""
    addr_parts = remaining[:-1]      if len(remaining) > 1 else []
    addr1     = addr_parts[0]        if len(addr_parts) > 0 else ""
    addr2     = addr_parts[1]        if len(addr_parts) > 1 else ""
    addr3     = ", ".join(addr_parts[2:]) if len(addr_parts) > 2 else ""

    return addr1, addr2, addr3, city, state, country, postal


def _run_single_correction(addr1, addr2, addr3, city, state, country, postal, use_ai=False):
    """Apply all corrections to a single address and return (original, corrected) dicts."""
    _null_upper = {s.upper() for s in _NULL_PLACEHOLDERS}

    # Step 1 — correct each field independently
    c_addr1  = correct_address_line(addr1)
    c_addr2  = correct_address_line(addr2)
    c_addr3  = correct_address_line(addr3)

    # Step 1b — AI spell-check street name words (optional)
    if use_ai:
        c_addr1 = _ai_correct_street(c_addr1) if c_addr1 else c_addr1
        c_addr2 = _ai_correct_street(c_addr2) if c_addr2 else c_addr2
        c_addr3 = _ai_correct_street(c_addr3) if c_addr3 else c_addr3
    c_city   = correct_city(city)
    c_postal = correct_postal_code(postal)
    c_country = correct_country(country)
    c_state  = correct_state(state)
    # Treat null placeholders as empty
    if c_state.upper() in _null_upper:
        c_state = ""

    # Step 2 — auto-fix country using postal code (strongest signal)
    postal_inferred = detect_country_from_postal(c_postal)
    if postal_inferred:
        c_country = postal_inferred

    # Step 3 — auto-fix country using state code if postal gave no answer
    if not postal_inferred and c_state:
        state_country = _STATE_CODE_TO_COUNTRY.get(c_state.upper(), "")
        if state_country:
            c_country = state_country

    # Step 4 — infer province from Canadian FSA when country is CA
    if c_country == "CA":
        prov = infer_province_from_canadian_postal(c_postal)
        if prov and (not c_state or c_state in _US_STATE_CODES):
            c_state = prov

    # Step 4b — correct US state from ZIP prefix (overrides wrong state entry)
    if c_country == "US":
        expected_state = infer_us_state_from_zip(c_postal)
        if expected_state and expected_state != c_state:
            c_state = expected_state

    # Step 5 — re-correct state now we know the definitive country
    #          (e.g. user typed "Ontario" → correct_state already gave "ON",
    #           but if it was a misspelling this pass cleans it up with country context)
    if c_state and len(c_state) > 2:
        c_state = correct_state(c_state, country_hint=c_country)

    originals  = {"Address Line 1": addr1,  "Address Line 2": addr2,
                  "Address Line 3": addr3,  "City": city,
                  "State": state, "Country": country, "Postal Code": postal}
    corrected  = {"Address Line 1": c_addr1, "Address Line 2": c_addr2,
                  "Address Line 3": c_addr3, "City": c_city,
                  "State": c_state, "Country": c_country, "Postal Code": c_postal}
    return originals, corrected


def _render_single_result(payload):
    originals, corrected = payload
    field_order = ["Address Line 1", "Address Line 2", "Address Line 3",
                   "City", "State", "Country", "Postal Code"]

    any_change  = False
    rows_html   = ""
    n_corrected = 0

    for label in field_order:
        orig_disp = (originals.get(label) or "").strip()
        corr_disp = (corrected.get(label) or "").strip()
        if not orig_disp and not corr_disp:
            continue
        changed = orig_disp != corr_disp and bool(corr_disp)
        if changed:
            any_change  = True
            n_corrected += 1

        if changed:
            # Amber original → green corrected
            orig_cell = (f'<span style="font-size:.87rem;color:#a16207;background:#fefce8;'
                         f'padding:2px 7px;border-radius:4px;font-weight:500">{orig_disp}</span>'
                         f'<span style="color:#c4c4cc;margin:0 6px">&#8594;</span>'
                         f'<span style="font-size:.87rem;color:#15803d;background:#f0fdf4;'
                         f'padding:2px 7px;border-radius:4px;font-weight:700">{corr_disp}</span>')
        else:
            # Unchanged — show corrected value once (clean, no duplication)
            disp = corr_disp if corr_disp else "—"
            orig_cell = f'<span style="font-size:.87rem;color:#111118;font-weight:500">{disp}</span>'

        rows_html += (
            f'<div style="display:flex;align-items:center;gap:.5rem;padding:.5rem 0;'
            f'border-bottom:1px solid #f4f4f5">'
            f'<span style="font-size:.71rem;color:#a1a1aa;min-width:120px;flex-shrink:0;'
            f'text-transform:uppercase;letter-spacing:.04em">{label}</span>'
            f'{orig_cell}</div>'
        )

    # ── Single combined markdown (avoids unclosed-tag rendering bugs) ──────────
    clean_badge = ('<div style="display:inline-flex;align-items:center;gap:.4rem;'
                   'background:#f0fdf4;border:1px solid #bbf7d0;border-radius:100px;'
                   'padding:.25rem .75rem;font-size:.72rem;color:#16a34a;font-weight:600">'
                   '✓ Already correct</div>') if not any_change else ""

    summary = (f'<div style="font-size:.72rem;color:#a1a1aa;margin-top:.8rem">'
               f'{n_corrected} field{"s" if n_corrected!=1 else ""} corrected</div>'
               if any_change else "")

    st.markdown(
        f'<div style="max-width:680px;margin:1.5rem auto 0">'
        f'<div style="background:#fff;border:1px solid rgba(0,0,0,.08);border-radius:14px;'
        f'padding:1.4rem 1.8rem;margin-bottom:1rem">'
        f'<div style="font-size:.72rem;font-weight:700;letter-spacing:.08em;'
        f'color:#6c47ff;text-transform:uppercase;margin-bottom:1.1rem">Corrected Address</div>'
        f'{rows_html}'
        f'{summary}'
        f'</div>'
        f'{clean_badge}'
        f'</div>',
        unsafe_allow_html=True,
    )

    # ── Legend ─────────────────────────────────────────────────────────────────
    if any_change:
        st.markdown(
            '<div style="max-width:680px;margin:.4rem auto 0" class="legend">'
            '<span><span class="ldot" style="background:#f0fdf4;border:1px solid #86efac"></span>Corrected</span>'
            '<span><span class="ldot" style="background:#fefce8;border:1px solid #fde047"></span>Original</span>'
            '</div>',
            unsafe_allow_html=True,
        )

    # ── Download buttons ───────────────────────────────────────────────────────
    st.markdown(
        '<div style="max-width:680px;margin:1rem auto 0">'
        '<p style="font-size:.72rem;font-weight:700;letter-spacing:.7px;'
        'text-transform:uppercase;color:#a1a1aa;margin-bottom:.6rem">Export</p>'
        '</div>',
        unsafe_allow_html=True,
    )
    # Build a one-row DataFrame for export
    export_df = pd.DataFrame([{
        **{f"Original {k}": v for k, v in originals.items()},
        **{f"Corrected {k}": v for k, v in corrected.items()},
    }])
    _dl1, _dl2, _dl3 = st.columns([2, 2, 5])
    with _dl1:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            export_df.to_excel(writer, index=False, sheet_name="Corrected Address")
        buf.seek(0)
        st.download_button("⬇  Excel (.xlsx)", buf,
                           "corrected_address.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True, type="primary")
    with _dl2:
        st.download_button("⬇  CSV (.csv)",
                           export_df.to_csv(index=False).encode("utf-8-sig"),
                           "corrected_address.csv", "text/csv",
                           use_container_width=True)


tab_single, tab_bulk = st.tabs(["  Single Address  ", "  Bulk (Paste / Upload)  "])
df_raw = None

# ── Tab 1: Single address ─────────────────────────────────────────────────────
with tab_single:
    st.markdown('<div style="max-width:680px;margin:1.5rem auto 0">', unsafe_allow_html=True)

    # ── Free-text entry (type the whole address at once) ──────────────────────
    st.markdown("""
    <div style="font-size:.73rem;font-weight:600;color:#71717a;
                text-transform:uppercase;letter-spacing:.06em;margin-bottom:.4rem">
        Full Address
    </div>""", unsafe_allow_html=True)
    full_addr = st.text_area(
        "full_addr", height=110, label_visibility="collapsed",
        placeholder=(
            "Type or paste the whole address — any format:\n"
            "123 Main St, Calgary, KS, United States, T3M 0V4\n"
            "or fill in the individual fields below instead"
        ),
        key="single_full_addr",
    )

    st.markdown('<div style="margin:1rem 0 .5rem;border-top:1px solid #f0f0f0;padding-top:1rem">'
                '<span style="font-size:.73rem;font-weight:600;color:#a1a1aa;'
                'text-transform:uppercase;letter-spacing:.06em">Or fill in individually</span>'
                '</div>', unsafe_allow_html=True)

    # ── Individual fields ─────────────────────────────────────────────────────
    c1, c2 = st.columns(2)
    with c1:
        s_addr1   = st.text_input("Address Line 1", placeholder="123 Main St",  key="s_a1")
        s_addr2   = st.text_input("Address Line 2", placeholder="Suite 400",    key="s_a2")
        s_city    = st.text_input("City",           placeholder="Calgary",       key="s_ci")
    with c2:
        s_state   = st.text_input("State / Province", placeholder="AB",         key="s_st")
        s_country = st.text_input("Country",          placeholder="Canada",      key="s_co")
        s_postal  = st.text_input("Postal / ZIP Code", placeholder="T3M 0V4",   key="s_po")

    use_ai_single = st.checkbox(
        "🤖 AI spell-check street names",
        value=True,
        key="single_ai_fix",
        help="Uses Claude AI to fix misspelled street name words (e.g. 'Plmdal' → 'Palmdale'). "
             "Requires ANTHROPIC_API_KEY to be set.",
    )
    run_single = st.button("Correct this address", type="primary",
                           use_container_width=True, key="btn_single")
    st.markdown('</div>', unsafe_allow_html=True)

    if run_single:
        # Prefer the free-text box; fall back to individual fields
        if full_addr.strip():
            p_a1, p_a2, p_a3, p_city, p_state, p_country, p_postal = \
                _parse_full_address(full_addr.strip())
            with st.spinner("Correcting address…"):
                st.session_state["single_result"] = _run_single_correction(
                    p_a1, p_a2, p_a3, p_city, p_state, p_country, p_postal,
                    use_ai=use_ai_single,
                )
        else:
            with st.spinner("Correcting address…"):
                st.session_state["single_result"] = _run_single_correction(
                    s_addr1, s_addr2, "", s_city, s_state, s_country, s_postal,
                    use_ai=use_ai_single,
                )

    if "single_result" in st.session_state:
        _render_single_result(st.session_state["single_result"])

# ── Tab 2: Bulk input (paste + upload) ───────────────────────────────────────
with tab_bulk:
    col_paste, col_or, col_upload = st.columns([10, 1, 10])

    with col_paste:
        st.markdown("""
        <div class="panel">
            <div class="panel-icon">📋</div>
            <div class="panel-title">Paste from Excel</div>
            <div class="panel-desc">
                Copy a block of cells with the header row and paste below.<br>
                Tab, comma &amp; semicolon delimiters auto-detected.
            </div>
        </div>
        """, unsafe_allow_html=True)
        pasted = st.text_area(
            "paste", height=200, label_visibility="collapsed",
            placeholder=(
                "Address Line 1\tCity\tState\tCountry\tPostal Code\n"
                "123 main st\tnew york\tnew york\tUSA\t10001\n"
                "10 Downing St\tlondon\t\tuk\tSW1A2AA"
            ),
        )
        if pasted.strip():
            sample = pasted.split("\n")[0]
            tabs, commas, semis = sample.count("\t"), sample.count(","), sample.count(";")
            delim = "\t" if tabs >= commas and tabs >= semis else (";" if semis > commas else ",")
            try:
                df_raw = pd.read_csv(io.StringIO(pasted), sep=delim, dtype=str, on_bad_lines="skip").fillna("")
                st.markdown(f'<div class="toast-ok">✓ &nbsp;{len(df_raw):,} rows ready</div>', unsafe_allow_html=True)
            except Exception as e:
                st.error(str(e))

    with col_or:
        st.markdown("""
        <div class="or-col" style="height:340px">
            <div class="or-line-seg"></div>
            <div class="or-circle">OR</div>
            <div class="or-line-seg"></div>
        </div>
        """, unsafe_allow_html=True)

    with col_upload:
        st.markdown("""
        <div class="panel">
            <div class="panel-icon">📁</div>
            <div class="panel-title">Upload a file</div>
            <div class="panel-desc">
                Drop in a CSV or Excel file.<br>
                All column formats handled automatically.
            </div>
        </div>
        """, unsafe_allow_html=True)
        uploaded = st.file_uploader("upload", type=["csv","xlsx","xls"], label_visibility="collapsed")
        if uploaded:
            try:
                df_raw = (
                    pd.read_csv(uploaded, dtype=str) if uploaded.name.endswith(".csv")
                    else pd.read_excel(uploaded, dtype=str)
                ).fillna("")
                st.markdown(
                    f'<div class="toast-ok">✓ &nbsp;{uploaded.name} &nbsp;·&nbsp; {len(df_raw):,} rows</div>',
                    unsafe_allow_html=True,
                )
            except Exception as e:
                st.error(str(e))

# ══════════════════════════════════════════════════════════════════════════════
# RESULTS
# ══════════════════════════════════════════════════════════════════════════════
if df_raw is not None and len(df_raw) > 0:

    col_map  = _detect_columns(df_raw.columns)
    detected = {f: c for f, c in col_map.items() if c is not None}

    if not detected:
        st.markdown("""
        <div style="background:#fef2f2;border:1px solid #fecaca;border-radius:12px;
                    padding:1.1rem 1.4rem;margin-top:1.5rem">
            <b style="color:#dc2626">⚠ No address columns detected</b><br>
            <span style="color:#a1a1aa;font-size:.82rem">
            Ensure your data has a header row with names like Address, City, State, Country, Postal Code.
            </span>
        </div>""", unsafe_allow_html=True)
        st.stop()

    # ── AI street-name spell-check toggle ────────────────────────────────────
    use_ai_bulk = st.checkbox(
        "🤖 AI spell-check street names",
        value=False,
        key="bulk_ai_fix",
        help="Uses Claude AI to fix misspelled street name words row-by-row. "
             "Adds ~1–2 s per row — best for smaller files (<200 rows). "
             "Requires ANTHROPIC_API_KEY to be set.",
    )
    if use_ai_bulk and len(df_raw) > 200:
        st.warning(f"⚠ AI spell-check is enabled on {len(df_raw):,} rows. This may take several minutes.")

    # Corrections
    result = df_raw.copy()
    corrected_col_map = {}
    for field, orig_col in col_map.items():
        if orig_col is None: continue
        if field not in CORRECTORS: continue   # hint-only fields (e.g. company_name)
        lbl = f"Corrected {DISPLAY_LABELS[field]}"
        result[lbl] = df_raw[orig_col].apply(CORRECTORS[field])
        corrected_col_map[lbl] = lbl

    # Auto-fix country & state from postal code + company name hints
    apply_autofix(result, col_map)

    # AI spell-check street name columns (bulk)
    if use_ai_bulk:
        addr_fields = ["address_line_1", "address_line_2", "address_line_3"]
        addr_labels = [f"Corrected {DISPLAY_LABELS[f]}" for f in addr_fields
                       if col_map.get(f) and f"Corrected {DISPLAY_LABELS[f]}" in result.columns]
        if addr_labels:
            with st.spinner(f"Running AI street spell-check on {len(df_raw):,} rows…"):
                for lbl in addr_labels:
                    result[lbl] = result[lbl].apply(
                        lambda v: _ai_correct_street(str(v)) if pd.notna(v) and str(v).strip() else v
                    )

    # Stats
    per_field, total_changed = {}, 0
    for field, orig_col in col_map.items():
        if orig_col is None: continue
        if field not in DISPLAY_LABELS: continue   # hint-only fields
        lbl = f"Corrected {DISPLAY_LABELS[field]}"
        if lbl not in result.columns: continue
        n = (result[orig_col].astype(str).str.strip() != result[lbl].astype(str).str.strip()).sum()
        per_field[DISPLAY_LABELS[field]] = int(n)
        total_changed += n
    total_rows, fields_found = len(df_raw), len(detected)

    # ── Step 2 heading ─────────────────────────────────────────────────────
    st.markdown("""
    <div class="tool-wrap" style="padding-bottom:1.4rem">
        <div class="tool-label">Step 2</div>
        <div class="tool-heading">Your corrected addresses</div>
        <div class="tool-sub">Review what changed, then download.</div>
    </div>
    """, unsafe_allow_html=True)

    # Wrap the whole results block in tool-wrap padding
    st.markdown('<div style="max-width:1060px;margin:0 auto;padding:0 2rem 6rem">', unsafe_allow_html=True)
    st.markdown('<div class="results-card">', unsafe_allow_html=True)

    # Metrics
    mc = st.columns(4)
    for col, (val, lbl) in zip(mc, [
        (f"{total_rows:,}", "Rows Processed"),
        (f"{fields_found}", "Fields Detected"),
        (f"{total_changed:,}", "Cells Corrected"),
        (f"{total_changed/max(total_rows*fields_found,1):.0%}", "Correction Rate"),
    ]):
        col.markdown(f'<div class="mtile"><div class="mtile-val">{val}</div><div class="mtile-lbl">{lbl}</div></div>',
                     unsafe_allow_html=True)

    st.markdown("<div style='height:.8rem'></div>", unsafe_allow_html=True)

    # Column badges
    bdg = '<div class="badge-row">'
    for field, label in DISPLAY_LABELS.items():
        orig = col_map.get(field)
        if orig:
            bdg += f'<span class="bdg bdg-ok">✓ {label} <span class="bdg-src">← {orig}</span></span>'
        else:
            bdg += f'<span class="bdg bdg-no">— {label}</span>'
    bdg += '</div>'
    st.markdown(bdg, unsafe_allow_html=True)

    # Bar chart
    if any(v > 0 for v in per_field.values()):
        with st.expander("Breakdown by field"):
            st.bar_chart(pd.DataFrame.from_dict(per_field, orient="index", columns=["Corrections"]),
                         color="#6c47ff")

    # View toggle + table
    vl, vr = st.columns([4, 3])
    with vl:
        st.markdown('<p style="font-size:.72rem;font-weight:700;letter-spacing:.7px;'
                    'text-transform:uppercase;color:#a1a1aa;margin-bottom:.3rem">View</p>',
                    unsafe_allow_html=True)
    with vr:
        view_mode = st.radio("v", ["Side by side", "Corrected only", "Changes only"],
                             horizontal=True, label_visibility="collapsed")

    orig_addr = [c for c in df_raw.columns if c in col_map.values()]
    corr_cols = list(corrected_col_map.keys())
    other     = [c for c in df_raw.columns if c not in orig_addr]

    if view_mode == "Corrected only":
        display_df = pd.concat([df_raw[other], result[corr_cols]], axis=1)
    elif view_mode == "Changes only":
        mask = pd.Series(False, index=result.index)
        for field, orig_col in col_map.items():
            if orig_col is None or field not in DISPLAY_LABELS: continue
            lbl = f"Corrected {DISPLAY_LABELS[field]}"
            if lbl not in result.columns: continue
            mask |= result[orig_col].astype(str).str.strip() != result[lbl].astype(str).str.strip()
        display_df = result[mask]
    else:
        display_df = result

    def highlight_changes(df):
        s = pd.DataFrame("", index=df.index, columns=df.columns)
        for field, orig_col in col_map.items():
            if field not in DISPLAY_LABELS: continue
            lbl = f"Corrected {DISPLAY_LABELS[field]}"
            if orig_col not in df.columns or lbl not in df.columns: continue
            changed = df[orig_col].astype(str).str.strip() != df[lbl].astype(str).str.strip()
            s.loc[changed, lbl]      = "background-color:#f0fdf4;color:#15803d;font-weight:600"
            s.loc[changed, orig_col] = "background-color:#fefce8;color:#a16207"
        return s

    st.dataframe(display_df.style.apply(highlight_changes, axis=None),
                 use_container_width=True, height=min(44 + 36*len(display_df), 520))

    st.markdown(f"""
    <div class="legend">
        <span><span class="ldot" style="background:#dcfce7;border:1px solid #86efac"></span>Corrected value</span>
        <span><span class="ldot" style="background:#fef9c3;border:1px solid #fde047"></span>Original (changed)</span>
        <span style="margin-left:auto">{len(display_df):,} of {total_rows:,} rows shown</span>
    </div>
    <div style="height:1.5rem"></div>
    """, unsafe_allow_html=True)

    # Download
    st.markdown("""
    <p style="font-size:.72rem;font-weight:700;letter-spacing:.7px;text-transform:uppercase;
              color:#a1a1aa;margin-bottom:.7rem">Export</p>
    """, unsafe_allow_html=True)
    d1, d2, _ = st.columns([2, 2, 5])
    with d1:
        buf = io.BytesIO()
        _write_excel(result, list(df_raw.columns), corrected_col_map, buf, col_map)
        buf.seek(0)
        st.download_button("⬇  Excel (.xlsx)", buf, "corrected_addresses.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True, type="primary")
    with d2:
        st.download_button("⬇  CSV (.csv)", result.to_csv(index=False).encode("utf-8-sig"),
                           "corrected_addresses.csv", "text/csv", use_container_width=True)

    st.markdown("</div></div>", unsafe_allow_html=True)

else:
    # ── Empty state ────────────────────────────────────────────────────────
    st.markdown("""
    <div style="max-width:1060px;margin:2rem auto 6rem;padding:0 2rem">
        <div style="background:#fff;border:1px solid rgba(0,0,0,.08);border-radius:18px;
                    text-align:center;padding:4.5rem 2rem;box-shadow:0 2px 16px rgba(0,0,0,.05)">
            <div style="font-size:2.8rem;margin-bottom:1rem">📍</div>
            <div style="font-size:1.15rem;font-weight:800;color:#111118;
                        letter-spacing:-.5px;margin-bottom:.5rem">
                Results will appear here
            </div>
            <p style="color:#a1a1aa;font-size:.85rem;max-width:380px;
                      margin:0 auto 2rem;line-height:1.75">
                Paste your addresses or upload a file above to see corrections,
                metrics, and download your cleaned data.
            </p>
            <div style="display:flex;justify-content:center;gap:.55rem;flex-wrap:wrap">
                <span style="background:#f5f4f0;border:1px solid rgba(0,0,0,.07);border-radius:100px;
                             padding:.35rem 1rem;font-size:.74rem;color:#6c47ff;font-weight:600">
                    ✓ Auto-detects column names
                </span>
                <span style="background:#f5f4f0;border:1px solid rgba(0,0,0,.07);border-radius:100px;
                             padding:.35rem 1rem;font-size:.74rem;color:#6c47ff;font-weight:600">
                    ✓ 249 countries
                </span>
                <span style="background:#f5f4f0;border:1px solid rgba(0,0,0,.07);border-radius:100px;
                             padding:.35rem 1rem;font-size:.74rem;color:#6c47ff;font-weight:600">
                    ✓ Catches misspellings
                </span>
                <span style="background:#f5f4f0;border:1px solid rgba(0,0,0,.07);border-radius:100px;
                             padding:.35rem 1rem;font-size:.74rem;color:#6c47ff;font-weight:600">
                    ✓ Excel &amp; CSV export
                </span>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
