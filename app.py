import io
import importlib
import pandas as pd
import streamlit as st

# Force fresh import every deploy — prevents Streamlit Cloud from serving stale module
import address_corrector as _ac
importlib.reload(_ac)
from address_corrector import (
    CORRECTORS, DISPLAY_LABELS, _detect_columns, _write_excel, apply_autofix,
    score_single_address, ai_enhance_address, validate_address_nominatim,
    infer_us_state_from_zip, infer_state_from_city,
    lookup_postal_from_address, correct_postal_code,
)


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
# API KEY — read from Streamlit secrets (set in app dashboard or secrets.toml)
# ══════════════════════════════════════════════════════════════════════════════
_openai_key = st.secrets.get("OPENAI_API_KEY", "")

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
    <div class="tool-heading">Correct your addresses</div>
    <div class="tool-sub">Fix a single address instantly, or process hundreds at once.</div>
</div>
""", unsafe_allow_html=True)

tab_single, tab_bulk = st.tabs(["Single Address", "Bulk Upload"])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — SINGLE ADDRESS
# ══════════════════════════════════════════════════════════════════════════════
with tab_single:
    import hashlib as _hashlib

    st.markdown('<div style="max-width:700px;margin:0 auto;padding:1.5rem 2rem 4rem">', unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        s_addr1   = st.text_input("Street Address",    placeholder="123 Main St",  key="s_addr1")
        s_addr2   = st.text_input("Address Line 2",    placeholder="APT 4B",       key="s_addr2")
        s_city    = st.text_input("City",              placeholder="Fort Mill",     key="s_city")
    with c2:
        s_state   = st.text_input("State / Province",  placeholder="NC",           key="s_state")
        s_country = st.text_input("Country",           placeholder="United States", key="s_country")
        s_zip     = st.text_input("ZIP / Postal Code", placeholder="29708",        key="s_zip")

    if any(x.strip() for x in [s_addr1, s_addr2, s_city, s_state, s_country, s_zip]):

        # Signature for this raw input — used to key the validation override
        _addr_sig = _hashlib.md5(
            f"{s_addr1}|{s_addr2}|{s_city}|{s_state}|{s_country}|{s_zip}".encode()
        ).hexdigest()
        # Validated fields from a previous "Validate on Map" click (if any)
        _val = st.session_state.get(f"_val_{_addr_sig}", {})

        # ── 1. Rule-based corrections ─────────────────────────────────────────
        scores    = score_single_address(s_addr1, s_addr2, s_city, s_state, s_country, s_zip)
        addr1_c   = scores["address"]["corrected"]  or s_addr1.strip()
        addr2_c   = scores["address2"]["corrected"] or s_addr2.strip()
        city_c    = scores["city"]["corrected"]     or s_city.strip()
        zip_c     = scores["postal"]["corrected"]   or s_zip.strip()
        state_c   = scores["state"]["corrected"]    or s_state.strip()
        country_c = scores["country"]["corrected"]  or s_country.strip()
        state_conf,   state_method   = scores["state"]["confidence"],   scores["state"]["method"]
        country_conf, country_method = scores["country"]["confidence"], scores["country"]["method"]
        zip_conf,     zip_method     = scores["postal"]["confidence"],  scores["postal"]["method"]
        _no_zip = not s_zip.strip()

        # ── 2. AI: correct spelling/format of the address (no ZIP inference yet) ─
        ai_result = None
        if _openai_key:
            if "_ai_cache" not in st.session_state:
                st.session_state["_ai_cache"] = {}
            _sig  = f"{s_addr1}|{s_addr2}|{s_city}|{s_state}|{s_country}|{s_zip}"
            _hash = _hashlib.md5(_sig.encode()).hexdigest()
            cache = st.session_state["_ai_cache"]
            if _hash not in cache:
                with st.spinner("Correcting address…"):
                    try:
                        cache[_hash] = ai_enhance_address(
                            s_addr1, s_addr2, s_city, s_state, s_country, s_zip,
                            _openai_key, infer_postal=False,
                        )
                    except Exception as _exc:
                        cache[_hash] = {"error": str(_exc)}
                if len(cache) >= 30:
                    for _k in list(cache.keys())[:10]:
                        del cache[_k]
            ai_result = cache[_hash]

        def _ai_val(v, fallback):
            clean = str(v or "").strip()
            return clean if clean and clean.lower() not in {"(blank)", "blank", "none", "nan", ""} else fallback

        if ai_result and "error" not in ai_result:
            addr1_c   = _ai_val(ai_result.get("address"),  addr1_c)
            addr2_c   = _ai_val(ai_result.get("address2"), addr2_c)
            city_c    = _ai_val(ai_result.get("city"),     city_c)
            state_c   = _ai_val(ai_result.get("state"),    state_c)
            country_c = _ai_val(ai_result.get("country"),  country_c)
            if not _no_zip:   # only reformat postal when user supplied one
                _ai_zip = _ai_val(ai_result.get("postal"), "")
                if _ai_zip:
                    zip_c = _ai_zip

        # ── 3. Authoritative ZIP → state when ZIP was already supplied ────────
        zip_derived_state = infer_us_state_from_zip(s_zip.strip()) or infer_us_state_from_zip(zip_c)
        if zip_derived_state:
            state_c       = zip_derived_state
            country_c     = "US"
            state_conf,   state_method   = 1.0, "zip_derived"
            country_conf, country_method = 1.0, "zip_derived"

        # ── 4. No ZIP supplied — ask ChatGPT with corrected fields ────────────
        elif _no_zip and _openai_key:
            # 4a. seed state/country from city so AI has best inputs
            city_state, city_country = infer_state_from_city(city_c, country_c)
            if city_state and not state_c:
                state_c     = city_state
                state_conf, state_method = 0.80, "city_derived"
            if city_country and not country_c:
                country_c   = city_country
                country_conf, country_method = 0.80, "city_derived"

            # 4b. ChatGPT ZIP lookup using the CORRECTED address fields
            _zip_ai_sig  = f"zipai4|{addr1_c}|{city_c}|{state_c}|{country_c}"
            _zip_ai_hash = _hashlib.md5(_zip_ai_sig.encode()).hexdigest()
            if "_zip_cache" not in st.session_state:
                st.session_state["_zip_cache"] = {}
            _zip_cache = st.session_state["_zip_cache"]
            if _zip_ai_hash not in _zip_cache:
                with st.spinner("Asking ChatGPT for ZIP code…"):
                    try:
                        _zr = ai_enhance_address(
                            addr1_c, addr2_c, city_c, state_c, country_c, "",
                            _openai_key, infer_postal=True,
                        )
                        _zip_cache[_zip_ai_hash] = _ai_val(_zr.get("postal"), "")
                    except Exception:
                        _zip_cache[_zip_ai_hash] = ""
            _ai_zip = _zip_cache.get(_zip_ai_hash, "")
            if _ai_zip:
                zip_c      = correct_postal_code(_ai_zip)
                zip_conf   = 0.88
                zip_method = "ai_inferred"
                # Re-pin state from ZIP if US
                _zdstate = infer_us_state_from_zip(zip_c)
                if _zdstate:
                    state_c       = _zdstate
                    country_c     = "US"
                    state_conf,   state_method   = 1.0, "zip_derived"
                    country_conf, country_method = 1.0, "zip_derived"
                    zip_conf,     zip_method     = 1.0, "zip_derived"

        # ── 5. Apply map-validation override (from a previous validate click) ─
        # When the user validates on the map, Nominatim's confirmed fields
        # replace the corrected fields so the output matches the map exactly.
        if _val.get("valid"):
            if _val.get("matched_city"):     city_c    = _val["matched_city"]
            if _val.get("matched_state"):    state_c   = _val["matched_state"]
            if _val.get("matched_postcode"): zip_c     = _val["matched_postcode"]
            zip_conf,     zip_method     = 1.0, "map_validated"
            state_conf,   state_method   = 1.0, "map_validated"
            country_conf, country_method = 1.0, "map_validated"

        # ── 5. Confidence badge helper ────────────────────────────────────────
        def _badge(conf, method=""):
            if method == "zip_derived":
                return ('<span style="background:#eff6ff;color:#1d4ed8;border:1px solid #bfdbfe;'
                        'border-radius:100px;padding:.12rem .55rem;font-size:.67rem;font-weight:700">'
                        '📌 ZIP</span>')
            if method == "city_derived":
                return ('<span style="background:#fdf4ff;color:#7c3aed;border:1px solid #e9d5ff;'
                        'border-radius:100px;padding:.12rem .55rem;font-size:.67rem;font-weight:700">'
                        '🏙 City</span>')
            if method == "ai_inferred":
                return ('<span style="background:#faf5ff;color:#7c3aed;border:1px solid #e9d5ff;'
                        'border-radius:100px;padding:.12rem .55rem;font-size:.67rem;font-weight:700">'
                        '✨ AI</span>')
            if method == "nominatim":
                return ('<span style="background:#f0fdf4;color:#15803d;border:1px solid #bbf7d0;'
                        'border-radius:100px;padding:.12rem .55rem;font-size:.67rem;font-weight:700">'
                        '🗺 OSM</span>')
            if method == "map_validated":
                return ('<span style="background:#f0fdf4;color:#15803d;border:1px solid #86efac;'
                        'border-radius:100px;padding:.12rem .55rem;font-size:.67rem;font-weight:700">'
                        '✅ Map</span>')
            pct = f"{conf:.0%}"
            if conf >= 0.85:
                return (f'<span style="background:#dcfce7;color:#15803d;border:1px solid #86efac;'
                        f'border-radius:100px;padding:.12rem .55rem;font-size:.67rem;font-weight:700">'
                        f'● {pct}</span>')
            elif conf >= 0.65:
                return (f'<span style="background:#fefce8;color:#a16207;border:1px solid #fde068;'
                        f'border-radius:100px;padding:.12rem .55rem;font-size:.67rem;font-weight:700">'
                        f'● {pct}</span>')
            return (f'<span style="background:#fef2f2;color:#dc2626;border:1px solid #fecaca;'
                    f'border-radius:100px;padding:.12rem .55rem;font-size:.67rem;font-weight:700">'
                    f'● {pct}</span>')

        # ── 6. Build field rows ───────────────────────────────────────────────
        field_rows = [
            ("Address", s_addr1.strip(),   addr1_c,   scores["address"]["confidence"],  scores["address"]["method"]),
            ("Addr 2",  s_addr2.strip(),   addr2_c,   scores["address2"]["confidence"], scores["address2"]["method"]),
            ("City",    s_city.strip(),    city_c,    scores["city"]["confidence"],      scores["city"]["method"]),
            ("State",   s_state.strip(),   state_c,   state_conf,                       state_method),
            ("ZIP",     s_zip.strip(),     zip_c,     zip_conf,                         zip_method),
            ("Country", s_country.strip(), country_c, country_conf,                     country_method),
        ]

        conf_html = changes_html = ""
        for label, orig, corr, conf, method in field_rows:
            display      = corr or orig
            is_empty     = not display
            display_html = (f'<span style="color:#d4d4d8">—</span>' if is_empty
                            else f'<span style="color:#111118">{display}</span>')
            badge_html   = _badge(conf, method) if display else ""
            conf_html += (
                f'<div style="display:flex;gap:.5rem;align-items:center;'
                f'font-size:.82rem;margin-bottom:.35rem">'
                f'<span style="color:#a1a1aa;min-width:72px">{label}</span>'
                f'<span style="flex:1">{display_html}</span>'
                f'{badge_html}'
                f'</div>'
            )
            if orig and orig.upper() != corr.upper():
                changes_html += (
                    f'<div style="display:flex;gap:.5rem;align-items:center;'
                    f'font-size:.82rem;margin-bottom:.4rem">'
                    f'<span style="color:#a1a1aa;min-width:72px">{label}</span>'
                    f'<span style="color:#a16207;background:#fefce8;border:1px solid #fde068;'
                    f'border-radius:6px;padding:.15rem .5rem">{orig}</span>'
                    f'<span style="color:#a1a1aa">→</span>'
                    f'<span style="color:#15803d;background:#f0fdf4;border:1px solid #86efac;'
                    f'border-radius:6px;padding:.15rem .5rem;font-weight:600">{corr}</span>'
                    f'</div>'
                )
            elif not orig and corr:
                changes_html += (
                    f'<div style="display:flex;gap:.5rem;align-items:center;'
                    f'font-size:.82rem;margin-bottom:.4rem">'
                    f'<span style="color:#a1a1aa;min-width:72px">{label}</span>'
                    f'<span style="color:#15803d;background:#f0fdf4;border:1px solid #86efac;'
                    f'border-radius:6px;padding:.15rem .5rem;font-weight:600">'
                    f'✦ {corr} <span style="font-weight:400;opacity:.7">(auto-filled)</span>'
                    f'</span></div>'
                )

        # ── 7. Render result card (all 6 fields always shown) ─────────────────
        def _fval(v): return v if v.strip() else '<span style="color:#d4d4d8">—</span>'
        addr_block = (
            f'<div style="display:grid;grid-template-columns:auto 1fr;gap:.25rem .9rem;'
            f'font-size:.9rem;line-height:1.75">'
            f'<span style="color:#a1a1aa">Address</span><span>{_fval(addr1_c)}</span>'
            + (f'<span style="color:#a1a1aa">Addr 2</span><span>{_fval(addr2_c)}</span>' if addr2_c.strip() else '')
            + f'<span style="color:#a1a1aa">City</span><span>{_fval(city_c)}</span>'
            f'<span style="color:#a1a1aa">State</span><span>{_fval(state_c)}</span>'
            f'<span style="color:#a1a1aa">ZIP</span><span>{_fval(zip_c)}</span>'
            f'<span style="color:#a1a1aa">Country</span><span>{_fval(country_c)}</span>'
            f'</div>'
        )

        ai_badge = ('<span style="background:#faf5ff;color:#7c3aed;border:1px solid #e9d5ff;'
                    'border-radius:100px;padding:.12rem .55rem;font-size:.67rem;font-weight:700;'
                    'margin-left:.4rem">✨ AI</span>' if (ai_result and "error" not in ai_result) else "")

        changes_section = (
            f'<div style="margin-top:1.2rem;padding-top:1.2rem;border-top:1px solid rgba(0,0,0,.06)">'
            f'<div style="font-size:.68rem;font-weight:700;letter-spacing:.7px;text-transform:uppercase;'
            f'color:#a1a1aa;margin-bottom:.6rem">Changes &amp; auto-filled fields</div>'
            f'{changes_html}</div>'
        ) if changes_html else ""

        if ai_result and "error" in ai_result:
            st.warning(f"AI correction failed: {ai_result['error']}")

        ai_note_html = ""
        if ai_result and "error" not in ai_result and ai_result.get("note"):
            ai_note_html = (f'<div style="margin-top:.8rem;padding-top:.8rem;border-top:1px solid rgba(0,0,0,.06);'
                            f'color:#7c3aed;font-size:.8rem">✨ {ai_result["note"]}</div>')

        st.markdown(f"""
        <div style="background:#fff;border:1px solid rgba(0,0,0,.08);border-radius:16px;
                    padding:1.8rem 2rem;margin-top:1.5rem;box-shadow:0 2px 16px rgba(0,0,0,.05)">
          <div style="font-size:.68rem;font-weight:700;letter-spacing:.7px;
                      text-transform:uppercase;color:#a1a1aa;margin-bottom:.9rem">
            Corrected address{ai_badge}
          </div>
          <div style="background:#f5f4f0;border-radius:10px;padding:1rem 1.2rem;margin-bottom:1.4rem">
            {addr_block}
          </div>
          <div style="font-size:.68rem;font-weight:700;letter-spacing:.7px;
                      text-transform:uppercase;color:#a1a1aa;margin-bottom:.6rem">Field confidence</div>
          {conf_html}
          {changes_section}
          {ai_note_html}
        </div>
        """, unsafe_allow_html=True)

        # ── 7. Validate on Map button ─────────────────────────────────────────
        st.markdown("<div style='height:.8rem'></div>", unsafe_allow_html=True)
        vbtn_col, _ = st.columns([2, 7])
        with vbtn_col:
            validate_clicked = st.button("🗺 Validate on Map", use_container_width=True, key="validate_btn")

        if validate_clicked:
            with st.spinner("Validating address on OpenStreetMap…"):
                vr = validate_address_nominatim(addr1_c, city_c, state_c, country_c, zip_c)
            st.session_state[f"_val_{_addr_sig}"] = vr
            st.rerun()

        # Show map result card whenever a stored validation result exists
        if _val.get("valid"):
            maps_url = f"https://www.openstreetmap.org/?mlat={_val['lat']}&mlon={_val['lon']}&zoom=16"
            st.markdown(
                f'<div style="background:#f0fdf4;border:1px solid #bbf7d0;border-radius:10px;'
                f'padding:.9rem 1.1rem;font-size:.84rem;margin-top:.6rem">'
                f'<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:.5rem">'
                f'<span style="color:#15803d;font-weight:700">✅ Validated — output updated to map result</span>'
                f'<a href="{maps_url}" target="_blank" style="color:#15803d;font-size:.78rem;'
                f'text-decoration:underline">View on map ↗</a></div>'
                f'<div style="color:#374151;font-size:.82rem;line-height:1.7">{_val["display_name"]}</div>'
                f'</div>',
                unsafe_allow_html=True,
            )
        elif _val.get("valid") is False:
            st.markdown(
                f'<div style="background:#fef2f2;border:1px solid #fecaca;border-radius:10px;'
                f'padding:.85rem 1.1rem;font-size:.84rem;margin-top:.6rem">'
                f'<span style="color:#dc2626;font-weight:700">⚠ {_val.get("message","Not found")}</span></div>',
                unsafe_allow_html=True,
            )

    else:
        st.markdown("""
        <div style="text-align:center;color:#a1a1aa;font-size:.88rem;padding:3rem 0">
            Enter an address above to see corrections instantly.
        </div>
        """, unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — BULK
# ══════════════════════════════════════════════════════════════════════════════
with tab_bulk:
    df_raw = None

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

    # ── RESULTS ────────────────────────────────────────────────────────────
    if df_raw is None or len(df_raw) == 0:
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
    else:
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
        else:
            # Corrections
            result = df_raw.copy()
            corrected_col_map = {}
            for field, orig_col in col_map.items():
                if orig_col is None: continue
                if field not in CORRECTORS: continue
                lbl = f"Corrected {DISPLAY_LABELS[field]}"
                result[lbl] = df_raw[orig_col].apply(CORRECTORS[field])
                corrected_col_map[lbl] = lbl

            apply_autofix(result, col_map)

            # Stats
            per_field, total_changed = {}, 0
            for field, orig_col in col_map.items():
                if orig_col is None: continue
                if field not in DISPLAY_LABELS: continue
                lbl = f"Corrected {DISPLAY_LABELS[field]}"
                if lbl not in result.columns: continue
                n = (result[orig_col].astype(str).str.strip() != result[lbl].astype(str).str.strip()).sum()
                per_field[DISPLAY_LABELS[field]] = int(n)
                total_changed += n
            total_rows, fields_found = len(df_raw), len(detected)

            st.markdown("""
            <div class="tool-wrap" style="padding-bottom:1.4rem">
                <div class="tool-label">Step 2</div>
                <div class="tool-heading">Your corrected addresses</div>
                <div class="tool-sub">Review what changed, then download.</div>
            </div>
            """, unsafe_allow_html=True)

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

