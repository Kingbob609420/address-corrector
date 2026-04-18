import io
import importlib
import pandas as pd
import streamlit as st

# Force fresh import every deploy — prevents Streamlit Cloud from serving stale module
import address_corrector as _ac
importlib.reload(_ac)
from address_corrector import (
    CORRECTORS, DISPLAY_LABELS, _detect_columns, _write_excel, apply_autofix,
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
    st.markdown('<div style="max-width:700px;margin:0 auto;padding:1.5rem 2rem 4rem">', unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        s_addr1   = st.text_input("Street Address",    placeholder="123 Main St",    key="s_addr1")
        s_city    = st.text_input("City",              placeholder="Fort Mill",       key="s_city")
        s_state   = st.text_input("State / Province",  placeholder="NC",             key="s_state")
    with c2:
        s_country = st.text_input("Country",           placeholder="United States",   key="s_country")
        s_zip     = st.text_input("ZIP / Postal Code", placeholder="29708",          key="s_zip")

    if any(x.strip() for x in [s_addr1, s_city, s_state, s_country, s_zip]):
        # Build one-row DataFrame and run the exact same correctors as bulk mode
        df_s = pd.DataFrame([{
            "Address":     s_addr1,
            "City":        s_city,
            "State":       s_state,
            "Country":     s_country,
            "Postal Code": s_zip,
        }])
        col_map_s = _detect_columns(df_s.columns)
        result_s  = df_s.copy()
        for field, orig_col in col_map_s.items():
            if orig_col is None or field not in CORRECTORS: continue
            lbl = f"Corrected {DISPLAY_LABELS[field]}"
            result_s[lbl] = df_s[orig_col].apply(CORRECTORS[field])
        apply_autofix(result_s, col_map_s)

        def _get_s(field):
            lbl  = f"Corrected {DISPLAY_LABELS[field]}"
            orig = col_map_s.get(field)
            if lbl in result_s.columns:
                v = str(result_s[lbl].iloc[0]).strip()
                return v if v else ""
            if orig and orig in result_s.columns:
                return str(result_s[orig].iloc[0]).strip()
            return ""

        addr1_c   = _get_s("address_line_1") or s_addr1.strip()
        city_c    = _get_s("city")
        state_c   = _get_s("state")
        zip_c     = _get_s("postal_code")
        country_c = _get_s("country")

        # Detect what changed to show the diff
        pairs = [
            ("Address", s_addr1.strip(),  addr1_c),
            ("City",    s_city.strip(),   city_c),
            ("State",   s_state.strip(),  state_c),
            ("ZIP",     s_zip.strip(),    zip_c),
            ("Country", s_country.strip(), country_c),
        ]
        changes_html = ""
        for label, orig, corr in pairs:
            if orig and corr and orig.lower() != corr.lower():
                changes_html += (
                    f'<div style="display:flex;gap:.5rem;align-items:center;'
                    f'font-size:.82rem;margin-bottom:.4rem">'
                    f'<span style="color:#a1a1aa;min-width:70px">{label}</span>'
                    f'<span style="color:#a16207;background:#fefce8;border:1px solid #fde068;'
                    f'border-radius:6px;padding:.15rem .5rem">{orig}</span>'
                    f'<span style="color:#a1a1aa">→</span>'
                    f'<span style="color:#15803d;background:#f0fdf4;border:1px solid #86efac;'
                    f'border-radius:6px;padding:.15rem .5rem;font-weight:600">{corr}</span>'
                    f'</div>'
                )

        # Format as mailing label block
        line2 = ", ".join(p for p in [city_c, state_c, zip_c] if p)
        addr_lines = [l for l in [addr1_c, line2, country_c] if l]
        addr_block = "<br>".join(addr_lines)

        changes_section = ""
        if changes_html:
            changes_section = f"""
            <div style="margin-top:1.2rem">
              <div style="font-size:.68rem;font-weight:700;letter-spacing:.7px;
                          text-transform:uppercase;color:#a1a1aa;margin-bottom:.6rem">
                Changes made
              </div>
              {changes_html}
            </div>"""

        st.markdown(f"""
        <div style="background:#fff;border:1px solid rgba(0,0,0,.08);border-radius:16px;
                    padding:1.8rem 2rem;margin-top:1.5rem;box-shadow:0 2px 16px rgba(0,0,0,.05)">
          <div style="font-size:.68rem;font-weight:700;letter-spacing:.7px;
                      text-transform:uppercase;color:#a1a1aa;margin-bottom:.9rem">
            Corrected address
          </div>
          <div style="font-size:1.05rem;line-height:1.9;color:#111118;
                      background:#f5f4f0;border-radius:10px;padding:1rem 1.2rem">
            {addr_block}
          </div>
          {changes_section}
        </div>
        """, unsafe_allow_html=True)
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

