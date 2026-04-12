import io

import pandas as pd
import streamlit as st

from address_corrector import (
    CORRECTORS,
    DISPLAY_LABELS,
    _detect_columns,
    _write_excel,
)

st.set_page_config(
    page_title="Address Corrector",
    page_icon="📍",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── Global CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* ── Page ── */
[data-testid="stAppViewContainer"] {
    background: #f7f8fc;
}
.block-container {
    padding: 2.5rem 3rem 3rem 3rem;
    max-width: 1400px;
}

/* ── Hero header ── */
.hero {
    background: linear-gradient(135deg, #1a1f36 0%, #2d3561 100%);
    border-radius: 16px;
    padding: 2.4rem 2.8rem;
    margin-bottom: 2rem;
    display: flex;
    align-items: center;
    gap: 1.2rem;
}
.hero-icon { font-size: 2.8rem; line-height: 1; }
.hero h1 {
    margin: 0 0 .3rem 0;
    font-size: 2rem;
    font-weight: 700;
    color: #ffffff;
    letter-spacing: -.5px;
}
.hero p {
    margin: 0;
    color: #a0aec0;
    font-size: .95rem;
}

/* ── Section cards ── */
.card {
    background: #ffffff;
    border-radius: 14px;
    border: 1px solid #e8eaf0;
    padding: 1.6rem 1.8rem;
    margin-bottom: 1.4rem;
    box-shadow: 0 1px 4px rgba(0,0,0,.05);
}
.section-label {
    font-size: .7rem;
    font-weight: 700;
    letter-spacing: 1.2px;
    text-transform: uppercase;
    color: #8892a4;
    margin-bottom: .7rem;
}
.section-title {
    font-size: 1.1rem;
    font-weight: 700;
    color: #1a1f36;
    margin-bottom: 1rem;
}

/* ── Metric tiles ── */
.metric-row { display: flex; gap: 1rem; margin-bottom: .5rem; }
.metric-tile {
    flex: 1;
    background: #f7f8fc;
    border: 1px solid #e8eaf0;
    border-radius: 12px;
    padding: 1.1rem 1.3rem;
    text-align: center;
}
.metric-value {
    font-size: 2rem;
    font-weight: 800;
    color: #2d3561;
    line-height: 1.1;
}
.metric-label {
    font-size: .78rem;
    color: #8892a4;
    font-weight: 600;
    letter-spacing: .4px;
    margin-top: .25rem;
    text-transform: uppercase;
}

/* ── Column detection badges ── */
.badge-row { display: flex; flex-wrap: wrap; gap: .6rem; margin-top: .2rem; }
.badge {
    display: inline-flex;
    flex-direction: column;
    padding: .55rem .9rem;
    border-radius: 10px;
    font-size: .8rem;
    min-width: 120px;
}
.badge-found {
    background: #e8f5e9;
    border: 1px solid #a5d6a7;
    color: #2e7d32;
}
.badge-missing {
    background: #fff8e1;
    border: 1px solid #ffe082;
    color: #f57f17;
}
.badge-field { font-weight: 700; font-size: .78rem; margin-bottom: .15rem; }
.badge-col   { font-size: .72rem; opacity: .8; font-style: italic; }

/* ── Tab styling ── */
[data-testid="stTabs"] button {
    font-weight: 600;
    font-size: .88rem;
    padding: .5rem 1.2rem;
}
[data-testid="stTabs"] button[aria-selected="true"] {
    color: #2d3561;
    border-bottom-color: #2d3561 !important;
}

/* ── Text area ── */
textarea {
    font-family: 'SF Mono', 'Fira Code', monospace !important;
    font-size: .82rem !important;
    background: #f7f8fc !important;
    border-radius: 10px !important;
}

/* ── Radio pills ── */
[data-testid="stRadio"] > div {
    display: flex;
    gap: .5rem;
    flex-wrap: wrap;
}
[data-testid="stRadio"] label {
    background: #f0f2f8;
    border: 1px solid #dde1ef;
    border-radius: 20px;
    padding: .3rem .9rem;
    font-size: .83rem;
    font-weight: 500;
    cursor: pointer;
    transition: all .15s;
}
[data-testid="stRadio"] label:hover { background: #e5e9f7; }

/* ── Download buttons ── */
[data-testid="stDownloadButton"] button {
    border-radius: 10px !important;
    font-weight: 600 !important;
    font-size: .88rem !important;
    padding: .6rem 1.2rem !important;
    transition: transform .1s !important;
}
[data-testid="stDownloadButton"] button:hover { transform: translateY(-1px); }

/* ── Dataframe ── */
[data-testid="stDataFrame"] {
    border-radius: 12px;
    overflow: hidden;
    border: 1px solid #e8eaf0 !important;
}

/* ── Legend pills ── */
.legend { display: flex; gap: .8rem; align-items: center; margin-top: .6rem; }
.legend-dot {
    width: 12px; height: 12px;
    border-radius: 3px;
    display: inline-block;
    margin-right: .3rem;
}
.legend-item { font-size: .78rem; color: #8892a4; display: flex; align-items: center; }

/* ── Info banner ── */
.info-banner {
    background: #eef2ff;
    border: 1px solid #c7d2fe;
    border-radius: 12px;
    padding: 1.2rem 1.5rem;
    color: #3730a3;
    font-size: .9rem;
    display: flex;
    align-items: center;
    gap: .8rem;
}

/* ── File uploader ── */
[data-testid="stFileUploader"] {
    background: #f7f8fc;
    border-radius: 12px;
    border: 2px dashed #dde1ef;
    padding: .5rem;
}

/* Hide Streamlit branding */
#MainMenu, footer, header { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ── Hero ──────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero">
    <div class="hero-icon">📍</div>
    <div>
        <h1>Address Corrector</h1>
        <p>Paste from Excel or upload a file &mdash; get standardised, corrected addresses instantly across all 249 countries.</p>
    </div>
</div>
""", unsafe_allow_html=True)

# ── Input card ────────────────────────────────────────────────────────────────
st.markdown('<div class="card">', unsafe_allow_html=True)
st.markdown('<div class="section-label">Input</div>', unsafe_allow_html=True)

tab_paste, tab_upload = st.tabs(["  📋  Paste from Excel", "  📁  Upload file"])

df_raw = None

with tab_paste:
    st.markdown(
        "<p style='color:#8892a4;font-size:.85rem;margin:.4rem 0 .8rem'>Copy cells directly from Excel — include the header row. "
        "Tab, comma and semicolon delimiters are all detected automatically.</p>",
        unsafe_allow_html=True,
    )
    pasted = st.text_area(
        "paste_area",
        height=220,
        label_visibility="collapsed",
        placeholder=(
            "Address Line 1\tAddress Line 2\tCity\tState\tCountry\tPostal Code\n"
            "123 main st\tapt 4b\tnew york\tnew york\tUSA\t10001\n"
            "456 oak ave\tsuite 200\tlos angeles\tCA\tus\t90210"
        ),
    )
    if pasted.strip():
        sample = pasted.split("\n")[0]
        tabs, commas, semis = sample.count("\t"), sample.count(","), sample.count(";")
        delim = "\t" if tabs >= commas and tabs >= semis else (";" if semis > commas else ",")
        try:
            df_raw = pd.read_csv(
                io.StringIO(pasted), sep=delim, dtype=str, on_bad_lines="skip"
            ).fillna("")
        except Exception as e:
            st.error(f"Could not parse pasted text: {e}")

with tab_upload:
    st.markdown("<div style='height:.4rem'></div>", unsafe_allow_html=True)
    uploaded = st.file_uploader(
        "upload",
        type=["csv", "xlsx", "xls"],
        label_visibility="collapsed",
    )
    if uploaded:
        try:
            df_raw = (
                pd.read_csv(uploaded, dtype=str)
                if uploaded.name.endswith(".csv")
                else pd.read_excel(uploaded, dtype=str)
            ).fillna("")
        except Exception as e:
            st.error(f"Could not read file: {e}")

st.markdown("</div>", unsafe_allow_html=True)

# ── Processing ────────────────────────────────────────────────────────────────
if df_raw is not None and len(df_raw) > 0:

    col_map  = _detect_columns(df_raw.columns)
    detected = {f: c for f, c in col_map.items() if c is not None}

    if not detected:
        st.markdown("""
        <div class="card">
            <div style="color:#c62828;font-weight:600">⚠️ No address columns detected</div>
            <p style="color:#8892a4;font-size:.85rem;margin-top:.4rem">
            Make sure your data includes a header row with column names like
            "Address", "City", "State", "Country", or "Postal Code".</p>
        </div>""", unsafe_allow_html=True)
        st.stop()

    # Apply corrections
    result = df_raw.copy()
    corrected_col_map = {}
    for field, orig_col in col_map.items():
        if orig_col is None:
            continue
        lbl = f"Corrected {DISPLAY_LABELS[field]}"
        result[lbl] = df_raw[orig_col].apply(CORRECTORS[field])
        corrected_col_map[lbl] = lbl

    # Compute stats
    total_rows   = len(df_raw)
    fields_found = len(detected)
    per_field    = {}
    total_changed = 0
    for field, orig_col in col_map.items():
        if orig_col is None:
            continue
        lbl = f"Corrected {DISPLAY_LABELS[field]}"
        n = (result[orig_col].astype(str).str.strip() != result[lbl].astype(str).str.strip()).sum()
        per_field[DISPLAY_LABELS[field]] = int(n)
        total_changed += n

    correction_rate = total_changed / max(total_rows * fields_found, 1)

    # ── Column detection card ──────────────────────────────────────────────
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="section-label">Detected columns</div>', unsafe_allow_html=True)

    badges_html = '<div class="badge-row">'
    for field, label in DISPLAY_LABELS.items():
        orig_col = col_map.get(field)
        if orig_col:
            badges_html += (
                f'<div class="badge badge-found">'
                f'<span class="badge-field">✓ {label}</span>'
                f'<span class="badge-col">{orig_col}</span>'
                f'</div>'
            )
        else:
            badges_html += (
                f'<div class="badge badge-missing">'
                f'<span class="badge-field">— {label}</span>'
                f'<span class="badge-col">not found</span>'
                f'</div>'
            )
    badges_html += '</div>'
    st.markdown(badges_html, unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    # ── Metrics card ──────────────────────────────────────────────────────
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="section-label">Summary</div>', unsafe_allow_html=True)

    m_cols = st.columns(4)
    tiles = [
        (f"{total_rows:,}",       "Rows Processed"),
        (f"{fields_found}",       "Fields Detected"),
        (f"{total_changed:,}",    "Cells Corrected"),
        (f"{correction_rate:.0%}","Correction Rate"),
    ]
    for col, (val, lbl) in zip(m_cols, tiles):
        col.markdown(f"""
        <div class="metric-tile">
            <div class="metric-value">{val}</div>
            <div class="metric-label">{lbl}</div>
        </div>""", unsafe_allow_html=True)

    # Per-field bar chart
    if any(v > 0 for v in per_field.values()):
        st.markdown("<div style='height:.8rem'></div>", unsafe_allow_html=True)
        with st.expander("View corrections by field"):
            bar_df = pd.DataFrame.from_dict(per_field, orient="index", columns=["Corrections"])
            st.bar_chart(bar_df, color="#2d3561")

    st.markdown("</div>", unsafe_allow_html=True)

    # ── Results card ──────────────────────────────────────────────────────
    st.markdown('<div class="card">', unsafe_allow_html=True)

    top_left, top_right = st.columns([3, 2])
    with top_left:
        st.markdown('<div class="section-label">Results</div>', unsafe_allow_html=True)
    with top_right:
        view_mode = st.radio(
            "view",
            ["Side by side", "Corrected only", "Changes only"],
            horizontal=True,
            label_visibility="collapsed",
        )

    orig_addr_cols = [c for c in df_raw.columns if c in col_map.values()]
    corr_addr_cols = list(corrected_col_map.keys())
    other_cols     = [c for c in df_raw.columns if c not in orig_addr_cols]

    if view_mode == "Corrected only":
        display_df = pd.concat([df_raw[other_cols], result[corr_addr_cols]], axis=1)
    elif view_mode == "Changes only":
        mask = pd.Series(False, index=result.index)
        for field, orig_col in col_map.items():
            if orig_col is None:
                continue
            lbl = f"Corrected {DISPLAY_LABELS[field]}"
            mask |= result[orig_col].astype(str).str.strip() != result[lbl].astype(str).str.strip()
        display_df = result[mask]
    else:
        display_df = result

    def highlight_changes(df):
        styles = pd.DataFrame("", index=df.index, columns=df.columns)
        for field, orig_col in col_map.items():
            lbl = f"Corrected {DISPLAY_LABELS[field]}"
            if orig_col not in df.columns or lbl not in df.columns:
                continue
            changed = df[orig_col].astype(str).str.strip() != df[lbl].astype(str).str.strip()
            styles.loc[changed, lbl]      = "background-color: #e8f5e9; font-weight: 600; color: #2e7d32"
            styles.loc[changed, orig_col] = "background-color: #fff8e1; color: #e65100"
        return styles

    styled = display_df.style.apply(highlight_changes, axis=None)
    st.dataframe(styled, use_container_width=True, height=min(42 + 36 * len(display_df), 580))

    st.markdown("""
    <div class="legend">
        <div class="legend-item"><span class="legend-dot" style="background:#c8e6c9;border:1px solid #a5d6a7"></span>Corrected value</div>
        <div class="legend-item"><span class="legend-dot" style="background:#fff9c4;border:1px solid #ffe082"></span>Original (changed)</div>
        <div class="legend-item" style="margin-left:auto;color:#b0bac8">
            Showing <b style="color:#2d3561">{shown:,}</b> of <b style="color:#2d3561">{total:,}</b> rows
        </div>
    </div>
    """.format(shown=len(display_df), total=total_rows), unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)

    # ── Download card ─────────────────────────────────────────────────────
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="section-label">Download</div>', unsafe_allow_html=True)
    st.markdown(
        "<p style='color:#8892a4;font-size:.85rem;margin:0 0 1rem'>Excel includes colour-coded headers and highlights. "
        "CSV is UTF-8 encoded for universal compatibility.</p>",
        unsafe_allow_html=True,
    )

    dl1, dl2, _ = st.columns([1, 1, 2])
    with dl1:
        buf_xlsx = io.BytesIO()
        _write_excel(result, list(df_raw.columns), corrected_col_map, buf_xlsx, col_map)
        buf_xlsx.seek(0)
        st.download_button(
            label="⬇  Download Excel",
            data=buf_xlsx,
            file_name="corrected_addresses.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )
    with dl2:
        st.download_button(
            label="⬇  Download CSV",
            data=result.to_csv(index=False).encode("utf-8-sig"),
            file_name="corrected_addresses.csv",
            mime="text/csv",
            use_container_width=True,
        )

    st.markdown("</div>", unsafe_allow_html=True)

else:
    # ── Empty state ───────────────────────────────────────────────────────
    st.markdown("""
    <div class="card" style="text-align:center;padding:3rem 2rem">
        <div style="font-size:3rem;margin-bottom:1rem">📋</div>
        <div style="font-size:1.1rem;font-weight:700;color:#1a1f36;margin-bottom:.5rem">
            Paste or upload your addresses to get started
        </div>
        <div style="color:#8892a4;font-size:.9rem;max-width:480px;margin:0 auto">
            Supports any column naming convention &mdash; Address Line 1/2/3, City, State,
            Country and Postal Code are all auto-detected. Covers all 249 countries.
        </div>
        <div style="margin-top:1.8rem;display:flex;justify-content:center;gap:2rem;flex-wrap:wrap">
            <div style="background:#f0f2f8;border-radius:10px;padding:.8rem 1.4rem;font-size:.82rem;color:#5c6bc0;font-weight:600">
                ✓  Auto-detects column names
            </div>
            <div style="background:#f0f2f8;border-radius:10px;padding:.8rem 1.4rem;font-size:.82rem;color:#5c6bc0;font-weight:600">
                ✓  249 countries supported
            </div>
            <div style="background:#f0f2f8;border-radius:10px;padding:.8rem 1.4rem;font-size:.82rem;color:#5c6bc0;font-weight:600">
                ✓  Fixes misspellings too
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
