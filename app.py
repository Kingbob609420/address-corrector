import io
import re

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
)

# ── Styles ────────────────────────────────────────────────────────────────────
st.markdown(
    """
    <style>
    .block-container { padding-top: 2rem; }
    .stDataFrame { font-size: 13px; }
    div[data-testid="stMetricValue"] { font-size: 1.6rem; font-weight: 700; }
    .changed-cell { background-color: #e2efda; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ── Header ────────────────────────────────────────────────────────────────────
st.title("Address Corrector")
st.caption(
    "Paste addresses copied from Excel / any spreadsheet, or upload a CSV / Excel file. "
    "Original and corrected columns are returned side-by-side."
)

# ── Input ─────────────────────────────────────────────────────────────────────
tab_paste, tab_upload = st.tabs(["Paste addresses", "Upload file"])

df_raw = None
source_label = ""

with tab_paste:
    st.markdown(
        "Copy a block of cells from Excel (including the header row) and paste below. "
        "Comma-separated, tab-separated, and semicolon-separated formats are all supported."
    )
    pasted = st.text_area(
        "Paste here",
        height=260,
        placeholder=(
            "Address Line 1\tAddress Line 2\tCity\tState\tCountry\tPostal Code\n"
            "123 main st\tapt 4b\tnew york\tnew york\tUSA\t10001\n"
            "456 oak ave\tsuite 200\tlos angeles\tCA\tus\t90210\n"
            "..."
        ),
    )
    if pasted.strip():
        # Auto-detect delimiter
        sample = pasted.split("\n")[0]
        tabs   = sample.count("\t")
        commas = sample.count(",")
        semis  = sample.count(";")
        delim  = "\t" if tabs >= commas and tabs >= semis else (";" if semis > commas else ",")

        try:
            df_raw = pd.read_csv(
                io.StringIO(pasted),
                sep=delim,
                dtype=str,
                on_bad_lines="skip",
            ).fillna("")
            source_label = f"Pasted text ({len(df_raw)} rows)"
        except Exception as e:
            st.error(f"Could not parse pasted text: {e}")

with tab_upload:
    uploaded = st.file_uploader(
        "Upload CSV or Excel file",
        type=["csv", "xlsx", "xls"],
        label_visibility="collapsed",
    )
    if uploaded:
        try:
            if uploaded.name.endswith(".csv"):
                df_raw = pd.read_csv(uploaded, dtype=str).fillna("")
            else:
                df_raw = pd.read_excel(uploaded, dtype=str).fillna("")
            source_label = f"{uploaded.name} ({len(df_raw)} rows)"
        except Exception as e:
            st.error(f"Could not read file: {e}")

# ── Processing ────────────────────────────────────────────────────────────────
if df_raw is not None and len(df_raw) > 0:

    col_map = _detect_columns(df_raw.columns)
    detected = {f: c for f, c in col_map.items() if c is not None}

    # ── Column mapping indicator ───────────────────────────────────────────
    st.divider()
    st.subheader("Detected columns")

    cols = st.columns(7)
    for idx, (field, label) in enumerate(DISPLAY_LABELS.items()):
        orig_col = col_map.get(field)
        with cols[idx]:
            if orig_col:
                st.success(f"**{label}**\n\n`{orig_col}`")
            else:
                st.warning(f"**{label}**\n\nnot found")

    if not detected:
        st.error("No address columns could be detected. Check that your data has a header row.")
        st.stop()

    # ── Apply corrections ──────────────────────────────────────────────────
    result = df_raw.copy()
    corrected_col_map = {}

    for field, orig_col in col_map.items():
        if orig_col is None:
            continue
        label = f"Corrected {DISPLAY_LABELS[field]}"
        result[label] = df_raw[orig_col].apply(CORRECTORS[field])
        corrected_col_map[label] = label

    # ── Summary metrics ────────────────────────────────────────────────────
    st.divider()
    st.subheader("Summary")

    total_rows    = len(df_raw)
    fields_found  = len(detected)
    total_changed = 0
    per_field     = {}

    for field, orig_col in col_map.items():
        if orig_col is None:
            continue
        corr_label = f"Corrected {DISPLAY_LABELS[field]}"
        changed = (
            result[orig_col].astype(str).str.strip()
            != result[corr_label].astype(str).str.strip()
        ).sum()
        total_changed += changed
        per_field[DISPLAY_LABELS[field]] = int(changed)

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Rows processed", f"{total_rows:,}")
    m2.metric("Fields detected", fields_found)
    m3.metric("Total corrections", f"{total_changed:,}")
    m4.metric("Correction rate", f"{total_changed / max(total_rows * fields_found, 1):.0%}")

    # Per-field breakdown
    with st.expander("Corrections per field"):
        bar_data = pd.DataFrame.from_dict(
            per_field, orient="index", columns=["Corrections"]
        )
        st.bar_chart(bar_data)

    # ── Results table ──────────────────────────────────────────────────────
    st.divider()
    st.subheader("Corrected addresses")

    view_mode = st.radio(
        "Show",
        ["Side by side (original + corrected)", "Corrected only", "Changes only"],
        horizontal=True,
        label_visibility="collapsed",
    )

    orig_addr_cols  = [c for c in df_raw.columns if c in col_map.values()]
    corr_addr_cols  = list(corrected_col_map.keys())
    other_cols      = [c for c in df_raw.columns if c not in orig_addr_cols]

    if view_mode == "Corrected only":
        display_df = pd.concat(
            [df_raw[other_cols], result[corr_addr_cols]], axis=1
        )
    elif view_mode == "Changes only":
        # Only rows where at least one field changed
        mask = pd.Series(False, index=result.index)
        for field, orig_col in col_map.items():
            if orig_col is None:
                continue
            corr_label = f"Corrected {DISPLAY_LABELS[field]}"
            mask |= (
                result[orig_col].astype(str).str.strip()
                != result[corr_label].astype(str).str.strip()
            )
        display_df = result[mask]
    else:
        display_df = result

    # Highlight changed cells
    def highlight_changes(df):
        styles = pd.DataFrame("", index=df.index, columns=df.columns)
        for field, orig_col in col_map.items():
            corr_label = f"Corrected {DISPLAY_LABELS[field]}"
            if orig_col not in df.columns or corr_label not in df.columns:
                continue
            changed_mask = (
                df[orig_col].astype(str).str.strip()
                != df[corr_label].astype(str).str.strip()
            )
            styles.loc[changed_mask, corr_label] = "background-color: #e2efda; font-weight: 600"
            styles.loc[changed_mask, orig_col]   = "background-color: #fff3cd"
        return styles

    styled = display_df.style.apply(highlight_changes, axis=None)

    st.dataframe(
        styled,
        use_container_width=True,
        height=min(40 + 35 * len(display_df), 600),
    )

    st.caption(
        f"Showing {len(display_df):,} of {total_rows:,} rows.  "
        "Green = corrected value · Yellow = original value that was changed."
    )

    # ── Download ───────────────────────────────────────────────────────────
    st.divider()
    st.subheader("Download")

    dl1, dl2 = st.columns(2)

    with dl1:
        buf_xlsx = io.BytesIO()
        _write_excel(result, list(df_raw.columns), corrected_col_map, buf_xlsx, col_map)
        buf_xlsx.seek(0)
        st.download_button(
            label="Download Excel (.xlsx)",
            data=buf_xlsx,
            file_name="corrected_addresses.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )

    with dl2:
        csv_bytes = result.to_csv(index=False).encode("utf-8-sig")  # utf-8-sig for Excel compat
        st.download_button(
            label="Download CSV (.csv)",
            data=csv_bytes,
            file_name="corrected_addresses.csv",
            mime="text/csv",
            use_container_width=True,
        )

else:
    st.info("Paste addresses above or upload a file to get started.")
