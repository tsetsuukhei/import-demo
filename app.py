"""
Streamlit UI for the Import File Processor.

Upload a raw import Excel file, process it through the matching pipeline,
and download the categorized result.
"""

import streamlit as st
import pandas as pd
import os
import io
import tempfile

from process_file import (
    DICT_DIR, THRESHOLD, DEL_THRESHOLD, COL_MAP, MATCH_COL_MAP, OUTPUT_COLS,
    norm, norm_code,
    load_business_dict, load_deleted_dict, load_hs_prefixes, load_sheet_names,
    filter_by_hs, load_raw, match_row, format_output,
)

# ── Page config ──────────────────────────────────────────
st.set_page_config(page_title="Import File Processor", page_icon="📦", layout="centered")

st.title("📦 Import File")

# ── Validate that dictionary files exist ─────────────────
REQUIRED_DICTS = ['business_dict.csv', 'deleted_dict.csv', 'hs_prefixes.csv', 'sheet_names.csv']
missing = [f for f in REQUIRED_DICTS if not os.path.exists(os.path.join(DICT_DIR, f))]
if missing:
    st.error(f"Missing dictionary files in `{DICT_DIR}/`: {', '.join(missing)}. Run `build_dict.py` first.")
    st.stop()


# ── Cache dictionary loading so it only runs once ────────
@st.cache_resource(show_spinner="Loading dictionaries…")
def load_dicts():
    biz_lookup, biz_registry = load_business_dict(DICT_DIR)
    del_lookup, del_registry = load_deleted_dict(DICT_DIR)
    hs_prefixes = load_hs_prefixes(DICT_DIR)
    biz_sheets = load_sheet_names(DICT_DIR)
    return biz_lookup, biz_registry, del_lookup, del_registry, hs_prefixes, biz_sheets


biz_lookup, biz_registry, del_lookup, del_registry, hs_prefixes, biz_sheets = load_dicts()


# ── File uploader ────────────────────────────────────────
uploaded = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])

if uploaded is not None:
    # Write uploaded bytes to a temp file so pandas can read multiple sheets
    suffix = os.path.splitext(uploaded.name)[1]
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(uploaded.getvalue())
        tmp_path = tmp.name

    try:
        # ── Step 1: Load raw data ────────────────────────
        with st.spinner("Loading input file…"):
            raw = load_raw(tmp_path)
        st.success(f"Loaded **{len(raw):,}** rows from **{uploaded.name}**")

        # ── Step 2: HS code filtering ────────────────────
        with st.spinner("Filtering by HS code…"):
            hs_kept, hs_deleted = filter_by_hs(raw, 'бараакод', hs_prefixes)

        col1, col2 = st.columns(2)
        col1.metric("HS Kept", f"{len(hs_kept):,}")
        col2.metric("HS Deleted", f"{len(hs_deleted):,}")

        # ── Step 3: Match rows ───────────────────────────
        match_count = len(hs_kept)
        progress_bar = st.progress(0, text="Matching rows…")
        processed = [0]

        def _match(row):
            result = match_row(
                row, biz_lookup, biz_registry,
                del_lookup, del_registry,
                THRESHOLD, DEL_THRESHOLD,
            )
            processed[0] += 1
            if processed[0] % 500 == 0 or processed[0] == match_count:
                progress_bar.progress(
                    processed[0] / match_count,
                    text=f"Matching rows… {processed[0]:,}/{match_count:,}",
                )
            return result

        results = hs_kept.apply(_match, axis=1)
        results_df = pd.DataFrame(list(results))
        combined = pd.concat([hs_kept.reset_index(drop=True), results_df], axis=1)
        progress_bar.progress(1.0, text="Matching complete ✓")

        # ── Step 4: Build output sheets ──────────────────
        sheet_data: dict[str, pd.DataFrame] = {}

        for sheet in biz_sheets:
            mask = combined['business'] == sheet
            if mask.any():
                sheet_data[sheet] = format_output(combined[mask].copy())

        del_mask = combined['business'] == 'Deleted'
        if del_mask.any():
            sheet_data['Deleted'] = format_output(combined[del_mask].copy())

        if len(hs_deleted) > 0:
            sheet_data['HS Deleted'] = format_output(hs_deleted.copy())

        unc_mask = combined['business'] == 'Uncertain'
        if unc_mask.any():
            sheet_data['Uncertain'] = format_output(combined[unc_mask].copy())

        # ── Summary table ────────────────────────────────
        st.subheader("Summary")
        summary_rows = []
        total_raw = len(raw)
        biz_total = sum((combined['business'] == s).sum() for s in biz_sheets)

        for sheet in biz_sheets:
            cnt = int((combined['business'] == sheet).sum())
            if cnt > 0:
                summary_rows.append({"Sheet": sheet, "Rows": cnt, "% of Total": f"{cnt/total_raw*100:.1f}%"})

        del_total = int(del_mask.sum())
        hs_del_total = len(hs_deleted)
        unc_total = int(unc_mask.sum())

        summary_rows.append({"Sheet": "Deleted (dict)", "Rows": del_total, "% of Total": f"{del_total/total_raw*100:.1f}%"})
        summary_rows.append({"Sheet": "HS Deleted", "Rows": hs_del_total, "% of Total": f"{hs_del_total/total_raw*100:.1f}%"})
        summary_rows.append({"Sheet": "Uncertain", "Rows": unc_total, "% of Total": f"{unc_total/total_raw*100:.1f}%"})

        summary_df = pd.DataFrame(summary_rows)
        st.table(summary_df)

        # ── Write to in-memory Excel & offer download ────
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as writer:
            for name, df in sheet_data.items():
                df.to_excel(writer, sheet_name=name[:31], index=False)
        buf.seek(0)

        base, ext = os.path.splitext(uploaded.name)
        output_name = base + '_cleaned.xlsx'

        st.download_button(
            label="⬇️ Download Processed File",
            data=buf,
            file_name=output_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

    finally:
        # Clean up temp file
        os.unlink(tmp_path)
