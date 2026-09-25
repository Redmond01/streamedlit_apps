from __future__ import annotations

import json
import streamlit as st

from core.pos_recon.cashbook_extractor import extract_cashbook_loop
from core.memory import get_memory_usage_mb
from utils.audit_viewer import render_audit_table
from utils.ui_components import (
    render_header,
    render_memory_badge,
    render_metric_cards,
)

try:
    st.set_page_config(
        page_title="Cashbook Extractor | POS Automation",
        page_icon="📋",
        layout="wide",
    )
except Exception:
    pass

render_memory_badge()

render_header(
    title="Cashbook Loop Extractor",
    subtitle="Extract active POS lists from Excel cashbook reports for the Moniepoint Chrome Extension",
    icon="📋",
)

st.markdown(
    """
    Upload an Excel cashbook reconciliation workbook (e.g. `Book21SEPT.xlsx` or multi-day `Book9-10SEPT.xlsx`). 
    The tool extracts the active terminal numbers for **SGR, OBX, AGS, and TA** groups formatted for extension injection.
    """
)

with st.form("cashbook_extractor_form"):
    col1, col2 = st.columns([1, 1])

    with col1:
        st.subheader("1. Reconciliation Workbook")
        workbook_file = st.file_uploader(
            "Upload Cashbook Workbook (.xlsx / .xlsm)",
            type=["xlsx", "xlsm"],
            help="Excel workbook containing 'cashbook' in row 4 headers.",
        )
        date_input = st.text_input(
            "Report Date (YYYY-MM-DD)",
            value="",
            placeholder="e.g. 2026-09-21 (leave blank to auto-detect)",
            help="Selects a specific date report block when the workbook covers multiple days.",
        )

    with col2:
        st.subheader("2. Filtering & Extraction Options")
        block_number = st.number_input(
            "Report Block Number",
            min_value=1,
            value=1,
            help="Used if the date is omitted and multiple blocks exist.",
        )
        min_number = st.number_input(
            "Minimum POS Number (Skip SGR/OBX below this)",
            min_value=1,
            value=11,
            help="Skips POS 1-10 for SGR and OBX to avoid inconsistent legacy labels.",
        )
        target_day_val = st.text_input(
            "Target Day Override (Optional)",
            value="",
            placeholder="e.g. 21 (defaults to report date day)",
        )

    col_t1, col_t2 = st.columns(2)
    with col_t1:
        compact_mode = st.toggle("Compact Single-Line Output", value=True, help="Recommended for pasting directly into the Chrome extension.")
    with col_t2:
        include_audit = st.checkbox("Include Audit Details (Skipped 1-10 & Ignored Rows)", value=False)

    submit_button = st.form_submit_button("🔍 Extract POS Loop Configuration", type="primary", use_container_width=True)

if submit_button:
    if not workbook_file:
        st.error("Please upload a Cashbook workbook before extracting.")
    else:
        target_day = int(target_day_val.strip()) if target_day_val.strip().isdigit() else None
        date_clean = date_input.strip() or None

        try:
            with st.spinner("Extracting POS terminal loop..."):
                result = extract_cashbook_loop(
                    workbook_input=workbook_file,
                    date=date_clean,
                    block_number=int(block_number),
                    min_number=int(min_number),
                    target_day=target_day,
                    include_audit=include_audit,
                )

            st.success(f"Successfully extracted POS loop for Report Date: **{result.get('date', 'N/A')}** (Target Day: {result.get('targetDay', 'N/A')})")

            # Metrics
            groups = result.get("groups", {})
            render_metric_cards([
                ("SGR Terminals", len(groups.get("SGR", [])), None),
                ("OBX Terminals", len(groups.get("OBX", [])), None),
                ("AGS Terminals", len(groups.get("AGS", [])), None),
                ("TA Terminals", len(groups.get("TA", [])), None),
            ])

            # JSON Preview
            json_text = json.dumps(result, separators=(",", ":")) if compact_mode else json.dumps(result, indent=2)

            st.markdown("### 📋 Generated Loop JSON")
            st.caption("Click the copy icon on the top right of the code block to copy into the Chrome extension:")
            st.code(json_text, language="json")

            download_filename = f"pos_loop_{result.get('date', 'export')}.json"
            st.download_button(
                label="📥 Download Loop JSON (.json)",
                data=json_text,
                file_name=download_filename,
                mime="application/json",
                type="primary",
                use_container_width=True,
            )

            # Audit Details
            if include_audit:
                st.markdown("### 🔍 Audit Information")
                skipped = result.get("skippedFirstTen", [])
                if skipped:
                    render_audit_table(skipped, title="Skipped Terminals (< Minimum Number)")
                ignored = result.get("ignoredOther", [])
                if ignored:
                    render_audit_table(ignored, title="Ignored Non-Terminal Rows")

        except Exception as exc:
            st.error(f"Extraction Error: {exc}")
