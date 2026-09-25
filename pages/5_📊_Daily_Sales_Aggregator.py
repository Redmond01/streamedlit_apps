from __future__ import annotations

from datetime import date, datetime
import streamlit as st

from core.sales.config import SalesConfig
from core.sales.engine import process_sales_batch
from core.memory import get_memory_usage_mb
from utils.audit_viewer import render_audit_table
from utils.ui_components import (
    render_download_button,
    render_header,
    render_memory_badge,
    render_metric_cards,
)

try:
    st.set_page_config(
        page_title="Daily Sales Aggregator | Station Sales Sync",
        page_icon="📊",
        layout="wide",
    )
except Exception:
    pass

render_memory_badge()

render_header(
    title="Daily Fuel Sales Aggregator",
    subtitle="Iterative batch extraction of PMS, AGO, and LPG daily sales from 50+ stations into master report",
    icon="📊",
)

st.markdown(
    """
    Upload your master monthly **Sales Report workbook** (containing `PMS`, `AGO`, and `LPG` sheets) along with daily **station workbooks**. 
    The engine streams station workbooks one-by-one, accurately detects dates and sales figures, maps station headers, and populates the master report while preserving all existing formulas.
    """
)

with st.form("daily_sales_aggregation_form"):
    col1, col2 = st.columns([1, 1])

    with col1:
        st.subheader("1. Master Monthly Sales Workbook")
        master_file = st.file_uploader(
            "Upload the prepared destination Sales Report workbook",
            type=["xlsx", "xlsm"],
            accept_multiple_files=False,
            help="The master monthly workbook (e.g. SEPTEMBER SALES REPORT 2026.xlsx) containing PMS, AGO, and LPG sheets with summary formulas.",
        )

        st.markdown("**Worksheet Names in Master**")
        sc1, sc2, sc3 = st.columns(3)
        with sc1:
            pms_sheet_name = st.text_input("PMS Sheet", value="PMS")
        with sc2:
            ago_sheet_name = st.text_input("AGO Sheet", value="AGO")
        with sc3:
            lpg_sheet_name = st.text_input("LPG Sheet", value="LPG")

    with col2:
        st.subheader("2. Station Workbooks (Batch)")
        source_files = st.file_uploader(
            "Upload daily station files (drag & drop 50+ files)",
            type=["xlsx", "xlsm"],
            accept_multiple_files=True,
            help="Upload all station sales workbooks. Files are processed iteratively to conserve RAM.",
        )

        sync_mode_choice = st.radio(
            "📅 Sync Mode",
            options=[
                "All Available Days (Full Sync)",
                "Latest Day Only (Last Sheet)",
                "Specific Calendar Date",
            ],
            index=0,
            help="Select whether to sync all historical daily sheets found in the workbooks, only the latest day, or a specific calendar date.",
        )

        target_date_val = None
        if sync_mode_choice == "Specific Calendar Date":
            target_date_val = st.date_input("Select Target Date to Extract", value=date.today())

        dry_run = st.checkbox(
            "Preview / Dry Run Only",
            value=False,
            help="Validates mappings and extracts values without modifying or generating the master workbook.",
        )

    with st.expander("⚙️ Advanced Destination Header Rows & Matching Threshold", expanded=False):
        st.info("Default settings correspond to the standard monthly sales report layout.")
        rcol1, rcol2, rcol3, rcol4 = st.columns(4)
        with rcol1:
            pms_header_row = st.number_input("PMS Station Header Row", min_value=1, value=2)
        with rcol2:
            ago_header_row = st.number_input("AGO Station Header Row", min_value=1, value=2)
        with rcol3:
            lpg_header_row = st.number_input("LPG Station Header Row", min_value=1, value=3)
        with rcol4:
            match_threshold = st.slider("Fuzzy Match Sensitivity", min_value=0.3, max_value=0.9, value=0.55, step=0.05)

    submit_button = st.form_submit_button("🚀 Run Daily Sales Aggregation Pipeline", type="primary", use_container_width=True)

if submit_button:
    if not master_file:
        st.error("Please upload the Master Sales Report Workbook before executing.")
    elif not source_files:
        st.error("Please upload at least one Station Workbook to process.")
    else:
        sync_mode_code = "all"
        if sync_mode_choice == "Latest Day Only (Last Sheet)":
            sync_mode_code = "latest"
        elif sync_mode_choice == "Specific Calendar Date":
            sync_mode_code = "specific_date"

        config = SalesConfig(
            pms_sheet_name=pms_sheet_name.strip(),
            ago_sheet_name=ago_sheet_name.strip(),
            lpg_sheet_name=lpg_sheet_name.strip(),
            pms_header_row=int(pms_header_row),
            ago_header_row=int(ago_header_row),
            lpg_header_row=int(lpg_header_row),
            sync_mode=sync_mode_code,
            target_date=target_date_val if isinstance(target_date_val, date) else None,
            match_threshold=float(match_threshold),
        )

        progress_bar = st.progress(0.0)
        status_text = st.empty()

        def on_progress(current: int, total: int, filename: str) -> None:
            pct = current / total
            progress_bar.progress(pct)
            status_text.markdown(f"**Processing:** `{filename}` ({current}/{total})")

        try:
            with st.spinner("Executing sales batch pipeline..."):
                result = process_sales_batch(
                    master_sales_file=master_file,
                    source_files=source_files,
                    config=config,
                    dry_run=dry_run,
                    progress_callback=on_progress,
                )

            progress_bar.progress(1.0)
            status_text.success(
                f"Processing complete: {result.processed_count} workbooks processed successfully. "
                f"({result.total_records_written} cell records {'projected' if dry_run else 'written'})."
            )

            # Summary Metrics
            pms_tot = result.summary_metrics.get("total_pms_litres", 0.0)
            ago_tot = result.summary_metrics.get("total_ago_litres", 0.0)
            lpg_tot = result.summary_metrics.get("total_lpg_kg", 0.0)

            render_metric_cards([
                ("Total Workbooks", len(source_files), None),
                ("Processed", result.processed_count, None),
                ("Skipped / Errors", result.skipped_count, None),
                ("Records Written", result.total_records_written, None),
                ("Total PMS", f"{pms_tot:,.2f} L", None),
                ("Total AGO", f"{ago_tot:,.2f} L", None),
                ("Total LPG", f"{lpg_tot:,.2f} KG", None),
                ("Current RAM", f"{get_memory_usage_mb():.1f} MB", None),
            ])

            # Download Action
            if result.output_path and not dry_run:
                st.markdown("### 📥 Download Updated Sales Report")
                render_download_button(
                    file_path=result.output_path,
                    label="📥 Download Updated Monthly Sales Report (.xlsx)",
                    download_filename=f"SalesReport_Updated_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
                )
            elif dry_run:
                st.info("Dry-run preview completed. Review extracted values and target cell coordinates below.")

            # Audit Table
            render_audit_table(result.logs, title="Station Sales Extraction & Mapping Log")

        except Exception as exc:
            st.error(f"Execution Error: {exc}")
