from __future__ import annotations

from datetime import datetime
import streamlit as st

from core.gear.config import GearConfig
from core.gear.engine import process_gear_batch
from core.memory import get_memory_usage_mb
from utils.audit_viewer import render_audit_table
from utils.ui_components import (
    render_download_button,
    render_header,
    render_memory_badge,
    render_metric_cards,
)

st.set_page_config(
    page_title="Gear Automation | Station Fuel Aggregator",
    page_icon="⛽",
    layout="wide",
)

render_memory_badge()

render_header(
    title="Station Fuel Gear Automation",
    subtitle="Iterative batch aggregation of 50+ station workbooks into the master Gear record",
    icon="⛽",
)

st.markdown(
    """
    Upload your master prepared **Gear workbook** along with all daily **station workbooks**. 
    The engine streams workbooks one-by-one with deterministic memory cleanup, matching station headers and computing formulas.
    """
)

with st.form("gear_automation_form"):
    col1, col2 = st.columns([1, 1])

    with col1:
        st.subheader("1. Master Gear Workbook")
        master_file = st.file_uploader(
            "Upload the prepared destination Gear workbook",
            type=["xlsx", "xlsm"],
            accept_multiple_files=False,
            help="The master monthly workbook containing the destination station columns and running formulas.",
        )
        gear_sheet_name = st.text_input(
            "Gear Sheet Name",
            value="Sheet1",
            help="Worksheet name in the master Gear workbook to write into.",
        )

    with col2:
        st.subheader("2. Station Workbooks (Batch)")
        source_files = st.file_uploader(
            "Upload daily station files (drag & drop 50+ files)",
            type=["xlsx", "xlsm"],
            accept_multiple_files=True,
            help="Upload all station sales workbooks. Files are processed iteratively to conserve RAM.",
        )
        dry_run = st.checkbox(
            "Preview / Dry Run Only",
            value=False,
            help="Validates mappings and extracts values without modifying or generating the master workbook.",
        )

    with st.expander("⚙️ Advanced Destination Row & Sheet Configuration", expanded=False):
        st.info("Default settings correspond to the standard monthly Gear template structure.")
        rcol1, rcol2, rcol3 = st.columns(3)
        with rcol1:
            header_row = st.number_input("Tank Header Row", min_value=1, value=2)
            sales_meter_row = st.number_input("Sales Meter Row", min_value=1, value=130)
            sales_dipping_row = st.number_input("Sales Dipping Row", min_value=1, value=131)
            source_sheet_index = st.number_input("Source Sheet Index (-1 = last sheet)", value=-1)
        with rcol2:
            actual_gear_row = st.number_input("Actual Gear Row", min_value=1, value=132)
            expected_gear_row = st.number_input("Expected Gear Row", min_value=1, value=133)
            gear_percent_row = st.number_input("Gear % Row", min_value=1, value=134)
        with rcol3:
            average_gear_row = st.number_input("Average Gear Row", min_value=1, value=135)
            prev_average_row = st.number_input("Previous Average Row", min_value=1, value=128)
            protected_cols = st.number_input("Protected Trailing Columns Count", min_value=0, value=2)

    submit_button = st.form_submit_button("🚀 Run Gear Aggregation Pipeline", type="primary", use_container_width=True)

if submit_button:
    if not master_file:
        st.error("Please upload the Master Gear Workbook before executing.")
    elif not source_files:
        st.error("Please upload at least one Station Workbook to process.")
    else:
        config = GearConfig(
            gear_sheet_name=gear_sheet_name.strip(),
            header_row_for_tanks=int(header_row),
            sales_meter_row=int(sales_meter_row),
            sales_dipping_row=int(sales_dipping_row),
            actual_gear_row=int(actual_gear_row),
            expected_gear_row=int(expected_gear_row),
            gear_percent_row=int(gear_percent_row),
            average_gear_row=int(average_gear_row),
            previous_average_gear_row=int(prev_average_row),
            protected_last_columns_count=int(protected_cols),
            source_sheet_index=int(source_sheet_index),
        )

        progress_bar = st.progress(0.0)
        status_text = st.empty()

        def on_progress(current: int, total: int, filename: str) -> None:
            pct = current / total
            progress_bar.progress(pct)
            status_text.markdown(f"**Processing:** `{filename}` ({current}/{total})")

        try:
            with st.spinner("Executing batch pipeline..."):
                result = process_gear_batch(
                    master_gear_file=master_file,
                    source_files=source_files,
                    config=config,
                    dry_run=dry_run,
                    progress_callback=on_progress,
                )

            progress_bar.progress(1.0)
            status_text.success(f"Processing complete: {result.processed_count} workbooks processed successfully.")

            # Summary Metrics
            render_metric_cards([
                ("Total Workbooks", len(source_files), None),
                ("Processed Successfully", result.processed_count, None),
                ("Skipped / Errors", result.skipped_count, None),
                ("Current RAM", f"{get_memory_usage_mb():.1f} MB", None),
            ])

            # Download Action
            if result.output_path and not dry_run:
                st.markdown("### 📥 Download Processed Workbook")
                render_download_button(
                    file_path=result.output_path,
                    label="📥 Download Updated Gear Master Workbook (.xlsx)",
                    download_filename=f"Gear_Updated_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
                )
            elif dry_run:
                st.info("Dry-run preview completed. Review mapping results in the table below.")

            # Audit Table
            render_audit_table(result.logs, title="Station Mapping & Extraction Log")

        except Exception as exc:
            st.error(f"Execution Error: {exc}")
