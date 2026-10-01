from __future__ import annotations

from datetime import datetime
import pandas as pd
import streamlit as st

from core.bank_lodgment.config import BankLodgmentConfig, DEFAULT_PHONE_MAP
from core.bank_lodgment.engine import process_bank_lodgment
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
        page_title="Bank Lodgment | Deposit Confirmation Sync",
        page_icon="🏦",
        layout="wide",
    )
except Exception:
    pass

render_memory_badge()

render_header(
    title="Bank Lodgment Automation",
    subtitle="Automated ingestion of bank deposit confirmations into the master lodgment workbook",
    icon="🏦",
)

st.markdown(
    """
    Upload your master monthly **Bank Lodgments workbook** (e.g. `SIFEM BANK LODGEMENTS - SEPTEMBER 2026.xlsx`) along with the 
    **confirmation JSON export** (e.g. `download.json`). 
    The engine accurately matches station deposits, deduplicates identical group confirmations, writes amounts, lodgement times, and confirmation officers, while keeping existing formulas and summary calculations 100% intact.
    """
)

with st.form("bank_lodgment_form"):
    col1, col2 = st.columns([1, 1])

    with col1:
        st.subheader("1. Master Bank Lodgments Workbook")
        master_file = st.file_uploader(
            "Upload the master Bank Lodgments workbook (.xlsx)",
            type=["xlsx", "xlsm"],
            accept_multiple_files=False,
            help="The master workbook containing daily sheets (e.g. 25-09-2026) and the SUMMARY sheet.",
        )

        auto_create = st.checkbox(
            "Auto-create new date sheet if not present",
            value=True,
            help="If the target date sheet does not exist in the workbook, clones the template from the previous day and links it to the SUMMARY sheet.",
        )

        color_mode_choice = st.radio(
            "🎨 Output Coloring Style",
            options=[
                "Highlight Entire Confirmed Row (Full Color)",
                "Highlight Confirmation Columns Only (Columns G-I)",
            ],
            index=0,
            help="Select whether to color the entire row for confirmed deposits or only the confirmation status columns.",
        )

    with col2:
        st.subheader("2. Confirmed Deposit JSON Export")
        json_file = st.file_uploader(
            "Upload the confirmation JSON dump (download.json)",
            type=["json"],
            accept_multiple_files=False,
            help="JSON export containing approved bank officer confirmations and manager receipt submissions.",
        )

        target_date_override = st.text_input(
            "Target Date Override (Optional)",
            value="",
            placeholder="e.g. 25-09-2026 or leave blank for auto-detection",
            help="Leave blank to automatically use the date recorded in the JSON metadata.",
        )

        dry_run = st.checkbox(
            "Preview / Dry Run Only",
            value=False,
            help="Validates matches and extracts values without modifying or generating the master workbook.",
        )

    with st.expander("📞 Verified Station Manager Phone Directory", expanded=False):
        st.info("The system automatically resolves station names from manager phone numbers even when the caption is 'Photo'.")
        phone_df = pd.DataFrame(
            [{"Phone Number": k, "Station Name": v} for k, v in sorted(DEFAULT_PHONE_MAP.items(), key=lambda x: x[1])]
        )
        st.dataframe(phone_df, use_container_width=True, hide_index=True)

    submit_button = st.form_submit_button("🚀 Run Bank Lodgment Pipeline", type="primary", use_container_width=True)

if submit_button:
    if not master_file:
        st.error("Please upload the Master Bank Lodgments Workbook before executing.")
    elif not json_file:
        st.error("Please upload the confirmed deposit JSON dump before executing.")
    else:
        color_mode_code = "row" if "Entire" in color_mode_choice else "columns"
        config = BankLodgmentConfig(
            target_date=target_date_override.strip() if target_date_override.strip() else None,
            auto_create_sheet=auto_create,
            color_mode=color_mode_code,
        )

        progress_bar = st.progress(0.0)
        status_text = st.empty()

        def on_progress(current: int, total: int, msg: str) -> None:
            pct = current / total
            progress_bar.progress(pct)
            status_text.markdown(f"**Processing:** `{msg}` ({current}/{total})")

        try:
            with st.spinner("Executing bank lodgment pipeline..."):
                result = process_bank_lodgment(
                    master_file=master_file,
                    json_file=json_file,
                    config=config,
                    dry_run=dry_run,
                    progress_callback=on_progress,
                )

            progress_bar.progress(1.0)
            status_text.success(
                f"Processing complete: {result.confirmed_count} deposits confirmed into sheet `{result.target_sheet_name}`. "
                f"({result.duplicate_count} duplicate confirmations skipped)."
            )

            # Summary Metrics
            render_metric_cards([
                ("Total Records", result.total_records, None),
                ("Deposits Confirmed", result.confirmed_count, None),
                ("Duplicates Skipped", result.duplicate_count, None),
                ("Unmatched", result.unmatched_count, None),
                ("Total Confirmed", f"₦{result.total_confirmed_amount:,.2f}", None),
                ("Union Bank", f"₦{result.union_confirmed_amount:,.2f}", None),
                ("FCMB / Access", f"₦{result.fcmb_confirmed_amount:,.2f}", None),
                ("Current RAM", f"{get_memory_usage_mb():.1f} MB", None),
            ])

            # Download Action
            if result.output_path and not dry_run:
                st.markdown("### 📥 Download Updated Lodgments Workbook")
                render_download_button(
                    file_path=result.output_path,
                    label="📥 Download Updated Bank Lodgments Workbook (.xlsx)",
                    download_filename=f"BankLodgements_{result.target_sheet_name}_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
                )
            elif dry_run:
                st.info("Dry-run preview completed. Review extracted values and target cell coordinates below.")

            # Audit Table
            render_audit_table(result.logs, title="Bank Lodgment Confirmation Log")

        except Exception as exc:
            st.error(f"Execution Error: {exc}")
