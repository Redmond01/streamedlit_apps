from __future__ import annotations

from datetime import datetime, date
import streamlit as st

from core.pos_recon.eod_extractor import write_eod_to_workbook
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
        page_title="EOD Cashbook Sync | POS Automation",
        page_icon="📥",
        layout="wide",
    )
except Exception:
    pass

render_memory_badge()

render_header(
    title="EOD Cashbook Synchronization",
    subtitle="Sync daily Moniepoint EOD transactions into the master reconciliation CASHBOOK column",
    icon="📥",
)

st.markdown(
    """
    Upload your **Moniepoint EOD export** (daily or multi-day monthly workbook) and the **Target Reconciliation workbook**. 
    The engine filters transactions matching the target date, maps normalized POS rows, calculates difference formulas, 
    and inserts an **EOD Extract Audit** sheet with full metric reporting.
    """
)

with st.form("eod_sync_form"):
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("1. Source & Target Workbooks")
        source_file = st.file_uploader(
            "Moniepoint EOD Source Workbook (.xlsx / .xlsm)",
            type=["xlsx", "xlsm"],
            help="Contains columns: Timestamp, Submission ID, Station, Report Date, Full POS ID, Amount.",
        )
        target_file = st.file_uploader(
            "Target Reconciliation Workbook (.xlsx / .xlsm)",
            type=["xlsx", "xlsm"],
            help="The master cashbook reconciliation workbook to be updated.",
        )

    with col2:
        st.subheader("2. Synchronization Parameters")
        target_date = st.date_input(
            "Target Report Date",
            value=date.today(),
            help="Select the exact report date to filter from the EOD source.",
        )
        source_sheet_override = st.text_input(
            "Source Sheet Name Override (Optional)",
            value="",
            placeholder="e.g. Daily_POS_2026_09_22 (auto-computed if blank)",
            help="Leave blank to use standard sheet naming format based on the selected date.",
        )
        block_number = st.number_input(
            "Target Report Block Number",
            min_value=1,
            value=1,
            help="1-based report block number in the target workbook if date matching has multiple matches.",
        )
        use_latest_duplicate = st.toggle(
            "🕒 Use Latest Submitted POS (by Timestamp)",
            value=True,
            help="When enabled (default), if duplicate POS submissions exist on the report date, the engine resolves duplicates by selecting the latest submission using its timestamp (IF POS DUPLICATE == TRUE -> LATEST SUBMITTED POS AMOUNT = TRUE AMOUNT). When disabled, duplicate amounts are summed together.",
        )
        remap_ags10 = st.toggle(
            "🔄 Map AGS 10 to SGR 10",
            value=True,
            help="When enabled (default), terminal AGS 10 (Igbesa AGS) in the Moniepoint export is automatically mapped to SGR 10 in the reconciliation workbook.",
        )

    submit_button = st.form_submit_button("🔄 Execute EOD Reconciliation Sync", type="primary", use_container_width=True)

if submit_button:
    if not source_file:
        st.error("Please upload the Moniepoint EOD Source workbook.")
    elif not target_file:
        st.error("Please upload the Target Reconciliation workbook.")
    else:
        date_str = target_date.isoformat()
        sheet_override = source_sheet_override.strip() or None

        try:
            with st.spinner("Processing EOD transactions and updating reconciliation workbook..."):
                output_path, stats = write_eod_to_workbook(
                    source_input=source_file,
                    target_workbook_input=target_file,
                    date_str=date_str,
                    source_sheet=sheet_override,
                    block_number=int(block_number),
                    use_latest_duplicate=use_latest_duplicate,
                    remap_ags10=remap_ags10,
                )

            st.success(f"EOD sync completed successfully for date **{date_str}** (Source sheet: `{stats['sourceSheet']}`).")

            # Metrics
            if stats.get("entryModeEnabled"):
                render_metric_cards([
                    ("Target DEFAULT", stats["defaultCount"], None),
                    ("Prior CURRENT Fill-ins", stats["currentFillinCount"], None),
                    ("Duplicates Resolved", stats["crossDayDuplicatesCount"], None),
                    ("Mapped into Workbook", stats["mappedRows"], None),
                    ("Unmatched Terminals", len(stats["unmatchedSourceTerminals"]), None),
                ])

                # Cross-Day Duplicates Breakdown
                if stats.get("crossDayDuplicates"):
                    cross_count = len(stats["crossDayDuplicates"])
                    with st.expander(f"🔁 Cross-Day Duplicates ({cross_count} resolved — DEFAULT Kept)", expanded=True):
                        st.info(
                            f"Found **{cross_count}** POS terminal(s) submitted on both days (as CURRENT yesterday and as DEFAULT today). "
                            f"Per directive, today's **DEFAULT** amount was kept as the TRUE amount, and yesterday's CURRENT submission was superseded."
                        )
                        cross_rows = [
                            {
                                "POS Terminal": d["pos"],
                                "Selected TRUE Amount (DEFAULT)": f"₦{d['defaultAmount']:,.2f}",
                                "Discarded Amount (Prior CURRENT)": f"₦{d['currentAmount']:,.2f}",
                                "DEFAULT Submission ID": d.get("defaultSubmissionId", "-"),
                                "Prior CURRENT Submission ID": d.get("currentSubmissionId", "-"),
                                "Station": d.get("defaultStation", "-"),
                            }
                            for d in stats["crossDayDuplicates"]
                        ]
                        st.dataframe(cross_rows, use_container_width=True)

                # Prior Day CURRENT Fill-ins Breakdown
                if stats.get("fillinTerminals"):
                    fillin_count = len(stats["fillinTerminals"])
                    prev_sheet = stats.get("prevSourceSheet", "prior day sheet")
                    with st.expander(f"📥 Prior Day Fill-ins ({fillin_count} adopted from `{prev_sheet}`)", expanded=True):
                        st.info(
                            f"Found **{fillin_count}** POS terminal(s) that were **not submitted today as DEFAULT**, but were submitted yesterday as **CURRENT**. "
                            f"These records have been automatically adopted to ensure complete reconciliation coverage."
                        )
                        fillin_rows = [
                            {
                                "POS Terminal": f["pos"],
                                "Adopted Amount": f"₦{f['amount']:,.2f}",
                                "Timestamp": f.get("timestamp", "-"),
                                "Station": f.get("station", "-"),
                                "Submission ID": f.get("submissionId", "-"),
                            }
                            for f in stats["fillinTerminals"]
                        ]
                        st.dataframe(fillin_rows, use_container_width=True)

            else:
                dup_count = len(stats.get("duplicateTerminals", []))
                dup_metric_title = "Duplicates Resolved (Latest)" if use_latest_duplicate else "Duplicates Summed"
                render_metric_cards([
                    ("Source Rows", stats["sourceRows"], None),
                    ("Kept for Date", stats["keptRows"], None),
                    ("Mapped into Workbook", stats["mappedRows"], None),
                    (dup_metric_title, dup_count, None),
                    ("Unmatched Terminals", len(stats["unmatchedSourceTerminals"]), None),
                ])

                # Legacy Duplicate Breakdown Table
                if stats.get("duplicateTerminals"):
                    dup_mode_label = "Latest Submission by Timestamp" if use_latest_duplicate else "Summed Together"
                    with st.expander(f"🔁 Duplicate POS Terminals ({dup_count} detected - {dup_mode_label})", expanded=True):
                        if use_latest_duplicate:
                            st.info(
                                f"Found **{dup_count}** POS terminal(s) with multiple submissions for **{date_str}**. "
                                f"Resolved by selecting the **latest submitted POS amount based on timestamp**."
                            )
                        else:
                            st.info(
                                f"Found **{dup_count}** POS terminal(s) with multiple submissions for **{date_str}**. "
                                f"Amounts were **summed together**."
                            )
                        dup_rows = [
                            {
                                "POS Terminal": d["pos"],
                                "Submission Count": d["count"],
                                "Latest Timestamp": d.get("latestTimestamp", "-"),
                                "Selected Cashbook Amount": f"₦{d['computedAmount']:,.2f}",
                                "All Submissions": ", ".join(f"₦{a:,.2f}" for a in d["amounts"]),
                                "Total If Summed": f"₦{d['totalSum']:,.2f}",
                            }
                            for d in stats["duplicateTerminals"]
                        ]
                        st.dataframe(dup_rows, use_container_width=True)

            # Download Action
            st.markdown("### 📥 Download Reconciled Workbook")
            render_download_button(
                file_path=output_path,
                label="📥 Download Updated Reconciliation Workbook (.xlsx)",
                download_filename=f"{target_file.name.rsplit('.', 1)[0]}_eod_synced_{date_str.replace('-', '')}.xlsx",
            )

            # Unmatched Terminals Warning
            if stats["unmatchedSourceTerminals"]:
                st.warning(f"Found {len(stats['unmatchedSourceTerminals'])} terminals in EOD source not present in target workbook:")
                st.write(", ".join(stats["unmatchedSourceTerminals"]))

            st.info("ℹ️ A styled 'EOD Extract Audit' sheet with complete metric details has been appended to the downloaded workbook.")

        except Exception as exc:
            st.error(f"Sync Error: {exc}")
