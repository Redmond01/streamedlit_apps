"""Streamlit Page: Station Report Auto-Sync System

Automates the daily rollover of station workbooks:
- Appends a new Day N+1 sheet at index -1
- Increments Sheet Date (+1 day)
- Resets Outer & Inner Meters (Closing -> Opening, Closing = Opening, Sales = 0)
- Resets Tank Dipping (Closing -> Opening, Closing = Opening, Dipping Sales = 0)
- Clears POS list and updates rollover date
- Clears Products Received
- Clears Cash-to-bank, Expenses, and Imprest
- Supports single and batch upload (up to 60+ stations) with ZIP archive packaging.
"""

from __future__ import annotations

from datetime import datetime
import pandas as pd
import streamlit as st

from core.memory import get_memory_usage_mb
from core.report_sync.report_syncer import sync_multiple_station_files
from utils.ui_components import (
    render_header,
    render_memory_badge,
    render_metric_cards,
)

try:
    st.set_page_config(
        page_title="Report Auto-Sync | Station Daily Rollover",
        page_icon="🔄",
        layout="wide",
    )
except Exception:
    pass

render_memory_badge()

render_header(
    title="Station Report Auto-Sync",
    subtitle="Automated Day N+1 sheet creation, meter & tank dipping reset, and POS clearing across all 60+ stations",
    icon="🔄",
)

st.markdown(
    """
    This tool automates the **daily rollover** of station daily report workbooks. For every uploaded workbook, 
    the system accesses the latest daily sheet (index `-1`), creates a new sheet at index `-1`, 
    increments the date, and resets all operational sectors for the new day:
    """
)

with st.expander("ℹ️ The 7 Key Operations Performed on the New Sheet", expanded=False):
    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown(
            """
            1. **Outer Meter (Left Table)**: Prior closing meters become new opening meters, and closing meters match opening so daily sales resolve to `0`.
            2. **Inner Meter (Right Table)**: Audited mirror closing meters become new opening meters, resetting sales to `0`.
            3. **Tank Table**: Prior closing dippings become new opening dippings, and closing dippings match opening so dipping sales & gear resolve to `0`.
            4. **Date**: Sheet date is incremented by `+1 day` (e.g. `10th` $\\rightarrow$ `11th`).
            """
        )
    with col_b:
        st.markdown(
            """
            5. **POS Table**: All terminal amounts (Default day, Rollover day, and Right-side) are cleared to empty. Rollover date is updated to `New Date + 1`.
            6. **Product Received**: Outer and inner received litres are cleared to empty.
            7. **Cash to Bank & Expenses**: Cash transfers, manual expenses, and imprest are cleared, setting remittance formulas to `0`. Available Cash carries over to ADD UP.
            """
        )

with st.form("report_auto_sync_form"):
    st.subheader("1. Upload Station Daily Report Workbooks")
    uploaded_files = st.file_uploader(
        "Upload Station Daily Reports (.xlsx) — drag & drop single or batch of 60+ workbooks",
        type=["xlsx", "xlsm"],
        accept_multiple_files=True,
        help="Select one, several, or all station workbooks. Each file will be rolled over to the next day.",
    )

    st.info("💡 The engine dynamically adapts to varying pump counts (4 to 12) and tank counts across all stations. ADD UP cash balances are statically brought forward without alteration.")

    submit_button = st.form_submit_button(
        "🚀 Run Report Auto-Sync Pipeline",
        type="primary",
        use_container_width=True,
    )

if submit_button:
    if not uploaded_files:
        st.error("Please upload at least one Station Daily Report workbook (.xlsx) before running.")
    else:
        progress_bar = st.progress(0.0)
        status_text = st.empty()

        def on_progress(idx: int, total: int, fname: str) -> None:
            pct = (idx + 1) / total if total > 0 else 1.0
            progress_bar.progress(pct)
            status_text.markdown(f"**Processing:** `{fname}` ({idx + 1}/{total})")

        try:
            with st.spinner("Processing station workbooks rollover..."):
                results, zip_buf, summaries = sync_multiple_station_files(
                    files=uploaded_files,
                    progress_callback=on_progress,
                )

            progress_bar.progress(1.0)
            success_count = sum(1 for s in summaries if s.status == "Success")
            error_count = sum(1 for s in summaries if s.status == "Error")

            status_text.success(
                f"Synchronization complete! {success_count} of {len(uploaded_files)} station workbooks successfully rolled over."
            )

            # Metric Cards
            tot_pumps = sum(s.outer_pumps_updated for s in summaries)
            tot_tanks = sum(s.tanks_updated for s in summaries)
            tot_pos = sum(s.pos_cleared for s in summaries)

            render_metric_cards([
                ("Total Workbooks", len(uploaded_files), None),
                ("Successful", success_count, None),
                ("Errors", error_count, None),
                ("Pumps Reset", tot_pumps, None),
                ("Tanks Reset", tot_tanks, None),
                ("POS Cleared", tot_pos, None),
                ("Current RAM", f"{get_memory_usage_mb():.1f} MB", None),
            ])

            # Download Section
            st.markdown("### 📥 Download Synchronized Workbooks")
            timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")

            if len(results) == 1:
                single_name, single_data = results[0]
                st.download_button(
                    label=f"📥 Download Updated Workbook: {single_name}",
                    data=single_data,
                    file_name=single_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                )
            elif len(results) > 1:
                zip_filename = f"Station_Reports_Synced_{timestamp_str}.zip"
                st.download_button(
                    label=f"📦 Download All {len(results)} Updated Workbooks (.zip)",
                    data=zip_buf.getvalue(),
                    file_name=zip_filename,
                    mime="application/zip",
                    type="primary",
                    use_container_width=True,
                )

                with st.expander("📁 Individual Workbook Downloads", expanded=False):
                    for fname, data in results:
                        st.download_button(
                            label=f"📥 Download {fname}",
                            data=data,
                            file_name=fname,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"dl_{fname}",
                        )

            # Audit Summary Table
            st.markdown("### 📋 Synchronization Log")
            table_data = []
            for s in summaries:
                table_data.append({
                    "Station": s.station_name,
                    "Previous Sheet": s.old_sheet,
                    "New Sheet": s.new_sheet,
                    "Previous Date": s.old_date,
                    "New Date": s.new_date,
                    "Outer Pumps": s.outer_pumps_updated,
                    "Inner Pumps": s.inner_pumps_updated,
                    "Tanks": s.tanks_updated,
                    "POS Cleared": s.pos_cleared,
                    "Status": "✅ Success" if s.status == "Success" else "❌ Error",
                    "Details": s.message,
                })

            df_summary = pd.DataFrame(table_data)
            st.dataframe(df_summary, use_container_width=True, hide_index=True)

        except Exception as exc:
            st.error(f"Execution failed: {exc}")
