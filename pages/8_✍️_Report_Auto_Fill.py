"""Streamlit Page: Station Report Auto-Fill System

Automatically copies operational data from Outer Meter (Left side) to Inner Meter (Right side):
1. Tank Dipping (Opening & Closing Dippings)
2. POS List (Matches terminal names from Left side to Right side)
3. Products Received (Litres received)
4. Expenses (Descriptions and amounts, including Imprest)
Supports single and batch processing (up to 60+ workbooks) with ZIP archive packaging.
"""

from __future__ import annotations

from datetime import datetime
import pandas as pd
import streamlit as st

from core.memory import get_memory_usage_mb
from core.report_sync.report_auto_fill import auto_fill_multiple_station_files
from utils.ui_components import (
    render_header,
    render_memory_badge,
    render_metric_cards,
)

try:
    st.set_page_config(
        page_title="Report Auto-Fill | Outer to Inner Sync",
        page_icon="✍️",
        layout="wide",
    )
except Exception:
    pass

render_memory_badge()

render_header(
    title="Station Report Auto-Fill",
    subtitle="Copy filled Outer Meter (Left Side) operational data to Inner Meter (Right Side) across 60+ stations",
    icon="✍️",
)

st.markdown(
    """
    Once a daily report has been filled on the **Outer Meter (Left Side)** by the station, this tool automatically
    mirrors the corresponding operational data to the **Inner Meter (Right Side)**:
    """
)

with st.expander("ℹ️ The 4 Sectors Copied from Left (Outer) to Right (Inner)", expanded=False):
    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown(
            """
            1. **Tank Dipping**: Copies Opening & Closing Dippings directly to the Inner Dipping table, automatically calculating inner sales, gear, and differences.
            2. **POS List**: Accurately maps each terminal header (e.g. `pos3`, `SGR 358`, `cash/transfer`) from Left side to Right side.
            """
        )
    with col_b:
        st.markdown(
            """
            3. **Products Received**: Copies received fuel volumes (PMS & AGO) to the Inner Product Received table.
            4. **Expenses & Imprest**: Copies expense item descriptions and amounts, as well as imprest figures, to the inner audit table.
            """
        )

with st.form("report_auto_fill_form"):
    st.subheader("1. Upload Filled Station Daily Report Workbooks")
    uploaded_files = st.file_uploader(
        "Upload Station Daily Reports (.xlsx) — drag & drop single or batch of 60+ workbooks",
        type=["xlsx", "xlsm"],
        accept_multiple_files=True,
        help="Upload the station reports that have been filled on the left side.",
    )

    st.subheader("2. Target Sheet Selection")
    col1, col2 = st.columns(2)
    with col1:
        target_mode = st.radio(
            "📅 Target Sheet to Fill",
            options=["Latest Sheet Only (Last Day)", "All Sheets in Workbook", "Specific Sheet Name"],
            index=0,
            help="Choose whether to mirror data on the most recent daily sheet, all sheets, or a specific sheet.",
        )
    with col2:
        specific_sheet_name = None
        if target_mode == "Specific Sheet Name":
            specific_sheet_name = st.text_input("Enter Worksheet Name", value="Sheet10")
        else:
            st.info("💡 By default, auto-fill targets the latest sheet (last day) in each workbook.")

    submit_button = st.form_submit_button(
        "🚀 Run Report Auto-Fill Pipeline",
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

        mode_code = "latest"
        if target_mode == "All Sheets in Workbook":
            mode_code = "all"
        elif target_mode == "Specific Sheet Name":
            mode_code = "specific"

        try:
            with st.spinner("Processing outer-to-inner auto-fill..."):
                results, zip_buf, summaries = auto_fill_multiple_station_files(
                    files=uploaded_files,
                    target_sheet_mode=mode_code,
                    specific_sheet_name=specific_sheet_name.strip() if specific_sheet_name else None,
                    progress_callback=on_progress,
                )

            progress_bar.progress(1.0)
            success_count = sum(1 for s in summaries if s.status == "Success")
            error_count = sum(1 for s in summaries if s.status == "Error")

            status_text.success(
                f"Auto-fill complete! Successfully processed {len(results)} of {len(uploaded_files)} workbooks."
            )

            # Metric Cards
            tot_tanks = sum(s.tanks_copied for s in summaries)
            tot_pos = sum(s.pos_copied for s in summaries)
            tot_prod = sum(s.products_copied for s in summaries)
            tot_exp = sum(s.expenses_copied for s in summaries)
            tot_pumps = sum(s.pumps_synced for s in summaries)

            render_metric_cards([
                ("Workbooks Processed", len(uploaded_files), None),
                ("Pumps Synced/Healed", tot_pumps, None),
                ("Tanks Copied", tot_tanks, None),
                ("POS Terminals Copied", tot_pos, None),
                ("Products Copied", tot_prod, None),
                ("Expenses Copied", tot_exp, None),
                ("Current RAM", f"{get_memory_usage_mb():.1f} MB", None),
            ])

            # Download Section
            st.markdown("### 📥 Download Filled Workbooks")
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
                zip_filename = f"Station_Reports_AutoFilled_{timestamp_str}.zip"
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
                            key=f"dl_filled_{fname}",
                        )

            # Audit Summary Table
            st.markdown("### 📋 Auto-Fill Audit Log")
            table_data = []
            for s in summaries:
                table_data.append({
                    "Station": s.station_name,
                    "Sheet": s.sheet_name,
                    "Date": s.date_str,
                    "Pumps Synced": s.pumps_synced,
                    "Tanks Copied": s.tanks_copied,
                    "POS Copied": s.pos_copied,
                    "Products Copied": s.products_copied,
                    "Expenses Copied": s.expenses_copied,
                    "Status": "✅ Success" if s.status == "Success" else "❌ Error",
                    "Details": s.message,
                })

            df_summary = pd.DataFrame(table_data)
            st.dataframe(df_summary, use_container_width=True, hide_index=True)

        except Exception as exc:
            st.error(f"Execution failed: {exc}")
