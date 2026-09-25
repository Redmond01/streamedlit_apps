from __future__ import annotations

import sys
import platform
import streamlit as st
import openpyxl
import pandas as pd

from core.memory import force_gc, get_memory_usage_mb
from utils.ui_components import render_header, render_memory_badge

st.set_page_config(
    page_title="Excel Automation & Reconciliation Suite",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

render_memory_badge()

render_header(
    title="Excel Automation & Reconciliation Suite",
    subtitle="High-performance, low-memory batch processing and POS reconciliation for cloud deployment",
    icon="⚡",
)

st.markdown(
    """
    Welcome to the cloud-optimized **Excel Automation & Reconciliation Suite**. This application consolidates 
    previously CLI-driven automation scripts into an intuitive, low-memory multi-page dashboard engineered to operate 
    comfortably under a **1.0 GB RAM ceiling**.
    """
)

st.markdown("### 🛠️ Available Automation Tools")

col1, col2 = st.columns(2)

with col1:
    st.markdown(
        """
        <div style="background: #1e293b; padding: 1.25rem; border-radius: 8px; border: 1px solid #334155; height: 100%;">
            <div style="display: flex; align-items: center; gap: 8px; margin-bottom: 8px;">
                <span style="font-size: 1.5rem;">⛽</span>
                <h4 style="margin: 0; color: #38bdf8;">1. Gear Automation</h4>
            </div>
            <p style="color: #94a3b8; font-size: 0.9rem; margin-bottom: 12px;">
                Batch process 50+ station sales workbooks simultaneously. Extracts dipping & meter sales, resolves station column mappings, calculates running averages, and compiles master total formulas.
            </p>
            <ul style="color: #cbd5e1; font-size: 0.85rem; padding-left: 20px;">
                <li>Iterative disk spooling (flat memory usage)</li>
                <li>Fuzzy station name matching</li>
                <li>Excel running average formula updates</li>
            </ul>
        </div>
        """,
        unsafe_allow_html=True,
    )

with col2:
    st.markdown(
        """
        <div style="background: #1e293b; padding: 1.25rem; border-radius: 8px; border: 1px solid #334155; height: 100%;">
            <div style="display: flex; align-items: center; gap: 8px; margin-bottom: 8px;">
                <span style="font-size: 1.5rem;">📋</span>
                <h4 style="margin: 0; color: #38bdf8;">2. EOD JONS EXTRACTOR</h4>
            </div>
            <p style="color: #94a3b8; font-size: 0.9rem; margin-bottom: 12px;">
                Scans reconciliation cashbook workbooks to extract active POS terminals (SGR, OBX, AGS, TA) into a clean JSON loop configuration for Moniepoint Chrome Extension automation.
            </p>
            <ul style="color: #cbd5e1; font-size: 0.85rem; padding-left: 20px;">
                <li>Automatic date & report block discovery</li>
                <li>Configurable POS skip threshold (default: skips 1-10)</li>
                <li>One-click JSON copy & file download</li>
            </ul>
        </div>
        """,
        unsafe_allow_html=True,
    )

st.markdown("<div style='height: 16px;'></div>", unsafe_allow_html=True)

col3, col4 = st.columns(2)

with col3:
    st.markdown(
        """
        <div style="background: #1e293b; padding: 1.25rem; border-radius: 8px; border: 1px solid #334155; height: 100%;">
            <div style="display: flex; align-items: center; gap: 8px; margin-bottom: 8px;">
                <span style="font-size: 1.5rem;">📥</span>
                <h4 style="margin: 0; color: #38bdf8;">3. EOD API LOADER</h4>
            </div>
            <p style="color: #94a3b8; font-size: 0.9rem; margin-bottom: 12px;">
                Ingests Moniepoint daily/monthly EOD transaction dumps and syncs amounts into the CASHBOOK column of your master reconciliation workbook.
            </p>
            <ul style="color: #cbd5e1; font-size: 0.85rem; padding-left: 20px;">
                <li>Automated formula generation (CASHBOOK - BANK)</li>
                <li>Generates formatted 'EOD Extract Audit' sheet</li>
                <li>Preserves custom styles, fonts, and formulas</li>
            </ul>
        </div>
        """,
        unsafe_allow_html=True,
    )

with col4:
    st.markdown(
        """
        <div style="background: #1e293b; padding: 1.25rem; border-radius: 8px; border: 1px solid #334155; height: 100%;">
            <div style="display: flex; align-items: center; gap: 8px; margin-bottom: 8px;">
                <span style="font-size: 1.5rem;">🏦</span>
                <h4 style="margin: 0; color: #38bdf8;">4. Bank Statement POS Loader</h4>
            </div>
            <p style="color: #94a3b8; font-size: 0.9rem; margin-bottom: 12px;">
                Loads Chrome extension POS JSON exports into the target reconciliation workbook's BANK STATEMENT column, injecting cell comments for failed records.
            </p>
            <ul style="color: #cbd5e1; font-size: 0.85rem; padding-left: 20px;">
                <li>Automatic report date inference from JSON runDate</li>
                <li>Detailed match statistics (OK, Failed, Missing)</li>
                <li>Hover comments with POS error details</li>
            </ul>
        </div>
        """,
        unsafe_allow_html=True,
    )

st.markdown("---")

st.markdown("### 🖥️ System Health & Cloud Constraints")

diag_col1, diag_col2, diag_col3, diag_col4 = st.columns(4)

with diag_col1:
    st.metric("Process RSS RAM", f"{get_memory_usage_mb():.1f} MB", help="Real-time resident memory used by this container.")
with diag_col2:
    st.metric("Python Version", platform.python_version())
with diag_col3:
    st.metric("Openpyxl Version", openpyxl.__version__)
with diag_col4:
    st.metric("Pandas Version", pd.__version__)

st.sidebar.markdown("---")
if st.sidebar.button("🧹 Run Garbage Collection", use_container_width=True):
    force_gc()
    st.sidebar.success("Garbage collection executed.")
    st.rerun()
