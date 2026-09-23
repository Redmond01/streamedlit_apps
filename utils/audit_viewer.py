from __future__ import annotations

from typing import Any
import pandas as pd
import streamlit as st


def render_audit_table(logs: list[dict[str, Any]], title: str = "Execution Log Preview") -> None:
    """Renders a formatted, searchable DataFrame of operation logs."""
    if not logs:
        st.info("No audit logs recorded for this operation.")
        return

    df = pd.DataFrame(logs)
    
    with st.expander(f"📋 {title} ({len(logs)} entries)", expanded=True):
        st.dataframe(
            df,
            use_container_width=True,
            hide_index=True,
        )
