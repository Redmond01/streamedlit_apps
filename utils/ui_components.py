from __future__ import annotations

from pathlib import Path
from typing import Any
import streamlit as st

from core.memory import get_memory_usage_mb


def render_header(title: str, subtitle: str, icon: str = "⚡") -> None:
    """Renders a styled header banner with icon and subtitle."""
    st.markdown(
        f"""
        <div style="padding: 1.2rem 1.5rem; border-radius: 10px; background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%); border-left: 5px solid #0284c7; margin-bottom: 1.5rem; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);">
            <div style="display: flex; align-items: center; gap: 12px;">
                <span style="font-size: 2rem;">{icon}</span>
                <div>
                    <h2 style="margin: 0; color: #f8fafc; font-weight: 700; font-size: 1.5rem;">{title}</h2>
                    <p style="margin: 4px 0 0 0; color: #94a3b8; font-size: 0.95rem;">{subtitle}</p>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_memory_badge() -> None:
    """Displays a memory usage indicator in the sidebar."""
    ram_mb = get_memory_usage_mb()
    if ram_mb < 350:
        color = "#10b981"  # Emerald
        status = "Optimal"
    elif ram_mb < 750:
        color = "#f59e0b"  # Amber
        status = "Moderate"
    else:
        color = "#ef4444"  # Red
        status = "Elevated"

    st.sidebar.markdown(
        f"""
        <div style="padding: 10px 12px; border-radius: 8px; background: #1e293b; border: 1px solid #334155; margin-bottom: 1rem;">
            <div style="display: flex; justify-content: space-between; align-items: center;">
                <span style="font-size: 0.8rem; color: #94a3b8; font-weight: 600;">ACTIVE RAM USAGE</span>
                <span style="font-size: 0.75rem; padding: 2px 6px; border-radius: 4px; background: {color}22; color: {color}; font-weight: 700;">{status}</span>
            </div>
            <div style="font-size: 1.25rem; font-weight: 700; color: #f8fafc; margin-top: 4px;">
                {ram_mb:.1f} <span style="font-size: 0.85rem; font-weight: 400; color: #64748b;">MB</span>
            </div>
            <div style="font-size: 0.75rem; color: #64748b; margin-top: 2px;">
                Strict ceiling: 1,024 MB
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_metric_cards(metrics: list[tuple[str, Any, str | None]]) -> None:
    """Renders a responsive row of st.metric cards."""
    if not metrics:
        return
    cols = st.columns(len(metrics))
    for col, (label, val, delta) in zip(cols, metrics):
        with col:
            st.metric(label=label, value=val, delta=delta)


def render_download_button(
    file_path: Path | str,
    label: str,
    download_filename: str,
    mime: str = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
) -> None:
    """Streams a temporary file to an st.download_button."""
    path = Path(file_path)
    if not path.exists():
        st.error("Output file not found for download.")
        return

    with open(path, "rb") as f:
        file_bytes = f.read()

    st.download_button(
        label=label,
        data=file_bytes,
        file_name=download_filename,
        mime=mime,
        type="primary",
        use_container_width=True,
    )
