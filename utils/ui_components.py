from __future__ import annotations

import streamlit as st


def section_header(title: str, subtitle: str) -> None:
    st.markdown(
        """
        <style>
        .js-section-title {font-size: 2rem; font-weight: 750; color: #0f172a; margin-bottom: .25rem;}
        .js-section-subtitle {font-size: 1rem; color: #475569; line-height: 1.45; max-width: 980px; margin-bottom: 1.25rem;}
        .js-feature-card {border: 1px solid #e5e7eb; border-radius: 8px; padding: 16px; background: #fff; min-height: 118px;}
        .js-feature-title {font-weight: 700; color: #0f172a; margin-bottom: 8px;}
        .js-feature-copy {color: #475569; font-size: .92rem; line-height: 1.4;}
        </style>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(f"<div class='js-section-title'>{title}</div>", unsafe_allow_html=True)
    st.markdown(f"<div class='js-section-subtitle'>{subtitle}</div>", unsafe_allow_html=True)


def feature_grid(items: list[tuple[str, str]]) -> None:
    for start in range(0, len(items), 3):
        cols = st.columns(3)
        for col, (title, copy) in zip(cols, items[start : start + 3]):
            with col:
                st.markdown(
                    f"""
                    <div class='js-feature-card'>
                        <div class='js-feature-title'>{title}</div>
                        <div class='js-feature-copy'>{copy}</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
