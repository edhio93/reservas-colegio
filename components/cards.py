"""Componentes visuales reutilizables (se migrarán progresivamente)."""

import streamlit as st


def action_card(title, caption, button_label, key, on_click=None, primary=False):
    with st.container(border=True):
        st.markdown(f"#### {title}")
        if caption:
            st.caption(caption)
        return st.button(
            button_label,
            key=key,
            type="primary" if primary else "secondary",
            use_container_width=True,
            on_click=on_click,
        )
