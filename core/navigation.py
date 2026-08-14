"""Estado de navegación programática seguro para Streamlit."""

import streamlit as st


def prepare_navigation(available_pages, default_page="Inicio"):
    """Aplica cambios pendientes ANTES de crear el widget de navegación."""
    if not available_pages:
        raise RuntimeError("El rol actual no tiene páginas habilitadas.")

    if default_page not in available_pages:
        default_page = available_pages[0]

    pending = st.session_state.pop("_pending_nav_page", None)
    if pending in available_pages:
        st.session_state.nav_page = pending

    if (
        "nav_page" not in st.session_state
        or st.session_state.nav_page not in available_pages
    ):
        st.session_state.nav_page = default_page

    return st.session_state.nav_page


def request_navigation(page, technical_module=None, open_tv_config=False):
    """Solicita un cambio para el siguiente rerun sin mutar una key ya renderizada."""
    st.session_state["_pending_nav_page"] = page

    if technical_module:
        st.session_state["_pending_tecnico_modulo"] = technical_module

    if open_tv_config:
        st.session_state.ver_pantalla_tv = False

    st.rerun()
