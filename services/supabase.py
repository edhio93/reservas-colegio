"""Acceso centralizado a Supabase y utilidades transversales."""

import datetime as dt
import logging
import traceback

import streamlit as st
from supabase import create_client, Client, ClientOptions

LOGGER = logging.getLogger("sistema_cav")


@st.cache_resource(show_spinner=False)
def get_supabase_client(url_supabase, clave_supabase):
    options = ClientOptions(
        postgrest_client_timeout=45,
        storage_client_timeout=60,
    )
    return create_client(url_supabase, clave_supabase, options=options)


def _load_client():
    try:
        url = st.secrets["SUPABASE_URL"]
        key = st.secrets["SUPABASE_KEY"]
    except KeyError as error:
        st.error(f"🚨 Falta configurar {error} en los Secrets de Streamlit.")
        st.stop()
    return get_supabase_client(url, key)


supabase: Client = _load_client()


def registrar_error(contexto, error):
    LOGGER.error("%s: %s\n%s", contexto, error, traceback.format_exc())
    st.session_state.setdefault("errores_sistema", [])
    st.session_state.errores_sistema.append({
        "fecha": dt.datetime.now().isoformat(timespec="seconds"),
        "contexto": contexto,
        "error": str(error),
    })
    st.session_state.errores_sistema = st.session_state.errores_sistema[-50:]


def select_paginado(
    tabla,
    columnas="*",
    filtros=None,
    orden=None,
    desc=False,
    pagina=1000,
):
    """Lee todos los registros evitando el límite por defecto de Supabase."""
    rows = []
    start = 0

    while True:
        query = supabase.table(tabla).select(columnas)

        for method, field, value in (filtros or []):
            query = getattr(query, method)(field, value)

        if orden:
            query = query.order(orden, desc=desc)

        block = (
            query.range(start, start + pagina - 1)
            .execute()
            .data
            or []
        )
        rows.extend(block)

        if len(block) < pagina:
            break

        start += pagina

    return rows


def registrar_auditoria(accion, modulo, registro_id=None, detalle=None):
    """Registra una acción sin interrumpir la app si auditoría falla."""
    try:
        supabase.table("auditoria").insert({
            "usuario": (
                st.session_state.get("profesor_name")
                or st.session_state.get("role")
                or "sistema"
            ),
            "accion": accion,
            "modulo": modulo,
            "registro_id": (
                str(registro_id)
                if registro_id is not None
                else None
            ),
            "detalle": detalle or {},
            "fecha": dt.datetime.now().isoformat(),
        }).execute()
    except Exception as error:
        registrar_error("auditoria", error)
