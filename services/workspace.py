"""Google Workspace Directory integrado mediante Supabase Edge Functions.

Las credenciales Google NO viven en Streamlit.
Las Edge Functions sincronizan Directory API -> Supabase.
Streamlit consume la copia local de Supabase, rápida y auditable.
"""

from __future__ import annotations

import datetime as dt
import unicodedata

import requests
import streamlit as st

from services.supabase import registrar_error, supabase


def _normalizar_nombre(valor: str) -> str:
    texto = unicodedata.normalize("NFKD", str(valor or ""))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return " ".join(texto.lower().strip().split())


def _function_url(nombre: str) -> str:
    return (
        str(st.secrets["SUPABASE_URL"]).rstrip("/")
        + f"/functions/v1/{nombre}"
    )


def _function_headers() -> dict:
    key = str(st.secrets["SUPABASE_KEY"])
    return {
        "Authorization": f"Bearer {key}",
        "apikey": key,
        "Content-Type": "application/json",
    }


def invocar_edge_function(nombre: str, payload: dict | None = None, timeout=120):
    respuesta = requests.post(
        _function_url(nombre),
        headers=_function_headers(),
        json=payload or {},
        timeout=timeout,
    )
    try:
        data = respuesta.json()
    except Exception:
        data = {"raw": respuesta.text}

    if not respuesta.ok:
        raise RuntimeError(
            f"{nombre} respondió HTTP {respuesta.status_code}: {data}"
        )

    return data


def sincronizar_workspace():
    """Solicita una sincronización inmediata a la Edge Function."""
    return invocar_edge_function(
        "workspace-sync",
        {"source": "streamlit_manual"},
        timeout=150,
    )


@st.cache_data(ttl=300, show_spinner=False)
def listar_workspace_users():
    try:
        return (
            supabase.table("workspace_users")
            .select(
                "google_id,primary_email,full_name,given_name,family_name,"
                "org_unit_path,suspended,is_admin,synced_at"
            )
            .eq("present_in_directory", True)
            .order("full_name")
            .execute()
            .data
            or []
        )
    except Exception as error:
        registrar_error("workspace_users", error)
        return []


@st.cache_data(ttl=300, show_spinner=False)
def listar_workspace_groups():
    try:
        return (
            supabase.table("workspace_groups")
            .select(
                "google_id,email,name,description,direct_members_count,synced_at"
            )
            .eq("present_in_directory", True)
            .order("name")
            .execute()
            .data
            or []
        )
    except Exception as error:
        registrar_error("workspace_groups", error)
        return []


@st.cache_data(ttl=300, show_spinner=False)
def listar_miembros_grupo(group_google_id: str):
    if not group_google_id:
        return []

    try:
        return (
            supabase.table("workspace_group_members")
            .select(
                "member_google_id,member_email,role,type,status,synced_at"
            )
            .eq("group_google_id", group_google_id)
            .order("member_email")
            .execute()
            .data
            or []
        )
    except Exception as error:
        registrar_error("workspace_group_members", error)
        return []


@st.cache_data(ttl=120, show_spinner=False)
def resumen_workspace():
    try:
        usuarios = (
            supabase.table("workspace_users")
            .select("google_id", count="exact")
            .eq("present_in_directory", True)
            .execute()
        )
        grupos = (
            supabase.table("workspace_groups")
            .select("google_id", count="exact")
            .eq("present_in_directory", True)
            .execute()
        )
        profesores = (
            supabase.table("profesores")
            .select(
                "id,nombre,email,workspace_user_id,workspace_primary_email,"
                "workspace_active,workspace_match_method,workspace_last_sync"
            )
            .order("nombre")
            .execute()
            .data
            or []
        )

        vinculados = [
            p
            for p in profesores
            if p.get("workspace_user_id")
            and p.get("workspace_primary_email")
        ]

        sync = (
            supabase.table("workspace_sync_log")
            .select(
                "id,status,source,users_count,groups_count,members_count,"
                "linked_professors_count,error,started_at,finished_at"
            )
            .order("started_at", desc=True)
            .limit(1)
            .execute()
            .data
            or []
        )

        return {
            "users": usuarios.count or 0,
            "groups": grupos.count or 0,
            "professors": len(profesores),
            "linked_professors": len(vinculados),
            "unlinked_professors": len(profesores) - len(vinculados),
            "professor_rows": profesores,
            "last_sync": sync[0] if sync else None,
        }
    except Exception as error:
        registrar_error("workspace_resumen", error)
        return {
            "users": 0,
            "groups": 0,
            "professors": 0,
            "linked_professors": 0,
            "unlinked_professors": 0,
            "professor_rows": [],
            "last_sync": None,
        }


def vincular_profesor_workspace(profesor_id: int, google_user_id: str):
    usuario = (
        supabase.table("workspace_users")
        .select(
            "google_id,primary_email,full_name,org_unit_path,suspended"
        )
        .eq("google_id", google_user_id)
        .single()
        .execute()
        .data
    )

    if not usuario:
        raise ValueError("No se encontró el usuario Workspace seleccionado.")

    ahora = dt.datetime.now(dt.timezone.utc).isoformat()

    (
        supabase.table("profesores")
        .update(
            {
                "email": usuario["primary_email"],
                "workspace_user_id": usuario["google_id"],
                "workspace_primary_email": usuario["primary_email"],
                "workspace_active": not bool(usuario.get("suspended")),
                "workspace_org_unit": usuario.get("org_unit_path"),
                "workspace_match_method": "manual",
                "workspace_last_sync": ahora,
            }
        )
        .eq("id", int(profesor_id))
        .execute()
    )

    st.cache_data.clear()
    return usuario


def desvincular_profesor_workspace(profesor_id: int):
    (
        supabase.table("profesores")
        .update(
            {
                "workspace_user_id": None,
                "workspace_primary_email": None,
                "workspace_active": False,
                "workspace_org_unit": None,
                "workspace_match_method": None,
                "workspace_last_sync": None,
            }
        )
        .eq("id", int(profesor_id))
        .execute()
    )

    st.cache_data.clear()


def sugerir_usuario_para_profesor(nombre_profesor: str, usuarios: list[dict]):
    objetivo = _normalizar_nombre(nombre_profesor)
    exactos = [
        u
        for u in usuarios
        if _normalizar_nombre(u.get("full_name")) == objetivo
    ]
    return exactos[0] if len(exactos) == 1 else None
