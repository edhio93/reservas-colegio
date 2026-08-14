"""Repositorio de profesores."""

from services.supabase import supabase


def listar_profesores():
    return (
        supabase.table("profesores")
        .select("id,nombre,email")
        .order("nombre")
        .execute()
        .data
        or []
    )
