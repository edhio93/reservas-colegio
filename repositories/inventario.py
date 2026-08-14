"""Repositorio de inventario/equipos."""

from services.supabase import select_paginado, supabase


def listar_equipos():
    return select_paginado("equipos", "*")


def obtener_equipo(equipo_id):
    data = (
        supabase.table("equipos")
        .select("*")
        .eq("id", int(equipo_id))
        .limit(1)
        .execute()
        .data
        or []
    )
    return data[0] if data else None
