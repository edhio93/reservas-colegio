"""Repositorio de tickets/mantenimientos."""

from services.supabase import select_paginado


def listar_tickets():
    return select_paginado(
        "mantenimientos",
        "*",
        orden="fecha",
        desc=True,
    )
