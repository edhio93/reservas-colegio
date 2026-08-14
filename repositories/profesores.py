"""Repositorio de profesores con vínculo Workspace."""

from services.supabase import supabase


def listar_profesores():
    try:
        return (
            supabase.table("profesores")
            .select(
                "id,nombre,email,workspace_user_id,workspace_primary_email,"
                "workspace_active,workspace_org_unit,workspace_match_method,"
                "workspace_last_sync"
            )
            .order("nombre")
            .execute()
            .data
            or []
        )
    except Exception:
        return (
            supabase.table("profesores")
            .select("id,nombre,email")
            .order("nombre")
            .execute()
            .data
            or []
        )
