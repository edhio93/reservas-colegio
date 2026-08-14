"""Repositorio de reservas: acceso a datos sin lógica visual."""

from services.supabase import supabase, select_paginado

SELECT_RESERVA = (
    "id,fecha,hora_inicio,hora_fin,observaciones,"
    "profesores(nombre),cursos(nombre),recursos(nombre)"
)


def listar_reservas():
    return select_paginado("reservas", SELECT_RESERVA, orden="fecha")


def listar_reservas_por_fecha(fecha_iso):
    return (
        supabase.table("reservas")
        .select(SELECT_RESERVA)
        .eq("fecha", fecha_iso)
        .order("hora_inicio")
        .execute()
        .data
        or []
    )


def actualizar_reserva(reserva_id, cambios):
    return (
        supabase.table("reservas")
        .update(cambios)
        .eq("id", int(reserva_id))
        .execute()
    )


def eliminar_reserva(reserva_id):
    return (
        supabase.table("reservas")
        .delete()
        .eq("id", int(reserva_id))
        .execute()
    )
