"""Cola de notificaciones institucionales V24.1.

Streamlit registra el evento y termina rápido.
Supabase Edge Function `process-notifications` realiza el envío Gmail API.
"""

from __future__ import annotations

import datetime as dt
import html
import uuid

from services.supabase import registrar_error, supabase


CAMBIOS_IMPORTANTES = {
    "fecha",
    "hora_inicio",
    "hora_fin",
    "profesor",
    "curso",
    "recurso",
}


def _fmt_fecha(valor):
    try:
        return dt.date.fromisoformat(str(valor)[:10]).strftime("%d/%m/%Y")
    except Exception:
        return str(valor or "—")


def _fmt_hora(valor):
    texto = str(valor or "")
    return texto[:5] if texto else "—"


def _safe(valor):
    return html.escape(str(valor or "—"))


def _obtener_correo_profesor(profesor_id):
    if not profesor_id:
        return None, None

    try:
        fila = (
            supabase.table("profesores")
            .select(
                "id,nombre,email,workspace_primary_email,workspace_active"
            )
            .eq("id", int(profesor_id))
            .single()
            .execute()
            .data
        )
        if not fila:
            return None, None

        correo = (
            fila.get("workspace_primary_email")
            or fila.get("email")
            or ""
        ).strip().lower()

        return correo or None, fila
    except Exception as error:
        registrar_error("obtener_correo_profesor_notificacion", error)
        return None, None


def _insertar_outbox(
    *,
    tipo,
    profesor_id,
    reserva_id,
    recipient_email,
    subject,
    html_body,
    metadata=None,
    available_at=None,
    dedupe_key=None,
):
    payload = {
        "type": tipo,
        "professor_id": int(profesor_id) if profesor_id else None,
        "reservation_id": int(reserva_id) if reserva_id else None,
        "recipient_email": recipient_email,
        "subject": subject,
        "html_body": html_body,
        "metadata": metadata or {},
        "status": "pending",
        "attempts": 0,
        "available_at": (
            available_at
            or dt.datetime.now(dt.timezone.utc).isoformat()
        ),
        "dedupe_key": dedupe_key,
    }

    try:
        return (
            supabase.table("notification_outbox")
            .insert(payload)
            .execute()
            .data
        )
    except Exception as error:
        # Una deduplicación no debe romper la operación principal.
        if "duplicate" not in str(error).lower():
            registrar_error("notification_outbox_insert", error)
        return None


def _actualizar_pendiente_existente(
    *,
    tipo,
    profesor_id,
    reserva_id,
    recipient_email,
    subject,
    html_body,
    metadata,
    available_at,
):
    """Consolida múltiples ediciones de la misma reserva en ~2 minutos."""
    try:
        limite = (
            dt.datetime.now(dt.timezone.utc)
            - dt.timedelta(minutes=5)
        ).isoformat()

        pendientes = (
            supabase.table("notification_outbox")
            .select("id")
            .eq("type", tipo)
            .eq("status", "pending")
            .eq("professor_id", int(profesor_id))
            .eq("reservation_id", int(reserva_id))
            .gte("created_at", limite)
            .order("created_at", desc=True)
            .limit(1)
            .execute()
            .data
            or []
        )

        if not pendientes:
            return False

        (
            supabase.table("notification_outbox")
            .update(
                {
                    "recipient_email": recipient_email,
                    "subject": subject,
                    "html_body": html_body,
                    "metadata": metadata,
                    "available_at": available_at,
                    "updated_at": dt.datetime.now(
                        dt.timezone.utc
                    ).isoformat(),
                }
            )
            .eq("id", pendientes[0]["id"])
            .execute()
        )
        return True
    except Exception as error:
        registrar_error("notification_outbox_consolidar", error)
        return False


def encolar_reservas_creadas_lote(
    *,
    profesor_id,
    profesor_nombre,
    curso,
    recursos,
    fechas,
    hora_inicio,
    hora_fin,
):
    """Un solo correo para una creación simple o recurrente."""
    correo, profesor = _obtener_correo_profesor(profesor_id)
    if not correo:
        return False

    fechas_limpias = [str(f)[:10] for f in (fechas or [])]
    recursos_limpios = [str(r) for r in (recursos or [])]

    lista_fechas = "".join(
        f"<li>{_safe(_fmt_fecha(fecha))}</li>"
        for fecha in fechas_limpias
    ) or "<li>—</li>"

    subject = f"✅ Reserva de Enlaces confirmada · {curso}"

    body = f"""
    <html>
    <body style="font-family:Arial,sans-serif;color:#172033;line-height:1.55">
        <h2 style="color:#800020">Reserva de Enlaces confirmada</h2>
        <p>Hola {_safe((profesor or {}).get('nombre') or profesor_nombre)},</p>
        <p>Se registró la siguiente reserva a tu nombre:</p>
        <table cellpadding="7" style="border-collapse:collapse">
            <tr><td><b>Horario</b></td><td>{_safe(_fmt_hora(hora_inicio))}–{_safe(_fmt_hora(hora_fin))}</td></tr>
            <tr><td><b>Curso</b></td><td>{_safe(curso)}</td></tr>
            <tr><td><b>Recurso(s)</b></td><td>{_safe(', '.join(recursos_limpios))}</td></tr>
        </table>
        <p><b>Fecha(s):</b></p>
        <ul>{lista_fechas}</ul>
        <p style="margin-top:20px">Departamento de Informática / Enlaces<br>
        Liceo Bicentenario de Excelencia Colegio Antonio Varas</p>
    </body>
    </html>
    """

    _insertar_outbox(
        tipo="reservation_created_batch",
        profesor_id=profesor_id,
        reserva_id=None,
        recipient_email=correo,
        subject=subject,
        html_body=body,
        metadata={
            "source": "streamlit",
            "fechas": fechas_limpias,
            "recursos": recursos_limpios,
        },
        dedupe_key=f"created-batch:{uuid.uuid4()}",
    )
    return True


def encolar_reserva_creada(
    *,
    profesor_id,
    reserva_id=None,
    profesor_nombre,
    curso,
    recurso,
    fecha,
    hora_inicio,
    hora_fin,
):
    correo, profesor = _obtener_correo_profesor(profesor_id)
    if not correo:
        return False

    subject = f"✅ Reserva de Enlaces confirmada · {curso}"

    body = f"""
    <html>
    <body style="font-family:Arial,sans-serif;color:#172033;line-height:1.55">
        <h2 style="color:#800020">Reserva de Enlaces confirmada</h2>
        <p>Hola {_safe((profesor or {}).get("nombre") or profesor_nombre)},</p>
        <p>Se registró una reserva a tu nombre:</p>
        <table cellpadding="7" style="border-collapse:collapse">
            <tr><td><b>Fecha</b></td><td>{_safe(_fmt_fecha(fecha))}</td></tr>
            <tr><td><b>Horario</b></td><td>{_safe(_fmt_hora(hora_inicio))}–{_safe(_fmt_hora(hora_fin))}</td></tr>
            <tr><td><b>Curso</b></td><td>{_safe(curso)}</td></tr>
            <tr><td><b>Recurso</b></td><td>{_safe(recurso)}</td></tr>
        </table>
        <p style="margin-top:20px">Departamento de Informática / Enlaces<br>
        Liceo Bicentenario de Excelencia Colegio Antonio Varas</p>
    </body>
    </html>
    """

    _insertar_outbox(
        tipo="reservation_created",
        profesor_id=profesor_id,
        reserva_id=reserva_id,
        recipient_email=correo,
        subject=subject,
        html_body=body,
        metadata={"source": "streamlit"},
        dedupe_key=(
            f"created:{reserva_id}:{correo}"
            if reserva_id
            else None
        ),
    )
    return True


def encolar_reserva_modificada(
    *,
    profesor_id,
    reserva_id,
    profesor_nombre,
    before,
    after,
):
    cambios = {
        campo: {
            "before": before.get(campo),
            "after": after.get(campo),
        }
        for campo in CAMBIOS_IMPORTANTES
        if str(before.get(campo)) != str(after.get(campo))
    }

    if not cambios:
        return False

    correo, profesor = _obtener_correo_profesor(profesor_id)
    if not correo:
        return False

    filas_cambios = "".join(
        f"""
        <tr>
          <td style="padding:7px"><b>{_safe(campo.replace('_', ' ').title())}</b></td>
          <td style="padding:7px;color:#8b1e2d">{_safe(datos['before'])}</td>
          <td style="padding:7px;color:#137333">{_safe(datos['after'])}</td>
        </tr>
        """
        for campo, datos in cambios.items()
    )

    subject = "🔔 Cambio en tu reserva de Enlaces"

    body = f"""
    <html>
    <body style="font-family:Arial,sans-serif;color:#172033;line-height:1.55">
        <h2 style="color:#800020">Tu reserva de Enlaces fue actualizada</h2>
        <p>Hola {_safe((profesor or {}).get("nombre") or profesor_nombre)},</p>
        <p>Se modificó una reserva registrada a tu nombre.</p>

        <table border="1" cellpadding="0"
               style="border-collapse:collapse;border-color:#e5e7eb;width:100%">
            <thead>
              <tr style="background:#f8fafc">
                <th style="padding:7px;text-align:left">Dato</th>
                <th style="padding:7px;text-align:left">Antes</th>
                <th style="padding:7px;text-align:left">Ahora</th>
              </tr>
            </thead>
            <tbody>{filas_cambios}</tbody>
        </table>

        <p style="margin-top:20px">
        Departamento de Informática / Enlaces<br>
        Liceo Bicentenario de Excelencia Colegio Antonio Varas
        </p>
    </body>
    </html>
    """

    available_at = (
        dt.datetime.now(dt.timezone.utc)
        + dt.timedelta(minutes=2)
    ).isoformat()

    metadata = {
        "source": "streamlit",
        "changes": cambios,
        "consolidation_window_minutes": 2,
    }

    consolidado = _actualizar_pendiente_existente(
        tipo="reservation_changed",
        profesor_id=profesor_id,
        reserva_id=reserva_id,
        recipient_email=correo,
        subject=subject,
        html_body=body,
        metadata=metadata,
        available_at=available_at,
    )

    if not consolidado:
        _insertar_outbox(
            tipo="reservation_changed",
            profesor_id=profesor_id,
            reserva_id=reserva_id,
            recipient_email=correo,
            subject=subject,
            html_body=body,
            metadata=metadata,
            available_at=available_at,
        )

    return True


def encolar_reserva_cancelada(
    *,
    profesor_id,
    reserva_id,
    profesor_nombre,
    curso,
    recurso,
    fecha,
    hora_inicio,
    hora_fin,
    motivo="Reserva eliminada desde el sistema",
):
    correo, profesor = _obtener_correo_profesor(profesor_id)
    if not correo:
        return False

    subject = "🗑️ Reserva de Enlaces cancelada"

    body = f"""
    <html>
    <body style="font-family:Arial,sans-serif;color:#172033;line-height:1.55">
        <h2 style="color:#800020">Reserva de Enlaces cancelada</h2>
        <p>Hola {_safe((profesor or {}).get("nombre") or profesor_nombre)},</p>
        <p>La siguiente reserva dejó de estar vigente:</p>
        <table cellpadding="7" style="border-collapse:collapse">
            <tr><td><b>Fecha</b></td><td>{_safe(_fmt_fecha(fecha))}</td></tr>
            <tr><td><b>Horario</b></td><td>{_safe(_fmt_hora(hora_inicio))}–{_safe(_fmt_hora(hora_fin))}</td></tr>
            <tr><td><b>Curso</b></td><td>{_safe(curso)}</td></tr>
            <tr><td><b>Recurso</b></td><td>{_safe(recurso)}</td></tr>
        </table>
        <p><b>Motivo:</b> {_safe(motivo)}</p>
        <p style="margin-top:20px">
        Departamento de Informática / Enlaces<br>
        Liceo Bicentenario de Excelencia Colegio Antonio Varas
        </p>
    </body>
    </html>
    """

    _insertar_outbox(
        tipo="reservation_cancelled",
        profesor_id=profesor_id,
        reserva_id=reserva_id,
        recipient_email=correo,
        subject=subject,
        html_body=body,
        metadata={"source": "streamlit"},
        dedupe_key=f"cancelled:{reserva_id}:{correo}",
    )
    return True


def listar_outbox_reciente(limit=100):
    try:
        return (
            supabase.table("notification_outbox")
            .select(
                "id,type,recipient_email,subject,status,attempts,error,"
                "created_at,available_at,sent_at"
            )
            .order("created_at", desc=True)
            .limit(limit)
            .execute()
            .data
            or []
        )
    except Exception as error:
        registrar_error("notification_outbox_list", error)
        return []
