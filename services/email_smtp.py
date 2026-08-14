"""Servicio SMTP centralizado del Sistema CAV.

Las credenciales se leen exclusivamente desde Streamlit Secrets.
Este módulo nunca imprime ni devuelve la contraseña SMTP.
"""

import re
import smtplib
import ssl
from email.message import EmailMessage
from email.utils import formataddr

import streamlit as st


_EMAIL_RE = re.compile(
    r"^[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@"
    r"[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?"
    r"(?:\.[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?)+$"
)

_DOMINIO_INSTITUCIONAL = "colegioantoniovaras.cl"

# Estructura institucional informada por el colegio:
# nombre.apellidopaterno.inicialapellidomaterno@colegioantoniovaras.cl
_EMAIL_INSTITUCIONAL_RE = re.compile(
    r"^[A-Za-z0-9_-]+\."
    r"[A-Za-z0-9_-]+\."
    r"[A-Za-z0-9_-]+@"
    r"colegioantoniovaras\.cl$",
    re.IGNORECASE,
)


def get_email_config():
    """Lee y normaliza [email_credentials] sin exponer secretos."""
    try:
        raw = st.secrets["email_credentials"]
    except Exception as exc:
        raise RuntimeError(
            "No existe la sección [email_credentials] en Streamlit Secrets."
        ) from exc

    smtp_server = str(raw.get("smtp_server", "smtp.gmail.com")).strip()
    smtp_port = int(raw.get("smtp_port", 587))
    smtp_username = str(raw.get("smtp_username", "")).strip()
    smtp_password = str(raw.get("smtp_password", "")).replace(" ", "")
    sender_email = str(raw.get("sender_email", smtp_username)).strip()
    sender_name = str(
        raw.get(
            "sender_name",
            "Liceo Bicentenario de Excelencia Colegio Antonio Varas",
        )
    ).strip()
    reply_to = str(raw.get("reply_to", sender_email)).strip()
    use_tls = bool(raw.get("use_tls", smtp_port != 465))
    use_ssl = bool(raw.get("use_ssl", smtp_port == 465))

    return {
        "smtp_server": smtp_server,
        "smtp_port": smtp_port,
        "smtp_username": smtp_username,
        "smtp_password": smtp_password,
        "sender_email": sender_email,
        "sender_name": sender_name,
        "reply_to": reply_to,
        "use_tls": use_tls,
        "use_ssl": use_ssl,
    }


def validate_email_config():
    """Retorna (ok, detalle) sin revelar la contraseña."""
    try:
        cfg = get_email_config()
    except Exception as exc:
        return False, str(exc)

    faltantes = []
    if not cfg["smtp_server"]:
        faltantes.append("smtp_server")
    if not cfg["sender_email"]:
        faltantes.append("sender_email/smtp_username")

    if cfg["smtp_server"].lower() == "smtp.gmail.com":
        if not cfg["smtp_username"]:
            faltantes.append("smtp_username")
        if not cfg["smtp_password"]:
            faltantes.append("smtp_password")

    if faltantes:
        return False, "Faltan: " + ", ".join(faltantes)

    return (
        True,
        f'{cfg["smtp_server"]}:{cfg["smtp_port"]} · remitente {cfg["sender_email"]}',
    )


def _normalizar_destinatario(destinatario):
    """Normaliza caracteres invisibles comunes al copiar/pegar un correo."""
    correo = str(destinatario or "").strip()

    for invisible in ("\u200b", "\u200c", "\u200d", "\ufeff", "\u00a0"):
        correo = correo.replace(invisible, "")

    return correo.strip()


def _validar_destinatario(destinatario):
    correo = _normalizar_destinatario(destinatario)

    if not correo:
        raise ValueError("Escribe un correo destinatario.")

    if any(caracter.isspace() for caracter in correo):
        raise ValueError(
            "El correo contiene espacios. "
            "Ejemplo institucional: nombre.apellido.i@colegioantoniovaras.cl"
        )

    if not _EMAIL_RE.fullmatch(correo):
        raise ValueError(
            "El correo destinatario no tiene un formato válido. "
            "Ejemplo: nombre.apellido.i@colegioantoniovaras.cl"
        )

    return correo.lower()


def es_correo_institucional(correo):
    """Indica si coincide con el dominio y estructura institucional CAV."""
    correo_normalizado = _normalizar_destinatario(correo).lower()
    return bool(_EMAIL_INSTITUCIONAL_RE.fullmatch(correo_normalizado))


def send_html_email(subject, html_body, recipient_email, text_body=None):
    """Envía un correo HTML usando [email_credentials].

    Lanza una excepción descriptiva si falla para que la interfaz pueda
    informar el problema en vez de ocultarlo.
    """
    destinatario = _validar_destinatario(recipient_email)
    cfg = get_email_config()
    ok, detalle = validate_email_config()
    if not ok:
        raise RuntimeError(detalle)

    mensaje = EmailMessage()
    mensaje["Subject"] = str(subject)
    mensaje["From"] = formataddr((cfg["sender_name"], cfg["sender_email"]))
    mensaje["To"] = destinatario
    if cfg["reply_to"]:
        mensaje["Reply-To"] = cfg["reply_to"]

    texto = text_body or "Mensaje enviado desde el Sistema CAV."
    mensaje.set_content(texto)
    mensaje.add_alternative(str(html_body), subtype="html")

    contexto_ssl = ssl.create_default_context()

    if cfg["use_ssl"] or cfg["smtp_port"] == 465:
        with smtplib.SMTP_SSL(
            cfg["smtp_server"],
            cfg["smtp_port"],
            context=contexto_ssl,
            timeout=30,
        ) as servidor:
            if cfg["smtp_username"] and cfg["smtp_password"]:
                servidor.login(
                    cfg["smtp_username"],
                    cfg["smtp_password"],
                )
            servidor.send_message(mensaje)
    else:
        with smtplib.SMTP(
            cfg["smtp_server"],
            cfg["smtp_port"],
            timeout=30,
        ) as servidor:
            servidor.ehlo()
            if cfg["use_tls"]:
                servidor.starttls(context=contexto_ssl)
                servidor.ehlo()
            if cfg["smtp_username"] and cfg["smtp_password"]:
                servidor.login(
                    cfg["smtp_username"],
                    cfg["smtp_password"],
                )
            servidor.send_message(mensaje)

    return True


def send_test_email(recipient_email):
    """Envía un mensaje de diagnóstico reconocible desde Configuración."""
    html = """
    <div style="font-family:Arial,sans-serif;max-width:620px;margin:auto;line-height:1.55">
      <div style="background:#800020;color:white;padding:20px 24px;border-radius:16px 16px 0 0">
        <div style="font-size:22px;font-weight:800">✅ Prueba de correo Sistema CAV</div>
        <div style="opacity:.9">Departamento de Informática / Enlaces</div>
      </div>
      <div style="border:1px solid #e5e7eb;border-top:0;padding:24px;border-radius:0 0 16px 16px">
        <p>Este mensaje confirma que el servicio SMTP del Sistema CAV está funcionando correctamente.</p>
        <p><strong>Resultado:</strong> conexión, autenticación y envío completados.</p>
        <p style="color:#64748b">No necesitas responder este correo.</p>
      </div>
    </div>
    """
    texto = (
        "Prueba de correo Sistema CAV. "
        "La conexión, autenticación y envío SMTP funcionaron correctamente."
    )
    return send_html_email(
        "✅ Prueba de correo · Sistema CAV",
        html,
        recipient_email,
        text_body=texto,
    )
