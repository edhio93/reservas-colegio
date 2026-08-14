"""Interfaz futura para Gmail API.

El envío SMTP existente permanece en schedule_app.py durante V24.0.
Este módulo será activado al migrar el Centro de Comunicaciones.
"""

class GmailServiceNotConfigured(RuntimeError):
    pass


def enviar_correo(*args, **kwargs):
    raise GmailServiceNotConfigured("Gmail API aún no está activada en V24.0.")


def enviar_correo_masivo(*args, **kwargs):
    raise GmailServiceNotConfigured("Gmail API aún no está activada en V24.0.")
