"""Interfaz futura para sincronización con Google Calendar."""

class CalendarServiceNotConfigured(RuntimeError):
    pass


def crear_evento_reserva(*args, **kwargs):
    raise CalendarServiceNotConfigured("Google Calendar aún no está activado en V24.0.")


def actualizar_evento_reserva(*args, **kwargs):
    raise CalendarServiceNotConfigured("Google Calendar aún no está activado en V24.0.")


def cancelar_evento_reserva(*args, **kwargs):
    raise CalendarServiceNotConfigured("Google Calendar aún no está activado en V24.0.")
