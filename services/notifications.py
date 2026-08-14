"""Frontera de notificaciones institucionales.

V24.0 no cambia el envío existente; prepara una única API para V24.x.
"""

def build_reservation_change_payload(before, after):
    return {
        "type": "reservation_changed",
        "before": before,
        "after": after,
    }


def build_reservation_cancelled_payload(reservation):
    return {
        "type": "reservation_cancelled",
        "reservation": reservation,
    }
