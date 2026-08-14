"""Configuración estable del Sistema CAV.

Este módulo NO contiene secretos. Las credenciales continúan en
Streamlit Secrets y son leídas por los servicios correspondientes.
"""

APP_NAME = "Sistema CAV"
APP_VERSION = "24.1.0"
APP_TITLE = "Sistema de Horarios CAV"
APP_ICON = "📅"
APP_LAYOUT = "wide"
APP_SIDEBAR_STATE = "expanded"
TIMEZONE = "America/Santiago"

CACHE_TTL = {
    "reservas": 120,
    "catalogos": 600,
    "login": 600,
    "inicio": 90,
    "clima": 1800,
}
