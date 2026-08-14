"""Matriz central de páginas y permisos.

V24.0 mantiene el radio de navegación actual para no cambiar la experiencia
visual. En V24.1 esta misma matriz será reutilizada por st.Page/st.navigation.
"""

PAGES_CONFIG = {
    "Inicio": {"icon": "🏠", "roles": ["admin", "profesor", "mensajeria"]},
    "Mis Reservas": {"icon": "👤", "roles": ["profesor"]},
    "Registrar": {"icon": "📝", "roles": ["admin"]},
    "Base de datos": {"icon": "🗃️", "roles": ["admin"]},
    "Semana": {"icon": "🗓️", "roles": ["admin", "profesor"]},
    "Dashboard": {"icon": "📈", "roles": ["admin"]},
    "Técnicos": {"icon": "🔧", "roles": ["admin"]},
    "Diplomas": {"icon": "🎓", "roles": ["admin", "profesor"]},
    "Inventario": {"icon": "💻", "roles": ["admin"]},
    "Mantención preventiva": {"icon": "🧰", "roles": ["admin"]},
    "Auditoría": {"icon": "🧾", "roles": ["admin"]},
    "Configuración": {"icon": "⚙️", "roles": ["admin"]},
    "Modo TV": {"icon": "📺", "roles": ["admin", "mensajeria"]},
}


def get_available_pages(role):
    return [
        page
        for page, config in PAGES_CONFIG.items()
        if role in config["roles"]
    ]


def role_can_access(role, page):
    config = PAGES_CONFIG.get(page)
    return bool(config and role in config["roles"])
