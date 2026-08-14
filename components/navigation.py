"""Helpers visuales de navegación.

V24.0 conserva la UI actual. La conversión completa a st.navigation se
realizará después de separar las páginas del monolito.
"""

from core.permissions import PAGES_CONFIG


def page_label(page):
    config = PAGES_CONFIG.get(page, {})
    return f"{config.get('icon', '•')} {page}"
