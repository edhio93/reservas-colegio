"""Contrato de integración con Google Workspace.

V24.0 define la frontera del servicio pero NO activa Admin SDK todavía.
La implementación real llegará en Workspace Core para no introducir
credenciales/permisos nuevos durante la primera migración estructural.
"""

class WorkspaceNotConfigured(RuntimeError):
    pass


def listar_usuarios(*args, **kwargs):
    raise WorkspaceNotConfigured("Google Workspace aún no está activado en V24.0.")


def listar_grupos(*args, **kwargs):
    raise WorkspaceNotConfigured("Google Workspace aún no está activado en V24.0.")


def listar_miembros_grupo(*args, **kwargs):
    raise WorkspaceNotConfigured("Google Workspace aún no está activado en V24.0.")
