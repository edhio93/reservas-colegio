# Instalación V24.0 Core

## Antes de cambiar

Haz un commit o crea una rama de respaldo de tu versión actual.

Sugerencia de rama:

`v24-core-migration`

## Qué debes subir

Copia al raíz del repositorio el contenido completo de este paquete.

El archivo:

`schedule_app.py`

sigue siendo el **Main file path** de Streamlit Cloud.

No cambies:

- URL de la app;
- proyecto Supabase;
- tablas Supabase;
- Streamlit Secrets;
- assets existentes;
- requirements existentes.

## Archivos nuevos

Sube además las carpetas:

- `core/`
- `services/`
- `repositories/`
- `components/`
- `pages/`
- `supabase/`
- `scripts/`

y opcionalmente el workflow:

`.github/workflows/v24_code_check.yml`

## Verificación local/GitHub

Ejecuta:

`python scripts/check_project.py`

Luego en Streamlit Cloud:

`Manage app → Reboot app`

## Señal visual

Después de iniciar sesión, el sidebar debe mostrar:

`Sistema CAV v24.0.0 · Core modular`

## Rollback

El paquete incluye en `legacy/` una copia del entrypoint V23.1. Si algo
inesperado ocurre, puedes devolver temporalmente esa copia a
`schedule_app.py`.
