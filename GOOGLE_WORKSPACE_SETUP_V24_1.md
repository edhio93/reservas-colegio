# Google Workspace — configuración V24.1

## Objetivo

El Sistema CAV utilizará Google Workspace como fuente oficial de:

- usuarios;
- correos institucionales;
- grupos;
- miembros de grupos.

Las credenciales Google se guardan exclusivamente como **Supabase Edge Function Secrets**.

## APIs que debes habilitar en Google Cloud

En el proyecto Google Cloud usado por CAV:

1. Admin SDK API
2. Gmail API

## Crear cuenta de servicio

Google Cloud Console:

`IAM y administración → Cuentas de servicio → Crear cuenta de servicio`

Nombre recomendado:

`cav-workspace-automation`

Después crea una clave JSON para esa cuenta y descárgala temporalmente.

**No subas ese JSON a GitHub.**

## Domain-Wide Delegation

En la cuenta de servicio, copia su **Client ID**.

Como Super Admin entra a Google Admin:

`Seguridad → Control de acceso y datos → Controles de API → Delegación en todo el dominio`

Agrega el Client ID y estos scopes, separados por coma:

- `https://www.googleapis.com/auth/admin.directory.user.readonly`
- `https://www.googleapis.com/auth/admin.directory.group.readonly`
- `https://www.googleapis.com/auth/admin.directory.group.member.readonly`
- `https://www.googleapis.com/auth/gmail.send`

## Cuenta delegada

Elige una cuenta administrativa Workspace para Directory API.

Ejemplo conceptual:

`administrador@colegioantoniovaras.cl`

Debe poder consultar usuarios y grupos.

Para el remitente Gmail elige la cuenta institucional que enviará notificaciones,
por ejemplo una cuenta dedicada de Enlaces.

## Supabase Edge Function Secrets

En Supabase:

`Edge Functions → Secrets`

Crea:

- `GOOGLE_SERVICE_ACCOUNT_JSON`
- `GOOGLE_DELEGATED_ADMIN`
- `GOOGLE_GMAIL_SENDER`
- `GOOGLE_WORKSPACE_CUSTOMER`

Para `GOOGLE_SERVICE_ACCOUNT_JSON`, pega el contenido JSON completo de la clave de la cuenta de servicio.

Para `GOOGLE_WORKSPACE_CUSTOMER`, usa normalmente:

`my_customer`

No publiques estos valores en GitHub ni en un chat.
