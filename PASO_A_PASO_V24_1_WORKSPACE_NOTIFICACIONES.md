# Sistema CAV V24.1 — Workspace + Notificaciones + Resumen Semanal

## Resultado de esta versión

V24.1 incorpora:

- sincronización automática de usuarios de Google Workspace;
- sincronización automática de grupos y miembros;
- vinculación automática profesor ↔ cuenta Workspace;
- vinculación manual para casos ambiguos;
- correo Workspace como correo oficial del profesor;
- cola `notification_outbox`;
- notificaciones por creación, modificación y cancelación de reservas;
- consolidación de múltiples modificaciones durante 2 minutos;
- Gmail API para el envío real;
- resumen semanal todos los lunes a las 07:30 hora de Chile;
- si el profesor no tiene reservas esa semana, **no se genera correo**;
- sincronización Workspace cada 6 horas;
- procesador de correos cada minuto;
- panel Workspace dentro de Configuración.

# ORDEN DE INSTALACIÓN

Es importante respetar este orden.

## PASO 0 — Respaldo

Antes de comenzar crea en GitHub una rama:

`backup-v24-0-2`

No elimines la versión actual hasta comprobar V24.1.

## PASO 1 — Ejecutar migración SQL en Supabase

Abre:

`Supabase → SQL Editor → New query`

Copia y ejecuta:

`supabase/migrations/20260814_v24_1_workspace_notifications.sql`

Debe terminar sin errores.

Esta migración crea:

- columnas Workspace en `profesores`;
- `workspace_users`;
- `workspace_groups`;
- `workspace_group_members`;
- `workspace_sync_log`;
- `notification_outbox`;
- `weekly_digest_log`.

Todavía **no ejecutes el archivo cron**.

## PASO 2 — Google Cloud

Lee:

`GOOGLE_WORKSPACE_SETUP_V24_1.md`

Habilita:

- Admin SDK API;
- Gmail API.

Crea una cuenta de servicio dedicada, por ejemplo:

`cav-workspace-automation`

Descarga una clave JSON.

**No la subas a GitHub.**

## PASO 3 — Domain-Wide Delegation

Copia el Client ID de la cuenta de servicio.

En Google Admin:

`Seguridad → Control de acceso y datos → Controles de API → Delegación en todo el dominio`

Autoriza:

- `https://www.googleapis.com/auth/admin.directory.user.readonly`
- `https://www.googleapis.com/auth/admin.directory.group.readonly`
- `https://www.googleapis.com/auth/admin.directory.group.member.readonly`
- `https://www.googleapis.com/auth/gmail.send`

## PASO 4 — Configurar secretos en Supabase

En:

`Supabase → Edge Functions → Secrets`

crea:

`GOOGLE_SERVICE_ACCOUNT_JSON`

Valor: contenido completo del JSON descargado.

`GOOGLE_DELEGATED_ADMIN`

Valor: correo de una cuenta Workspace administrativa que la automatización impersonará para consultar Directory.

`GOOGLE_GMAIL_SENDER`

Valor: cuenta institucional que enviará los correos.

`GOOGLE_WORKSPACE_CUSTOMER`

Valor recomendado:

`my_customer`

No agregues estos secretos a Streamlit Secrets y no los subas a GitHub.

## PASO 5 — Desplegar Edge Functions

Tienes tres funciones:

- `supabase/functions/workspace-sync/index.ts`
- `supabase/functions/process-notifications/index.ts`
- `supabase/functions/weekly-digest/index.ts`

Y dos helpers compartidos:

- `supabase/functions/_shared/google_auth.ts`
- `supabase/functions/_shared/gmail.ts`

### Recomendado: Supabase CLI

Desde la raíz del repositorio:

`supabase login`

`supabase link --project-ref TU_PROJECT_REF`

`supabase functions deploy workspace-sync`

`supabase functions deploy process-notifications`

`supabase functions deploy weekly-digest`

Supabase también permite crear funciones desde el Dashboard, pero para esta versión la CLI resulta más cómoda porque conserva automáticamente la carpeta `_shared`.

## PASO 6 — Subir V24.1 a GitHub

Descomprime el paquete y sube el **contenido**, no el ZIP.

Reemplaza al menos:

- `schedule_app.py`
- `core/config.py`
- `services/workspace.py`
- `services/notifications.py`
- `repositories/profesores.py`
- `V24_BUILD_INFO.json`

Agrega:

- `supabase/functions/...`
- `supabase/migrations/20260814_v24_1_workspace_notifications.sql`
- `supabase/cron/20260814_v24_1_cron_setup.sql`
- `GOOGLE_WORKSPACE_SETUP_V24_1.md`

## PASO 7 — Reiniciar Streamlit

`Manage app → Reboot app`

Luego:

`Ctrl + F5`

En:

`Configuración → Google Workspace`

debe aparecer el nuevo panel.

## PASO 8 — Probar Workspace manualmente

En:

`Configuración → Google Workspace`

presiona:

`🔄 Sincronizar Workspace ahora`

Resultado esperado:

- usuarios > 0;
- grupos > 0;
- profesores vinculados;
- última sincronización en verde.

La sincronización automática intenta vincular en este orden:

1. Google ID ya guardado;
2. correo actual exacto;
3. nombre completo exacto cuando existe una sola coincidencia.

Los profesores ambiguos quedan disponibles para vinculación manual.

## PASO 9 — Probar una notificación antes del cron

Edita una reserva con la casilla de notificación activada.

En:

`Configuración → Google Workspace → Cola de notificaciones`

debe aparecer un registro con estado:

`pending`

La edición de la reserva no espera el envío del correo.

## PASO 10 — Probar Gmail API manualmente

En Supabase:

`Edge Functions → process-notifications → Invoke/Test`

La notificación debe pasar:

`pending → sending → sent`

y llegar al correo Workspace del profesor.

Si falla, revisa:

- `notification_outbox.error`;
- logs de `process-notifications`;
- scopes de Domain-Wide Delegation;
- `GOOGLE_GMAIL_SENDER`.

## PASO 11 — Probar resumen semanal

La función `weekly-digest` solo actúa automáticamente los lunes entre 07:30 y 07:44 `America/Santiago`.

Para probarla cualquier día desde el panel de Edge Functions, invócala con este body:

`{"force": true}`

Opcionalmente puedes probar una semana concreta:

`{"force": true, "week_start": "2026-08-17"}`

La prueba sigue respetando el control de duplicados de `weekly_digest_log`.

La lógica garantiza:

- profesor con reservas → correo;
- profesor sin reservas → ningún correo;
- mismo profesor + misma semana → máximo un resumen.

## PASO 12 — Instalar Cron

Solo cuando las tres funciones funcionen manualmente.

Abre:

`supabase/cron/20260814_v24_1_cron_setup.sql`

Antes de ejecutarlo reemplaza:

- `TU_PROJECT_REF`;
- `PEGA_AQUI_TU_SERVICE_ROLE_SOLO_EN_SQL_EDITOR`.

El script guarda los valores en Supabase Vault y configura:

- `process-notifications`: cada minuto;
- `weekly-digest`: cada 15 minutos;
- `workspace-sync`: cada 6 horas.

`weekly-digest` comprueba la zona `America/Santiago`, por lo que el cambio de horario de invierno/verano en Chile no obliga a cambiar el cron.

# Cómo funciona una modificación

`Editar reserva`

→ actualiza Supabase

→ crea/actualiza registro `notification_outbox`

→ Streamlit termina rápidamente

→ `process-notifications` recoge el trabajo

→ Gmail API envía el mensaje.

Si una misma reserva se modifica varias veces dentro de aproximadamente 2 minutos, la cola intenta consolidar esas ediciones para reducir correos repetidos.

Si además se cambia el profesor:

- el profesor anterior recibe aviso de reasignación;
- el profesor nuevo recibe confirmación de la reserva asignada.

# Cómo funciona el lunes 07:30

Supabase invoca `weekly-digest` cada 15 minutos.

La función solo continúa cuando:

- es lunes;
- hora local de Chile está entre 07:30 y 07:44.

Busca las reservas de lunes a viernes y agrupa por profesor.

Si un profesor tiene 0 reservas, ni siquiera se crea una notificación.

`weekly_digest_log` evita duplicados.

# Seguridad

La clave JSON de Google vive únicamente en:

`Supabase Edge Function Secrets`

Las nuevas tablas tienen RLS habilitado sin políticas públicas.

Los scopes Directory son de solo lectura y Gmail usa exclusivamente `gmail.send`.
