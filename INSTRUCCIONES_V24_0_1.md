# V24.0.1 — Hotfix Gemini + prueba de correo SMTP

## Qué corrige

### Gemini

El error `name 'model' is not defined` era un falso negativo del panel
**Estado del sistema**. En V24.0 el modelo Gemini fue movido a
`services/gemini.py`, pero el panel antiguo seguía buscando la variable global
`model` del monolito.

V24.0.1 usa `estado_gemini()` y agrega un botón **Probar Gemini**.

### Correo SMTP

Se agrega `services/email_smtp.py` y en:

`Configuración → Estado del sistema`

aparece una tarjeta **Correo SMTP** con:

1. campo `Enviar prueba a`;
2. botón `Enviar correo de prueba`;
3. resultado visible de éxito/error.

El correo de prueba confirma conexión, login SMTP y envío.

## Instalación

Reemplaza/añade estos archivos del paquete completo en tu repositorio:

- `schedule_app.py`
- `core/config.py`
- `services/gemini.py`
- `services/email_smtp.py` (nuevo)
- `V24_BUILD_INFO.json`

También puedes copiar todo el contenido del ZIP sobre el repositorio V24.0.

No requiere SQL ni nuevos Secrets.

## Secrets esperados

La configuración SMTP existente sigue usando:

```toml
[email_credentials]
smtp_server = "smtp.gmail.com"
smtp_port = 587
smtp_username = "..."
smtp_password = "..."
sender_email = "..."
sender_name = "Liceo Bicentenario de Excelencia Colegio Antonio Varas"
reply_to = "..."
use_tls = true
use_ssl = false
```

No publiques este archivo ni pegues sus valores en GitHub.
