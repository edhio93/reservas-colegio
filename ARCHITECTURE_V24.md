# Arquitectura Sistema CAV — V24.0 Core

```text
schedule_app.py                  ← sigue siendo el entrypoint de Streamlit Cloud
│
├── core/
│   ├── config.py                ← metadata y configuración no sensible
│   ├── permissions.py           ← roles y permisos de páginas
│   └── navigation.py            ← navegación programática segura
│
├── services/
│   ├── supabase.py              ← conexión, paginado, auditoría y errores
│   ├── gemini.py                ← cliente Gemini
│   ├── workspace.py             ← contrato futuro Workspace
│   ├── gmail.py                 ← contrato futuro Gmail API
│   ├── calendar.py              ← contrato futuro Calendar API
│   └── notifications.py         ← contratos de notificación
│
├── repositories/
│   ├── reservas.py
│   ├── profesores.py
│   ├── inventario.py
│   └── tickets.py
│
├── components/
│   ├── responsive.py            ← CSS responsivo ya extraído y activo
│   ├── navigation.py
│   ├── cards.py
│   └── assistant.py             ← reservado para migración posterior
│
├── pages/
│   └── README.md                ← intencionalmente sin .py en V24.0
│
├── supabase/migrations/
├── scripts/check_project.py
└── .github/workflows/v24_code_check.yml
```

## Decisión clave

En V24.0 **no se separan todavía las páginas visuales**. El objetivo es
introducir las fronteras de arquitectura sin modificar el comportamiento
de reservas, diplomas, TV, inventario o base de datos.

La carpeta `pages/` no contiene `.py` todavía para evitar que Streamlit
active una navegación multipágina distinta antes de migrar a
`st.Page`/`st.navigation` de forma controlada en V24.1.
