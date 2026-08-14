# Pages — reservado para V24.1

V24.0 no incluye archivos `.py` dentro de `pages/` de forma intencional.

El sistema actual todavía usa navegación interna en `schedule_app.py`.
Agregar scripts de página ahora podría activar la navegación multipágina
automática de Streamlit antes de terminar la migración.

En V24.1 se moverán primero:

- Inicio
- Registrar / Reservas
- Semana
- Base de datos

y el entrypoint cambiará de forma controlada a `st.Page` + `st.navigation`.
