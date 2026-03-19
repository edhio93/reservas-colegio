import datetime as dt
from datetime import date, datetime as dt_datetime
import time
from io import BytesIO
from pathlib import Path
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
import xlsxwriter
import zipfile
import hashlib
import urllib.parse

# ──────────────────────────────────────────────────────────────────────────────
# 0) CONFIGURACIÓN GLOBAL Y ESTILO
# ──────────────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Sistema de Horarios CAV",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ---- Inyección de CSS para un rediseño completo ----
st.markdown("""
<style>
    /* Paleta de colores y variables */
    :root {
        --primary-color: #0072C6; /* Azul más profesional */
        --background-color: #F8F9FA; /* Fondo muy claro */
        --sidebar-background: #FFFFFF;
        --text-color: #343A40;
        --border-color: #DEE2E6;
        --card-shadow: 0 4px 6px rgba(0,0,0,0.05);
        --card-radius: 0.75rem;
    }

    /* Estilo general del cuerpo */
    .main .block-container {
        padding: 2rem;
    }
    
    /* Títulos de las páginas */
    h1 {
        color: var(--primary-color);
        font-weight: 600;
        border-bottom: 2px solid var(--border-color);
        padding-bottom: 0.5rem;
        margin-bottom: 1.5rem;
    }

    /* Sidebar */
    [data-testid="stSidebar"] {
        background-color: var(--sidebar-background);
        border-right: 1px solid var(--border-color);
        padding: 1rem;
    }
    
    /* Contenedores y formularios (Tarjetas) */
    .st-emotion-cache-1r4qj8v, [data-testid="stForm"], [data-testid="stExpander"] {
        border: 1px solid var(--border-color);
        border-radius: var(--card-radius);
        padding: 1.5rem;
        box-shadow: var(--card-shadow);
        background-color: #FFFFFF;
    }
    
    /* Tarjetas de métricas en el Dashboard */
    [data-testid="stMetric"] {
        background-color: #FFFFFF;
        border: 1px solid var(--border-color);
        border-radius: var(--card-radius);
        padding: 1.5rem;
        box-shadow: var(--card-shadow);
    }

    /* Estilo de los botones de navegación en la barra lateral */
    [data-testid="stSidebarNav"] ul {
        padding: 0;
    }
    [data-testid="stSidebarNav"] li {
        list-style-type: none;
        margin-bottom: 0.5rem;
    }
    [data-testid="stSidebarNav"] a {
        display: flex;
        align-items: center;
        padding: 0.75rem 1rem;
        border-radius: 0.5rem;
        color: var(--text-color);
        text-decoration: none;
        transition: background-color 0.2s, color 0.2s;
        font-weight: 500;
    }
    [data-testid="stSidebarNav"] a:hover {
        background-color: #E9ECEF;
        color: var(--primary-color);
    }
    [data-testid="stSidebarNav"] a[aria-current="page"] {
        background-color: var(--primary-color);
        color: white;
    }
</style>
""", unsafe_allow_html=True)

# ---- Funciones de utilidad ----
def parse_date(val):
    if isinstance(val, date) and not isinstance(val, dt_datetime): return val
    if isinstance(val, dt_datetime): return val.date()
    if isinstance(val, pd.Timestamp): return val.to_pydatetime().date()
    s = str(val).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S"):
        try: return dt_datetime.strptime(s, fmt).date()
        except (ValueError, TypeError): continue
    raise ValueError(f"Formato de fecha inválido: {val!r}")

def as_time(val):
    if isinstance(val, dt.time): return val
    if isinstance(val, dt.datetime): return val.time()
    if isinstance(val, str):
        s = val.strip()
        for fmt in ('%H:%M:%S', '%H:%M'):
            try: return dt_datetime.strptime(s, fmt).time()
            except ValueError: continue
    try: return pd.to_datetime(val).time()
    except (ValueError, TypeError): raise ValueError(f"Formato de hora inválido: {val!r}")

def overlap(s1, e1, s2, e2):
    return max(s1, s2) < min(e1, e2)

def init_empty_db(file_path, sheet_name):
    df_init = pd.DataFrame(columns=['Fecha', 'Hora inicio', 'Hora fin', 'Profesor', 'Curso', 'Recurso', 'Observaciones'])
    df_init.to_excel(file_path, index=False, sheet_name=sheet_name)
    st.success(f"Archivo '{file_path.name}' no encontrado. Se ha creado uno nuevo.")
    time.sleep(2); st.rerun()

def get_color_from_string(input_string: str) -> str:
    hash_obj = hashlib.md5(input_string.encode())
    hash_int = int(hash_obj.hexdigest(), 16)
    hue = hash_int % 360; saturation = 75; lightness = 90
    return f"hsl({hue}, {saturation}%, {lightness}%)"

def sort_time_key(time_string):
    try:
        start_time_str = time_string.split(' a ')[0].strip()
        return dt.datetime.strptime(start_time_str, '%H:%M').time()
    except (ValueError, IndexError): return dt.time(23, 59)

# ------------------------------------------------------------------
# 1) AUTENTICACIÓN Y CARGA DE DATOS
# ------------------------------------------------------------------

if "logged" not in st.session_state:
    st.session_state.logged = False
    st.session_state.role = 'profesor'

if not st.session_state.logged:
    st.title('🔐 Acceso al Sistema de Horarios')
    st.write("Bienvenido, por favor introduce la contraseña para continuar.")
    password = st.text_input("Contraseña", type="password", label_visibility="collapsed")
    if password == "admin":
        st.session_state.logged = True; st.session_state.role = 'admin'; st.rerun()
    elif password == "profesor":
        st.session_state.logged = True; st.session_state.role = 'profesor'; st.rerun()
    elif password:
        st.error("Contraseña incorrecta.")
    st.stop()

# ---- Carga de datos ----
BASE_DIR = Path(__file__).parent
EXCEL = BASE_DIR / "Recursos.xlsx"
if not Path(EXCEL).exists(): init_empty_db(EXCEL, 'Reservas')

try:
    df = pd.read_excel(EXCEL, sheet_name='Reservas', dtype=str).fillna('')
    df['Fecha'] = df['Fecha'].apply(parse_date)
    df['Hora inicio'] = df['Hora inicio'].apply(as_time)
    df['Hora fin'] = df['Hora fin'].apply(as_time)
except Exception as e:
    st.error(f"Error al cargar el archivo Excel: {e}"); st.stop()

def recalc_lists(excel_path):
    try:
        with pd.ExcelFile(excel_path) as xls:
            recursos_df = pd.read_excel(xls, 'Recursos') if 'Recursos' in xls.sheet_names else pd.DataFrame()
            profesores_df = pd.read_excel(xls, 'Profesores') if 'Profesores' in xls.sheet_names else pd.DataFrame()
            cursos_df = pd.read_excel(xls, 'Cursos') if 'Cursos' in xls.sheet_names else pd.DataFrame()
            horas_df = pd.read_excel(xls, 'Horas') if 'Horas' in xls.sheet_names else pd.DataFrame()
        recursos = sorted([str(r[0]) for r in recursos_df.values.tolist()])
        profesores = sorted([str(p[0]) for p in profesores_df.values.tolist()])
        cursos = sorted([str(c[0]) for c in cursos_df.values.tolist()])
        horas_raw = [str(h[0]) for h in horas_df.values.tolist()]
        horas = sorted(horas_raw, key=sort_time_key)
        return recursos, profesores, cursos, horas
    except Exception: return [], [], [], []

RECURSOS, PROFESORES, CURSOS, HORAS = recalc_lists(EXCEL)

# ------------------------------------------------------------------
# 2) NAVEGACIÓN Y VISTAS
# ------------------------------------------------------------------

st.sidebar.title(f"Panel de {st.session_state.role.capitalize()}")

logo_path = BASE_DIR / "logocav.png"
if logo_path.exists():
    st.sidebar.image(str(logo_path), width=180)

st.sidebar.markdown("---")

PAGES_CONFIG = {
    "Registrar": {"icon": "📝", "roles": ["admin", "profesor"]},
    "Base de datos": {"icon": "🗃️", "roles": ["admin"]},
    "Semana": {"icon": "🗓️", "roles": ["admin", "profesor"]},
    "Dashboard": {"icon": "📈", "roles": ["admin"]},
    "Mantenimiento": {"icon": "⚙️", "roles": ["admin"]},
}

available_pages = [page for page, conf in PAGES_CONFIG.items() if st.session_state.role in conf["roles"]]
page = st.sidebar.radio("Navegación", available_pages, format_func=lambda page: f"{PAGES_CONFIG[page]['icon']} {page}", label_visibility="collapsed")

st.sidebar.markdown("---")
if st.sidebar.button("🚪 Cerrar Sesión", use_container_width=True):
    st.session_state.logged = False; st.rerun()

# ------------------------------------------------------------------
# PÁGINAS
# ------------------------------------------------------------------

if page == "Registrar":
    st.title("📝 Registrar Nuevo Horario")
    
    with st.container(border=True):
        with st.form("entry_form"):
            c1, c2 = st.columns(2)
            default_date = df['Fecha'].max() if not df.empty else dt.date.today()
            fecha = c1.date_input('Fecha', value=default_date)
            hora  = c1.selectbox('Bloque Horario', HORAS)
            obs   = c1.text_area('Observaciones (Opcional)')
            prof   = c2.selectbox('Profesor', PROFESORES)
            curso  = c2.selectbox('Curso',    CURSOS)
            
            unavail = []
            try:
                df_mant = pd.read_excel(EXCEL, sheet_name='Mantenimientos').fillna('')
                if not df_mant.empty:
                    unavail = df_mant[(df_mant['FechaInicio'].apply(parse_date) <= fecha) & (df_mant['FechaFin'].apply(parse_date) >= fecha)]['Recurso'].astype(str).tolist()
            except ValueError: pass
            if unavail: c2.warning(f"⚠️ Mantenimiento: {', '.join(unavail)}")
            available = [r for r in RECURSOS if r not in unavail]
            recs = c2.multiselect('Recursos', available, placeholder="Selecciona uno o más recursos")

            if st.form_submit_button('💾 Guardar Registro', use_container_width=True, type="primary"):
                if recs:
                    h_inicio, h_fin = [dt.datetime.strptime(t.strip(), '%H:%M').time() for t in hora.split(' a ')]
                    # ... (resto de la lógica de guardado) ...
                    new_rows = [{'Fecha': fecha, 'Hora inicio': h_inicio, 'Hora fin': h_fin, 'Profesor': prof, 'Curso': curso, 'Recurso': r, 'Observaciones': obs} for r in recs]
                    df_updated = pd.concat([df, pd.DataFrame(new_rows)], ignore_index=True)
                    
                    # ---- NUEVO: Manejo de PermissionError ----
                    try:
                        with pd.ExcelWriter(EXCEL, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                            df_updated.to_excel(writer, sheet_name='Reservas', index=False)
                        st.success("✅ ¡Registro(s) guardado(s)!")
                        time.sleep(1)
                        # No es necesario st.rerun() aquí, el clear_on_submit limpia el form
                    except PermissionError:
                        st.error(f"❌ Error de Permiso: No se pudo guardar. Asegúrate de que el archivo '{EXCEL.name}' no esté abierto en Excel u otro programa e inténtalo de nuevo.")
                    except Exception as e:
                        st.error(f"Ocurrió un error inesperado al guardar: {e}")
                else:
                    st.warning("Por favor, selecciona al menos un recurso.")

if page == "Base de datos":
    st.title("🗃️ Base de Datos de Reservas")
    
    with st.container(border=True):
        st.info("💡 Edita las celdas directamente. Para eliminar filas, usa la papelera en el editor y luego presiona 'Guardar Cambios'.")
        edited_df = st.data_editor(df.copy(), hide_index=True, use_container_width=True, num_rows="dynamic",
            column_config={
                "Fecha": st.column_config.DateColumn("Fecha", format="DD/MM/YYYY"),
                "Hora inicio": st.column_config.TimeColumn("Hora Inicio", format="HH:mm"),
                "Hora fin": st.column_config.TimeColumn("Hora Fin", format="HH:mm"),
                "Profesor": st.column_config.SelectboxColumn("Profesor", options=PROFESORES, required=True),
                "Curso": st.column_config.SelectboxColumn("Curso", options=CURSOS, required=True),
                "Recurso": st.column_config.SelectboxColumn("Recurso", options=RECURSOS, required=True)})
        
        if 'confirm_db_save' not in st.session_state: st.session_state.confirm_db_save = False

        if st.button("💾 Guardar Cambios", use_container_width=True, type="primary"):
            if len(edited_df) < len(df):
                st.session_state.confirm_db_save = True
                st.session_state.df_to_save = edited_df
                st.rerun()
            else:
                try:
                    with pd.ExcelWriter(EXCEL, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        edited_df.to_excel(writer, sheet_name='Reservas', index=False)
                    st.success("✅ Cambios guardados."); time.sleep(1); st.rerun()
                except PermissionError:
                    st.error(f"❌ Error de Permiso: No se pudo guardar. Asegúrate de que el archivo '{EXCEL.name}' no esté abierto.")
        
        if st.session_state.confirm_db_save:
            st.warning(f"Has eliminado {len(df) - len(st.session_state.df_to_save)} fila(s). ¿Confirmas guardar todos los cambios?")
            c1, c2 = st.columns(2)
            if c1.button("✅ Sí, guardar y eliminar", use_container_width=True):
                try:
                    with pd.ExcelWriter(EXCEL, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        st.session_state.df_to_save.to_excel(writer, sheet_name='Reservas', index=False)
                    st.success("✅ Cambios guardados y filas eliminadas.")
                    st.session_state.confirm_db_save = False; del st.session_state.df_to_save
                    time.sleep(1); st.rerun()
                except PermissionError:
                    st.error(f"❌ Error de Permiso: No se pudo guardar. Asegúrate de que el archivo '{EXCEL.name}' no esté abierto.")
            if c2.button("❌ Cancelar", use_container_width=True):
                st.session_state.confirm_db_save = False; del st.session_state.df_to_save
                st.rerun()

if page == "Semana":
    st.title("🗓️ Vista Semanal")
    
    with st.container(border=True):
        c1, c2 = st.columns(2)
        selected_recursos = c1.multiselect("Filtrar por Recurso", RECURSOS, placeholder="Mostrar todos los recursos")
        default_date_week = df['Fecha'].max() if not df.empty else dt.date.today()
        selected_date = c2.date_input("Seleccionar semana", value=default_date_week)
    
    st.markdown("---")
    
    start_of_week = selected_date - dt.timedelta(days=selected_date.weekday())
    week_days = [start_of_week + dt.timedelta(days=i) for i in range(5)]
    df_week = df[(df['Fecha'] >= week_days[0]) & (df['Fecha'] <= week_days[-1]) & (df['Recurso'].isin(selected_recursos) if selected_recursos else df['Recurso'].notna())]

    calendar_style = """
    <style>
        .calendar-container { display: grid; grid-template-columns: 100px repeat(5, 1fr); gap: 2px; background-color: #f0f2f6; border-radius: 0.75rem; overflow: hidden; border: 1px solid var(--border-color); box-shadow: var(--card-shadow); }
        .header-day, .header-time { background-color: #e9ecef; text-align: center; font-weight: 600; padding: 8px 4px; font-size: 0.9em; }
        .time-slot { display: flex; align-items: center; justify-content: center; font-size: 0.8em; background-color: #FFFFFF; border-top: 1px solid var(--border-color); font-weight: 500; }
        .calendar-cell { background-color: white; min-height: 120px; padding: 4px; border-top: 1px solid var(--border-color); border-left: 1px solid var(--border-color); position: relative; }
        .reservation-card { border-radius: 5px; padding: 6px; margin-bottom: 4px; font-size: 0.8em; line-height: 1.3; overflow: hidden; word-wrap: break-word; border: 1px solid rgba(0,0,0,0.1); }
        .fill-link { position: absolute; top: 0; left: 0; width: 100%; height: 100%; display: block; text-decoration: none; border-radius: 3px; }
        .fill-link:hover { background-color: rgba(0, 114, 198, 0.1); border: 1px dashed var(--primary-color); }
    </style>
    """
    
    html = "<div class='calendar-container'>"
    html += "<div class='header-time'></div>"
    dias_semana_es = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]
    for i, day in enumerate(week_days):
        html += f"<div class='header-day'>{dias_semana_es[i]}<br>{day.strftime('%d/%m')}</div>"

    for bloque in HORAS:
        html += f"<div class='time-slot'>{bloque.replace(' a ', '<br>')}</div>"
        for day in week_days:
            html += "<div class='calendar-cell'>"
            try:
                h_inicio_b, h_fin_b = [dt.datetime.strptime(t.strip(), '%H:%M').time() for t in bloque.split(' a ')]
                reservas_celda = df_week[(df_week['Fecha'] == day) & (df_week['Hora inicio'] == h_inicio_b) & (df_week['Hora fin'] == h_fin_b)]
                if reservas_celda.empty:
                    date_str = day.strftime("%Y-%m-%d")
                    hora_str_encoded = urllib.parse.quote(bloque)
                    link = f"?page=Registrar&date={date_str}&hora={hora_str_encoded}"
                    html += f"<a href='{link}' target='_self' class='fill-link'></a>"
                else:
                    for _, row in reservas_celda.iterrows():
                        prof_color = get_color_from_string(row['Profesor'])
                        html += f"<div class='reservation-card' style='background-color:{prof_color};'><strong>{row['Recurso']}</strong><br>{row['Profesor']}<br><em>{row['Curso']}</em></div>"
            except Exception: continue
            html += "</div>"
    html += "</div>"

    st.markdown(calendar_style + html, unsafe_allow_html=True)

if page == "Dashboard":
    st.title("📈 Dashboard y Reportes")
    
    with st.container(border=True):
        st.subheader("Métricas Generales")
        c1, c2, c3 = st.columns(3)
        c1.metric("Total de Reservas", len(df))
        if not df.empty:
            try:
                recurso_mas_usado = df['Recurso'].mode()[0]
                conteo_recurso = df['Recurso'].value_counts().iloc[0]
                c2.metric("Recurso Más Usado", recurso_mas_usado, f"{conteo_recurso} veces")
                
                profesor_mas_activo = df['Profesor'].mode()[0]
                conteo_profesor = df['Profesor'].value_counts().iloc[0]
                c3.metric("Profesor Más Activo", profesor_mas_activo, f"{conteo_profesor} reservas")
            except IndexError:
                c2.info("No hay datos suficientes para mostrar métricas.")
    
    st.markdown("---")
    
    with st.expander("📥 Opciones de Exportación"):
        c_excel, c_ics = st.columns(2)
        with c_excel:
            st.subheader("Exportar a Excel")
            if st.button("📊 Generar Informe Excel", use_container_width=True):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='Reservas')
                    df_summary = df['Recurso'].value_counts().reset_index()
                    df_summary.columns = ['Recurso', 'Usos']; df_summary.to_excel(writer, index=False, sheet_name='Resumen')
                    chart = writer.book.add_chart({'type': 'bar'}); chart.add_series({'categories': f'=Resumen!$A$2:$A${len(df_summary)+1}', 'values': f'=Resumen!$B$2:$B${len(df_summary)+1}'})
                    writer.sheets['Resumen'].insert_chart('D2', chart)
                st.download_button("📥 Descargar Informe", output, "Reporte_Horarios.xlsx", use_container_width=True, type="primary")
        with c_ics:
            st.subheader("Exportar a Calendario (.ics)")
            prof_to_export = st.selectbox("Seleccionar Profesor", PROFESORES, label_visibility="collapsed")
            if st.button("🗓️ Generar .ics", use_container_width=True):
                df_prof = df[df['Profesor'] == prof_to_export]
                ics = "BEGIN:VCALENDAR\nVERSION:2.0\nPRODID:-//ScheduleApp//\n"
                for _, r in df_prof.iterrows():
                    start, end = dt.datetime.combine(r['Fecha'], r['Hora inicio']), dt.datetime.combine(r['Fecha'], r['Hora fin'])
                    ics += f"BEGIN:VEVENT\nUID:{start.strftime('%Y%m%dT%H%M%S')}-{r['Recurso']}\nDTSTAMP:{dt.datetime.now().strftime('%Y%m%dT%H%M%S')}Z\n"
                    ics += f"DTSTART:{start.strftime('%Y%m%dT%H%M%S')}\nDTEND:{end.strftime('%Y%m%dT%H%M%S')}\n"
                    ics += f"SUMMARY:{r['Recurso']} - {r['Curso']}\nDESCRIPTION:{r['Observaciones']}\nEND:VEVENT\n"
                ics += "END:VCALENDAR"
                st.download_button(f"📥 Descargar .ics para {prof_to_export}", ics, f"horario_{prof_to_export}.ics", use_container_width=True, type="primary")

if page == "Mantenimiento":
    st.title("⚙️ Gestión de Mantenimientos")
    
    with st.expander("➕ Agregar Nuevo Mantenimiento", expanded=True):
        with st.form("maint_form", clear_on_submit=True):
            rsrc_maint = st.selectbox("Recurso", RECURSOS)
            c1, c2 = st.columns(2)
            start_date, start_time = c1.date_input("Fecha Inicio"), c1.time_input("Hora Inicio", dt.time(8,0))
            end_date, end_time = c2.date_input("Fecha Fin"), c2.time_input("Hora Fin", dt.time(18,0))
            if st.form_submit_button("💾 Guardar Mantenimiento", use_container_width=True, type="primary"):
                try:
                    try: mant_df = pd.read_excel(EXCEL, sheet_name='Mantenimientos').fillna("")
                    except ValueError: mant_df = pd.DataFrame(columns=MANT_COLS)
                    new_row = {"Recurso":rsrc_maint, "FechaInicio":start_date.strftime('%d/%m/%Y'), "HoraInicio":start_time.strftime('%H:%M'), "FechaFin":end_date.strftime('%d/%m/%Y'), "HoraFin":end_time.strftime('%H:%M')}
                    updated_mant = pd.concat([mant_df, pd.DataFrame([new_row])], ignore_index=True)
                    with pd.ExcelWriter(EXCEL, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                        updated_mant.to_excel(writer, sheet_name='Mantenimientos', index=False)
                    st.success("✅ Mantenimiento registrado."); time.sleep(1); st.rerun()
                except PermissionError:
                    st.error(f"❌ Error de Permiso: No se pudo guardar. Asegúrate de que el archivo '{EXCEL.name}' no esté abierto.")

    st.markdown("---")
    st.subheader("📋 Mantenimientos Programados")
    try: mant_df = pd.read_excel(EXCEL, sheet_name='Mantenimientos').fillna("")
    except ValueError: mant_df = pd.DataFrame(columns=MANT_COLS)

    if mant_df.empty:
        st.info("No hay mantenimientos programados.")
    else:
        with st.container(border=True):
            to_drop = st.multiselect("Seleccionar para eliminar:", mant_df.index, format_func=lambda i: f"ID {i}: {mant_df.at[i, 'Recurso']} @ {mant_df.at[i, 'FechaInicio']}", placeholder="Haz clic para seleccionar...")
            st.dataframe(mant_df, hide_index=True, use_container_width=True)
            if st.button("🗑️ Eliminar Seleccionados", use_container_width=True) and to_drop:
                st.session_state.confirm_delete = True
                st.session_state.items_to_delete = to_drop
                st.rerun()
            
            if st.session_state.get('confirm_delete', False):
                st.warning(f"**¿Eliminar {len(st.session_state.get('items_to_delete', []))} registro(s)?**")
                c1, c2 = st.columns(2)
                if c1.button("✅ Sí, eliminar", use_container_width=True):
                    try:
                        kept = mant_df.drop(index=st.session_state.items_to_delete)
                        with pd.ExcelWriter(EXCEL, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                            kept.to_excel(writer, sheet_name='Mantenimientos', index=False)
                        st.success("✅ Eliminado(s).")
                        st.session_state.confirm_delete = False; st.session_state.items_to_delete = []
                        time.sleep(1); st.rerun()
                    except PermissionError:
                        st.error(f"❌ Error de Permiso: No se pudo guardar. Asegúrate de que el archivo '{EXCEL.name}' no esté abierto.")
                if c2.button("❌ Cancelar", use_container_width=True):
                    st.session_state.confirm_delete = False; st.session_state.items_to_delete = []
                    st.rerun()

