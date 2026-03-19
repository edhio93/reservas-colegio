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
import re
import html as html_sanitizer
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import plotly.express as px

# ──────────────────────────────────────────────────────────────────────────────
# 0) CONFIGURACIÓN GLOBAL Y ESTILO
# ──────────────────────────────────────────────────────────────────────────────

st.set_page_config(page_title="Sistema de Horarios CAV", page_icon="📅", layout="wide", initial_sidebar_state="expanded")

# ---- Inyección de CSS ----
st.markdown("""
<style>
    :root {
        --primary-color: #0072C6;
        --background-color: #F8F9FA;
        --sidebar-background: #FFFFFF;
        --card-background: #FFFFFF;
        --text-color: #343A40;
        --subtle-text-color: #6C757D;
        --border-color: #DEE2E6;
        --hover-color: #E9ECEF;
        --card-shadow: 0 4px 6px rgba(0,0,0,0.05);
    }
    .main .block-container { padding: 2rem; }
    h1 {
        color: var(--primary-color);
        font-weight: 600;
        border-bottom: 2px solid var(--border-color);
        padding-bottom: 0.5rem;
        margin-bottom: 1.5rem;
    }
    [data-testid="stSidebar"] {
        background-color: var(--sidebar-background);
        border-right: 1px solid var(--border-color);
        padding: 1rem;
    }
    .st-emotion-cache-1r4qj8v, [data-testid="stForm"], [data-testid="stExpander"] {
        border: 1px solid var(--border-color);
        border-radius: 0.75rem;
        padding: 1.5rem;
        box-shadow: var(--card-shadow);
        background-color: var(--card-background);
    }
    [data-testid="stMetric"] {
        background-color: var(--card-background);
        border: 1px solid var(--border-color);
        border-radius: 0.75rem;
        padding: 1.5rem;
        box-shadow: var(--card-shadow);
    }
    [data-testid="stSidebarNav"] a:hover {
        background-color: var(--hover-color);
        color: var(--primary-color);
    }
    [data-testid="stSidebarNav"] a[aria-current="page"] {
        background-color: var(--primary-color);
        color: white;
    }
    @media (prefers-color-scheme: dark) {
        :root {
            --primary-color: #58A6FF;
            --background-color: #0D1117;
            --sidebar-background: #161B22;
            --card-background: #161B22;
            --text-color: #C9D1D9;
            --subtle-text-color: #8B949E;
            --border-color: #30363D;
            --hover-color: #252b33;
        }
        body, .stApp { background-color: var(--background-color); color: var(--text-color); }
        .st-emotion-cache-1r4qj8v, [data-testid="stForm"], [data-testid="stExpander"], [data-testid="stMetric"] { border-color: var(--border-color); }
        .tooltip-text { background-color: #f0f2f6 !important; color: #111 !important; }
    }
    .reservation-card { 
        border-radius: 5px; 
        padding: 6px; 
        margin-bottom: 4px;
        font-size: 0.8em; 
        line-height: 1.3; 
        word-wrap: break-word; 
        border: 1px solid rgba(0,0,0,0.1);
        position: relative;
        cursor: default;
    }
    .tooltip-text {
        visibility: hidden; width: 220px; background-color: #333; color: #fff; text-align: left;
        border-radius: 6px; padding: 10px; position: absolute; z-index: 10; bottom: 105%;
        left: 50%; margin-left: -110px; opacity: 0; transition: opacity 0.3s; box-shadow: 0 4px 8px rgba(0,0,0,0.2);
        pointer-events: none;
    }
    .reservation-card:hover .tooltip-text { visibility: visible; opacity: 1; }
    .dataframe td { vertical-align: top; }
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

def overlap(s1, e1, s2, e2): return max(s1, s2) < min(e1, e2)

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

def save_excel_sheet(df_to_save, sheet_name, excel_path):
    try:
        try:
            with pd.ExcelFile(excel_path) as xls: all_sheets = {sheet: pd.read_excel(xls, sheet) for sheet in xls.sheet_names}
        except FileNotFoundError: all_sheets = {}
        all_sheets[sheet_name] = df_to_save
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            for sheet, df_sheet in all_sheets.items(): df_sheet.to_excel(writer, sheet_name=sheet, index=False)
        st.success(f"✅ Cambios en '{sheet_name}' guardados correctamente.")
        time.sleep(1); st.rerun()
    except PermissionError: st.error(f"❌ Error de Permiso: No se pudo guardar. Asegúrate de que '{excel_path.name}' no esté abierto en Excel.")
    except Exception as e: st.error(f"Ocurrió un error inesperado al guardar: {e}")

def format_date_es(date_obj):
    dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
    meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    return f"{dias[date_obj.weekday()]}, {date_obj.day} de {meses[date_obj.month - 1]} de {date_obj.year}"

def custom_course_sort_key(course_name):
    course_name = str(course_name).strip()
    if 'Dif' in course_name: return (3, 0, course_name)
    match = re.match(r"(\d+)°\s*(BÁSICO|MEDIO)\s*([A-Z])?", course_name, re.IGNORECASE)
    if match:
        num, level, letter = match.groups()
        level_priority = 0 if 'BÁSICO' in level.upper() else 1
        return (level_priority, int(num), letter or '')
    return (4, 0, course_name)

def send_email(subject, body, recipient_email):
    try:
        creds = st.secrets["email_credentials"]
        sender_email, password = creds["smtp_username"], creds["smtp_password"]
        msg = MIMEMultipart(); msg['From'] = sender_email; msg['To'] = recipient_email; msg['Subject'] = subject
        msg.attach(MIMEText(body, 'html'))
        server = smtplib.SMTP(creds["smtp_server"], creds["smtp_port"]); server.starttls()
        server.login(sender_email, password); server.send_message(msg); server.quit()
        st.toast(f"📧 Notificación enviada a {recipient_email}")
    except Exception as e:
        st.warning(f"⚠️ No se pudo enviar la notificación por correo: {e}")

# ------------------------------------------------------------------
# 1) AUTENTICACIÓN Y CARGA DE DATOS
# ------------------------------------------------------------------

if "logged" not in st.session_state:
    st.session_state.logged = False
    st.session_state.role = 'profesor'
    st.session_state.profesor_name = None

if not st.session_state.logged:
    st.title('🔐 Acceso al Sistema de Horarios')
    password = st.text_input("Contraseña", type="password", label_visibility="collapsed")
    if password == "admin": st.session_state.logged = True; st.session_state.role = 'admin'; st.rerun()
    elif password == "profesor": st.session_state.logged = True; st.session_state.role = 'profesor'; st.rerun()
    elif password: st.error("Contraseña incorrecta.")
    st.stop()

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
            recursos_df = pd.read_excel(xls, 'Recursos') if 'Recursos' in xls.sheet_names else pd.DataFrame({'Recursos': []})
            profesores_df = pd.read_excel(xls, 'Profesores') if 'Profesores' in xls.sheet_names else pd.DataFrame([[]])
            cursos_df = pd.read_excel(xls, 'Cursos') if 'Cursos' in xls.sheet_names else pd.DataFrame({'Cursos': []})
            horas_df = pd.read_excel(xls, 'Horas') if 'Horas' in xls.sheet_names else pd.DataFrame({'Horas': []})
        recursos = sorted(recursos_df.iloc[:, 0].dropna().astype(str).tolist())
        profesores = sorted(profesores_df.iloc[:, 0].dropna().astype(str).tolist())
        profesor_data = {}
        if not profesores_df.empty and len(profesores_df.columns) > 1:
            profesor_data = pd.Series(profesores_df.iloc[:, 1].values, index=profesores_df.iloc[:, 0]).to_dict()
        cursos_raw = cursos_df.iloc[:, 0].dropna().astype(str).tolist()
        cursos = sorted(cursos_raw, key=custom_course_sort_key)
        horas_raw = horas_df.iloc[:, 0].dropna().astype(str).tolist()
        horas = sorted(horas_raw, key=sort_time_key)
        return recursos, profesores, cursos, horas, profesor_data
    except Exception as e:
        st.error(f"Error al leer las listas maestras del Excel: {e}")
        return [], [], [], [], {}

RECURSOS, PROFESORES, CURSOS, HORAS, PROFESOR_DATA = recalc_lists(EXCEL)

if st.session_state.role == 'profesor' and not st.session_state.profesor_name:
    st.title("👤 Selección de Perfil")
    st.write("Por favor, selecciona tu nombre para continuar.")
    prof_name = st.selectbox("Nombre del Profesor", PROFESORES, index=None, placeholder="Selecciona tu nombre...")
    if prof_name:
        st.session_state.profesor_name = prof_name
        st.rerun()
    st.stop()

# ------------------------------------------------------------------
# 2) NAVEGACIÓN Y VISTAS
# ------------------------------------------------------------------

sidebar_title = f"Panel de {st.session_state.role.capitalize()}"
if st.session_state.role == 'profesor':
    sidebar_title = f"Hola, {st.session_state.profesor_name.split(' ')[0]}"
st.sidebar.title(sidebar_title)

logo_path = BASE_DIR / "logocav.png"
if logo_path.exists(): st.sidebar.image(str(logo_path), width=180)
st.sidebar.markdown("---")

PAGES_CONFIG = {
    "Mis Reservas": {"icon": "👤", "roles": ["profesor"]},
    "Registrar": {"icon": "📝", "roles": ["admin"]},
    "Base de datos": {"icon": "🗃️", "roles": ["admin"]},
    "Semana": {"icon": "🗓️", "roles": ["admin", "profesor"]},
    "Dashboard": {"icon": "📈", "roles": ["admin"]},
    "Mantenimiento": {"icon": "⚙️", "roles": ["admin"]},
}
available_pages = [page for page, conf in PAGES_CONFIG.items() if st.session_state.role in conf["roles"]]
default_page = "Mis Reservas" if st.session_state.role == 'profesor' else "Registrar"
page = st.sidebar.radio("Navegación", available_pages, index=available_pages.index(default_page), format_func=lambda p: f"{PAGES_CONFIG[p]['icon']} {p}", label_visibility="collapsed")

st.sidebar.markdown("---")
if st.sidebar.button("🚪 Cerrar Sesión", use_container_width=True):
    for key in st.session_state.keys(): del st.session_state[key]
    st.rerun()

# ------------------------------------------------------------------
# PÁGINAS
# ------------------------------------------------------------------

if page == "Mis Reservas":
    st.title("👤 Mis Próximas Reservas")
    prof_df = df[df['Profesor'] == st.session_state.profesor_name]
    future_reservas = prof_df[prof_df['Fecha'] >= dt.date.today()].sort_values(by="Fecha")
    if future_reservas.empty:
        st.info("No tienes reservas programadas para el futuro.")
    else:
        for _, row in future_reservas.iterrows():
            with st.container(border=True):
                st.markdown(f"#### {format_date_es(row['Fecha'])}")
                st.markdown(f"**Hora:** {row['Hora inicio'].strftime('%H:%M')} - {row['Hora fin'].strftime('%H:%M')}<br>**Curso:** {row['Curso']}<br>**Recurso:** {row['Recurso']}", unsafe_allow_html=True)
                if row['Observaciones']:
                    st.markdown(f"> *{row['Observaciones']}*")

if page == "Registrar":
    st.title("📝 Registrar Nuevo Horario")
    def check_all_conflicts(fechas, recursos, hora_inicio, hora_fin, df_reservas, df_mantenimiento):
        conflictos_reserva, conflictos_mantenimiento = [], []
        for fecha in fechas:
            for rec in recursos:
                df_check = df_reservas[(df_reservas['Fecha'] == fecha) & (df_reservas['Recurso'] == rec)]
                for _, row in df_check.iterrows():
                    if overlap(hora_inicio, hora_fin, row['Hora inicio'], row['Hora fin']):
                        conflictos_reserva.append(f"<li>{rec} el {fecha.strftime('%d/%m/%Y')} (con {row['Profesor']})</li>")
        if not df_mantenimiento.empty:
            for fecha in fechas:
                for rec in recursos:
                    mant_check = df_mantenimiento[(df_mantenimiento['Recurso'] == rec) & (df_mantenimiento['FechaInicio_dt'] <= fecha) & (df_mantenimiento['FechaFin_dt'] >= fecha)]
                    for _, m_row in mant_check.iterrows():
                        if overlap(hora_inicio, hora_fin, as_time(m_row['HoraInicio']), as_time(m_row['HoraFin'])):
                             conflictos_mantenimiento.append(f"<li>{rec} el {fecha.strftime('%d/%m/%Y')} (en mantenimiento)</li>")
        return conflictos_reserva, list(set(conflictos_mantenimiento))
    with st.container(border=True):
        tipo_reserva = st.radio("Tipo de Reserva", ["Única", "Múltiples Fechas", "Semanal Recurrente"], horizontal=True, key="tipo_reserva")
        with st.form("entry_form"):
            st.markdown("---")
            c1, c2 = st.columns(2)
            fechas_a_registrar = []
            if tipo_reserva == "Única":
                default_date = df['Fecha'].max() if not df.empty else dt.date.today()
                fecha = c1.date_input('Fecha', value=default_date, format="DD/MM/YYYY")
                fechas_a_registrar.append(fecha)
            elif tipo_reserva == "Múltiples Fechas":
                today = dt.date.today(); date_range = [today + dt.timedelta(days=i) for i in range(180)]
                df_dates = pd.DataFrame({"Seleccionar": [False] * len(date_range), "Fecha Disponible": [format_date_es(d) for d in date_range], "_date_obj": date_range})
                with c1:
                    st.write("Selecciona las fechas deseadas:")
                    edited_dates_df = st.data_editor(df_dates, column_config={"Seleccionar": st.column_config.CheckboxColumn(required=True), "_date_obj": None}, hide_index=True, height=200, use_container_width=True)
                selected_dates_df = edited_dates_df[edited_dates_df["Seleccionar"]]
                fechas_a_registrar = sorted(selected_dates_df["_date_obj"].tolist())
            else:
                fecha_inicio = c1.date_input('Fecha de Inicio', value=dt.date.today(), format="DD/MM/YYYY")
                num_semanas = c1.number_input('Repetir durante (semanas)', min_value=1, max_value=52, value=4)
                fechas_a_registrar = [fecha_inicio + dt.timedelta(weeks=i) for i in range(num_semanas)]
                if fechas_a_registrar: c1.info(f"Se registrarán {len(fechas_a_registrar)} fechas.")
            
            hora  = c1.selectbox('Bloque Horario', HORAS)
            obs   = c1.text_area('Observaciones (Opcional)')
            prof   = c2.selectbox('Profesor', PROFESORES)
            curso  = c2.selectbox('Curso',    CURSOS)
            recs = c2.multiselect('Recursos', RECURSOS, placeholder="Selecciona uno o más recursos")
            
            if st.form_submit_button('💾 Guardar Registro', use_container_width=True, type="primary"):
                if recs and fechas_a_registrar:
                    h_inicio, h_fin = [dt.datetime.strptime(t.strip(), '%H:%M').time() for t in hora.split(' a ')]
                    try:
                        df_mant = pd.read_excel(EXCEL, sheet_name='Mantenimientos').fillna('')
                        if not df_mant.empty: df_mant['FechaInicio_dt'] = df_mant['FechaInicio'].apply(parse_date); df_mant['FechaFin_dt'] = df_mant['FechaFin'].apply(parse_date)
                    except ValueError: df_mant = pd.DataFrame()
                    conflictos_r, conflictos_m = check_all_conflicts(fechas_a_registrar, recs, h_inicio, h_fin, df, df_mant)
                    if conflictos_r or conflictos_m:
                        if conflictos_r: st.error(f"❌ **Conflicto de Reserva:**"); st.markdown(f"<ul>{''.join(conflictos_r)}</ul>", unsafe_allow_html=True)
                        if conflictos_m: st.error(f"❌ **Conflicto de Mantenimiento:**"); st.markdown(f"<ul>{''.join(conflictos_m)}</ul>", unsafe_allow_html=True)
                    else:
                        nuevas_reservas = []
                        for fecha in fechas_a_registrar:
                            for rec in recs:
                                nuevas_reservas.append({'Fecha': fecha, 'Hora inicio': h_inicio, 'Hora fin': h_fin, 'Profesor': prof, 'Curso': curso, 'Recurso': rec, 'Observaciones': obs})
                        df_updated = pd.concat([df, pd.DataFrame(nuevas_reservas)], ignore_index=True)
                        save_excel_sheet(df_updated, 'Reservas', EXCEL)
                        email_to = PROFESOR_DATA.get(prof)
                        if email_to:
                            subject = f"Confirmación de Reserva de Recursos - {curso}"
                            body = f"""<html><body><p>Hola {prof.split(' ')[0]},</p><p>Se ha(n) confirmado la(s) siguiente(s) reserva(s) a tu nombre:</p><ul><li><b>Curso:</b> {curso}</li><li><b>Recurso(s):</b> {', '.join(recs)}</li><li><b>Horario:</b> {hora}</li></ul><p><b>Fechas Registradas:</b></p><ul>{''.join([f'<li>{format_date_es(f)}</li>' for f in fechas_a_registrar])}</ul>{f"<p><b>Observaciones:</b> {obs}</p>" if obs else ""}<p>Saludos,<br>Sistema de Horarios CAV</p></body></html>"""
                            send_email(subject, body, email_to)

if page == "Base de datos":
    st.title("🗃️ Base de Datos de Reservas")
    with st.container(border=True):
        edited_df = st.data_editor(df.copy(), hide_index=True, use_container_width=True, num_rows="dynamic", column_config={"Fecha": st.column_config.DateColumn("Fecha", format="DD/MM/YYYY"), "Hora inicio": st.column_config.TimeColumn("Hora Inicio", format="HH:mm"), "Hora fin": st.column_config.TimeColumn("Hora Fin", format="HH:mm"), "Profesor": st.column_config.SelectboxColumn("Profesor", options=PROFESORES, required=True), "Curso": st.column_config.SelectboxColumn("Curso", options=CURSOS, required=True), "Recurso": st.column_config.SelectboxColumn("Recurso", options=RECURSOS, required=True)})
        if 'confirm_db_save' not in st.session_state: st.session_state.confirm_db_save = False
        if st.button("💾 Guardar Cambios", use_container_width=True, type="primary"):
            if len(edited_df) < len(df):
                original_indices = set(df.index); edited_indices = set(edited_df.index)
                deleted_indices = original_indices - edited_indices
                deleted_rows = df.loc[list(deleted_indices)]
                st.session_state.deleted_rows = deleted_rows
                st.session_state.confirm_db_save = True; st.session_state.df_to_save = edited_df; st.rerun()
            else: save_excel_sheet(edited_df, 'Reservas', EXCEL)
        if st.session_state.confirm_db_save:
            st.warning(f"Has eliminado {len(df) - len(st.session_state.df_to_save)} fila(s). ¿Confirmas guardar todos los cambios?")
            c1, c2 = st.columns(2)
            if c1.button("✅ Sí, guardar y eliminar", use_container_width=True):
                save_excel_sheet(st.session_state.df_to_save, 'Reservas', EXCEL)
                for _, row in st.session_state.deleted_rows.iterrows():
                    prof_name = row['Profesor']
                    email_to = PROFESOR_DATA.get(prof_name)
                    if email_to:
                        subject = f"Cancelación de Reserva de Recursos - {row['Curso']}"
                        body = f"""<html><body><p>Hola {prof_name.split(' ')[0]},</p><p>Te informamos que la siguiente reserva ha sido <b>cancelada</b>:</p><ul><li><b>Fecha:</b> {format_date_es(row['Fecha'])}</li><li><b>Horario:</b> {row['Hora inicio'].strftime('%H:%M')} - {row['Hora fin'].strftime('%H:%M')}</li><li><b>Curso:</b> {row['Curso']}</li><li><b>Recurso:</b> {row['Recurso']}</li></ul><p>Saludos,<br>Sistema de Horarios CAV</p></body></html>"""
                        send_email(subject, body, email_to)
                st.session_state.confirm_db_save = False
            if c2.button("❌ Cancelar", use_container_width=True): st.session_state.confirm_db_save = False; st.rerun()

if page == "Semana":
    st.title("🗓️ Vista Semanal")
    with st.container(border=True):
        c1, c2 = st.columns(2)
        selected_recursos = c1.multiselect("Filtrar por Recurso", RECURSOS, placeholder="Mostrar todos los recursos")
        default_date_week = df['Fecha'].max() if not df.empty else dt.date.today()
        selected_date = c2.date_input("Seleccionar semana", value=default_date_week, format="DD/MM/YYYY")
    
    start_of_week = selected_date - dt.timedelta(days=selected_date.weekday())
    week_days = [start_of_week + dt.timedelta(days=i) for i in range(5)]
    df_week = df[(df['Fecha'] >= week_days[0]) & (df['Fecha'] <= week_days[-1]) & (df['Recurso'].isin(selected_recursos) if selected_recursos else df['Recurso'].notna())]

    st.markdown("---")
    
    # ---- LÓGICA DE CALENDARIO DEL ARCHIVO FUNCIONAL, FUSIONADA CON NUEVAS CARACTERÍSTICAS ----
    
    schedule = pd.DataFrame(index=HORAS, columns=[d.strftime('%A %d/%m') for d in week_days]).fillna('')

    for _, row in df_week.iterrows():
        day_str = row['Fecha'].strftime('%A %d/%m')
        
        for bloque in HORAS:
            try:
                h_inicio_b, h_fin_b = [dt.datetime.strptime(t.strip(), '%H:%M').time() for t in bloque.split(' a ')]
                if row['Hora inicio'] >= h_inicio_b and row['Hora fin'] <= h_fin_b:
                    prof_color = get_color_from_string(row['Profesor'])
                    observacion = row['Observaciones']
                    icon = " 📝" if observacion and observacion.strip() != '' else ""

                    card_content = f"<strong>{row['Recurso']}</strong>{icon}<br>{row['Profesor']}<br><em>{row['Curso']}</em>"
                    
                    if icon:
                        safe_observacion = html_sanitizer.escape(observacion)
                        card_html = f"<div class='reservation-card' style='background-color:{prof_color};'>{card_content}<span class='tooltip-text'>{safe_observacion}</span></div>"
                    else:
                        card_html = f"<div class='reservation-card' style='background-color:{prof_color};'>{card_content}</div>"

                    if schedule.at[bloque, day_str] == '':
                        schedule.at[bloque, day_str] = card_html
                    else:
                        schedule.at[bloque, day_str] += card_html
            except Exception:
                continue

    st.markdown(schedule.to_html(escape=False), unsafe_allow_html=True)


if page == "Dashboard":
    st.title("📈 Dashboard Analítico")
    with st.container(border=True):
        st.subheader("Filtrar por Rango de Fechas")
        today = dt.date.today()
        c1, c2 = st.columns(2)
        start_date = c1.date_input("Fecha de Inicio", today - dt.timedelta(days=30), format="DD/MM/YYYY")
        end_date = c2.date_input("Fecha de Fin", today, format="DD/MM/YYYY")
        if start_date > end_date:
            st.error("Error: La fecha de inicio no puede ser posterior a la fecha de fin.")
        else:
            df_filtered = df[(df['Fecha'] >= start_date) & (df['Fecha'] <= end_date)]
            st.markdown("---")
            st.subheader("Métricas Generales del Periodo")
            c1, c2, c3 = st.columns(3)
            c1.metric("Total de Reservas en Periodo", len(df_filtered))
            if not df_filtered.empty:
                try:
                    recurso_mas_usado = df_filtered['Recurso'].mode()[0]
                    c2.metric("Recurso Más Usado", recurso_mas_usado)
                    profesor_mas_activo = df_filtered['Profesor'].mode()[0]
                    c3.metric("Profesor Más Activo", profesor_mas_activo)
                except IndexError:
                    c2.info("No hay datos para mostrar métricas.")
            else:
                c2.info("No hay reservas en el periodo seleccionado.")
    if not df_filtered.empty:
        st.markdown("---")
        st.subheader("Mapa de Calor de Ocupación")
        with st.container(border=True):
            df_heatmap = df_filtered.copy()
            dias_semana_es_cat = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
            df_heatmap['DiaSemana'] = pd.Categorical(df_heatmap['Fecha'].apply(lambda x: dias_semana_es_cat[x.weekday()]), categories=dias_semana_es_cat, ordered=True)
            df_heatmap['BloqueHorario'] = df_heatmap.apply(lambda row: f"{row['Hora inicio'].strftime('%H:%M')} a {row['Hora fin'].strftime('%H:%M')}", axis=1)
            heatmap_data = df_heatmap.groupby(['DiaSemana', 'BloqueHorario']).size().reset_index(name='count')
            heatmap_pivot = heatmap_data.pivot_table(index='BloqueHorario', columns='DiaSemana', values='count', fill_value=0)
            heatmap_pivot = heatmap_pivot.loc[sorted(heatmap_pivot.index, key=sort_time_key)]
            heatmap_pivot = heatmap_pivot.reindex(columns=dias_semana_es_cat[:5], fill_value=0)
            fig = px.imshow(heatmap_pivot, labels=dict(x="Día de la Semana", y="Bloque Horario", color="N° de Reservas"), color_continuous_scale=px.colors.sequential.Reds)
            fig.update_layout(xaxis_title="", yaxis_title="")
            st.plotly_chart(fig, use_container_width=True)
        st.markdown("---")
        c_cursos, c_recursos = st.columns(2)
        with c_cursos:
            with st.container(border=True):
                st.subheader("Top 5 Cursos con más Reservas")
                cursos_count = df_filtered['Curso'].value_counts().nlargest(5)
                if not cursos_count.empty:
                    fig_cursos = px.bar(cursos_count, x=cursos_count.index, y=cursos_count.values, labels={'x': 'Curso', 'y': 'Cantidad de Reservas'})
                    st.plotly_chart(fig_cursos, use_container_width=True)
                else: st.info("No hay datos de cursos en este periodo.")
        with c_recursos:
            with st.container(border=True):
                st.subheader("Top 5 Recursos más Solicitados")
                recursos_count = df_filtered['Recurso'].value_counts().nlargest(5)
                if not recursos_count.empty:
                    fig_recursos = px.bar(recursos_count, x=recursos_count.index, y=recursos_count.values, labels={'x': 'Recurso', 'y': 'Cantidad de Reservas'})
                    st.plotly_chart(fig_recursos, use_container_width=True)
                else: st.info("No hay datos de recursos en este periodo.")

if page == "Mantenimiento":
    st.title("⚙️ Centro de Gestión")
    
    tab_periodos, tab_recursos, tab_profesores, tab_cursos, tab_horas = st.tabs(["Periodos", "Recursos", "Profesores", "Cursos", "Horas"])

    with tab_periodos:
        st.subheader("Gestionar Periodos de Mantenimiento")
        with st.expander("➕ Agregar Nuevo Periodo", expanded=True):
            with st.form("maint_form", clear_on_submit=True):
                rsrc_maint = st.selectbox("Recurso", RECURSOS)
                c1, c2 = st.columns(2)
                start_date = c1.date_input("Fecha Inicio", format="DD/MM/YYYY")
                start_time = c1.time_input("Hora Inicio", dt.time(8,0))
                end_date = c2.date_input("Fecha Fin", format="DD/MM/YYYY")
                end_time = c2.time_input("Hora Fin", dt.time(18,0))
                if st.form_submit_button("💾 Guardar Periodo", use_container_width=True, type="primary"):
                    try: mant_df = pd.read_excel(EXCEL, sheet_name='Mantenimientos').fillna("")
                    except ValueError: mant_df = pd.DataFrame(columns=["Recurso", "FechaInicio", "HoraInicio", "FechaFin", "HoraFin"])
                    new_row = {"Recurso":rsrc_maint, "FechaInicio":start_date.strftime('%d/%m/%Y'), "HoraInicio":start_time.strftime('%H:%M'), "FechaFin":end_date.strftime('%d/%m/%Y'), "HoraFin":end_time.strftime('%H:%M')}
                    updated_mant = pd.concat([mant_df, pd.DataFrame([new_row])], ignore_index=True)
                    save_excel_sheet(updated_mant, 'Mantenimientos', EXCEL)
        st.markdown("---")
        st.subheader("📋 Periodos Programados")
        try: mant_df = pd.read_excel(EXCEL, sheet_name='Mantenimientos').fillna("")
        except ValueError: mant_df = pd.DataFrame()
        if mant_df.empty:
            st.info("No hay mantenimientos programados.")
        else:
            edited_mant = st.data_editor(mant_df, hide_index=True, use_container_width=True, num_rows="dynamic")
            if st.button("💾 Guardar cambios en Periodos", use_container_width=True):
                save_excel_sheet(edited_mant, 'Mantenimientos', EXCEL)

    def create_list_manager(tab, sheet_name):
        with tab:
            st.subheader(f"Gestionar Lista de {sheet_name}")
            with st.container(border=True):
                try:
                    with pd.ExcelFile(EXCEL) as xls:
                        df_list = pd.read_excel(xls, sheet_name) if sheet_name in xls.sheet_names else pd.DataFrame(columns=[sheet_name] if sheet_name != "Profesores" else ["Profesores", "Email"])
                except Exception:
                    df_list = pd.DataFrame(columns=[sheet_name] if sheet_name != "Profesores" else ["Profesores", "Email"])
                
                edited_list = st.data_editor(df_list, num_rows="dynamic", use_container_width=True, hide_index=True)
                
                if st.button(f"💾 Guardar Cambios en {sheet_name}", use_container_width=True, type="primary"):
                    if sheet_name == "Horas": final_df = edited_list.dropna()
                    else: final_df = edited_list.dropna().sort_values(by=sheet_name)
                    save_excel_sheet(final_df, sheet_name, EXCEL)

    create_list_manager(tab_recursos, "Recursos")
    create_list_manager(tab_profesores, "Profesores")
    create_list_manager(tab_cursos, "Cursos")
    create_list_manager(tab_horas, "Horas")

