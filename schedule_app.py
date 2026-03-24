import datetime as dt
from datetime import date, datetime as dt_datetime
import time
from io import BytesIO
from pathlib import Path
import pandas as pd
import streamlit as st
import hashlib
import re
import html as html_sanitizer
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import plotly.express as px
import streamlit.components.v1 as components
from streamlit_autorefresh import st_autorefresh
import qrcode
import io
import urllib.parse
import zipfile

# ------------------------------------------------------------------
# CONFIGURACIÓN SUPABASE (NUEVO MOTOR DE BASE DE DATOS)
# ------------------------------------------------------------------
from supabase import create_client, Client, ClientOptions

URL_SUPABASE = "https://zxzpaubemwpwgvswvwjh.supabase.co"
CLAVE_SUPABASE = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inp4enBhdWJlbXdwd2d2c3d2d2poIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc3Mzg1NzMzMiwiZXhwIjoyMDg5NDMzMzMyfQ.CGWbTQprQaAhYruqlIkmMAMhx7EzD9hJ8QnJ7wCBxto"

opciones = ClientOptions(postgrest_client_timeout=60, storage_client_timeout=60)
supabase: Client = create_client(URL_SUPABASE, CLAVE_SUPABASE, options=opciones)

# ──────────────────────────────────────────────────────────────────────────────
# 0) CONFIGURACIÓN GLOBAL Y ESTILO
# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Sistema de Horarios CAV", page_icon="📅", layout="wide", initial_sidebar_state="expanded")

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

def get_color_from_string(input_string: str) -> str:
    if not input_string: return "#CCCCCC"
    hash_obj = hashlib.md5(input_string.encode())
    hash_int = int(hash_obj.hexdigest(), 16)
    hue = hash_int % 360; saturation = 75; lightness = 90
    return f"hsl({hue}, {saturation}%, {lightness}%)"

def sort_time_key(time_string):
    try:
        start_time_str = time_string.split(' a ')[0].strip()
        return dt.datetime.strptime(start_time_str, '%H:%M').time()
    except (ValueError, IndexError): return dt.time(23, 59)

def format_date_es(date_obj):
    dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
    meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    return f"{dias[date_obj.weekday()]}, {date_obj.day} de {meses[date_obj.month - 1]} de {date_obj.year}"

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
        pass 

# ------------------------------------------------------------------
# 0.5) VISTA PÚBLICA DE REPORTE (ACCESO POR CÓDIGO QR MÓVIL)
# ------------------------------------------------------------------
if "reportar" in st.query_params:
    st.markdown("<style>.block-container { padding-top: 2rem !important; }</style>", unsafe_allow_html=True)
    recurso_qr = st.query_params["reportar"]
    
    with st.container(border=True):
        st.markdown(f"<h2 style='text-align: center; color: var(--primary-color);'>🚨 Reporte de Falla</h2>", unsafe_allow_html=True)
        st.markdown(f"<h4 style='text-align: center; color: gray;'>Equipo: {recurso_qr}</h4>", unsafe_allow_html=True)
        st.markdown("---")
        
        try:
            rec_data = supabase.table("recursos").select("id").eq("nombre", recurso_qr).execute().data
            if not rec_data:
                st.error("❌ El equipo escaneado no existe en la base de datos.")
                st.stop()
            recurso_id = rec_data[0]["id"]
        except:
            st.error("Error de conexión al verificar el equipo.")
            st.stop()

        with st.form("qr_report_form"):
            detalle = st.text_area("Describe el problema detalladamente (Ej. 'El cable HDMI está roto', 'No enciende', etc.)")
            if st.form_submit_button("📤 Enviar Reporte a Técnicos", type="primary", use_container_width=True):
                if detalle.strip():
                    datos_mant = {
                        "recurso_id": recurso_id,
                        "fecha": dt.date.today().strftime("%Y-%m-%d"),
                        "descripcion": detalle.strip(),
                        "estado": "Reportado (Vía QR)"
                    }
                    try:
                        supabaseF.table("mantenimientos").insert(datos_mant).execute()
                        st.success("✅ ¡Gracias! Tu reporte ha sido enviado al equipo técnico.")
                        st.balloons()
                    except Exception as e:
                        st.error("Ocurrió un error al enviar el reporte.")
                else:
                    st.error("⚠️ Debes escribir una descripción del problema.")
        
    st.info("💡 Ya puedes cerrar esta pestaña en tu celular.")
    st.stop()

# ------------------------------------------------------------------
# 1) INICIALIZACIÓN DE DATOS
# ------------------------------------------------------------------
if 'PROFESORES' not in globals(): PROFESORES = []
if 'RECURSOS' not in globals(): RECURSOS = []
if 'CURSOS' not in globals(): CURSOS = []

def custom_course_sort_key(course_name):
    course_name = str(course_name).strip()
    if 'Dif' in course_name: return (3, 0, course_name)
    match = re.match(r"(\d+)°\s*(BÁSICO|MEDIO)\s*([A-Z])?", course_name, re.IGNORECASE)
    if match:
        num, level, letter = match.groups()
        level_priority = 0 if 'BÁSICO' in level.upper() else 1
        return (level_priority, int(num), letter or '')
    return (4, 0, course_name)

@st.cache_data(ttl=60)
def cargar_datos_login():
    try:
        p_res = supabase.table("profesores").select("nombre").execute().data
        profs = sorted([p["nombre"] for p in p_res]) if p_res else []
        r_res = supabase.table("recursos").select("nombre").execute().data
        recs = sorted([r["nombre"] for r in r_res]) if r_res else []
        c_res = supabase.table("cursos").select("nombre").execute().data
        curs = sorted([c["nombre"] for c in c_res], key=custom_course_sort_key) if c_res else []
        return profs, recs, curs
    except:
        return [], [], []

PROFESORES, RECURSOS, CURSOS = cargar_datos_login()

# ------------------------------------------------------------------
# 2) SISTEMA DE LOGIN HORIZONTAL
# ------------------------------------------------------------------
if "logged" not in st.session_state:
    st.session_state.logged = False
    st.session_state.role = None
    st.session_state.profesor_name = None

if not st.session_state.logged:
    st.markdown("""
        <style>
            .block-container { padding-top: 3rem !important; }
            [data-testid="stVerticalBlock"] { gap: 0.5rem !important; }
            .login-card { padding: 20px; border-radius: 15px; border: 1px solid #eeeeee; background-color: white; box-shadow: 0 4px 6px rgba(0,0,0,0.05); }
            .stTextInput, .stSelectbox { margin-bottom: -10px; }
            label { font-size: 0.85rem !important; font-weight: 600 !important; }
        </style>
    """, unsafe_allow_html=True)

    main_container = st.container()
    
    with main_container:
        col_logo, col_form = st.columns([1, 1.8], gap="large")

        with col_logo:
            st.markdown("<br><br>", unsafe_allow_html=True)
            BASE_DIR = Path(__file__).parent
            logo_path = BASE_DIR / "logocav.png"
            if logo_path.exists(): st.image(str(logo_path), use_container_width=True)
            else: st.info("Logo CAV")

        with col_form:
            st.markdown("<h2 style='text-align: center; color: #1E3A8A; margin-bottom: 0px;'>SISTEMA CAV</h2>", unsafe_allow_html=True)
            st.markdown("<p style='text-align: center; color: gray; font-size: 0.9rem;'>Reserva de Recursos y Espacios</p>", unsafe_allow_html=True)
            
            with st.container(border=True):
                tipo_user = st.radio("Acceder como:", ["Profesor", "Administrador"], horizontal=True)
                st.markdown("---")
                
                if tipo_user == "Administrador":
                    with st.form("admin_form", clear_on_submit=True):
                        u_adm = st.text_input("Nombre de Administrador", placeholder="Ej: Edgar")
                        p_adm = st.text_input("Contraseña", type="password", placeholder="••••••••")
                        if st.form_submit_button("INICIAR SESIÓN ADMIN", use_container_width=True, type="primary"):
                            if u_adm.strip().upper() in ["EDGAR", "GLORIA", "CARLOS", "ALEXIS"] and p_adm == "cav690":
                                st.session_state.logged = True
                                st.session_state.role = "admin"
                                st.session_state.profesor_name = u_adm.strip().capitalize()
                                st.rerun()
                            else:
                                st.error("Acceso denegado")
                else:
                    with st.form("profe_form", clear_on_submit=True):
                        u_profe = st.selectbox("Busca tu nombre", PROFESORES, index=None, placeholder="Selecciona...")
                        p_profe = st.text_input("Clave de Acceso", type="password", placeholder="6904")
                        if st.form_submit_button("ENTRAR AL PANEL", use_container_width=True, type="primary"):
                            if u_profe and p_profe == "6904":
                                st.session_state.logged = True
                                st.session_state.role = "profesor"
                                st.session_state.profesor_name = u_profe
                                st.rerun()
                            elif not u_profe: st.warning("Por favor selecciona tu nombre")
                            else: st.error("Contraseña incorrecta")
    st.stop() 

# ------------------------------------------------------------------
# 3) CARGA DE LA BASE DE DATOS PRINCIPAL 
# ------------------------------------------------------------------
@st.cache_data(ttl=30)
def cargar_reservas_y_datos():
    horas_corregidas = [
        '8:00 a 8:45', '8:45 a 9:30', '8:00 a 9:30',
        '9:45 a 10:30', '10:30 a 11:15', '9:45 a 11:15',
        '11:30 a 12:15', '12:15 a 13:00', '11:30 a 13:00',
        '14:00 a 14:45', '14:45 a 15:30', '14:00 a 15:30',
        '14:00 a 16:30', '14:45 a 16:30', '15:45 a 16:30',
        '17:00 a 18:30', '17:30 a 18:30'
    ]
    
    try:
        res_data = supabase.table("reservas").select("id, fecha, hora_inicio, hora_fin, observaciones, profesores(nombre), cursos(nombre), recursos(nombre)").execute().data
        reservas_limpias = []
        for r in res_data:
            reservas_limpias.append({
                "id": r["id"],
                "Fecha": parse_date(r["fecha"]),
                "Hora inicio": as_time(r["hora_inicio"]),
                "Hora fin": as_time(r["hora_fin"]),
                "Profesor": r["profesores"]["nombre"] if r.get("profesores") else "",
                "Curso": r["cursos"]["nombre"] if r.get("cursos") else "",
                "Recurso": r["recursos"]["nombre"] if r.get("recursos") else "",
                "Observaciones": r["observaciones"]
            })
        df_res = pd.DataFrame(reservas_limpias) if reservas_limpias else pd.DataFrame(columns=['id', 'Fecha', 'Hora inicio', 'Hora fin', 'Profesor', 'Curso', 'Recurso', 'Observaciones'])
            
        try:
            # CORRECCIÓN DEFINITIVA DE MANTENIMIENTO: Carga simple sin JOIN para evitar errores
            mant_data = supabase.table("mantenimientos").select("*").execute().data
            rec_data = supabase.table("recursos").select("id, nombre").execute().data
            mant_map_rec = {r['id']: r['nombre'] for r in rec_data} if rec_data else {}
            
            df_mant = pd.DataFrame(mant_data) if mant_data else pd.DataFrame()
            
            if not df_mant.empty:
                df_mant = df_mant[df_mant['estado'] != 'Reparado']
                df_mant['FechaInicio_dt'] = df_mant['fecha'].apply(parse_date) if 'fecha' in df_mant.columns else dt.date.today()
                df_mant['FechaFin_dt'] = df_mant['FechaInicio_dt']
                df_mant['HoraInicio'] = dt.time(0, 0)
                df_mant['HoraFin'] = dt.time(23, 59)
                if 'recurso_id' in df_mant.columns:
                    df_mant['Recurso'] = df_mant['recurso_id'].apply(lambda x: mant_map_rec.get(x, 'Desconocido'))
                else: df_mant['Recurso'] = 'Desconocido'
            else:
                df_mant = pd.DataFrame(columns=['Recurso', 'FechaInicio_dt', 'HoraInicio', 'FechaFin_dt', 'HoraFin'])
        except Exception as e:
            df_mant = pd.DataFrame(columns=['Recurso', 'FechaInicio_dt', 'HoraInicio', 'FechaFin_dt', 'HoraFin'])

        return df_res, horas_corregidas, df_mant
    except Exception as e:
        return pd.DataFrame(columns=['id', 'Fecha', 'Hora inicio', 'Hora fin', 'Profesor', 'Curso', 'Recurso', 'Observaciones']), horas_corregidas, pd.DataFrame()

df, HORAS, df_mantenimiento = cargar_reservas_y_datos()

map_prof, map_cur, map_rec, PROFESOR_DATA = {}, {}, {}, {}
try: 
    prof_data_db = supabase.table("profesores").select("id, nombre, email").execute().data
    map_prof = {p["nombre"]: p["id"] for p in prof_data_db}; PROFESOR_DATA = {p["nombre"]: p.get("email", "") for p in prof_data_db}
except: pass
try: map_cur = {c["nombre"]: c["id"] for c in supabase.table("cursos").select("id, nombre").execute().data}
except: pass
try: map_rec = {r["nombre"]: r["id"] for r in supabase.table("recursos").select("id, nombre").execute().data}
except: pass

# ------------------------------------------------------------------
# 4) NAVEGACIÓN Y VISTAS
# ------------------------------------------------------------------
st_autorefresh(interval=300000, key="data_refresh")

sidebar_title = f"Panel de {st.session_state.role.capitalize()}"
if st.session_state.role == 'profesor': sidebar_title = f"Hola, {st.session_state.profesor_name.split(' ')[0]}"

BASE_DIR = Path(__file__).parent
logo_path = BASE_DIR / "logocav.png"

if logo_path.exists():
    st.markdown("""<style>[data-testid="stSidebarNav"]::before { content: ""; display: none; margin-top: 0px; } section[data-testid="stSidebar"] div.st-emotion-cache-16t70r2 { padding-top: 0.5rem !important; }</style>""", unsafe_allow_html=True)
    col1, col2, col3 = st.sidebar.columns([0.1, 2.8, 0.1])
    with col2: st.image(str(logo_path), use_container_width=True)

st.sidebar.markdown(f"<div style='text-align: center; color: var(--primary-color); font-weight: bold; margin-bottom: 0px; font-size: 1.1em; letter-spacing: 1px;'>{sidebar_title.upper()}</div>", unsafe_allow_html=True)
st.sidebar.markdown("<hr style='margin: 8px 0px 5px 0px; padding: 0;'>", unsafe_allow_html=True)

html_reloj = """
<!DOCTYPE html>
<html>
<head>
<style>
    body { margin: 0; padding: 0; font-family: "Source Sans Pro", sans-serif; text-align: center; display: flex; justify-content: center; align-items: center; overflow: hidden; }
    .container { padding: 0px; width: 100%; }
    @media (prefers-color-scheme: dark) { .container { background-color: transparent; color: #FAFAFA; } .reloj { color: #ff4b4b; } .mensaje { color: #c6c6d1; } }
    @media (prefers-color-scheme: light) { .container { background-color: transparent; color: #31333F; } .reloj { color: #ff4b4b; } .mensaje { color: #555555; } }
    .fecha { font-weight: 600; font-size: 0.9em; text-transform: capitalize; margin-bottom: 2px;}
    .reloj { font-size: 1.6em; font-weight: 700; margin: 2px 0; font-variant-numeric: tabular-nums;}
    .mensaje { font-size: 0.8em; font-style: italic; margin-top: 2px;}
</style>
</head>
<body>
    <div class="container">
        <div id="fecha" class="fecha"></div>
        <div id="reloj" class="reloj"></div>
        <div id="mensaje" class="mensaje"></div>
    </div>
    <script>
        const mensajes = ["¡Que tengas un excelente día! ☀️", "Cada día es una nueva oportunidad para brillar. ✨", "Tu esfuerzo de hoy es el éxito de mañana. 💪", "Haz que las cosas pasen. ¡Tú puedes! 🚀", "Pequeños pasos todos los días llevan a grandes resultados. 🏔️", "Sonríe, respira y sigue adelante. 🌻", "La actitud lo es todo. ¡A dar el 100%! 💯", "Hoy es un buen día para hacer la diferencia. 🌟", "Tu trabajo y dedicación son muy valiosos. 🤝", "¡Mucho éxito en todas tus tareas de hoy! 🎯"];
        document.getElementById("mensaje").innerText = mensajes[Math.floor(Math.random() * mensajes.length)];
        function actualizarReloj() {
            const ahora = new Date();
            document.getElementById("reloj").innerText = ahora.toLocaleTimeString('es-CL', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
            document.getElementById("fecha").innerText = ahora.toLocaleDateString('es-CL', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });
        }
        setInterval(actualizarReloj, 1000); actualizarReloj();
    </script>
</body>
</html>
"""
with st.sidebar:
    components.html(html_reloj, height=85)
    st.markdown("<hr style='margin: 0px 0px 10px 0px; padding: 0;'>", unsafe_allow_html=True)

PAGES_CONFIG = {
    "Mis Reservas": {"icon": "👤", "roles": ["profesor"]},
    "Registrar": {"icon": "📝", "roles": ["admin"]},
    "Base de datos": {"icon": "🗃️", "roles": ["admin"]},
    "Semana": {"icon": "🗓️", "roles": ["admin", "profesor"]},
    "Dashboard": {"icon": "📈", "roles": ["admin"]},
    "Técnicos": {"icon": "🔧", "roles": ["admin"]},
    "Configuración": {"icon": "⚙️", "roles": ["admin"]},
}

available_pages = [p for p, conf in PAGES_CONFIG.items() if st.session_state.role in conf["roles"]]
default_page = "Mis Reservas" if st.session_state.role == 'profesor' else "Registrar"
page = st.sidebar.radio("Navegación", available_pages, index=available_pages.index(default_page), format_func=lambda p: f"{PAGES_CONFIG[p]['icon']} {p}", label_visibility="collapsed")

st.sidebar.markdown("---")

if st.sidebar.button("🔄 Refrescar Pantalla", use_container_width=True):
    st.cache_data.clear(); st.rerun()

if st.sidebar.button("🚪 Cerrar Sesión", use_container_width=True):
    for key in st.session_state.keys(): del st.session_state[key]
    st.rerun()

# ------------------------------------------------------------------
# PÁGINAS
# ------------------------------------------------------------------

if page == "Mis Reservas":
    st.title("👤 Mis Próximas Reservas")
    if not df.empty:
        prof_df = df[df['Profesor'] == st.session_state.profesor_name]
        future_reservas = prof_df[prof_df['Fecha'] >= dt.date.today()].sort_values(by="Fecha")
    else: future_reservas = pd.DataFrame()

    if future_reservas.empty: st.info("No tienes reservas programadas para el futuro.")
    else:
        for _, row in future_reservas.iterrows():
            with st.container(border=True):
                st.markdown(f"#### {format_date_es(row['Fecha'])}")
                st.markdown(f"**Hora:** {row['Hora inicio'].strftime('%H:%M')} - {row['Hora fin'].strftime('%H:%M')}<br>**Curso:** {row['Curso']}<br>**Recurso:** {row['Recurso']}", unsafe_allow_html=True)
                if row['Observaciones']: st.markdown(f"> *{row['Observaciones']}*")

if page == "Registrar":
    st.title("📝 Registrar Nuevo Horario")
    def check_all_conflicts(fechas, recursos, hora_inicio, hora_fin, df_reservas, df_mantenimiento):
        conflictos_reserva, conflictos_mantenimiento = [], []
        if not df_reservas.empty:
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
                    conflictos_r, conflictos_m = check_all_conflicts(fechas_a_registrar, recs, h_inicio, h_fin, df, df_mantenimiento)
                    if conflictos_r or conflictos_m:
                        if conflictos_r: st.error(f"❌ **Conflicto de Reserva:**"); st.markdown(f"<ul>{''.join(conflictos_r)}</ul>", unsafe_allow_html=True)
                        if conflictos_m: st.error(f"❌ **Conflicto de Mantenimiento:**"); st.markdown(f"<ul>{''.join(conflictos_m)}</ul>", unsafe_allow_html=True)
                    else:
                        nuevas_reservas = []
                        for fecha in fechas_a_registrar:
                            for rec in recs:
                                nuevas_reservas.append({
                                    'fecha': str(fecha), 
                                    'hora_inicio': str(h_inicio), 
                                    'hora_fin': str(h_fin), 
                                    'profesor': map_prof.get(prof), 
                                    'curso': map_cur.get(curso), 
                                    'recurso': map_rec.get(rec), 
                                    'observaciones': obs
                                })
                        try:
                            supabase.table("reservas").insert(nuevas_reservas).execute()
                            st.success("✅ ¡Reservas guardadas exitosamente!")
                            st.cache_data.clear()
                            
                            email_to = PROFESOR_DATA.get(prof)
                            if email_to:
                                subject = f"Confirmación de Reserva de Recursos - {curso}"
                                body = f"""<html><body><p>Hola {prof.split(' ')[0]},</p><p>Se ha(n) confirmado la(s) siguiente(s) reserva(s) a tu nombre:</p><ul><li><b>Curso:</b> {curso}</li><li><b>Recurso(s):</b> {', '.join(recs)}</li><li><b>Horario:</b> {hora}</li></ul><p><b>Fechas Registradas:</b></p><ul>{''.join([f'<li>{format_date_es(f)}</li>' for f in fechas_a_registrar])}</ul>{f"<p><b>Observaciones:</b> {obs}</p>" if obs else ""}<p>Saludos,<br>Sistema de Horarios CAV</p></body></html>"""
                                send_email(subject, body, email_to)
                                
                            time.sleep(1); st.rerun()
                        except Exception as e:
                            st.error(f"Error al guardar en la nube: {e}")

if page == "Base de datos":
    st.title("🗃️ Base de Datos de Reservas")
    st.info("Nota: Para eliminar una fila selecciónala y presiona Supr/Delete, luego guarda.")
    with st.container(border=True):
        if not df.empty:
            df_display = df.drop(columns=['id'])
            edited_df = st.data_editor(df_display, hide_index=True, use_container_width=True, num_rows="dynamic", column_config={"Fecha": st.column_config.DateColumn("Fecha", format="DD/MM/YYYY"), "Hora inicio": st.column_config.TimeColumn("Hora Inicio", format="HH:mm"), "Hora fin": st.column_config.TimeColumn("Hora Fin", format="HH:mm"), "Profesor": st.column_config.SelectboxColumn("Profesor", options=PROFESORES, required=True), "Curso": st.column_config.SelectboxColumn("Curso", options=CURSOS, required=True), "Recurso": st.column_config.SelectboxColumn("Recurso", options=RECURSOS, required=True)})
            
            if st.button("💾 Guardar Cambios en la Nube", use_container_width=True, type="primary"):
                with st.spinner("Sincronizando con Supabase..."):
                    try:
                        original_indices = set(df.index)
                        edited_indices = set(edited_df.index)
                        deleted_indices = original_indices - edited_indices
                        
                        for idx in deleted_indices:
                            id_borrar = df.loc[idx, 'id']
                            supabase.table("reservas").delete().eq("id", id_borrar).execute()
                            
                            prof_name = df.loc[idx, 'Profesor']
                            email_to = PROFESOR_DATA.get(prof_name)
                            if email_to:
                                subject = f"Cancelación de Reserva de Recursos - {df.loc[idx, 'Curso']}"
                                body = f"""<html><body><p>Hola {prof_name.split(' ')[0]},</p><p>Te informamos que la siguiente reserva ha sido <b>cancelada</b>:</p><ul><li><b>Fecha:</b> {format_date_es(df.loc[idx, 'Fecha'])}</li><li><b>Horario:</b> {df.loc[idx, 'Hora inicio'].strftime('%H:%M')} - {df.loc[idx, 'Hora fin'].strftime('%H:%M')}</li><li><b>Curso:</b> {df.loc[idx, 'Curso']}</li><li><b>Recurso:</b> {df.loc[idx, 'Recurso']}</li></ul><p>Saludos,<br>Sistema de Horarios CAV</p></body></html>"""
                                send_email(subject, body, email_to)

                        new_rows = edited_df[~edited_df.index.isin(original_indices)]
                        if not new_rows.empty:
                            nuevas_inserciones = []
                            for _, r in new_rows.iterrows():
                                nuevas_inserciones.append({
                                    "fecha": str(r["Fecha"]),
                                    "hora_inicio": str(r["Hora inicio"]),
                                    "hora_fin": str(r["Hora fin"]),
                                    "profesor": map_prof.get(r["Profesor"]),
                                    "curso": map_cur.get(r["Curso"]),
                                    "recurso": map_rec.get(r["Recurso"]),
                                    "observaciones": r["Observaciones"]
                                })
                            supabase.table("reservas").insert(nuevas_inserciones).execute()
                            
                        st.success("Sincronización completa.")
                        st.cache_data.clear(); time.sleep(1); st.rerun()
                    except Exception as e:
                        st.error(f"Error al sincronizar: {e}")
        else:
            st.write("No hay datos registrados.")

# --- VISTA SEMANAL ---
elif page == "Semana": 
    import json # Necesario para la animación del monitor
    st.header("🗓️ Vista Semanal")
    
    def get_color_from_string(s):
        import hashlib
        hash_val = int(hashlib.md5(s.encode('utf-8')).hexdigest(), 16)
        hue = hash_val % 360
        return f"hsl({hue}, 75%, 50%)"

    with st.container(border=True):
        st.write("🔍 **Filtros de Búsqueda Avanzados**")
        col_d, col_r, col_p, col_c = st.columns(4)
        
        recursos_list = df['Recurso'].dropna().unique().tolist() if not df.empty else []
        profesores_list = df['Profesor'].dropna().unique().tolist() if not df.empty else []
        cursos_list = df['Curso'].dropna().unique().tolist() if not df.empty else []
        
        if not df.empty:
            df['fecha_obj'] = pd.to_datetime(df['Fecha']).dt.date
            default_date_week = df['fecha_obj'].max()
        else:
            default_date_week = dt.date.today()
            
        selected_date = col_d.date_input("Semana del", value=default_date_week, format="DD/MM/YYYY")
        
        selected_recursos = col_r.multiselect("Filtrar Recursos", recursos_list, placeholder="Todos")
        selected_profesores = col_p.multiselect("Filtrar Profesores", profesores_list, placeholder="Todos")
        selected_cursos = col_c.multiselect("Filtrar Cursos", cursos_list, placeholder="Todos")

    start_of_week = selected_date - dt.timedelta(days=selected_date.weekday())
    week_days = [start_of_week + dt.timedelta(days=i) for i in range(5)]
    
    if not df.empty:
        mask = (df['fecha_obj'] >= week_days[0]) & (df['fecha_obj'] <= week_days[-1])
        if selected_recursos: mask &= df['Recurso'].isin(selected_recursos)
        if selected_profesores: mask &= df['Profesor'].isin(selected_profesores)
        if selected_cursos: mask &= df['Curso'].isin(selected_cursos)
        df_week = df[mask]
    else:
        df_week = pd.DataFrame()

    st.markdown("---")
    
    # --- INTERRUPTOR MODO TV ---
    modo_tv = st.toggle("📺 Activar Modo Monitor (Pantalla Automática)", value=False)
    
    dias_es = {0: 'Lunes', 1: 'Martes', 2: 'Miércoles', 3: 'Jueves', 4: 'Viernes', 5: 'Sábado', 6: 'Domingo'}

    if modo_tv:
        # ======= MODO PANTALLA ANIMADA Y FULLSCREEN =======
        if df_week.empty:
            st.info("No hay reservas esta semana para mostrar en la pantalla.")
        else:
            df_tv = df_week.sort_values(by=['fecha_obj', 'Hora inicio'])
            
            # Preparar los datos para Javascript
            tv_data = []
            for _, row in df_tv.iterrows():
                tv_data.append({
                    "dia": dias_es[row['fecha_obj'].weekday()],
                    "fecha": row['fecha_obj'].strftime('%d/%m'),
                    "horario": f"{str(row['Hora inicio'])[:5]} a {str(row['Hora fin'])[:5]}",
                    "recurso": str(row['Recurso']),
                    "profesor": str(row['Profesor']),
                    "curso": str(row['Curso']),
                    "color": get_color_from_string(str(row['Profesor']))
                })
            
            tv_data_json = json.dumps(tv_data)
            
            # Código HTML/JS con Carrusel Automático y Botón FullScreen
            tv_html = f"""
            <!DOCTYPE html>
            <html>
            <head>
            <style>
                @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;700;900&display=swap');
                body {{ margin: 0; padding: 0; background-color: #0f172a; font-family: 'Inter', sans-serif; color: white; overflow: hidden; }}
                .tv-wrapper {{ display: flex; flex-direction: column; height: 100vh; width: 100vw; background: radial-gradient(circle at top right, #1e293b, #0f172a); }}
                
                .header {{ display: flex; justify-content: space-between; align-items: center; padding: 25px 40px; background: rgba(15, 23, 42, 0.8); border-bottom: 2px solid #334155; box-shadow: 0 4px 15px rgba(0,0,0,0.5); }}
                .title {{ font-size: 32px; font-weight: 900; color: #38bdf8; letter-spacing: 2px; text-transform: uppercase; }}
                .fullscreen-btn {{ background: #0284c7; color: white; border: none; padding: 12px 24px; font-size: 18px; border-radius: 8px; cursor: pointer; font-weight: bold; transition: background 0.3s; }}
                .fullscreen-btn:hover {{ background: #38bdf8; color: #0f172a; }}
                
                .content {{ flex-grow: 1; padding: 40px; display: flex; flex-direction: column; justify-content: center; gap: 20px; }}
                
                .tv-card {{ background: rgba(255, 255, 255, 0.05); backdrop-filter: blur(10px); border-left: 10px solid; border-radius: 15px; padding: 30px; display: flex; align-items: center; box-shadow: 0 10px 25px rgba(0,0,0,0.3); opacity: 0; transform: translateY(30px); transition: all 0.6s cubic-bezier(0.4, 0, 0.2, 1); }}
                .tv-card.visible {{ opacity: 1; transform: translateY(0); }}
                
                .time-box {{ min-width: 280px; border-right: 3px solid rgba(255,255,255,0.1); margin-right: 40px; padding-right: 20px; }}
                .date-text {{ font-size: 22px; color: #94a3b8; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 8px; }}
                .time-text {{ font-size: 40px; font-weight: 900; color: #e2e8f0; }}
                
                .details-box {{ flex-grow: 1; }}
                .resource-text {{ font-size: 45px; font-weight: 900; color: #38bdf8; margin-bottom: 12px; text-shadow: 0 2px 5px rgba(0,0,0,0.5); }}
                .prof-text {{ font-size: 28px; color: #cbd5e1; font-weight: 400; }}
                
                .progress-bar {{ position: absolute; bottom: 0; left: 0; height: 6px; background: #38bdf8; width: 0%; transition: width 8s linear; }}
            </style>
            </head>
            <body>
                <div class="tv-wrapper" id="tv-wrapper">
                    <div class="header">
                        <div class="title">📡 HORARIO DE ENLACES Y RECURSOS</div>
                        <button class="fullscreen-btn" onclick="toggleFullScreen()">🔲 Pantalla Completa</button>
                    </div>
                    <div class="content" id="cards-container"></div>
                    <div class="progress-bar" id="progress"></div>
                </div>

                <script>
                    const data = {tv_data_json};
                    const container = document.getElementById('cards-container');
                    const progress = document.getElementById('progress');
                    let currentIndex = 0;
                    const itemsPerPage = 4; // Muestra 4 bloques por pantalla
                    const slideDuration = 8000; // 8 segundos por pantalla

                    function toggleFullScreen() {{
                        const elem = document.documentElement;
                        if (!document.fullscreenElement) {{
                            elem.requestFullscreen().catch(err => console.log(err));
                        }} else {{
                            if (document.exitFullscreen) document.exitFullscreen();
                        }}
                    }}

                    function renderCards() {{
                        container.innerHTML = '';
                        
                        // Reiniciar barra de progreso
                        progress.style.transition = 'none';
                        progress.style.width = '0%';
                        setTimeout(() => {{
                            progress.style.transition = `width ${{slideDuration}}ms linear`;
                            progress.style.width = '100%';
                        }}, 50);

                        // Seleccionar los datos de esta página
                        const pageData = [];
                        for(let i=0; i<itemsPerPage; i++) {{
                            if(data.length > 0) {{
                                pageData.push(data[(currentIndex + i) % data.length]);
                            }}
                        }}
                        
                        // Eliminar duplicados visuales si hay menos reservas que itemsPerPage
                        const uniquePageData = [...new Set(pageData)];

                        uniquePageData.forEach((item, index) => {{
                            const card = document.createElement('div');
                            card.className = 'tv-card';
                            card.style.borderLeftColor = item.color;
                            card.innerHTML = `
                                <div class="time-box">
                                    <div class="date-text">${{item.dia}} ${{item.fecha}}</div>
                                    <div class="time-text">🕒 ${{item.horario}}</div>
                                </div>
                                <div class="details-box">
                                    <div class="resource-text">${{item.recurso}}</div>
                                    <div class="prof-text">👨‍🏫 ${{item.profesor}} &nbsp;|&nbsp; 📚 ${{item.curso}}</div>
                                </div>
                            `;
                            container.appendChild(card);
                            
                            // Animación en cascada para entrar
                            setTimeout(() => {{ card.classList.add('visible'); }}, index * 200);
                        }});

                        // Avanzar el índice
                        if (data.length > itemsPerPage) {{
                            currentIndex = (currentIndex + itemsPerPage) % data.length;
                        }}
                    }}

                    renderCards();
                    if(data.length > itemsPerPage) {{
                        setInterval(renderCards, slideDuration);
                    }} else {{
                        // Si hay pocas reservas, dejar barra llena y no rotar
                        progress.style.transition = 'width 1s linear';
                        progress.style.width = '100%';
                    }}
                </script>
            </body>
            </html>
            """
            
            # Usamos una altura alta para visualizar bien antes del fullscreen
            components.html(tv_html, height=800)

    else:
        # ======= MODO TABLA NORMAL =======
        column_names = [f"{dias_es[d.weekday()]}<br><span style='font-size:0.8em; font-weight:normal;'>{d.strftime('%d/%m')}</span>" for d in week_days]
        
        base_horas = [
            "08:00 a 08:45", "08:00 a 09:30", "08:45 a 09:30", "09:45 a 10:30",
            "09:45 a 11:15", "10:30 a 11:15", "11:30 a 12:15", "11:30 a 13:00",
            "12:15 a 13:00", "14:00 a 14:45", "14:00 a 15:30", "14:00 a 16:30",
            "14:45 a 15:30", "14:45 a 16:30", "15:45 a 16:30", "17:00 a 18:30",
            "17:30 a 18:30"
        ]
        
        if not df.empty:
            df['bloque_hora'] = df['Hora inicio'].astype(str).str[:5] + " a " + df['Hora fin'].astype(str).str[:5]
            dynamic_horas = df['bloque_hora'].unique().tolist()
        else:
            dynamic_horas = []

        HORAS = sorted(list(set(base_horas + dynamic_horas)))
        schedule = pd.DataFrame(index=HORAS, columns=column_names).fillna('')

        if not df_week.empty:
            for _, row in df_week.iterrows():
                day_col = f"{dias_es[row['fecha_obj'].weekday()]}<br><span style='font-size:0.8em; font-weight:normal;'>{row['fecha_obj'].strftime('%d/%m')}</span>"
                bloque_actual = f"{str(row['Hora inicio'])[:5]} a {str(row['Hora fin'])[:5]}"
                
                if day_col in column_names:
                    prof_color = get_color_from_string(str(row['Profesor']))
                    observacion = str(row['Observaciones']) if pd.notna(row['Observaciones']) and str(row['Observaciones']).strip() != '' else ""
                    icon = " 📌" if observacion else ""

                    card_content = f"<div style='font-weight:bold; color:#1e293b; margin-bottom:4px;'>{row['Recurso']}{icon}</div><div style='color:#334155; margin-bottom:2px; font-size:0.95em;'>👨‍🏫 {row['Profesor']}</div><div style='color:#64748b; font-style:italic; font-size:0.9em;'>📚 {row['Curso']}</div>"
                    
                    safe_observacion = html_sanitizer.escape(observacion)
                    card_html = f"<div style='background-color:#ffffff; padding:12px; margin-bottom:8px; border-radius:6px; border-left: 5px solid {prof_color}; box-shadow: 0 2px 5px rgba(0,0,0,0.08); font-size:0.85em; text-align:left; cursor:help;' title='{safe_observacion}'>{card_content}</div>"

                    if schedule.at[bloque_actual, day_col] == '': 
                        schedule.at[bloque_actual, day_col] = card_html
                    else: 
                        schedule.at[bloque_actual, day_col] += card_html

        st.markdown("""
            <style>
            .dataframe { width: 100%; border-collapse: separate; border-spacing: 0; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.05); }
            .dataframe th { background: linear-gradient(135deg, #4A90E2 0%, #357ABD 100%); color: white; padding: 15px 10px; text-align: center; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; border-right: 1px solid rgba(255,255,255,0.2); }
            .dataframe th:last-child { border-right: none; }
            .dataframe td { border-bottom: 1px solid #E2E8F0; border-right: 1px solid #E2E8F0; padding: 12px; vertical-align: top; min-width: 170px; background-color: #F8FAFC; }
            .dataframe td:last-child { border-right: none; }
            .dataframe tbody tr:hover td { background-color: #F1F5F9; }
            .dataframe th[scope="row"] { background: #EDF2F7; color: #2D3748; font-weight: bold; text-align: center; width: 110px; border-right: 1px solid #CBD5E1; }
            </style>
        """, unsafe_allow_html=True)
        
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
            if not df.empty: df_filtered = df[(df['Fecha'] >= start_date) & (df['Fecha'] <= end_date)]
            else: df_filtered = pd.DataFrame()
                
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

# ------------------------------------------------------------------
# SECCIÓN: TÉCNICOS
# ------------------------------------------------------------------
if page == "Técnicos":
    st.title("🔧 Área de Técnicos y Mantenimiento")
    
    tab_mant, tab_qr = st.tabs(["🛠️ Gestión de Reportes", "📲 Generador de Códigos QR"])
    
    with tab_mant:
        c1, c2 = st.columns([1, 2])
        with c1:
            st.write("#### Actualizar / Nuevo Reporte")
            with st.form("form_mant_config"):
                rec_mant = st.selectbox("Selecciona el Recurso", list(map_rec.keys()) if map_rec else ["No hay recursos"])
                fecha_mant = st.date_input("Fecha del reporte", dt.date.today())
                estado = st.selectbox("Estado", ["Reportado (Vía QR)", "En Reparación", "Dado de Baja", "Reparado"])
                detalle = st.text_area("Descripción de la falla")
                
                if st.form_submit_button("Guardar/Actualizar Reporte", use_container_width=True, type="primary"):
                    if rec_mant == "No hay recursos" or not detalle.strip():
                        st.error("Por favor completa la descripción.")
                    else:
                        datos_mant = {
                            "recurso_id": map_rec[rec_mant],
                            "fecha": fecha_mant.strftime("%Y-%m-%d"),
                            "descripcion": detalle,
                            "estado": estado
                        }
                        try:
                            supabase.table("mantenimientos").insert(datos_mant).execute()
                            st.success(f"Reporte guardado para {rec_mant}.")
                            st.cache_data.clear(); time.sleep(1); st.rerun()
                        except Exception as e:
                            st.error(f"Error al guardar: {e}")

        with c2:
            st.write("#### 📋 Historial de Equipos (Editable)")
            st.caption("✏️ Doble clic en una celda para editar. Para eliminar, selecciona la fila (casilla izquierda) y presiona **Suprimir/Delete**. Luego haz clic en Guardar.")
            
            try:
                mants = supabase.table("mantenimientos").select("*, recursos(nombre)").order("fecha", desc=True).execute().data
                
                if mants:
                    df_mants = pd.DataFrame(mants)
                    df_mants['Recurso'] = df_mants['recursos'].apply(lambda x: x['nombre'] if isinstance(x, dict) else 'Desconocido')
                    
                    df_mostrar = df_mants[['id', 'fecha', 'Recurso', 'descripcion', 'estado']].copy()
                    df_mostrar = df_mostrar.rename(columns={'fecha': 'Fecha', 'descripcion': 'Detalle', 'estado': 'Estado'})

                    # 🔥 SOLUCIÓN AQUÍ: Convertimos los textos a Objetos de Fecha Reales
                    df_mostrar['Fecha'] = df_mostrar['Fecha'].apply(parse_date)

                    editado = st.data_editor(
                        df_mostrar,
                        column_config={
                            "id": None, 
                            "Recurso": st.column_config.TextColumn("Recurso", disabled=True), 
                            "Fecha": st.column_config.DateColumn("Fecha", format="DD/MM/YYYY"),
                            "Estado": st.column_config.SelectboxColumn(
                                "Estado",
                                help="Selecciona el estado actual",
                                options=["Reportado (Vía QR)", "En Revisión", "Reparado", "Dado de Baja"],
                                required=True,
                            )
                        },
                        use_container_width=True,
                        hide_index=False,
                        num_rows="dynamic",
                        key="editor_mantenimientos"
                    )

                    if st.button("💾 Guardar Cambios en la Base de Datos", type="primary", use_container_width=True):
                        cambios = st.session_state["editor_mantenimientos"]
                        
                        if cambios.get("deleted_rows"):
                            for row_index in cambios["deleted_rows"]:
                                registro_id = int(df_mostrar.iloc[row_index]['id'])
                                supabase.table("mantenimientos").delete().eq("id", registro_id).execute()
                        
                        if cambios.get("edited_rows"):
                            for row_index, modificaciones in cambios["edited_rows"].items():
                                registro_id = int(df_mostrar.iloc[row_index]['id'])
                                datos_actualizar = {}
                                
                                # 🔥 SOLUCIÓN AL GUARDAR: Convertimos la fecha de vuelta a texto para Supabase
                                if "Fecha" in modificaciones: 
                                    datos_actualizar["fecha"] = str(modificaciones["Fecha"])[:10] 
                                if "Detalle" in modificaciones: 
                                    datos_actualizar["descripcion"] = modificaciones["Detalle"]
                                if "Estado" in modificaciones: 
                                    datos_actualizar["estado"] = modificaciones["Estado"]
                                
                                if datos_actualizar:
                                    supabase.table("mantenimientos").update(datos_actualizar).eq("id", registro_id).execute()

                        st.success("✅ ¡Base de datos actualizada correctamente!")
                        time.sleep(1)
                        st.rerun()

                else:
                    st.info("No hay registros de mantenimiento activos.")
            except Exception as e:
                st.warning(f"Error al cargar o modificar el historial: {e}")
    with tab_qr:
        st.subheader("🖨️ Generador de Códigos QR para Equipos")
        st.write("Escribe el enlace de la aplicación. Al generarse los códigos, podrás descargar un archivo `.zip` con todos los QR listos para imprimir.")
        
        url_base = st.text_input("Enlace Público de la Aplicación (Ej: https://tu-colegio.streamlit.app):")
        
        # Muestra automáticamente los QR si se ha escrito un enlace
        if url_base:
            url_base = url_base.strip()
            if not url_base.endswith("/"):
                url_base += "/"
                
            # Creamos el archivo ZIP en memoria
            zip_buffer = io.BytesIO()
            
            st.markdown("---")
            cols = st.columns(4)
            
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for i, recurso in enumerate(RECURSOS):
                    qr_url = f"{url_base}?reportar={urllib.parse.quote(recurso)}"
                    qr = qrcode.QRCode(version=1, box_size=10, border=2)
                    qr.add_data(qr_url)
                    qr.make(fit=True)
                    img = qr.make_image(fill_color="black", back_color="white")
                    
                    buf = io.BytesIO()
                    img.save(buf, format="PNG")
                    
                    # Añadir al ZIP limpiando caracteres raros para el nombre del archivo
                    safe_name = str(recurso).replace("/", "-").replace("\\", "-").replace(":", "-")
                    zip_file.writestr(f"QR_{safe_name}.png", buf.getvalue())
                    
                    # Mostrar en pantalla
                    with cols[i % 4]:
                        with st.container(border=True):
                            st.image(buf.getvalue(), use_container_width=True)
                            st.markdown(f"<p style='text-align:center; font-weight:bold; font-size:14px; margin-top:-10px;'>{recurso}</p>", unsafe_allow_html=True)
            
            st.markdown("---")
            c1, c2, c3 = st.columns([1, 2, 1])
            with c2:
                # BOTÓN DE DESCARGA ZIP
                st.download_button(
                    label="📦 Descargar Todos los QRs (.zip)",
                    data=zip_buffer.getvalue(),
                    file_name="Codigos_QR_CAV.zip",
                    mime="application/zip",
                    use_container_width=True,
                    type="primary"
                )

# ------------------------------------------------------------------
# SECCIÓN: CONFIGURACIÓN
# ------------------------------------------------------------------
if page == "Configuración":
    st.title("⚙️ Configuración del Sistema")
    st.write("Desde aquí puedes administrar los elementos centrales de la aplicación.")
    
    tab_prof, tab_cur, tab_rec = st.tabs(["Profesores", "Cursos", "Recursos"])
    
    with tab_prof:
        st.write("### 👥 Administración de Profesores")
        col_add, col_list = st.columns([1, 2])
        with col_add:
            with st.form("form_add_prof"):
                nuevo_prof = st.text_input("Nombre y Apellidos")
                nuevo_email = st.text_input("Correo Electrónico (Opcional)")
                if st.form_submit_button("➕ Agregar Profesor", use_container_width=True):
                    if nuevo_prof.strip():
                        try:
                            supabase.table("profesores").insert({"nombre": nuevo_prof.strip().upper(), "email": nuevo_email.strip()}).execute()
                            st.success("¡Profesor agregado!")
                            st.cache_data.clear(); time.sleep(0.5); st.rerun()
                        except Exception as e: st.error(f"Error al agregar: {e}")
                    else: st.error("El nombre es obligatorio.")
        with col_list:
            prof_data = supabase.table("profesores").select("*").order("nombre").execute().data
            if prof_data:
                df_p = pd.DataFrame(prof_data)
                st.dataframe(df_p[['nombre', 'email']], use_container_width=True, hide_index=True)
                with st.expander("🗑️ Eliminar un Profesor"):
                    st.warning("⚠️ Nota: No puedes eliminar a un profesor si ya tiene reservas en el sistema.")
                    prof_borrar = st.selectbox("Selecciona el profesor a eliminar", df_p['nombre'].tolist(), key="del_prof")
                    if st.button("Eliminar Profesor Definitivamente", type="primary"):
                        try:
                            id_b = int(df_p[df_p['nombre'] == prof_borrar]['id'].values[0])
                            supabase.table("profesores").delete().eq("id", id_b).execute()
                            st.success(f"Profesor {prof_borrar} eliminado.")
                            st.cache_data.clear(); time.sleep(0.5); st.rerun()
                        except Exception as e: st.error("No se puede eliminar porque este profesor tiene reservas asociadas.")

    with tab_cur:
        st.write("### 📚 Administración de Cursos")
        col_add_c, col_list_c = st.columns([1, 2])
        with col_add_c:
            with st.form("form_add_cur"):
                nuevo_curso = st.text_input("Nombre del Curso (ej. 1° BÁSICO A)")
                if st.form_submit_button("➕ Agregar Curso", use_container_width=True):
                    if nuevo_curso.strip():
                        try:
                            supabase.table("cursos").insert({"nombre": nuevo_curso.strip().upper()}).execute()
                            st.success("¡Curso agregado!")
                            st.cache_data.clear(); time.sleep(0.5); st.rerun()
                        except Exception as e: st.error(f"Error al agregar: {e}")
                    else: st.error("El nombre es obligatorio.")
        with col_list_c:
            cur_data = supabase.table("cursos").select("*").order("nombre").execute().data
            if cur_data:
                df_c = pd.DataFrame(cur_data)
                st.dataframe(df_c[['nombre']], use_container_width=True, hide_index=True)
                with st.expander("🗑️ Eliminar un Curso"):
                    cur_borrar = st.selectbox("Selecciona el curso a eliminar", df_c['nombre'].tolist(), key="del_cur")
                    if st.button("Eliminar Curso Definitivamente", type="primary"):
                        try:
                            id_b = int(df_c[df_c['nombre'] == cur_borrar]['id'].values[0])
                            supabase.table("cursos").delete().eq("id", id_b).execute()
                            st.success(f"Curso {cur_borrar} eliminado.")
                            st.cache_data.clear(); time.sleep(0.5); st.rerun()
                        except Exception as e: st.error("No se puede eliminar porque este curso tiene reservas asociadas.")

    with tab_rec:
        st.write("### 💻 Administración de Recursos")
        col_add_r, col_list_r = st.columns([1, 2])
        with col_add_r:
            with st.form("form_add_rec"):
                nuevo_rec = st.text_input("Nombre del Recurso (ej. Proyector 5)")
                if st.form_submit_button("➕ Agregar Recurso", use_container_width=True):
                    if nuevo_rec.strip():
                        try:
                            supabase.table("recursos").insert({"nombre": nuevo_rec.strip().upper()}).execute()
                            st.success("¡Recurso agregado!")
                            st.cache_data.clear(); time.sleep(0.5); st.rerun()
                        except Exception as e: st.error(f"Error al agregar: {e}")
                    else: st.error("El nombre es obligatorio.")
        with col_list_r:
            rec_data = supabase.table("recursos").select("*").order("nombre").execute().data
            if rec_data:
                df_r = pd.DataFrame(rec_data)
                st.dataframe(df_r[['nombre']], use_container_width=True, hide_index=True)
                with st.expander("🗑️ Eliminar un Recurso"):
                    rec_borrar = st.selectbox("Selecciona el recurso a eliminar", df_r['nombre'].tolist(), key="del_rec")
                    if st.button("Eliminar Recurso Definitivamente", type="primary"):
                        try:
                            id_b = int(df_r[df_r['nombre'] == rec_borrar]['id'].values[0])
                            supabase.table("recursos").delete().eq("id", id_b).execute()
                            st.success(f"Recurso {rec_borrar} eliminado.")
                            st.cache_data.clear(); time.sleep(0.5); st.rerun()
                        except Exception as e: st.error("No se puede eliminar porque tiene reservas o reportes de mantenimiento asociados.")
