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
import google.generativeai as genai

# --- CONFIGURACIÓN DE GEMINI ---
# Usamos st.secrets para que no te vuelvan a bloquear la llave
try:
    GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=GEMINI_API_KEY)
    
    # AQUÍ ESTÁ LA MAGIA: Usamos el nuevo modelo que tu llave sí reconoce
    model = genai.GenerativeModel('gemini-2.5-flash') 
    
except KeyError:
    st.error("🚨 Error: No se encontró la API Key en los secretos de Streamlit.")
    st.stop()

# --- TU FUNCIÓN PARA LLAMAR A LA IA ---
def consultar_gemini(prompt):
    try:
        response = model.generate_content(prompt) 
        return response.text
    except Exception as e:
        return f"Error con la IA: {e}"
        st.header("💬 Asistente Técnico IA")

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

# ==============================================================================
# --- MODO PÚBLICO: ENRUTAMIENTO VÍA CÓDIGO QR ---
# ==============================================================================
# 1. Captura limpia de parámetros (SIN imprimir nada en pantalla)
query_params = st.query_params
page_param = query_params.get("page", "")

if page_param == "reporte":
    # 2. ESTILOS AESTHETIC (Ocultar menús y decorar)
    st.markdown("""
        <style>
            [data-testid="stHeader"] {display: none;}
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            .stApp { background-color: #F8FAFC; }
            [data-testid="stForm"] {
                background-color: #FFFFFF;
                border-radius: 16px;
                padding: 25px;
                box-shadow: 0 4px 15px rgba(0,0,0,0.05);
                border: 1px solid #E2E8F0;
            }
        </style>
    """, unsafe_allow_html=True)

    # 3. ENCABEZADO (He quitado la imagen que causaba el cuadro con el "0")
    st.markdown("<h1 style='text-align: center; font-size: 50px;'>🏫</h1>", unsafe_allow_html=True)
    st.markdown("<h2 style='text-align: center; color: #1E3A8A; font-weight: 800; margin-bottom: 0px;'>🚨 Reportar Falla</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #64748b; font-size: 14px; margin-top: 5px;'>Colegio Antonio Varas - Depto. Enlaces</p>", unsafe_allow_html=True)
    st.write("") 
    
    # 4. LÓGICA DEL REPORTE
    recurso_id = query_params.get("id", "")
    
    if recurso_id:
        try:
            recurso = supabase.table("recursos").select("nombre").eq("id", recurso_id).execute().data
            if recurso:
                st.success(f"🛠️ **Equipo Identificado:** {recurso[0]['nombre']}")
                
                with st.form("form_reporte_publico", clear_on_submit=True):
                    st.markdown("#### 📝 Detalles del Problema")
                    st.markdown("<span style='color: #64748b; font-size: 14px;'>Ayúdanos con estos datos para que Soporte Técnico lo resuelva rápido.</span>", unsafe_allow_html=True)
                    st.write("")
                    
                    nombre_reporta = st.text_input("👤 Tu Nombre Completo:", placeholder="Ej. Juan Pérez")
                    descripcion = st.text_area("🔧 Describe el problema detalladamente:", height=120, placeholder="Ej. El proyector no enciende y parpadea una luz roja...")
                    
                    st.write("")
                    submit = st.form_submit_button("🚀 Enviar Reporte Técnico", type="primary", use_container_width=True)
                    
                    if submit:
                        if not nombre_reporta.strip() or not descripcion.strip():
                            st.error("⚠️ Faltan datos: Por favor ingresa tu nombre y la descripción.")
                        else:
                            with st.spinner("Enviando reporte a Enlaces..."):
                                supabase.table("mantenimientos").insert({
    "recurso_id": recurso_id,
    "descripcion": descripcion.strip(),
    "estado": "Reportado (Vía QR)",
    "reportado_por": nombre_reporta.strip(),
    "fecha": str(dt.date.today())  # <-- ¡Esta línea soluciona el error!
}).execute()
                                
                                st.success("✅ ¡Reporte enviado con éxito! Gracias por avisarnos.")
                                st.balloons() # 🎈 ¡AQUÍ ESTÁN LOS GLOBOS! 🎈
                                time.sleep(4) 
            else:
                st.error("❌ El equipo que intentas reportar no existe o fue dado de baja.")
        except Exception as e:
            st.error(f"Error de conexión con la base de datos: {e}")
    else:
        st.error("❌ Enlace no válido. Falta el identificador del equipo.")
        
    st.stop() # Detiene la app aquí para que no cargue el Login de los profes
# ==============================================================================
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
# ==============================================================================
# --- MODO PÚBLICO: ENRUTAMIENTO VÍA CÓDIGO QR ---
# ==============================================================================
# Si la URL contiene "?page=reporte", mostramos la interfaz pública y bloqueamos el panel admin
if "page" in st.query_params and st.query_params["page"] == "reporte":
    st.image("https://images.vexels.com/content/135222/preview/university-building-simple-icon-135222.png", width=80)
    st.header("🚨 Reportar Falla de Equipo")
    
    recurso_id = st.query_params.get("id")
    
    if recurso_id:
        try:
            # Buscar el nombre del equipo en la base de datos
            recurso = supabase.table("recursos").select("nombre").eq("id", recurso_id).execute().data
            
            if recurso:
                st.info(f"Estás reportando una falla para el equipo: **{recurso[0]['nombre']}**")
                
                with st.form("form_reporte_publico", clear_on_submit=True):
                    st.write("Por favor, completa los siguientes datos para que Soporte Técnico pueda ayudarte.")
                    
                    # NUEVO CAMPO OBLIGATORIO
                    nombre_reporta = st.text_input("👤 Tu Nombre Completo (Obligatorio):", placeholder="Ej. Juan Pérez")
                    descripcion = st.text_area("📝 Describe detalladamente el problema (Obligatorio):", height=150, placeholder="Ej. El proyector no enciende y parpadea una luz roja...")
                    
                    submit = st.form_submit_button("🚀 Enviar Reporte al Equipo Técnico", type="primary", use_container_width=True)
                    
                    if submit:
                        if not nombre_reporta.strip() or not descripcion.strip():
                            st.warning("⚠️ Debes ingresar tu nombre y la descripción del problema para enviar el reporte.")
                        else:
                            with st.spinner("Enviando reporte..."):
                                # Guardar en la tabla mantenimientos
                                supabase.table("mantenimientos").insert({
                                    "recurso_id": recurso_id,
                                    "descripcion": descripcion.strip(),
                                    "estado": "Reportado (Vía QR)",
                                    "reportado_por": nombre_reporta.strip() # Guardamos el nombre aquí
                                }).execute()
                                
                                st.success("✅ ¡Reporte enviado con éxito! El Departamento de Enlaces ha sido notificado.")
                                st.balloons() # ¡GLOBITOS ACTIVADOS! 🎈
            else:
                st.error("❌ El equipo que intentas reportar no existe o fue dado de baja.")
        except Exception as e:
            st.error(f"Error de conexión con la base de datos: {e}")
    else:
        st.error("❌ Enlace no válido. Falta el identificador del equipo.")
        
    # DETENEMOS LA APP AQUÍ: Esto evita que el usuario público vea el menú lateral de administrador
    st.stop()
# ==============================================================================
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
    "Modo TV": {"icon": "📺", "roles": ["admin"]},  # <-- AQUÍ AGREGAMOS MODO TV
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
# ==============================================================================
# 🤖 WIDGET DE CHATBOT (VENTANA FLOTANTE / MODAL)
# ==============================================================================
st.sidebar.markdown("---")

# 1. Definimos la estructura de la ventana flotante (con ancho "large")
@st.dialog("🤖 Asistente Técnico IA", width="large")
def abrir_ventana_chat():
    st.info("💡 Pregúntame sobre reparaciones, fallas, configuraciones o cómo usar el sistema. Cierra esta ventana con la 'X' arriba a la derecha.")
    
    # Inicializar la memoria de la IA con CONTEXTO TOTAL DEL SISTEMA Y PERSONALIDAD
    if "chat_session" not in st.session_state:
        
        instrucciones_sistema = """
        Eres el Asistente de IA exclusivo del Departamento de Informática/Enlaces del 'Liceo Bicentenario Colegio Antonio Varas'.
        Tu misión es ayudar a los técnicos y administradores con el soporte diario y guiarlos en el uso de esta plataforma.
        
        Conoces a la perfección todas las secciones de nuestro sistema:
        1. Mis Reservas / Registrar: Módulo donde los profesores agendan los recursos del colegio (laboratorios, proyectores, carritos de tablets, etc.).
        2. Semana (Calendario): La vista semanal para organizar los horarios y ver qué recursos están ocupados y por quién.
        3. Dashboard: Panel de estadísticas donde vemos gráficos de uso, métricas de fallas, equipos con más problemas y profesores que más reservan.
        4. Base de Datos: Nuestro motor es Supabase. Ahí se guarda todo nuestro inventario, los usuarios, el historial de reservas y los reportes de mantenimiento.
        5. Panel Técnico: El centro de mando. Aquí gestionamos los Tickets (fallas reportadas por QR que pasan a Pendiente, Revisión o Resuelto), creamos Códigos QR para pegar en los equipos nuevos, y procesamos las Bajas de Equipos obsoletos generando informes automáticos en Word.
        6. Configuración: Módulo de administración para agregar/eliminar recursos del inventario o gestionar usuarios.
        
        Tu tono y personalidad:
        Háblame como un colega informático más del equipo. Sé directo, empático, resolutivo y nada robótico. Puedes usar términos informáticos y modismos chilenos de forma natural (ej: 'apañar', 'la pega', 'dar jugo', 'sacar el cacho'). 
        Si te pregunto por un problema técnico (ej. un notebook no conecta o una impresora mancha) o cómo usar una sección del sistema, dame soluciones prácticas, paso a paso, y pensadas para la realidad de nuestro colegio.
        """
        
        st.session_state.chat_session = model.start_chat(
            history=[
                {"role": "user", "parts": [instrucciones_sistema]},
                {"role": "model", "parts": ["¡Todo anotado! Ya tengo el mapa completo del sistema del colegio en mi cabeza (Reservas, Dashboard, Supabase, Tickets...). Listo para apañar con la pega técnica. ¿Qué revisamos hoy?"]}
            ]
        )
        
    # Contenedor para que el chat tenga barra de desplazamiento
    contenedor_chat = st.container(height=400) 
    
    with contenedor_chat:
        # Mostrar el historial (Filtramos las instrucciones para que no se vean en pantalla)
        for mensaje in st.session_state.chat_session.history:
            if "Eres el Asistente de IA exclusivo" not in mensaje.parts[0].text and "¡Todo anotado!" not in mensaje.parts[0].text:
                rol = "assistant" if mensaje.role == "model" else "user"
                with st.chat_message(rol):
                    st.markdown(mensaje.parts[0].text)
                    
    # Caja de texto para el usuario
    if pregunta := st.chat_input("Escribe tu duda técnica o del sistema aquí..."):
        with contenedor_chat:
            with st.chat_message("user"):
                st.markdown(pregunta)
            
            with st.chat_message("assistant"):
                with st.spinner("Pensando..."):
                    try:
                        respuesta = st.session_state.chat_session.send_message(pregunta)
                        st.markdown(respuesta.text)
                    except Exception as e:
                        st.error(f"Error de red: {e}")

# 2. El botón en la barra lateral que activa la ventana flotante
if st.sidebar.button("💬 Abrir Asistente IA", type="primary", use_container_width=True):
    abrir_ventana_chat()


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
            edited_df = st.data_editor(
                df_display, 
                hide_index=True, 
                use_container_width=True, 
                num_rows="dynamic", 
                column_config={
                    "Fecha": st.column_config.DateColumn("Fecha", format="DD/MM/YYYY"), 
                    "Hora inicio": st.column_config.TimeColumn("Hora Inicio", format="HH:mm"), 
                    "Hora fin": st.column_config.TimeColumn("Hora Fin", format="HH:mm"), 
                    "Profesor": st.column_config.SelectboxColumn("Profesor", options=PROFESORES, required=True), 
                    "Curso": st.column_config.SelectboxColumn("Curso", options=CURSOS, required=True), 
                    "Recurso": st.column_config.SelectboxColumn("Recurso", options=RECURSOS, required=True)
                }
            )
            
            if st.button("💾 Guardar Cambios en la Nube", use_container_width=True, type="primary"):
                with st.spinner("Sincronizando con Supabase..."):
                    try:
                        original_indices = set(df.index)
                        edited_indices = set(edited_df.index)
                        
                        # ---------------------------------------------------------
                        # 1. ELIMINACIONES
                        # ---------------------------------------------------------
                        deleted_indices = original_indices - edited_indices
                        for idx in deleted_indices:
                            id_borrar = int(df.loc[idx, 'id'])
                            supabase.table("reservas").delete().eq("id", id_borrar).execute()
                            
                            prof_name = df.loc[idx, 'Profesor']
                            email_to = PROFESOR_DATA.get(prof_name)
                            if email_to:
                                subject = f"Cancelación de Reserva de Recursos - {df.loc[idx, 'Curso']}"
                                body = f"""<html><body><p>Hola {prof_name.split(' ')[0]},</p><p>Te informamos que la siguiente reserva ha sido <b>cancelada</b>:</p><ul><li><b>Fecha:</b> {format_date_es(df.loc[idx, 'Fecha'])}</li><li><b>Horario:</b> {df.loc[idx, 'Hora inicio'].strftime('%H:%M')} - {df.loc[idx, 'Hora fin'].strftime('%H:%M')}</li><li><b>Curso:</b> {df.loc[idx, 'Curso']}</li><li><b>Recurso:</b> {df.loc[idx, 'Recurso']}</li></ul><p>Saludos,<br>Sistema de Horarios CAV</p></body></html>"""
                                send_email(subject, body, email_to)

                        # ---------------------------------------------------------
                        # 2. MODIFICACIONES (¡Lo que nos faltaba!)
                        # ---------------------------------------------------------
                        common_indices = original_indices.intersection(edited_indices)
                        for idx in common_indices:
                            r_orig = df_display.loc[idx]
                            r_edit = edited_df.loc[idx]
                            
                            # Comparamos si cambió alguna celda (Fecha, hora, profe, etc.)
                            if str(r_orig["Fecha"]) != str(r_edit["Fecha"]) or \
                               str(r_orig["Hora inicio"]) != str(r_edit["Hora inicio"]) or \
                               str(r_orig["Hora fin"]) != str(r_edit["Hora fin"]) or \
                               str(r_orig["Profesor"]) != str(r_edit["Profesor"]) or \
                               str(r_orig["Curso"]) != str(r_edit["Curso"]) or \
                               str(r_orig["Recurso"]) != str(r_edit["Recurso"]) or \
                               str(r_orig.get("Observaciones", "")) != str(r_edit.get("Observaciones", "")):
                               
                                # Si algo cambió, actualizamos esa fila específica en Supabase
                                id_actualizar = int(df.loc[idx, 'id'])
                                datos_actualizados = {
                                    "fecha": str(r_edit["Fecha"]),
                                    "hora_inicio": str(r_edit["Hora inicio"]),
                                    "hora_fin": str(r_edit["Hora fin"]),
                                    "profesor": map_prof.get(r_edit["Profesor"]),
                                    "curso": map_cur.get(r_edit["Curso"]),
                                    "recurso": map_rec.get(r_edit["Recurso"]),
                                    "observaciones": r_edit.get("Observaciones", "")
                                }
                                supabase.table("reservas").update(datos_actualizados).eq("id", id_actualizar).execute()

                        # ---------------------------------------------------------
                        # 3. NUEVAS INSERCIONES
                        # ---------------------------------------------------------
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
                            
                        # Fin del proceso: recargamos la interfaz
                        st.success("Sincronización completa.")
                        st.cache_data.clear()
                        time.sleep(0.5)
                        st.rerun()
                        
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
            default_date_week = dt.date.today() # ✅ CORREGIDO: Siempre usa la fecha de hoy
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
        
# --- PANEL PRINCIPAL (DASHBOARD) ---
elif page == "Dashboard":
    import tempfile
    import unicodedata
    import os
    import matplotlib.pyplot as plt
    import plotly.express as px
    
    st.header("📊 Panel de Supervisión General")

    # Definir colores institucionales
    COLOR_PRIMARIO = '#1E3A8A' 
    COLOR_SECUNDARIO = '#10B981' 

    # --- 1. FUNCIÓN MAESTRA: GENERAR PDF PROFESIONAL CON GRÁFICOS ---
    def generar_pdf_profesional_con_graficos(df_datos):
        from fpdf import FPDF
        
        # Plantilla del Documento Estilizada
        class PDF(FPDF):
            def header(self):
                self.set_fill_color(30, 58, 138)
                self.rect(0, 0, 210, 15, 'F')
                self.set_y(15)
                self.set_font('Arial', 'B', 20)
                self.set_text_color(255, 255, 255)
                self.cell(0, 15, 'REPORTE EJECUTIVO DE GESTION', 0, 1, 'C')
                self.set_text_color(31, 41, 55)
                self.ln(10)

            def footer(self):
                self.set_y(-20)
                self.set_font('Arial', 'I', 9)
                self.set_text_color(100, 100, 100)
                hoy_str = dt.date.today().strftime('%d/%m/%Y')
                self.cell(0, 10, f'Generado el: {hoy_str} | Sistema de Reservas Enlaces', 0, 0, 'L')
                self.cell(0, 10, f'Pagina {self.page_no()}', 0, 0, 'R')

            def section_title(self, label):
                self.set_font('Arial', 'B', 14)
                self.set_fill_color(243, 244, 246)
                self.set_text_color(30, 58, 138)
                self.cell(0, 10, f"  {label}", 0, 1, 'L', 1)
                self.set_text_color(31, 41, 55)
                self.ln(4)

        # Limpiar acentos para evitar errores en PDF
        def s(texto):
            if not texto: return "N/A"
            texto_str = str(texto)
            nfkd = unicodedata.normalize('NFKD', texto_str)
            return u"".join([c for c in nfkd if not unicodedata.combining(c)])

        pdf = PDF()
        pdf.set_auto_page_break(auto=True, margin=25)
        pdf.add_page()
        pdf.set_font('Arial', '', 12)

        if df_datos.empty:
            pdf.cell(0, 10, s('No hay datos registrados para analizar.'), 0, 1)
        else:
            C_FECHA = 'Fecha'
            C_REC = 'Recurso'
            C_PROF = 'Profesor'

            # Cálculos rápidos
            total_res = len(df_datos)
            df_datos['fecha_obj'] = pd.to_datetime(df_datos[C_FECHA]).dt.date
            hoy_res = len(df_datos[df_datos['fecha_obj'] == dt.date.today()])
            recurso_top = df_datos[C_REC].mode()[0] if not df_datos[C_REC].empty else "N/A"
            profesor_top = df_datos[C_PROF].mode()[0] if not df_datos[C_PROF].empty else "N/A"

            # Calcular rango de fechas para el subtítulo
            fecha_min = df_datos['fecha_obj'].min().strftime('%d/%m/%Y')
            fecha_max = df_datos['fecha_obj'].max().strftime('%d/%m/%Y')

            # --- TÍTULO Y RANGO DE FECHAS EN EL DOCUMENTO ---
            pdf.set_font('Arial', 'B', 16)
            pdf.set_text_color(30, 58, 138) # Azul institucional
            pdf.cell(0, 10, s('INFORME ESTADISTICO DE USO DE RECURSOS'), 0, 1, 'C')
            
            pdf.set_font('Arial', 'I', 11)
            pdf.set_text_color(100, 100, 100) # Gris oscuro
            pdf.cell(0, 6, s(f'Periodo analizado: {fecha_min} al {fecha_max}'), 0, 1, 'C')
            pdf.ln(8)
            pdf.set_text_color(31, 41, 55) # Resetear a color texto normal

            # =========================================================
            # SECCIÓN 1: RESUMEN EJECUTIVO
            # =========================================================
            pdf.section_title('1. Resumen de Metricas Clave')
            
            pdf.set_fill_color(255, 255, 255)
            pdf.set_font('Arial', 'B', 12)
            pdf.cell(95, 10, s('Concepto'), 1, 0, 'C', 1)
            pdf.cell(95, 10, s('Valor'), 1, 1, 'C', 1)
            
            pdf.set_font('Arial', '', 11)
            pdf.cell(95, 9, s('Total de Reservas Historicas'), 1, 0, 'L')
            pdf.cell(95, 9, f"{total_res}", 1, 1, 'R')
            pdf.cell(95, 9, s('Reservas Agendadas para Hoy'), 1, 0, 'L')
            pdf.cell(95, 9, f"{hoy_res}", 1, 1, 'R')
            
            pdf.set_font('Arial', 'B', 11)
            pdf.cell(95, 9, s('Recurso Mas Ocupado (Estrella)'), 1, 0, 'L')
            pdf.cell(95, 9, s(recurso_top), 1, 1, 'R')
            pdf.cell(95, 9, s('Profesor con Mas Solicitudes'), 1, 0, 'L')
            pdf.cell(95, 9, s(profesor_top), 1, 1, 'R')
            
            pdf.ln(12)

            # =========================================================
            # SECCIÓN 2: GRÁFICOS CON ARCHIVOS TEMPORALES
            # =========================================================
            pdf.section_title('2. Analisis Visual de Ocupacion')
            pdf.set_font('Arial', '', 11)
            pdf.multi_cell(0, 7, s('A continuacion se presentan los graficos detallados de uso de recursos y distribucion temporal de las reservas.'), 0, 'L')
            pdf.ln(5)

            # --- GRÁFICO A: TOP 5 RECURSOS ---
            pdf.set_font('Arial', 'I', 10)
            pdf.cell(0, 8, s('Grafico A: Top 5 Recursos Mas Utilizados'), 0, 1, 'C')
            
            plt.figure(figsize=(8, 4))
            top5_data = df_datos[C_REC].value_counts().nlargest(5).sort_values(ascending=True)
            names_clean = [s(name) for name in top5_data.index]
            
            bars = plt.barh(names_clean, top5_data.values, color='#3B82F6', edgecolor='#1E3A8A', height=0.7)
            plt.xlabel('Cantidad de Reservas', fontsize=10, fontweight='bold', color='#4B5563')
            plt.title('Top 5 Recursos', fontsize=12, fontweight='bold', color='#1E3A8A')
            plt.gca().spines['top'].set_visible(False)
            plt.gca().spines['right'].set_visible(False)
            plt.grid(axis='x', linestyle='--', alpha=0.5)
            
            for bar in bars:
                width = bar.get_width()
                plt.text(width + (max(top5_data.values)*0.01), bar.get_y() + bar.get_height()/2, 
                         f'{int(width)}', va='center', fontsize=9, fontweight='bold', color='#1E3A8A')
            plt.tight_layout()
            
            # Guardar en archivo temporal físico
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_a:
                plt.savefig(tmp_a.name, format='png', dpi=150, bbox_inches='tight')
                img_path_a = tmp_a.name
            plt.close()
            
            # Insertar en PDF y eliminar archivo
            current_y = pdf.get_y()
            pdf.image(img_path_a, x=25, y=current_y, w=160)
            os.remove(img_path_a)
            
            # --- CORRECCIÓN VISUAL: FORZAR PÁGINA 2 PARA EL GRÁFICO DE PASTEL ---
            pdf.add_page() 
            
            # --- GRÁFICO B: DISTRIBUCIÓN DE USO ---
            pdf.set_font('Arial', 'I', 10)
            pdf.cell(0, 8, s('Grafico B: Distribucion de Uso por Tipo de Recurso (Total)'), 0, 1, 'C')
            
            plt.figure(figsize=(6, 6))
            uso_total = df_datos[C_REC].value_counts()
            colores_pie = ['#1E3A8A', '#10B981', '#F59E0B', '#EF4444', '#8B5CF6', '#EC4899', '#6B7280']
            labels_clean = [s(name) for name in uso_total.index]
            
            patches, texts, autotexts = plt.pie(uso_total.values, labels=labels_clean, autopct='%1.1f%%', 
                                              startangle=140, colors=colores_pie, shadow=False,
                                              wedgeprops={'edgecolor': 'white', 'linewidth': 2})
            
            plt.title('Uso General de Recursos', fontsize=12, fontweight='bold', color='#1E3A8A')
            for text in texts: text.set_color('#4B5563'); text.set_fontsize(9)
            for autotext in autotexts: autotext.set_color('white'); autotext.set_fontweight('bold'); autotext.set_fontsize(9)
            
            plt.axis('equal')
            plt.tight_layout()
            
            # Guardar en archivo temporal físico
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_b:
                plt.savefig(tmp_b.name, format='png', dpi=150, bbox_inches='tight')
                img_path_b = tmp_b.name
            plt.close()
            
            # Insertar en PDF (ahora en la parte superior de la página 2)
            current_y_b = pdf.get_y()
            pdf.image(img_path_b, x=55, y=current_y_b, w=100)
            os.remove(img_path_b)

        return pdf.output(dest='S').encode('latin-1')

    # --- 2. BOTÓN DE DESCARGA PDF ---
    col_titulo, col_boton = st.columns([2, 1])
    with col_boton:
        if not df.empty:
            with st.spinner("Preparando reporte PDF..."):
                try:
                    pdf_data = generar_pdf_profesional_con_graficos(df)
                    st.download_button(
                        label="📄 Descargar Reporte PDF",
                        data=pdf_data,
                        file_name=f"Reporte_Gestión_Enlaces_{dt.date.today().strftime('%d_%m_%Y')}.pdf",
                        mime="application/pdf",
                        type="primary",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"⚠️ Error generando PDF: {e}")

    # --- 3. MÉTRICAS Y GRÁFICOS EN PANTALLA ---
    if not df.empty:
        C_FECHA = 'Fecha'
        C_REC = 'Recurso'
        C_PROF = 'Profesor'
        C_HORA_I = 'Hora inicio'
        C_HORA_F = 'Hora fin'
        
        df['fecha_obj'] = pd.to_datetime(df[C_FECHA]).dt.date
        hoy = dt.date.today()
        reservas_hoy = df[df['fecha_obj'] == hoy]
        
        c1, c2, c3, c4 = st.columns(4)
        style_kpi = "background:white; border-radius:12px; padding: 25px; box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1); border-top: 5px solid;"

        with c1:
            st.markdown(f"""
                <div style="{style_kpi} border-top-color: #3B82F6;">
                    <div style="font-size:1em; color:#64748b; font-weight:bold;">TOTAL RESERVAS</div>
                    <div style="font-size:3em; font-weight:900; color:#1E3A8A;">{len(df)}</div>
                </div>
            """, unsafe_allow_html=True)
        with c2:
            st.markdown(f"""
                <div style="{style_kpi} border-top-color: #EF4444;">
                    <div style="font-size:1em; color:#64748b; font-weight:bold;">RESERVAS HOY</div>
                    <div style="font-size:3em; font-weight:900; color:#B91C1C;">{len(reservas_hoy)}</div>
                </div>
            """, unsafe_allow_html=True)
        with c3:
            recurso_top = df[C_REC].mode()[0] if not df[C_REC].empty else "N/A"
            st.markdown(f"""
                <div style="{style_kpi} border-top-color: #F59E0B;">
                    <div style="font-size:1em; color:#64748b; font-weight:bold;">RECURSO ESTRELLA</div>
                    <div style="font-size:1.8em; font-weight:800; color:#B45309;">{recurso_top}</div>
                </div>
            """, unsafe_allow_html=True)
        with c4:
            profesor_top = df[C_PROF].mode()[0] if not df[C_PROF].empty else "N/A"
            prof_display = str(profesor_top)[:16] + ".." if len(str(profesor_top)) > 18 else str(profesor_top)
            st.markdown(f"""
                <div style="{style_kpi} border-top-color: #8B5CF6;">
                    <div style="font-size:1em; color:#64748b; font-weight:bold;">PROF. FRECUENTE</div>
                    <div style="font-size:1.8em; font-weight:800; color:#5B21B6;">{prof_display}</div>
                </div>
            """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        col_graf_1, col_graf_2 = st.columns([2, 1])
        
        with col_graf_1:
            st.markdown("### 🗓️ Mapa de Calor: Ocupación Semanal")
            df['fecha_dt'] = pd.to_datetime(df[C_FECHA])
            dias_map = {0: '1-Lunes', 1: '2-Martes', 2: '3-Miércoles', 3: '4-Jueves', 4: '5-Viernes', 5: '6-Sábado', 6: '7-Domingo'}
            df['Dia_Semana'] = df['fecha_dt'].dt.dayofweek.map(dias_map)
            df['Bloque_Ordenado'] = df[C_HORA_I].astype(str).str[:5] + " - " + df[C_HORA_F].astype(str).str[:5]
            
            df_heatmap_filter = df[df['fecha_dt'].dt.dayofweek < 5]
            heatmap_data = df_heatmap_filter.groupby(['Bloque_Ordenado', 'Dia_Semana']).size().reset_index(name='Cantidad')
            
            if not heatmap_data.empty:
                fig_heat = px.density_heatmap(heatmap_data, x='Dia_Semana', y='Bloque_Ordenado', z='Cantidad',
                                             text_auto=True, color_continuous_scale='Viridis')
                fig_heat.update_layout(xaxis_nticks=5, yaxis={'categoryorder':'category ascending'}, height=450)
                fig_heat.update_xaxes(tickformat="%A", labelalias={d: d.split('-')[1] for d in dias_map.values()})
                st.plotly_chart(fig_heat, use_container_width=True, config={'displayModeBar': False})
            else:
                st.info("Datos insuficientes para el mapa de calor.")

        with col_graf_2:
            st.markdown("### 🏆 Top 5 Recursos")
            uso_recursos = df[C_REC].value_counts().reset_index()
            uso_recursos.columns = ['Recurso', 'Cantidad']
            top5_recursos = uso_recursos.nlargest(5, 'Cantidad').sort_values(by='Cantidad', ascending=True)
            
            if not top5_recursos.empty:
                fig_top5 = px.bar(top5_recursos, x='Cantidad', y='Recurso', orientation='h', text='Cantidad',
                                 color='Cantidad', color_continuous_scale='Blues')
                fig_top5.update_layout(showlegend=False, coloraxis_showscale=False, height=450)
                fig_top5.update_traces(textposition='outside')
                st.plotly_chart(fig_top5, use_container_width=True, config={'displayModeBar': False})

        st.markdown("<br>### 📊 Distribución Global de Recursos", unsafe_allow_html=True)
        fig_full = px.bar(uso_recursos, x='Recurso', y='Cantidad', color='Recurso', text_auto=True)
        fig_full.update_layout(showlegend=False, height=350, xaxis_tickangle=-45)
        st.plotly_chart(fig_full, use_container_width=True)
        
    else:
        st.info("No hay reservas registradas en el sistema todavía.")
# ==============================================================================
# --- SECCIÓN TÉCNICOS (TICKETS, BAJAS INDEPENDIENTES Y CÓDIGOS QR) ---
# ==============================================================================
elif page == "Técnicos":
    import io
    from io import BytesIO
    import qrcode
    try:
        import docx
        from docx import Document
        from docx.shared import Pt, Inches, RGBColor
        from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
    except ImportError:
        st.error("🚨 Falta la librería 'python-docx'. Instálala con: pip install python-docx")
        st.stop()

    st.header("🛠️ Panel de Soporte Técnico")

   
    
    modulo_tec = st.radio("Selecciona el módulo de trabajo:", 
                          ["🎫 Tickets", "🗑️ Baja de Equipos", "📋 Generador QR"], 
                          horizontal=True)
    st.markdown("---")

    # ---------------------------------------------------------
    # MÓDULO 1: TICKETS (GESTIÓN DE FALLAS)
    # ---------------------------------------------------------
    if modulo_tec == "🎫 Tickets":
        st.subheader("Gestión de Tickets ingresados vía QR")
        try:
            mant_data = supabase.table("mantenimientos").select("*, recursos(nombre)").order("fecha", desc=True).execute().data
        except Exception as e:
            st.error(f"Error consultando mantenimientos: {e}")
            mant_data = []

        if mant_data:
            df_mant = pd.DataFrame(mant_data)
            df_mant['Recurso'] = df_mant['recursos'].apply(lambda x: x.get('nombre', 'Desconocido') if isinstance(x, dict) else "Desconocido")
            
            if 'notas_tecnico' not in df_mant.columns: 
                df_mant['notas_tecnico'] = ""
            else: 
                df_mant['notas_tecnico'] = df_mant['notas_tecnico'].fillna("")
            
            # Métricas
            pendientes = len(df_mant[df_mant['estado'] == 'Reportado (Vía QR)'])
            en_revision = len(df_mant[df_mant['estado'] == 'En Revisión'])
            resueltos = len(df_mant[df_mant['estado'] == 'Resuelto'])
            
            c1, c2, c3 = st.columns(3)
            estilo_metrica = "background:white; border-radius:12px; padding: 15px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); border-left: 5px solid"
            with c1: st.markdown(f'<div style="{estilo_metrica} #EF4444;"><div style="color:#64748b; font-size:0.85em; font-weight:bold;">🔴 PENDIENTES</div><div style="font-size:2.2em; font-weight:900; color:#B91C1C;">{pendientes}</div></div>', unsafe_allow_html=True)
            with c2: st.markdown(f'<div style="{estilo_metrica} #F59E0B;"><div style="color:#64748b; font-size:0.85em; font-weight:bold;">🟡 EN REVISIÓN</div><div style="font-size:2.2em; font-weight:900; color:#D97706;">{en_revision}</div></div>', unsafe_allow_html=True)
            with c3: st.markdown(f'<div style="{estilo_metrica} #10B981;"><div style="color:#64748b; font-size:0.85em; font-weight:bold;">🟢 RESUELTOS</div><div style="font-size:2.2em; font-weight:900; color:#047857;">{resueltos}</div></div>', unsafe_allow_html=True)
            
            st.markdown("<br>", unsafe_allow_html=True)
            
            t_pendientes, t_revision, t_resueltos = st.tabs(["🔴 Pendientes", "🟡 En Revisión", "🟢 Resueltos"])
            
            def renderizar_tickets(df_filtrado, color_icon, estados_destino):
                if df_filtrado.empty:
                    st.info("✨ No hay tickets en esta categoría actualmente.")
                    return
                for _, row in df_filtrado.iterrows():
                    with st.expander(f"{color_icon} Ticket #{row['id']} | {row['Recurso']} | Fecha: {row['fecha']}"):
                        quien_reporto = row.get('reportado_por', 'No registrado')
                        st.markdown(f"**👤 Reportado por:** {quien_reporto}")
                        st.markdown(f"**📝 Falla:**\n> {row['descripcion']}")
                        
                        # --- INTEGRACIÓN GEMINI EN TICKETS ---
                        if st.button("🤖 Obtener Diagnóstico Sugerido (IA)", key=f"ai_btn_{row['id']}"):
                            with st.spinner("Gemini está analizando la falla..."):
                                prompt_tecnico = f"Actúa como un experto en soporte técnico informático. El equipo '{row['Recurso']}' tiene este problema: '{row['descripcion']}'. Dame una solución técnica breve en 3 puntos."
                                sugerencia = consultar_gemini(prompt_tecnico)
                                st.info(f"**💡 Sugerencia de Gemini:**\n\n{sugerencia}")
                        
                        if row['notas_tecnico']: 
                            st.markdown(f"**🛠️ Notas Previas:**\n> {row['notas_tecnico']}")
                        
                        st.markdown("---")
                        col_a, col_b = st.columns([1, 2])
                        with col_a: 
                            nuevo_est = st.selectbox("Cambiar estado a:", estados_destino, key=f"est_{row['id']}")
                        with col_b: 
                            nueva_nota = st.text_area("Notas de Reparación:", value=row['notas_tecnico'], key=f"not_{row['id']}", height=68)

                        if st.button("💾 Guardar Cambios", key=f"btn_{row['id']}", type="primary", use_container_width=True):
                            try:
                                supabase.table("mantenimientos").update({
                                    "estado": nuevo_est, 
                                    "notas_tecnico": html_sanitizer.escape(nueva_nota).strip()
                                }).eq("id", row['id']).execute()
                                st.success("Ticket actualizado.")
                                time.sleep(1); st.rerun()
                            except Exception as e: 
                                st.error(f"Error técnico: {e}")

            with t_pendientes: renderizar_tickets(df_mant[df_mant['estado'] == 'Reportado (Vía QR)'], "🔴", ["En Revisión", "Resuelto", "Reportado (Vía QR)"])
            with t_revision: renderizar_tickets(df_mant[df_mant['estado'] == 'En Revisión'], "🟡", ["Resuelto", "En Revisión", "Reportado (Vía QR)"])
            with t_resueltos: renderizar_tickets(df_mant[df_mant['estado'] == 'Resuelto'], "🟢", ["Resuelto", "En Revisión", "Reportado (Vía QR)"])
        else: 
            st.info("No hay reportes de mantenimiento.")

    # ---------------------------------------------------------
    # MÓDULO 2: BAJA DE EQUIPOS
    # ---------------------------------------------------------
    elif modulo_tec == "🗑️ Baja de Equipos":
        st.subheader("🗑️ Procesar Baja y Generar Informe")
        
        # Función para generar Word (Se mantiene tu lógica estética)
        def generar_docx_baja(datos, foto_data=None):
            try:
                document = Document()
                # ... (Tu configuración de estilos burdeo y fuentes igual que antes) ...
                titulo = document.add_paragraph()
                titulo.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                r_tit = titulo.add_run('INFORME TÉCNICO DE BAJA DE EQUIPO')
                r_tit.bold = True
                r_tit.font.size = Pt(16)
                r_tit.font.color.rgb = RGBColor(128, 0, 32)

                document.add_heading('1. IDENTIFICACIÓN DEL EQUIPO', level=1)
                table = document.add_table(rows=6, cols=2)
                table.style = 'Table Grid'
                # ... Rellenar tabla ...
                document.add_heading('2. ESTADO TÉCNICO', level=1)
                document.add_paragraph(datos['diagnosis'])
                document.add_heading('3. JUSTIFICACIÓN', level=1)
                document.add_paragraph(datos['justificacion'])

                docx_buf = BytesIO()
                document.save(docx_buf)
                docx_buf.seek(0)
                return docx_buf.read()
            except Exception as e:
                st.error(f"Error generando Word: {e}")
                return None

        tab_baja, tab_historial = st.tabs(["🆕 Procesar Nueva Baja", "📋 Ver Historial"])
        
        with tab_baja:
            res_data_raw = supabase.table("recursos").select("*").execute().data
            if res_data_raw:
                nombres_recursos = sorted(list(set([r.get('nombre', 'Sin nombre') for r in res_data_raw])))
                
                with st.form("form_baja_indep", clear_on_submit=False): 
                    col_sel, col_datos_m = st.columns([1, 2])
                    with col_sel:
                        st.markdown("**Identificación**")
                        recurso_cat_nom = st.selectbox("Categoría Maestro:", nombres_recursos)
                        marca_mod = st.text_input("Marca / Modelo específico")
                        cantidad_baja = st.number_input("Cantidad:", min_value=1, value=1)
                        num_serie = st.text_input("N° de Serie / Inventario")
                        ubicacion = st.text_input("Ubicación habitual")
                        fecha_adq = st.text_input("Fecha Adquisición")
                        uploaded_file = st.file_uploader("Subir foto (Opcional):", type=['png', 'jpg', 'jpeg'])
                    
                    with col_datos_m:
                        st.markdown("**Diagnosis y Justificación**")
                        diagnosis = st.text_area("Diagnosis Técnico / Daño detectado", height=120)
                        
                        # --- INTEGRACIÓN GEMINI EN BAJAS ---
                        # Nota: Dentro de un st.form, el botón debe ser submit o usarse fuera.
                        # Para que funcione mejor, usaremos un botón especial.
                        btn_ia = st.form_submit_button("✨ IA: Mejorar Redacción Técnica")
                        if btn_ia:
                            if diagnosis:
                                with st.spinner("Redactando profesionalmente..."):
                                    prompt_baja = f"Convierte esta nota técnica en un párrafo formal para un informe de baja: '{diagnosis}'"
                                    texto_mejorado = consultar_gemini(prompt_baja)
                                    st.info(f"**Sugerencia:**\n{texto_mejorado}")
                            else:
                                st.warning("Escribe algo en diagnosis primero.")

                        justificacion = st.text_area("Justificación de la Baja", height=100)
                        recomendacion = st.text_area("Recomendación Técnica", height=80)
                        tecnico = st.text_input("Técnico Responsable")

                    submit_baja = st.form_submit_button("🚫 Registrar Baja e Historial", type="primary", use_container_width=True)

                    if submit_baja and not btn_ia:
                        if not diagnosis or not tecnico:
                            st.warning("⚠️ Rellena los campos obligatorios.")
                        else:
                            # Lógica de guardado en Supabase (tu código original)
                            datos_bd = {
                                "recurso_nombre": recurso_cat_nom,
                                "marca": marca_mod,
                                "modelo": marca_mod,
                                "cantidad_baja": cantidad_baja,
                                "serie": num_serie,
                                "diagnosis": diagnosis,
                                "justificacion": justificacion,
                                "recomendacion": recomendacion,
                                "tecnico_responsable": tecnico,
                                "fecha_baja": str(dt.date.today())
                            }
                            supabase.table("equipos").insert(datos_bd).execute()
                            st.success("✅ Baja registrada.")
                            st.balloons()

        with tab_historial:
            st.subheader("📋 Historial de Bajas")
            # ... (Aquí va tu código de historial que ya funciona) ...

    # ---------------------------------------------------------
    # MÓDULO 3: GENERADOR QR
    # ---------------------------------------------------------
    elif modulo_tec == "📋 Generador QR":
        st.subheader("Generador de Códigos QR")
        res_data = supabase.table("recursos").select("*").execute().data
        if res_data:
            res_nombres = {r['nombre']: r['id'] for r in res_data}
            sel_res = st.selectbox("Selecciona el equipo para generar QR:", sorted(list(res_nombres.keys())))
            
            # URL de tu app (Asegúrate que sea la correcta)
            base_url = "https://tu-app.streamlit.app/" 
            final_url = f"{base_url}?page=reporte&id={res_nombres[sel_res]}"
            
            qr = qrcode.make(final_url)
            buf = BytesIO()
            qr.save(buf, format="PNG")
            
            st.image(buf.getvalue(), width=250, caption=f"Código QR para {sel_res}")
            st.download_button("⬇️ Descargar Código QR", data=buf.getvalue(), file_name=f"QR_{sel_res}.png", mime="image/png")
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

elif page == "Modo TV":
    # --- PESTAÑAS DEL MODO TV ---
    tab_tv, tab_gestion = st.tabs(["📺 Pantalla de Visualización", "📝 Gestión de Eventos y Anuncios"])

    # ==========================================
    # PESTAÑA 1: PANTALLA TV (VISUALIZACIÓN)
    # ==========================================
    with tab_tv:
        # Auto-refresco cada 60 segundos (60000 ms) solo en esta pestaña
        count = st_autorefresh(interval=60000, key="tv_autorefresh")
        
        # Ocultar menú superior y padding para aprovechar toda la pantalla (Modo Kiosco simulado)
        st.markdown("""
            <style>
                .block-container { padding-top: 1rem; padding-bottom: 0rem; }
                header { visibility: hidden; }
                .evento-card { background-color: #f8f9fa; border-left: 5px solid #0056b3; padding: 15px; margin-bottom: 10px; border-radius: 5px; box-shadow: 2px 2px 5px rgba(0,0,0,0.1); }
                .anuncio-alto { background-color: #ffeaea; border-left: 5px solid #dc3545; padding: 15px; margin-bottom: 10px; border-radius: 5px; }
                .anuncio-medio { background-color: #fff8ea; border-left: 5px solid #ffc107; padding: 15px; margin-bottom: 10px; border-radius: 5px; }
            </style>
        """, unsafe_allow_html=True)

        st.markdown("<h1 style='text-align: center; color: #1f2937;'>Panel Informativo CAV</h1>", unsafe_allow_html=True)
        st.markdown("---")

        col_horario, col_eventos = st.columns([2, 1], gap="large")

        with col_horario:
            st.subheader("📅 Horario de Hoy")
            # Aquí filtramos las reservas para mostrar solo las del día de hoy
            hoy_str = dt.date.today().strftime("%Y-%m-%d")
            reservas_hoy = df[df['Fecha'].astype(str) == hoy_str] if not df.empty else pd.DataFrame()
            
            if not reservas_hoy.empty:
                # Mostrar tabla limpia sin ID y ordenada por hora
                st.dataframe(
                    reservas_hoy[['Hora inicio', 'Hora fin', 'Curso', 'Profesor', 'Recurso']].sort_values('Hora inicio'),
                    hide_index=True, use_container_width=True, height=500
                )
            else:
                st.info("No hay reservas programadas para el día de hoy.")

        with col_eventos:
            st.subheader("📢 Anuncios y Eventos")
            
            # Consultar Anuncios Activos
           # Usamos strftime para darle a Supabase un formato que entiende sin problemas (YYYY-MM-DD HH:MM:SS)
fecha_actual_sql = dt_datetime.now().strftime("%Y-%m-%d %H:%M:%S")
anuncios_data = supabase.table("anuncios_urgentes").select("*").eq("is_active", True).gte("expiracion", fecha_actual_sql).execute().data
            if anuncios_data:
                for a in anuncios_data:
                    clase_css = "anuncio-alto" if a['prioridad'] == 1 else "anuncio-medio"
                    icono = "🚨" if a['prioridad'] == 1 else "⚠️"
                    st.markdown(f"""
                        <div class="{clase_css}">
                            <h4 style="margin:0;">{icono} {a['titulo']}</h4>
                            <p style="margin:5px 0 0 0;">{a['descripcion']}</p>
                        </div>
                    """, unsafe_allow_html=True)

            # Consultar Eventos Activos (Hoy y Futuros)
            eventos_data = supabase.table("eventos_tv").select("*").eq("is_active", True).gte("fecha_evento", hoy_str).order("fecha_evento").execute().data
            if eventos_data:
                for e in eventos_data:
                    iconos_cat = {"Reunión": "🤝", "Examen": "📝", "Efeméride": "📅", "Taller": "🛠️", "Otro": "📌"}
                    st.markdown(f"""
                        <div class="evento-card">
                            <h5 style="margin:0; color: #0056b3;">{iconos_cat.get(e['categoria'], '📌')} {e['titulo']}</h5>
                            <p style="margin:5px 0; font-size: 0.9em;"><b>Fecha:</b> {e['fecha_evento']}</p>
                            <p style="margin:0;">{e['descripcion']}</p>
                        </div>
                    """, unsafe_allow_html=True)
            
            if not anuncios_data and not eventos_data:
                st.write("No hay eventos ni anuncios activos.")

    # ==========================================
    # PESTAÑA 2: GESTIÓN DE EVENTOS (FORMULARIOS)
    # ==========================================
    with tab_gestion:
        st.header("Gestión de Pantalla TV")
        
        col_form1, col_form2 = st.columns(2)
        
        with col_form1:
            with st.container(border=True):
                st.subheader("🗓️ Registrar Nuevo Evento")
                with st.form("form_evento_tv"):
                    titulo_ev = st.text_input("Título del Evento", max_chars=50)
                    desc_ev = st.text_area("Descripción corta", max_chars=200)
                    cat_ev = st.selectbox("Categoría", ["Reunión", "Examen", "Efeméride", "Taller", "Otro"])
                    fecha_ev = st.date_input("Fecha del Evento", min_value=dt.date.today())
                    
                    if st.form_submit_button("Guardar Evento", use_container_width=True):
                        if titulo_ev.strip():
                            supabase.table("eventos_tv").insert({
                                "titulo": titulo_ev, "descripcion": desc_ev, 
                                "fecha_evento": str(fecha_ev), "categoria": cat_ev, "is_active": True
                            }).execute()
                            st.success("Evento guardado. Aparecerá en la TV en el próximo refresco.")
                            time.sleep(1)
                            st.rerun()
                        else:
                            st.error("El título es obligatorio.")

        with col_form2:
            with st.container(border=True):
                st.subheader("🚨 Registrar Anuncio Urgente")
                with st.form("form_anuncio_tv"):
                    titulo_an = st.text_input("Título del Anuncio", max_chars=50)
                    desc_an = st.text_area("Detalles del anuncio")
                    prio_an = st.radio("Prioridad", [("Alta (Rojo)", 1), ("Media (Amarillo)", 2)], format_func=lambda x: x[0])
                    expira_an = st.date_input("Válido hasta", min_value=dt.date.today())
                    
                    if st.form_submit_button("Publicar Anuncio", type="primary", use_container_width=True):
                        if titulo_an.strip():
                            # Añadimos 23:59:59 a la fecha de expiración para que dure todo el día seleccionado
                            exp_dt = dt_datetime.combine(expira_an, dt.time(23, 59, 59)).isoformat()
                            supabase.table("anuncios_urgentes").insert({
                                "titulo": titulo_an, "descripcion": desc_an, 
                                "prioridad": prio_an[1], "expiracion": exp_dt, "is_active": True
                            }).execute()
                            st.success("Anuncio publicado en la TV.")
                            time.sleep(1)
                            st.rerun()
                        else:
                            st.error("El título es obligatorio.")
                        
