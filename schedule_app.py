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

# ------------------------------------------------------------------
# CONFIGURACIÓN DE PÁGINA
# ------------------------------------------------------------------
st.set_page_config(page_title="Sistema de Reservas - Enlaces", page_icon="📅", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #f4f6f9; }
    h1, h2, h3 { color: #2c3e50; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    .stButton>button { border-radius: 5px; font-weight: bold; width: 100%; transition: all 0.3s; }
    .stButton>button:hover { transform: translateY(-2px); box-shadow: 0 4px 8px rgba(0,0,0,0.1); }
    .stSelectbox>div>div, .stDateInput>div>div, .stTimeInput>div>div, .stTextInput>div>div { border-radius: 5px; border: 1px solid #bdc3c7; }
    .card { background-color: white; padding: 20px; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); margin-bottom: 20px; border-left: 5px solid #3498db; }
    .metric-value { font-size: 2em; font-weight: bold; color: #2980b9; }
    .metric-label { font-size: 1em; color: #7f8c8d; text-transform: uppercase; letter-spacing: 1px; }
    .status-badge { padding: 5px 10px; border-radius: 15px; font-size: 0.8em; font-weight: bold; display: inline-block; }
    .status-activa { background-color: #d4edda; color: #155724; }
    .status-completada { background-color: #cce5ff; color: #004085; }
    .status-cancelada { background-color: #f8d7da; color: #721c24; }
    .grid-container {
        display: grid;
        grid-template-columns: 80px repeat(5, 1fr);
        gap: 10px;
        background-color: #ffffff;
        padding: 15px;
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        overflow-x: auto;
    }
    .grid-header { background-color: #34495e; color: white; padding: 10px; text-align: center; font-weight: bold; border-radius: 5px; }
    .time-column { background-color: #ecf0f1; color: #2c3e50; padding: 10px; text-align: center; font-weight: bold; border-radius: 5px; display: flex; align-items: center; justify-content: center;}
    .day-column { min-height: 100px; padding: 10px; border: 1px dashed #bdc3c7; border-radius: 5px; display: flex; flex-direction: column; gap: 10px;}
    .reservation-card { font-size: 0.9em; transition: transform 0.2s; cursor: pointer; }
    .reservation-card:hover { transform: scale(1.02); }
    </style>
""", unsafe_allow_html=True)

# ------------------------------------------------------------------
# FUNCIONES DE BASE DE DATOS (SUPABASE)
# ------------------------------------------------------------------
@st.cache_data(ttl=10)
def cargar_reservas():
    try:
        res = supabase.table("reservas").select("*").execute()
        return pd.DataFrame(res.data) if res.data else pd.DataFrame(columns=['id', 'fecha', 'hora_inicio', 'hora_fin', 'profesor', 'curso', 'recurso', 'observaciones'])
    except Exception as e:
        st.error(f"Error cargando reservas: {e}")
        return pd.DataFrame()

# ------------------------------------------------------------------
# AUTENTICACIÓN SIMPLE
# ------------------------------------------------------------------
if 'autenticado' not in st.session_state:
    st.session_state.autenticado = False

def login():
    st.markdown("<h1 style='text-align: center; color: #3498db;'>🔐 Acceso al Sistema</h1>", unsafe_allow_html=True)
    with st.form("login_form"):
        st.markdown("<div class='card'>", unsafe_allow_html=True)
        usuario = st.text_input("👤 Usuario", placeholder="Ingresa tu nombre de usuario")
        password = st.text_input("🔑 Contraseña", type="password", placeholder="Ingresa tu contraseña")
        submit = st.form_submit_button("Ingresar", use_container_width=True)
        if submit:
            if usuario == "admin" and password == "1234":
                st.session_state.autenticado = True
                st.rerun()
            else:
                st.error("Credenciales incorrectas")
        st.markdown("</div>", unsafe_allow_html=True)

if not st.session_state.autenticado:
    login()
    st.stop()

# ------------------------------------------------------------------
# BARRA LATERAL (NAVEGACIÓN)
# ------------------------------------------------------------------
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/2830/2830305.png", width=100)
    st.title("📚 Gestión Enlaces")
    st.markdown("---")
    
    opciones_menu = {
        "📊 Panel Principal": "dashboard",
        "➕ Nueva Reserva": "nueva",
        "📅 Vista Semanal": "semana",
        "📋 Vista de Horarios": "lista",
        "⚙️ Mantenedores": "mantenedores",
        "🚪 Salir": "salir"
    }
    
    tab_seleccionado = st.radio("Navegación", list(opciones_menu.keys()))
    
    if tab_seleccionado == "🚪 Salir":
        st.session_state.autenticado = False
        st.rerun()
        
    st.markdown("---")
    st.info("💡 **Tip:** Usa la vista semanal para encontrar bloques disponibles rápidamente.")

# ------------------------------------------------------------------
# CARGA DE DATOS PRINCIPAL
# ------------------------------------------------------------------
with st.spinner('Cargando datos desde la nube...'):
    df = cargar_reservas()

# ------------------------------------------------------------------
# RENDERIZADO DE VISTAS
# ------------------------------------------------------------------

# --- PANEL PRINCIPAL ---
if tab_seleccionado == "📊 Panel Principal":
    st.header("📊 Panel de Control")
    if not df.empty:
        col1, col2, col3, col4 = st.columns(4)
        hoy = dt.date.today().strftime('%Y-%m-%d')
        reservas_hoy = df[df['fecha'] == hoy]
        
        with col1:
            st.markdown(f"""
                <div class="card" style="border-left-color: #2ecc71;">
                    <div class="metric-label">Total Reservas</div>
                    <div class="metric-value">{len(df)}</div>
                </div>
            """, unsafe_allow_html=True)
        with col2:
            st.markdown(f"""
                <div class="card" style="border-left-color: #e74c3c;">
                    <div class="metric-label">Reservas Hoy</div>
                    <div class="metric-value">{len(reservas_hoy)}</div>
                </div>
            """, unsafe_allow_html=True)
        with col3:
            recurso_top = df['recurso'].mode()[0] if not df['recurso'].empty else "N/A"
            st.markdown(f"""
                <div class="card" style="border-left-color: #f39c12;">
                    <div class="metric-label">Recurso más usado</div>
                    <div class="metric-value" style="font-size: 1.5em;">{recurso_top}</div>
                </div>
            """, unsafe_allow_html=True)
        with col4:
            profesor_top = df['profesor'].mode()[0] if not df['profesor'].empty else "N/A"
            st.markdown(f"""
                <div class="card" style="border-left-color: #9b59b6;">
                    <div class="metric-label">Profesor Frecuente</div>
                    <div class="metric-value" style="font-size: 1.2em;">{profesor_top[:15]}...</div>
                </div>
            """, unsafe_allow_html=True)

        st.markdown("### Uso de Recursos")
        uso_recursos = df['recurso'].value_counts().reset_index()
        uso_recursos.columns = ['Recurso', 'Cantidad']
        fig = px.bar(uso_recursos, x='Recurso', y='Cantidad', color='Recurso', 
                     template='plotly_white', text_auto=True)
        fig.update_layout(showlegend=False, margin=dict(t=20, b=20, l=20, r=20))
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("No hay reservas registradas en el sistema todavía.")

# --- NUEVA RESERVA ---
elif tab_seleccionado == "➕ Nueva Reserva":
    st.header("➕ Registrar Nueva Reserva")
    
    # Cargar listas desde BD para selects
    profesores_list = [p['nombre'] for p in supabase.table("profesores").select("nombre").order("nombre").execute().data]
    cursos_list = [c['nombre'] for c in supabase.table("cursos").select("nombre").order("nombre").execute().data]
    recursos_list = [r['nombre'] for r in supabase.table("recursos").select("nombre").order("nombre").execute().data]
    
    st.markdown("<div class='card'>", unsafe_allow_html=True)
    with st.form("form_nueva_reserva"):
        col1, col2 = st.columns(2)
        with col1:
            fecha_res = st.date_input("📅 Fecha de Reserva", dt.date.today())
            hora_inicio = st.time_input("⏰ Hora de Inicio", dt.time(8, 0))
            hora_fin = st.time_input("⌛ Hora de Término", dt.time(9, 30))
        with col2:
            prof_res = st.selectbox("👨‍🏫 Profesor", profesores_list)
            curso_res = st.selectbox("📚 Curso", cursos_list)
            rec_res = st.selectbox("💻 Recurso", recursos_list)
            obs_res = st.text_area("📝 Observaciones (Opcional)")
        
        btn_guardar = st.form_submit_button("💾 Guardar Reserva", type="primary")
        
        if btn_guardar:
            # Validar choque
            choque = False
            if not df.empty:
                fechas_df = df[df['fecha'] == fecha_res.strftime('%Y-%m-%d')]
                for _, row in fechas_df.iterrows():
                    if row['recurso'] == rec_res:
                        # Convertir a objetos time para comparar
                        h_ini_bd = dt.datetime.strptime(str(row['hora_inicio'])[:5], '%H:%M').time()
                        h_fin_bd = dt.datetime.strptime(str(row['hora_fin'])[:5], '%H:%M').time()
                        
                        # Lógica de solapamiento de tiempo
                        if max(hora_inicio, h_ini_bd) < min(hora_fin, h_fin_bd):
                            choque = True
                            st.error(f"⚠️ El recurso {rec_res} ya está reservado por {row['profesor']} en ese horario.")
                            break
            
            if not choque:
                nueva_data = {
                    "fecha": fecha_res.strftime('%Y-%m-%d'),
                    "hora_inicio": hora_inicio.strftime('%H:%M:%S'),
                    "hora_fin": hora_fin.strftime('%H:%M:%S'),
                    "profesor": prof_res,
                    "curso": curso_res,
                    "recurso": rec_res,
                    "observaciones": obs_res
                }
                try:
                    supabase.table("reservas").insert(nueva_data).execute()
                    st.success("✅ ¡Reserva registrada con éxito!")
                    st.cache_data.clear() # Limpiar caché para refrescar
                    time.sleep(1)
                    st.rerun()
                except Exception as e:
                    st.error(f"Error al guardar: {e}")
    st.markdown("</div>", unsafe_allow_html=True)

# --- VISTA SEMANAL ---
elif tab_seleccionado == "📅 Vista Semanal":
    st.header("📅 Vista Semanal")
    
    col_d1, col_d2 = st.columns([1, 3])
    with col_d1:
        fecha_seleccionada = st.date_input("Selecciona una fecha para ver su semana", dt.date.today())
    
    # Calcular lunes de esa semana
    inicio_semana = fecha_seleccionada - dt.timedelta(days=fecha_seleccionada.weekday())
    dias_semana_nombres = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]
    
    st.markdown("### 🔍 Filtros de Búsqueda")
    col_f1, col_f2, col_f3 = st.columns(3)
    
    # Obtener valores únicos para los filtros manejando nulos
    opciones_prof = ["Todos"] + sorted([str(p) for p in df['profesor'].dropna().unique()])
    opciones_curso = ["Todos"] + sorted([str(c) for c in df['curso'].dropna().unique()])
    opciones_rec = ["Todos"] + sorted([str(r) for r in df['recurso'].dropna().unique()])

    with col_f1:
        filtro_profesor = st.selectbox("Filtrar por Profesor", opciones_prof, key="sem_prof")
    with col_f2:
        filtro_curso = st.selectbox("Filtrar por Curso", opciones_curso, key="sem_curso")
    with col_f3:
        filtro_recurso = st.selectbox("Filtrar por Recurso", opciones_rec, key="sem_rec")
    
    if df.empty:
        st.warning("No hay datos para mostrar.")
    else:
        df_filtrado = df.copy()
        if filtro_profesor != "Todos": df_filtrado = df_filtrado[df_filtrado['profesor'] == filtro_profesor]
        if filtro_curso != "Todos": df_filtrado = df_filtrado[df_filtrado['curso'] == filtro_curso]
        if filtro_recurso != "Todos": df_filtrado = df_filtrado[df_filtrado['recurso'] == filtro_recurso]
        
        # Agrupar por fecha en un diccionario (asegurando que la fecha sea string)
        reservas_por_fecha = {}
        for idx, row in df_filtrado.iterrows():
            fecha_str = str(row['fecha'])
            if fecha_str not in reservas_por_fecha:
                reservas_por_fecha[fecha_str] = []
            reservas_por_fecha[fecha_str].append(row)
            
        colores_recursos = {
            "ENLACE MEDIA": "#3498db",
            "ENLACE BASICA": "#2ecc71",
            "MOVIL 1": "#e74c3c",
            "MOVIL 4": "#f1c40f",
            "TABLETS": "#9b59b6"
        }

        # ¡AQUÍ ESTÁ LA MAGIA Y LA SOLUCIÓN AL BUG!
        # Construimos todo el HTML de la cuadrícula en un solo texto antes de renderizarlo
        html_grid = '<div class="grid-container">'
        
        # Fila de Cabeceras
        html_grid += '<div class="grid-header">Hora</div>'
        for i in range(5):
            dia = inicio_semana + dt.timedelta(days=i)
            html_grid += f'<div class="grid-header">{dias_semana_nombres[i]}<br>{dia.strftime("%d %m %Y")}</div>'
        
        # Extraer y ordenar bloques de tiempo (quitando nulos)
        tiempos_validos = df_filtrado['hora_inicio'].dropna().unique().tolist()
        bloques_tiempo = sorted([str(t) for t in tiempos_validos])
        
        # Rellenar la cuadrícula
        for hora_inicio in bloques_tiempo:
            html_grid += f'<div class="time-column">{hora_inicio[:5]}</div>'
            for i in range(5):
                dia = inicio_semana + dt.timedelta(days=i)
                fecha_actual_str = dia.strftime("%Y-%m-%d")
                
                # Filtrar las reservas que coinciden en fecha y hora
                reservas_celda = [r for r in reservas_por_fecha.get(fecha_actual_str, []) if str(r['hora_inicio']) == hora_inicio]
                
                if reservas_celda:
                    contenido_celda = ""
                    for r in reservas_celda:
                        rec_name = str(r.get('recurso', ''))
                        color = colores_recursos.get(rec_name, "#95a5a6")
                        
                        prof_texto = str(r.get('profesor', 'N/A'))
                        cur_texto = str(r.get('curso', 'N/A'))
                        rec_texto = str(r.get('recurso', 'N/A'))
                        hora_i = str(r.get('hora_inicio', ''))[:5]
                        hora_f = str(r.get('hora_fin', ''))[:5]
                        
                        contenido_celda += f"""
                            <div class="reservation-card" style="border-left: 4px solid {color}; padding: 10px; margin-bottom: 5px; background-color: #f8f9fa; border-radius: 5px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                                <div style="font-weight: bold; color: #2c3e50;">{prof_texto}</div>
                                <div style="font-size: 0.9em; color: #34495e;">📚 {cur_texto}</div>
                                <div style="font-size: 0.9em; color: #7f8c8d;">💻 {rec_texto}</div>
                                <div style="font-size: 0.8em; color: #95a5a6; margin-top: 5px;">{hora_i} - {hora_f}</div>
                            </div>
                        """
                    html_grid += f'<div class="day-column">{contenido_celda}</div>'
                else:
                    html_grid += '<div class="day-column" style="background-color: #fdfdfd;"></div>'
                    
        html_grid += '</div>' # Cerrar el grid-container
        
        # Imprimir todo en Streamlit de una sola vez
        st.markdown(html_grid, unsafe_allow_html=True)

# --- VISTA DE HORARIOS (LISTA) ---
elif tab_seleccionado == "📋 Vista de Horarios":
    st.header("📋 Vista de Horarios (Detallada)")
    
    if df.empty:
        st.info("No hay reservas registradas.")
    else:
        st.markdown("### 🔍 Filtros de Búsqueda")
        col_f1, col_f2, col_f3, col_f4 = st.columns(4)
        
        opc_prof = ["Todos"] + sorted([str(p) for p in df['profesor'].dropna().unique()])
        opc_cur = ["Todos"] + sorted([str(c) for c in df['curso'].dropna().unique()])
        opc_rec = ["Todos"] + sorted([str(r) for r in df['recurso'].dropna().unique()])

        with col_f1:
            filtro_fecha = st.date_input("Filtrar por Fecha", value=None)
        with col_f2:
            filtro_prof = st.selectbox("Filtrar por Profesor", opc_prof)
        with col_f3:
            filtro_cur = st.selectbox("Filtrar por Curso", opc_cur)
        with col_f4:
            filtro_rec = st.selectbox("Filtrar por Recurso", opc_rec)
            
        df_mostrar = df.copy()
        if filtro_fecha: df_mostrar = df_mostrar[df_mostrar['fecha'] == filtro_fecha.strftime('%Y-%m-%d')]
        if filtro_prof != "Todos": df_mostrar = df_mostrar[df_mostrar['profesor'] == filtro_prof]
        if filtro_cur != "Todos": df_mostrar = df_mostrar[df_mostrar['curso'] == filtro_cur]
        if filtro_rec != "Todos": df_mostrar = df_mostrar[df_mostrar['recurso'] == filtro_rec]
        
        # Renombrar columnas para la visualización en español
        columnas_espanol = {
            'fecha': 'Fecha',
            'hora_inicio': 'Hora Inicio',
            'hora_fin': 'Hora Fin',
            'profesor': 'Profesor',
            'curso': 'Curso',
            'recurso': 'Recurso',
            'observaciones': 'Observaciones'
        }
        
        # Mostrar la tabla ordenada por Fecha y Hora de Inicio
        df_mostrar = df_mostrar.sort_values(by=['fecha', 'hora_inicio'], ascending=[False, True])
        
        st.dataframe(
            df_mostrar[['fecha', 'hora_inicio', 'hora_fin', 'profesor', 'curso', 'recurso', 'observaciones']].rename(columns=columnas_espanol), 
            use_container_width=True, 
            hide_index=True
        )
        
        with st.expander("🗑️ Eliminar una reserva"):
            if not df_mostrar.empty:
                opciones_borrar = [f"ID: {row['id']} | {row['fecha']} - {row['profesor']} - {row['recurso']}" for _, row in df_mostrar.iterrows()]
                res_borrar = st.selectbox("Selecciona la reserva a eliminar", opciones_borrar)
                if st.button("Eliminar Reserva", type="primary"):
                    try:
                        id_borrar = int(res_borrar.split("|")[0].replace("ID:", "").strip())
                        supabase.table("reservas").delete().eq("id", id_borrar).execute()
                        st.success("Reserva eliminada.")
                        st.cache_data.clear()
                        time.sleep(1)
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error al eliminar: {e}")

# --- MANTENEDORES ---
elif tab_seleccionado == "⚙️ Mantenedores":
    st.header("⚙️ Mantenedores del Sistema")
    st.markdown("Administra los catálogos base de la aplicación.")
    
    tab_p, tab_c, tab_r = st.tabs(["👨‍🏫 Profesores", "📚 Cursos", "💻 Recursos"])
    
    with tab_p:
        col_form_p, col_list_p = st.columns([1, 2])
        with col_form_p:
            with st.form("form_nuevo_prof"):
                st.subheader("Agregar Profesor")
                nuevo_prof = st.text_input("Nombre Completo")
                if st.form_submit_button("Guardar"):
                    if nuevo_prof.strip():
                        try:
                            supabase.table("profesores").insert({"nombre": nuevo_prof.strip().upper()}).execute()
                            st.success("¡Profesor agregado!")
                            st.cache_data.clear(); time.sleep(0.5); st.rerun()
                        except Exception as e: st.error(f"Error al agregar: {e}")
                    else: st.error("El nombre es obligatorio.")
        with col_list_p:
            prof_data = supabase.table("profesores").select("*").order("nombre").execute().data
            if prof_data:
                df_p = pd.DataFrame(prof_data)
                st.dataframe(df_p[['nombre']], use_container_width=True, hide_index=True)
                with st.expander("🗑️ Eliminar un Profesor"):
                    prof_borrar = st.selectbox("Selecciona el profesor a eliminar", df_p['nombre'].tolist(), key="del_prof")
                    if st.button("Eliminar Profesor Definitivamente", type="primary"):
                        try:
                            id_b = int(df_p[df_p['nombre'] == prof_borrar]['id'].values[0])
                            supabase.table("profesores").delete().eq("id", id_b).execute()
                            st.success(f"Profesor {prof_borrar} eliminado.")
                            st.cache_data.clear(); time.sleep(0.5); st.rerun()
                        except Exception as e: st.error("No se puede eliminar, posiblemente tenga reservas asociadas.")
                        
    with tab_c:
        col_form_c, col_list_c = st.columns([1, 2])
        with col_form_c:
            with st.form("form_nuevo_curso"):
                st.subheader("Agregar Curso")
                nuevo_curso = st.text_input("Nombre del Curso (Ej: 5° BÁSICO A)")
                if st.form_submit_button("Guardar"):
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
                        except Exception as e: st.error("No se puede eliminar, posiblemente tenga reservas asociadas.")

    with tab_r:
        col_form_r, col_list_r = st.columns([1, 2])
        with col_form_r:
            with st.form("form_nuevo_recurso"):
                st.subheader("Agregar Recurso")
                nuevo_rec = st.text_input("Nombre del Recurso")
                if st.form_submit_button("Guardar"):
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
                        except Exception as e: st.error("No se puede eliminar porque tiene reservas asociadas.")
