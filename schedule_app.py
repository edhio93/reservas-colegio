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
import base64
import os
from streamlit_autorefresh import st_autorefresh
import base64
import os

import requests
from icalendar import Calendar
import pytz
from datetime import datetime, date

@st.cache_data(ttl=1800) # Se actualiza cada 30 minutos para no saturar el servidor
def obtener_clima_vicuna():
    try:
        # Coordenadas exactas de Vicuña, Chile
        lat = -30.0319
        lon = -70.7081
        url = f"https://api.open-meteo.com/v1/forecast?latitude={lat}&longitude={lon}&current_weather=true"
        
        respuesta = requests.get(url, timeout=5).json()
        temp = respuesta["current_weather"]["temperature"]
        codigo_clima = respuesta["current_weather"]["weathercode"]
        
        # Elegir el ícono Phosphor correcto según el clima
        if codigo_clima in [0, 1]: # Despejado
            icon = "ph-sun"
            color = "#f59e0b" # Naranja sol
        elif codigo_clima in [2, 3]: # Nublado parcial
            icon = "ph-cloud-sun"
            color = "#94a3b8" # Gris azulado
        elif codigo_clima in [45, 48]: # Niebla
            icon = "ph-cloud-fog"
            color = "#cbd5e1"
        elif codigo_clima in [51, 53, 55, 61, 63, 65, 80, 81, 82]: # Lluvia
            icon = "ph-cloud-rain"
            color = "#3b82f6" # Azul agua
        else: # Otros (tormenta, nieve, etc)
            icon = "ph-cloud"
            color = "#64748b"

        # Devolvemos el HTML listo para insertar
        return f"<i class='ph-fill {icon}' style='color: {color}; font-size: 1.3rem; margin-right: 5px; vertical-align: middle;'></i> <span style='vertical-align: middle;'>{temp}°C</span>"
    except Exception as e:
        return "" # Si no hay internet, no mostramos nada para que no salga error

st.set_page_config(page_title="Sistema de Horarios CAV", page_icon="📅", layout="wide", initial_sidebar_state="expanded")


@st.cache_data(ttl=1800) # Se actualiza cada 30 minutos
def obtener_eventos_google_calendar(url_ics):
    if not url_ics:
        return []
    try:
        respuesta = requests.get(url_ics)
        respuesta.raise_for_status()
        
        # Leer el calendario con la nueva librería
        calendario = Calendar.from_ical(respuesta.text)
        zona_horaria = pytz.timezone('America/Santiago')
        ahora = dt_datetime.now(zona_horaria).date()
        
        eventos_hoy = []
        
        # Buscar todos los eventos (VEVENT)
        for componente in calendario.walk('VEVENT'):
            inicio = componente.get('dtstart').dt
            
            # Verificar si es un evento con hora específica
            if isinstance(inicio, datetime):
                inicio_dt = inicio.astimezone(zona_horaria)
                es_hoy = (inicio_dt.date() == ahora)
                hora_str = inicio_dt.strftime("%H:%M")
                hora_sort = hora_str
            # O si es un evento de "Todo el día" (solo fecha)
            else:
                es_hoy = (inicio == ahora)
                hora_str = "TODO EL DÍA"
                hora_sort = "00:00" # Para que salga primero en la lista

            if es_hoy:
                titulo = str(componente.get('summary', 'Evento Especial'))
                descripcion = str(componente.get('description', ''))
                
                eventos_hoy.append({
                    "hora_sort": hora_sort,
                    "display_hora": hora_str,
                    "titulo": titulo,
                    "descripcion": descripcion,
                    "categoria": "Evento Especial"
                })
                
        # Ordenamos los eventos por hora
        return sorted(eventos_hoy, key=lambda x: x['hora_sort'])
    except Exception as e:
        return []
        
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


# ==============================================================================
# 📺 PANTALLA INFORMATIVA PÚBLICA (MODO KIOSCO SIN LOGIN)
# ==============================================================================
if "ver_pantalla_tv" in st.session_state and st.session_state.ver_pantalla_tv:

    # Inicializamos el zoom en el session state si no existe
    if "tv_scale" not in st.session_state:
        st.session_state.tv_scale = 100

    # Temporizador de recarga de la página: 20000ms = 20 segundos
    refresh_count = st_autorefresh(interval=20000, limit=None, key="tv_refresh_timer")
    
    ruta_logo = "logotv.png"
    logo_src_html = ""
    if os.path.exists(ruta_logo):
        with open(ruta_logo, "rb") as image_file:
            encoded_string = base64.b64encode(image_file.read()).decode()
            logo_src_html = f"<img src='data:image/png;base64,{encoded_string}' class='header-logo-img'/>"
    else:
        logo_src_html = "<i class='ph-fill ph-airplane-landing header-logo-fallback'></i>"

    now_dt = dt_datetime.now()
    hoy_str = now_dt.strftime("%Y-%m-%d")

    dias_es = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
    meses_es = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    fecha_es_formateada = f"{dias_es[now_dt.weekday()]}, {now_dt.day} de {meses_es[now_dt.month - 1]} de {now_dt.year}"

    # Calculamos el multiplicador de escala
    escala = st.session_state.tv_scale / 100.0

    # === ESTILO: FONDO OSCURO, TARJETAS BLANCAS ===
    aesthetic_style = f"""            
    <style>
        @import url('https://unpkg.com/@phosphor-icons/web@2.1.1/src/fill/style.css');

        :root {{
            --tv-scale: {escala};
        }}

        /* Fondo general oscuro */
        .stApp {{ background-color: #0f172a; color: #f8fafc; font-family: 'Inter', sans-serif; }}
        [data-testid="stHeader"] {{ background: rgba(0,0,0,0); }}
        [data-testid="stToolbar"] {{ display: none; }}
        [data-testid="stSidebar"] {{ display: none; }}
        
        /* Cabecera Principal - Blanca/Gris Clara */
        .tv-header-container {{ 
            background: linear-gradient(135deg, #f8fafc 0%, #e2e8f0 100%);
            color: #0f172a; padding: 15px 25px 0 25px; border-radius: 20px; 
            margin-bottom: 25px; border: 1px solid #cbd5e1; 
            box-shadow: 0 4px 15px rgba(0,0,0,0.15); overflow: hidden; 
            display: flex; flex-direction: column; align-items: center;
        }}
        .header-content-layout {{ display: flex; align-items: center; justify-content: space-between; width: 100%; padding-bottom: 15px; }}
        .header-logo-img {{ height: calc(85px * var(--tv-scale)); width: auto; display: block; }}
        .header-logo-fallback {{ font-size: calc(4rem * var(--tv-scale)); color: #64748b; line-height: 1; display: block; }}

        .header-info-group {{ display: flex; align-items: center; gap: 15px; font-size: calc(1.2rem * var(--tv-scale)); font-weight: 600; color: #1e293b; }}
        .header-divider {{ opacity: 0.3; font-weight: 300; font-size: calc(1.5rem * var(--tv-scale)); color: #94a3b8; }}
        .header-status {{ display: flex; align-items: center; color: #475569; }}
        .status-icon {{ margin-right: 8px; font-size: calc(1.3rem * var(--tv-scale)); color: #10b981; }}
        
        /* Barra de progreso 20s */
        .progress-container {{ width: 100%; height: 6px; background-color: #cbd5e1; }}
        .progress-bar {{ height: 100%; background-color: #3b82f6; width: 0%; animation: loadBar 20s linear infinite; }}
        @keyframes loadBar {{ 0% {{ width: 0%; }} 100% {{ width: 100%; }} }}
        
        /* Títulos de sección adaptados para fondo oscuro */
        .tv-sub-header {{ color: #f8fafc; font-weight: 800; font-size: calc(1.6rem * var(--tv-scale)); margin-top: 5px; margin-bottom: 15px; text-transform: uppercase; letter-spacing: 1px; border-bottom: 2px solid #334155; padding-bottom: 10px;}}
        
        /* Animaciones */
        @keyframes cascadeIn {{ 0% {{ opacity: 0; transform: translateY(30px) scale(0.98); }} 100% {{ opacity: 1; transform: translateY(0) scale(1); }} }}
        @keyframes pulseAlert {{ 0% {{ box-shadow: 0 0 0 0 rgba(239, 68, 68, 0.4); }} 70% {{ box-shadow: 0 0 0 15px rgba(239, 68, 68, 0); }} 100% {{ box-shadow: 0 0 0 0 rgba(239, 68, 68, 0); }} }}

        /* Tarjetas Cronograma - Blancas */
        .block-card {{ padding: 22px; border-radius: 16px; border-left: 8px solid; margin-bottom: 18px; background-color: white; box-shadow: 0 6px 15px rgba(0,0,0,0.1); }}
        
        /* Nuevo título con hora a la derecha */
        .block-title-row {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px; width: 100%; }}
        .block-title-text {{ font-weight: 700; font-size: calc(1.35rem * var(--tv-scale)); text-transform: uppercase; }}
        .block-time-badge {{ display: flex; align-items: center; gap: 6px; background-color: #f1f5f9; color: #475569; padding: 6px 12px; border-radius: 8px; font-weight: 700; font-size: calc(1.1rem * var(--tv-scale)); border: 1px solid #e2e8f0; }}
        
        .block-info {{ font-size: calc(1.05rem * var(--tv-scale)); color: #475569; }}
        .block-info-row {{ display: flex; gap: 20px; align-items: center; margin-top: 12px; font-size: calc(1rem * var(--tv-scale));}}
        .block-info-item {{ display: flex; align-items: center; color: #64748b; }}
        
        .info-icon {{ font-size: calc(1.25rem * var(--tv-scale)); margin-right: 8px; }}
        .icon-profesor {{ color: #4ade80; }} 
        .icon-observaciones {{ color: #fbbf24; }} 
        
        .block-hora-pill {{ margin-top: 15px; font-weight: 600; color: #64748b; background: #f8fafc; display: inline-flex; align-items: center; padding: 6px 12px; border-radius: 10px; font-size: calc(0.9rem * var(--tv-scale)); border: 1px solid #e2e8f0;}}
        .icon-categoria {{ color: #818cf8; margin-right: 7px; font-size: calc(1.1rem * var(--tv-scale));}} 

        /* Contenedor Avisos - Blanco */
        .announcements-container {{ background-color: white; border-radius: 20px; padding: 25px; border: 1px solid #e2e8f0; box-shadow: 0 4px 15px rgba(0,0,0,0.1); height: 100%; }}
        .announcement-card {{ padding: 20px; border-radius: 16px; border: 1px solid #e2e8f0; margin-bottom: 18px; border-left: 6px solid;}}
        .announcement-title {{ font-weight: 700; margin-bottom: 7px; font-size: calc(1.2rem * var(--tv-scale)); text-transform: uppercase;}}
        .announcement-desc {{ font-size: calc(1rem * var(--tv-scale)); color: #334155; line-height: 1.5; }}
    </style>
    """
    st.markdown(aesthetic_style, unsafe_allow_html=True) 

    clima_html = obtener_clima_vicuna()

    html_cabecera = f"""
<div class="tv-header-container">
    <div class="header-content-layout">
        <div class="header-logo-section">
            {logo_src_html}
        </div>
        <div class="header-info-group">
            <div class="header-date">{fecha_es_formateada}</div>
            <div class="header-divider">|</div>
            <div class="header-weather">
                {clima_html}
            </div>
            <div class="header-divider">|</div>
            <div class="header-status">
                <i class="ph-fill ph-check-circle status-icon"></i> Siguiente en 20s
            </div>
        </div>
    </div>
    <div class="progress-container"><div class="progress-bar"></div></div>
</div>
"""
    st.markdown(html_cabecera, unsafe_allow_html=True)
    
    col_main, col_ann = st.columns([2.5, 1], gap="large")
    
        # ==========================================
        # COLUMNA IZQUIERDA (CRONOGRAMA DE EVENTOS)
        # ==========================================
    with col_main:
            try:
                # Capturamos la hora y minuto actual en formato HH:MM (Ej: "14:30")
                hora_actual = dt_datetime.now().strftime("%H:%M")
                perfil_actual = st.session_state.tv_profile.upper()
                
                # 1. Inicializar lista maestra
                events_hoy_list = []
                
                # 2. Cargar Google Calendar (Todos los perfiles)
                url_guardada = st.session_state.get('url_calendario_tv', '')
                eventos_calendar = obtener_eventos_google_calendar(url_guardada)
                events_hoy_list.extend(eventos_calendar) 
                
                # 3. Cargar Eventos TV de Supabase (Todos los perfiles)
                res_tv = supabase.table("eventos_tv").select("*").eq("fecha_evento", hoy_str).eq("is_active", True).execute()
                res_tv_hoy = res_tv.data if res_tv.data else []
                
                for ev in res_tv_hoy:
                    h_ini = str(ev.get("hora_inicio", ev.get("hora", "00:00")))[:5]
                    h_fin = str(ev.get("hora_fin", "23:59"))[:5]
                    if not h_fin or h_fin == "None" or h_fin.strip() == "": 
                        h_fin = "23:59"
                    
                    if hora_actual > h_fin and h_fin != "23:59":
                        continue
                        
                    disp_hora = f"{h_ini} - {h_fin}" if h_fin and h_fin != h_ini and h_fin != "23:59" else f"{h_ini}"
                    if not h_ini or h_ini == "None" or h_ini == "00:00": disp_hora = "TODO EL DÍA"
                    
                    events_hoy_list.append({
                        "hora_sort": h_ini if h_ini and h_ini != "None" and h_ini != "00:00" else "00:00", 
                        "display_hora": disp_hora,
                        "titulo": ev.get("titulo", "Evento"), 
                        "descripcion": ev.get("descripcion", ""),
                        "categoria": ev.get("categoria", "Evento")
                    })
                    
                # 4. Cargar Reservas de Enlaces (SOLO Perfil Profesores/PIE)
                if "PROFESORES" in perfil_actual or "PIE" in perfil_actual:
                    res_supabase = supabase.table("reservas").select("*, profesores(nombre), recursos(nombre), cursos(nombre)").eq("fecha", hoy_str).execute()
                    res_reservas_hoy = res_supabase.data if res_supabase.data else []
                    
                    for r in res_reservas_hoy:
                        h_ini = str(r.get("hora_inicio", r.get("hora", "00:00")))[:5]
                        h_fin = str(r.get("hora_fin", "23:59"))[:5]
                        if not h_fin or h_fin == "None" or h_fin.strip() == "": 
                            h_fin = "23:59"
                        
                        if hora_actual > h_fin and h_fin != "23:59":
                            continue
                            
                        disp_hora = f"{h_ini} - {h_fin}" if h_fin and h_fin != h_ini and h_fin != "23:59" else f"{h_ini}"
                        if not h_ini or h_ini == "None" or h_ini == "00:00": disp_hora = "RESERVA"
                        
                        prof = r.get("profesores", {}).get("nombre", "Docente") if r.get("profesores") else "Docente"
                        rec = r.get("recursos", {}).get("nombre", "Recurso") if r.get("recursos") else "Recurso"
                        curso = r.get("cursos", {}).get("nombre", "Curso") if r.get("cursos") else "Curso"
                        obs = r.get("observaciones", "")
                        
                        events_hoy_list.append({
                            "hora_sort": h_ini if h_ini and h_ini != "None" and h_ini != "00:00" else "23:59", 
                            "display_hora": disp_hora,
                            "titulo": f"{rec} ➔ {curso}", 
                            "profesor": prof,
                            "observaciones": obs, 
                            "categoria": "Clase / Uso Recurso"
                        })
                
                # 5. Ordenar cronológicamente
                events_hoy_list = sorted(events_hoy_list, key=lambda x: str(x.get("hora_sort", "99:99")))
                
                if not events_hoy_list:
                    st.markdown("<div class='tv-sub-header'>⏱️ Cronograma de Hoy</div>", unsafe_allow_html=True)
                    st.info(f"No hay eventos activos para el resto del día en este perfil.")
                else:
                    ITEMS_POR_PAGINA = 3
                    total_paginas = max(1, (len(events_hoy_list) + ITEMS_POR_PAGINA - 1) // ITEMS_POR_PAGINA)
                    pagina_actual = refresh_count % total_paginas 
                    
                    inicio_idx = pagina_actual * ITEMS_POR_PAGINA
                    fin_idx = inicio_idx + ITEMS_POR_PAGINA
                    eventos_a_mostrar = events_hoy_list[inicio_idx:fin_idx]
                    
                    st.markdown(f"<div class='tv-sub-header'>⏱️ Cronograma de Hoy (Pág. {pagina_actual + 1}/{total_paginas})</div>", unsafe_allow_html=True)
                    
                    paleta_colores = ["#0ea5e9", "#10b981", "#f59e0b", "#ec4899", "#8b5cf6", "#14b8a6"]
                    html_cronograma = ""
                    
                    for i, item in enumerate(eventos_a_mostrar):
                        color_tema = "#6366f1" if item['categoria'] == "Evento Especial" else paleta_colores[i % len(paleta_colores)]
                        delay = i * 0.15 
                        
                        info_row_html = ""
                        if item.get("profesor") or item.get("observaciones"):
                            info_row_html = "<div class='block-info-row'>"
                            if item.get("profesor"):
                                info_row_html += f"<div class='block-info-item'><i class='ph-fill ph-user-graduate info-icon'></i> {item['profesor']}</div>"
                            if item.get("observaciones"):
                                info_row_html += f"<div class='block-info-item'><i class='ph-fill ph-clipboard-text info-icon'></i> {item['observaciones']}</div>"
                            info_row_html += "</div>"
                        
                        html_cronograma += (
                            f"<div class='block-card' style='border-left-color: {color_tema}; animation: cascadeIn 0.8s forwards; animation-delay: {delay}s; opacity: 0;'>"
                            f"  <div class='block-title-row'>"
                            f"      <div class='block-title-text' style='color: {color_tema};'>{item['titulo']}</div>"
                            f"      <div class='block-time-badge'><i class='ph-fill ph-clock'></i> {item['display_hora']}</div>"
                            f"  </div>"
                            f"  <div class='block-info'>{item.get('descripcion', '')}</div>"
                            f"  {info_row_html}"
                            f"  <div class='block-hora-pill'><i class='ph-fill ph-tag icon-categoria'></i> {item['categoria']}</div>"
                            f"</div>"
                        )
                    st.markdown(html_cronograma, unsafe_allow_html=True)
                    
            except Exception as e:
                st.error(f"Error al cargar cronograma: {e}")

        # ==========================================
        # COLUMNA DERECHA (ANUNCIOS Y CONTROLES)
        # ==========================================
    with col_ann:
            # --- SELECTOR DE PERFILES ---
            st.selectbox(
                "👁️ Perfil de Visualización", 
                ["Inspectoría / UTP", "Profesores / PIE", "Apoderados"], 
                key="tv_profile"
            )
            st.markdown("<div style='margin-bottom: 15px;'></div>", unsafe_allow_html=True)
            
            # --- CONTROLES SECUNDARIOS ---
            with st.expander("⚙️ Controles Extra", expanded=False):
                st.slider("🔍 Tamaño del texto (%)", min_value=50, max_value=250, value=st.session_state.tv_scale, step=5, key="tv_scale")
                st.markdown("<hr style='margin: 10px 0;'>", unsafe_allow_html=True)
                if st.button("🔙 Volver al Menú", use_container_width=True):
                    st.session_state.ver_pantalla_tv = False
                    st.rerun()
                    
                components.html(
                    """
                    <style>
                        body { margin: 0; padding: 0; font-family: 'Inter', sans-serif; }
                        button {
                            width: 100%; height: 38px; background-color: #ffffff;
                            border: 1px solid #cbd5e1; border-radius: 8px; color: #0f172a;
                            font-size: 14px; font-weight: 500; cursor: pointer;
                            display: flex; align-items: center; justify-content: center; gap: 8px;
                            transition: all 0.2s;
                        }
                        button:hover { border-color: #94a3b8; background-color: #f8fafc; }
                    </style>
                    <button onclick="
                        const doc = window.parent.document;
                        if (!doc.fullscreenElement) {
                            doc.documentElement.requestFullscreen();
                            this.innerHTML = '🗗 Salir Pantalla Completa';
                        } else {
                            doc.exitFullscreen();
                            this.innerHTML = '🔲 Pantalla Completa';
                        }
                    ">
                        🔲 Pantalla Completa
                    </button>
                    """,
                    height=40
                )

            # --- SECCIÓN DE ANUNCIOS ---
            try:
                now = dt_datetime.now()
                ann_data = supabase.table("anuncios_urgentes").select("id, titulo, descripcion, prioridad, expiracion").eq("is_active", True).execute().data
                
                active_ann = []
                for ann in ann_data:
                    try:
                        exp_dt = pd.to_datetime(ann['expiracion']).tz_localize(None)
                        if exp_dt > now: 
                            active_ann.append(ann)
                    except:
                        pass
                
                active_ann = sorted(active_ann, key=lambda x: x['prioridad'])
                
                if st.session_state.tv_profile == "Apoderados":
                    titulo_panel = "📰 Noticias y Comunicados"
                    texto_vacio = "No hay comunicados vigentes en este momento."
                else:
                    titulo_panel = "🚨 Avisos Urgentes y Alertas"
                    texto_vacio = "No hay avisos en este momento."

                html_anuncios = '<div class="announcements-container">'
                
                if not active_ann:
                    st.markdown(f"<div class='tv-sub-header'>{titulo_panel}</div>", unsafe_allow_html=True)
                    html_anuncios += f"<p style='color: #64748b; text-align:center; font-style:italic; margin-top: 10px;'>{texto_vacio}</p>"
                else:
                    ITEMS_POR_PAGINA_ANN = 3
                    total_paginas_ann = max(1, (len(active_ann) + ITEMS_POR_PAGINA_ANN - 1) // ITEMS_POR_PAGINA_ANN)
                    pagina_actual_ann = refresh_count % total_paginas_ann 
                    
                    inicio_idx_ann = pagina_actual_ann * ITEMS_POR_PAGINA_ANN
                    fin_idx_ann = inicio_idx_ann + ITEMS_POR_PAGINA_ANN
                    anuncios_a_mostrar = active_ann[inicio_idx_ann:fin_idx_ann]
                    
                    st.markdown(f"<div class='tv-sub-header'>{titulo_panel} (Pág. {pagina_actual_ann + 1}/{total_paginas_ann})</div>", unsafe_allow_html=True)

                    for i, ann in enumerate(anuncios_a_mostrar):
                        delay_ann = i * 0.15 
                        
                        if st.session_state.tv_profile == "Apoderados":
                            bg_color = "#f0f9ff"; border_color = "#38bdf8"; title_color = "#0369a1"; desc_color = "#334155"
                            animacion_extra = ""
                        else:
                            if ann['prioridad'] == 1:
                                bg_color = "#fef2f2"; border_color = "#ef4444"; title_color = "#dc2626"; desc_color = "#334155"
                                animacion_extra = ", pulseAlert 2s infinite"
                            else:
                                bg_color = "#fffbeb"; border_color = "#f59e0b"; title_color = "#d97706"; desc_color = "#334155"
                                animacion_extra = "" 
                        
                        html_anuncios += (
                            f"<div class='announcement-card' style='border-left-color: {border_color}; background-color: {bg_color}; animation: cascadeIn 0.8s cubic-bezier(0.2, 0.8, 0.2, 1) forwards{animacion_extra}; animation-delay: {delay_ann}s; opacity: 0;'>"
                            f"<div class='announcement-title' style='color: {title_color};'>{ann['titulo']}</div>"
                            f"<div class='announcement-desc' style='color: {desc_color};'>{ann['descripcion']}</div>"
                            f"</div>"
                        )
                
                html_anuncios += '</div>'
                st.markdown(html_anuncios, unsafe_allow_html=True)
                
            except Exception as e:
                st.error(f"Error técnico al consultar anuncios: {e}")

        # MUY IMPORTANTE: Este stop() frena que se dibuje el resto del menú cuando estamos en modo TV
    st.stop()
# ──────────────────────────────────────────────────────────────────────────────
# 0) CONFIGURACIÓN GLOBAL Y ESTILO
# ──────────────────────────────────────────────────────────────────────────────
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

    # 3. ENCABEZADO
    st.markdown("<h1 style='text-align: center; font-size: 50px;'>🏫</h1>", unsafe_allow_html=True)
    st.markdown("<h2 style='text-align: center; color: #1E3A8A; font-weight: 800; margin-bottom: 0px;'>🚨 Reportar Falla</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #64748b; font-size: 14px; margin-top: 5px;'>Colegio Antonio Varas - Depto. Enlaces</p>", unsafe_allow_html=True)
    st.write("") 
    
    # 4. LÓGICA DEL REPORTE
    recurso_id_str = query_params.get("id", "")
    
    if recurso_id_str:
        try:
            # ¡CLAVE AQUÍ! Convertimos el ID de texto a número entero
            recurso_id = int(recurso_id_str)
            
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
                                # ¡CORREGIDA LA INDENTACIÓN AQUÍ!
                                supabase.table("mantenimientos").insert({
                                    "recurso_id": recurso_id,
                                    "descripcion": descripcion.strip(),
                                    "estado": "Reportado (Vía QR)",
                                    "reportado_por": nombre_reporta.strip(),
                                    "fecha": str(dt.date.today())  
                                }).execute()
                                
                                st.success("✅ ¡Reporte enviado con éxito! Gracias por avisarnos.")
                                st.balloons() # 🎈 ¡AQUÍ ESTÁN LOS GLOBOS! 🎈
                                time.sleep(4) 
            else:
                st.error("❌ El equipo que intentas reportar no existe o fue dado de baja.")
        except ValueError:
            st.error("❌ El enlace del QR es inválido (el ID debe ser un número).")
        except Exception as e:
            st.error(f"Error técnico al guardar: {e}")
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
            if logo_path.exists(): 
                st.image(str(logo_path), use_container_width=True)
            else: 
                st.info("Logo CAV")

        with col_form:
            st.markdown("<h2 style='text-align: center; color: #1E3A8A; margin-bottom: 0px;'>SISTEMA CAV</h2>", unsafe_allow_html=True)
            st.markdown("<p style='text-align: center; color: gray; font-size: 0.9rem;'>Reserva de Recursos y Espacios</p>", unsafe_allow_html=True)
            
            with st.container(border=True):
                tipo_user = st.radio("Acceder como:", ["Profesor", "Administrador", "Mensajería Interna"], horizontal=True)
                st.markdown("---")
                
                if tipo_user == "Administrador":
                    with st.form("admin_form", clear_on_submit=True):
                        u_adm = st.text_input("Nombre de Administrador", placeholder="Ej: Edgar")
                        p_adm = st.text_input("Contraseña", type="password", placeholder="••••••••")
                        
                        # --- Aquí estaba el error. Ahora está correctamente indentado hacia adentro ---
                        if st.form_submit_button("INICIAR SESIÓN ADMIN", use_container_width=True, type="primary"):
                            if u_adm.strip().upper() in ["EDGAR", "GLORIA", "CARLOS", "ALEXIS"] and p_adm == "cav690":
                                st.session_state.logged = True
                                st.session_state.role = "admin"
                                st.session_state.profesor_name = u_adm.strip().capitalize()
                                st.rerun()
                            else:
                                st.error("Acceso denegado")
                elif tipo_user == "Mensajería Interna":
                    st.info("💡 Acceso rápido libre para la gestión de las Pantallas Informativas.")
            
                    # Un simple botón que los deja entrar directo sin preguntar nada más
                    if st.button("ENTRAR AL PANEL DE MENSAJERÍA", use_container_width=True, type="primary"):
                        st.session_state.logged = True
                        st.session_state.role = "mensajeria" 
                        st.rerun() 
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
                            elif not u_profe: 
                                st.warning("Por favor selecciona tu nombre")
                            else: 
                                st.error("Contraseña incorrecta")
                
            # --- AQUÍ VA EL ACCESO PÚBLICO (Fuera de los formularios) ---
            st.markdown("---")
            st.markdown("<h4 style='text-align:center;'>Acceso Público</h4>", unsafe_allow_html=True)

            # Este botón activa la pantalla TV sin necesidad de contraseñas
            if st.button("📺 Abrir Pantalla Informativa", use_container_width=True):
                st.session_state.ver_pantalla_tv = True
                st.rerun()
            # -------------------------------------------------------------

    # ESTO SIEMPRE DEBE IR AL FINAL DEL BLOQUE DE LOGIN
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
        "Modo TV": {"icon": "📺", "roles": ["admin", "mensajeria"]}, # <-- El nuevo rol ahora tiene acceso
    }

    available_pages = [p for p, conf in PAGES_CONFIG.items() if st.session_state.role in conf["roles"]]
    default_page = "Mis Reservas" if st.session_state.role == 'profesor' else "Registrar"
    
    # --- AQUÍ ESTÁ LA SOLUCIÓN ---
    # Si la página por defecto ("Registrar") no está permitida para este usuario, 
    # entonces selecciona la primera que sí tenga disponible en su lista.
    if default_page not in available_pages:
        default_page = available_pages[0]
    # -----------------------------

    page = st.sidebar.radio("Navegación", available_pages, index=available_pages.index(default_page), format_func=lambda p: f"{PAGES_CONFIG[p]['icon']} {p}", label_visibility="collapsed")

    st.sidebar.markdown("---")

    if st.sidebar.button("🔄 Refrescar Pantalla", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    if st.sidebar.button("🚪 Cerrar Sesión", use_container_width=True):
        for key in st.session_state.keys(): 
            del st.session_state[key]
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
            base_url = "https://enlaces.streamlit.app/" 
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

# ==============================================================================
# 📺 PÁGINA: GESTIÓN DE TV Y MENSAJERÍA
# ==============================================================================
elif page == "Modo TV":

    # ==============================================================================
    # 📺 1. PANTALLA INFORMATIVA PÚBLICA (MODO KIOSCO)
    # ==============================================================================
    # Si apretaste "Iniciar Pantalla Pública", solo se dibuja esto y se frena.
    if st.session_state.get("ver_pantalla_tv", False):
        import json
        import os
        import datetime as dt
        from datetime import datetime as dt_datetime
        from streamlit_autorefresh import st_autorefresh
        import base64
        import pandas as pd

        now_dt = dt_datetime.now()

        # 🚨 RECEPTOR DE ALERTA ROJA (VÍA SUPABASE)
        try:
            # Busca la alerta de máxima prioridad (Ajusta a 1 si no pudiste usar 999 en Supabase)
            alertas_rojas = supabase.table("anuncios_urgentes").select("*").eq("is_active", True).eq("prioridad", 999).execute().data
            if alertas_rojas and len(alertas_rojas) > 0:
                alerta_activa = alertas_rojas[0]
                
                st.markdown(f"""
                    <style>
                    .alerta-fullscreen {{ position: fixed; top: 0; left: 0; width: 100vw; height: 100vh; background-color: rgba(220, 38, 38, 0.95); color: white; z-index: 999999; display: flex; flex-direction: column; justify-content: center; align-items: center; text-align: center; padding: 50px; backdrop-filter: blur(15px); }}
                    .alerta-fullscreen h1 {{ font-size: 5rem !important; font-weight: 900; margin-bottom: 20px; text-transform: uppercase; color: white; }}
                    .alerta-fullscreen p {{ font-size: 3rem; font-weight: 500; color: white; line-height: 1.2; }}
                    header {{visibility: hidden;}} .stApp {{overflow: hidden;}} [data-testid="stSidebar"] {{ display: none !important; }}
                    </style>
                    <div class="alerta-fullscreen">
                        <h1>⚠️ AVISO IMPORTANTE ⚠️</h1>
                        <p>{alerta_activa['descripcion']}</p>
                    </div>
                """, unsafe_allow_html=True)

                id_unica_alerta = str(alerta_activa["id"])
                if st.session_state.get("ultima_alerta_sonada") != id_unica_alerta:
                    st.audio("alarma.mp3", format="audio/mp3", autoplay=True)
                    st.session_state["ultima_alerta_sonada"] = id_unica_alerta
                    st.markdown("<style>audio { display: none !important; }</style>", unsafe_allow_html=True)

                st.stop() # Frena TODO aquí. No dibuja el resto del colegio.
        except Exception as e:
            pass

        # 📺 CÓDIGO NORMAL DE LA TV (Solo si NO hay alertas rojas)
        if "tv_scale" not in st.session_state: st.session_state.tv_scale = 100
        if "tv_profile" not in st.session_state: st.session_state.tv_profile = "Profesores / PIE"

        refresh_count = st_autorefresh(interval=20000, limit=None, key="tv_refresh_timer")
        
        ruta_logo = "logotv.png"
        logo_src_html = ""
        if os.path.exists(ruta_logo):
            with open(ruta_logo, "rb") as image_file:
                encoded_string = base64.b64encode(image_file.read()).decode()
                logo_src_html = f"<img src='data:image/png;base64,{encoded_string}' class='header-logo-img'/>"
        else:
            logo_src_html = "<i class='ph-fill ph-airplane-landing header-logo-fallback'></i>"

        hoy_str = now_dt.strftime("%Y-%m-%d")
        hora_actual = now_dt.strftime("%H:%M")

        dias_es = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
        meses_es = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
        fecha_es_formateada = f"{dias_es[now_dt.weekday()]}, {now_dt.day} de {meses_es[now_dt.month - 1]} de {now_dt.year}"
        escala = st.session_state.tv_scale / 100.0

        aesthetic_style = f"""            
        <style>
            @import url('https://unpkg.com/@phosphor-icons/web@2.1.1/src/fill/style.css');
            :root {{ --tv-scale: {escala}; }}
            .stApp {{ background-color: #0f172a; color: #f8fafc; font-family: 'Inter', sans-serif; }}
            [data-testid="stHeader"] {{ background: rgba(0,0,0,0); }}
            [data-testid="stToolbar"] {{ display: none; }}
            [data-testid="stSidebar"] {{ display: none !important; }}
            .tv-header-container {{ background: linear-gradient(135deg, #f8fafc 0%, #e2e8f0 100%); color: #0f172a; padding: 15px 25px 0 25px; border-radius: 20px; margin-bottom: 25px; border: 1px solid #cbd5e1; box-shadow: 0 4px 15px rgba(0,0,0,0.15); display: flex; flex-direction: column; align-items: center; }}
            .header-content-layout {{ display: flex; align-items: center; justify-content: space-between; width: 100%; padding-bottom: 15px; }}
            .header-logo-img {{ height: calc(85px * var(--tv-scale)); width: auto; display: block; }}
            .header-logo-fallback {{ font-size: calc(4rem * var(--tv-scale)); color: #64748b; line-height: 1; display: block; }}
            .header-info-group {{ display: flex; align-items: center; gap: 15px; font-size: calc(1.2rem * var(--tv-scale)); font-weight: 600; color: #1e293b; }}
            .header-divider {{ opacity: 0.3; font-weight: 300; font-size: calc(1.5rem * var(--tv-scale)); color: #94a3b8; }}
            .header-status {{ display: flex; align-items: center; color: #475569; }}
            .status-icon {{ margin-right: 8px; font-size: calc(1.3rem * var(--tv-scale)); color: #10b981; }}
            .progress-container {{ width: 100%; height: 6px; background-color: #cbd5e1; }}
            .progress-bar {{ height: 100%; background-color: #3b82f6; width: 0%; animation: loadBar 20s linear infinite; }}
            @keyframes loadBar {{ 0% {{ width: 0%; }} 100% {{ width: 100%; }} }}
            .tv-sub-header {{ color: #f8fafc; font-weight: 800; font-size: calc(1.6rem * var(--tv-scale)); margin-top: 5px; margin-bottom: 15px; text-transform: uppercase; letter-spacing: 1px; border-bottom: 2px solid #334155; padding-bottom: 10px;}}
            @keyframes cascadeIn {{ 0% {{ opacity: 0; transform: translateY(30px) scale(0.98); }} 100% {{ opacity: 1; transform: translateY(0) scale(1); }} }}
            @keyframes pulseAlert {{ 0% {{ box-shadow: 0 0 0 0 rgba(239, 68, 68, 0.4); }} 70% {{ box-shadow: 0 0 0 15px rgba(239, 68, 68, 0); }} 100% {{ box-shadow: 0 0 0 0 rgba(239, 68, 68, 0); }} }}
            .block-card {{ padding: 22px; border-radius: 16px; border-left: 8px solid; margin-bottom: 18px; background-color: white; box-shadow: 0 6px 15px rgba(0,0,0,0.1); }}
            .block-title-row {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px; width: 100%; }}
            .block-title-text {{ font-weight: 700; font-size: calc(1.35rem * var(--tv-scale)); text-transform: uppercase; }}
            .block-time-badge {{ display: flex; align-items: center; gap: 6px; background-color: #f1f5f9; color: #475569; padding: 6px 12px; border-radius: 8px; font-weight: 700; font-size: calc(1.1rem * var(--tv-scale)); border: 1px solid #e2e8f0; }}
            .block-info {{ font-size: calc(1.05rem * var(--tv-scale)); color: #475569; }}
            .block-info-row {{ display: flex; gap: 20px; align-items: center; margin-top: 12px; font-size: calc(1rem * var(--tv-scale));}}
            .block-info-item {{ display: flex; align-items: center; color: #64748b; }}
            .info-icon {{ font-size: calc(1.25rem * var(--tv-scale)); margin-right: 8px; }}
            .icon-profesor {{ color: #4ade80; }} 
            .icon-observaciones {{ color: #fbbf24; }} 
            .block-hora-pill {{ margin-top: 15px; font-weight: 600; color: #64748b; background: #f8fafc; display: inline-flex; align-items: center; padding: 6px 12px; border-radius: 10px; font-size: calc(0.9rem * var(--tv-scale)); border: 1px solid #e2e8f0;}}
            .icon-categoria {{ color: #818cf8; margin-right: 7px; font-size: calc(1.1rem * var(--tv-scale));}} 
            .announcements-container {{ background-color: white; border-radius: 20px; padding: 25px; border: 1px solid #e2e8f0; box-shadow: 0 4px 15px rgba(0,0,0,0.1); height: 100%; }}
            .announcement-card {{ padding: 20px; border-radius: 16px; border: 1px solid #e2e8f0; margin-bottom: 18px; border-left: 6px solid;}}
            .announcement-title {{ font-weight: 700; margin-bottom: 7px; font-size: calc(1.2rem * var(--tv-scale)); text-transform: uppercase;}}
            .announcement-desc {{ font-size: calc(1rem * var(--tv-scale)); color: #334155; line-height: 1.5; }}
        </style>
        """
        st.markdown(aesthetic_style, unsafe_allow_html=True) 

        modo_texto = f"MODO: {st.session_state.tv_profile.upper()}"
        clima_html = obtener_clima_vicuna()

        st.markdown(f"""
            <div class="tv-header-container">
                <div class="header-content-layout">
                    <div class="header-logo-section">{logo_src_html}</div>
                    <div class="header-info-group">
                        <div class="header-date">{fecha_es_formateada}</div>
                        <div class="header-divider">|</div>
                        <div class="header-weather">{clima_html}</div>
                        <div class="header-divider">|</div>
                        <div class="header-status"><i class="ph-fill ph-users-three status-icon"></i> {modo_texto}</div>
                    </div>
                </div>
                <div class="progress-container"><div class="progress-bar"></div></div>
            </div>
        """, unsafe_allow_html=True)
        
        col_main, col_ann = st.columns([2.5, 1], gap="large")
        
        with col_main:
            try:
                perfil_actual = st.session_state.tv_profile.upper()
                events_hoy_list = []
                
                url_guardada = st.session_state.get('url_calendario_tv', '')
                eventos_calendar = obtener_eventos_google_calendar(url_guardada)
                events_hoy_list.extend(eventos_calendar) 
                
                res_tv = supabase.table("eventos_tv").select("*").eq("fecha_evento", hoy_str).eq("is_active", True).execute()
                res_tv_hoy = res_tv.data if res_tv.data else []
                
                for ev in res_tv_hoy:
                    h_ini = str(ev.get("hora_inicio", ev.get("hora", "00:00")))[:5]
                    h_fin = str(ev.get("hora_fin", "23:59"))[:5]
                    if not h_fin or h_fin == "None" or h_fin.strip() == "": h_fin = "23:59"
                    if hora_actual > h_fin and h_fin != "23:59": continue
                        
                    disp_hora = f"{h_ini} - {h_fin}" if h_fin and h_fin != h_ini and h_fin != "23:59" else f"{h_ini}"
                    if not h_ini or h_ini == "None" or h_ini == "00:00": disp_hora = "TODO EL DÍA"
                    events_hoy_list.append({"hora_sort": h_ini if h_ini and h_ini != "None" and h_ini != "00:00" else "00:00", "display_hora": disp_hora, "titulo": ev.get("titulo", "Evento"), "descripcion": ev.get("descripcion", ""), "categoria": ev.get("categoria", "Evento General")})
                    
                if "PROFESORES" in perfil_actual or "PIE" in perfil_actual:
                    res_supabase = supabase.table("reservas").select("*, profesores(nombre), recursos(nombre), cursos(nombre)").eq("fecha", hoy_str).execute()
                    res_reservas_hoy = res_supabase.data if res_supabase.data else []
                    
                    for r in res_reservas_hoy:
                        h_ini = str(r.get("hora_inicio", r.get("hora", "00:00")))[:5]
                        h_fin = str(r.get("hora_fin", "23:59"))[:5]
                        if not h_fin or h_fin == "None" or h_fin.strip() == "": h_fin = "23:59"
                        if hora_actual > h_fin and h_fin != "23:59": continue
                            
                        disp_hora = f"{h_ini} - {h_fin}" if h_fin and h_fin != h_ini and h_fin != "23:59" else f"{h_ini}"
                        if not h_ini or h_ini == "None" or h_ini == "00:00": disp_hora = "RESERVA"
                        
                        prof = r.get("profesores", {}).get("nombre", "Docente") if r.get("profesores") else "Docente"
                        rec = r.get("recursos", {}).get("nombre", "Recurso") if r.get("recursos") else "Recurso"
                        curso = r.get("cursos", {}).get("nombre", "Curso") if r.get("cursos") else "Curso"
                        events_hoy_list.append({"hora_sort": h_ini if h_ini and h_ini != "None" and h_ini != "00:00" else "23:59", "display_hora": disp_hora, "titulo": f"{rec} ➔ {curso}", "profesor": prof, "observaciones": r.get("observaciones", ""), "categoria": "Uso Recurso / Sala"})
                
                events_hoy_list = sorted(events_hoy_list, key=lambda x: str(x.get("hora_sort", "99:99")))
                
                if not events_hoy_list:
                    st.markdown("<div class='tv-sub-header'>⏱️ Cronograma de Hoy</div>", unsafe_allow_html=True)
                    st.info(f"No hay eventos activos para el resto del día en este perfil.")
                else:
                    ITEMS_POR_PAGINA = 3
                    total_paginas = max(1, (len(events_hoy_list) + ITEMS_POR_PAGINA - 1) // ITEMS_POR_PAGINA)
                    pagina_actual = refresh_count % total_paginas 
                    inicio_idx = pagina_actual * ITEMS_POR_PAGINA
                    fin_idx = inicio_idx + ITEMS_POR_PAGINA
                    eventos_a_mostrar = events_hoy_list[inicio_idx:fin_idx]
                    
                    st.markdown(f"<div class='tv-sub-header'>⏱️ Cronograma de Hoy (Pág. {pagina_actual + 1}/{total_paginas})</div>", unsafe_allow_html=True)
                    paleta_colores = ["#0ea5e9", "#10b981", "#f59e0b", "#ec4899", "#8b5cf6", "#14b8a6"]
                    html_cronograma = ""
                    for i, item in enumerate(eventos_a_mostrar):
                        color_tema = "#6366f1" if item['categoria'] == "Evento Especial" else paleta_colores[i % len(paleta_colores)]
                        delay = i * 0.15 
                        info_row_html = ""
                        if item.get("profesor") or item.get("observaciones"):
                            info_row_html = "<div class='block-info-row'>"
                            if item.get("profesor"): info_row_html += f"<div class='block-info-item'><i class='ph-fill ph-user-graduate info-icon icon-profesor'></i> {item['profesor']}</div>"
                            if item.get("observaciones"): info_row_html += f"<div class='block-info-item'><i class='ph-fill ph-clipboard-text info-icon icon-observaciones'></i> {item['observaciones']}</div>"
                            info_row_html += "</div>"
                        
                        html_cronograma += (f"<div class='block-card' style='border-left-color: {color_tema}; animation: cascadeIn 0.8s forwards; animation-delay: {delay}s; opacity: 0;'>  <div class='block-title-row'>      <div class='block-title-text' style='color: {color_tema};'>{item['titulo']}</div>      <div class='block-time-badge'><i class='ph-fill ph-clock'></i> {item['display_hora']}</div>  </div>  <div class='block-info'>{item.get('descripcion', '')}</div>  {info_row_html}  <div class='block-hora-pill'><i class='ph-fill ph-tag icon-categoria'></i> {item['categoria']}</div></div>")
                    st.markdown(html_cronograma, unsafe_allow_html=True)
            except Exception as e:
                pass

        with col_ann:
            st.selectbox("👁️ Perfil de Visualización", ["Inspectoría / UTP", "Profesores / PIE", "Apoderados"], key="tv_profile")
            st.markdown("<div style='margin-bottom: 15px;'></div>", unsafe_allow_html=True)
            with st.expander("⚙️ Controles Extra", expanded=False):
                st.slider("🔍 Tamaño del texto (%)", min_value=50, max_value=250, value=st.session_state.tv_scale, step=5, key="tv_scale")
                st.markdown("<hr style='margin: 10px 0;'>", unsafe_allow_html=True)
                if st.button("🔙 Volver al Menú", use_container_width=True):
                    st.session_state.ver_pantalla_tv = False
                    st.rerun()
                import streamlit.components.v1 as components
                components.html("""<style>body { margin: 0; padding: 0; font-family: 'Inter', sans-serif; } button { width: 100%; height: 38px; background-color: #ffffff; border: 1px solid #cbd5e1; border-radius: 8px; color: #0f172a; font-size: 14px; font-weight: 500; cursor: pointer; display: flex; align-items: center; justify-content: center; gap: 8px; transition: all 0.2s; } button:hover { border-color: #94a3b8; background-color: #f8fafc; }</style><button onclick="const doc = window.parent.document; if (!doc.fullscreenElement) { doc.documentElement.requestFullscreen(); this.innerHTML = '🗗 Salir Pantalla Completa'; } else { doc.exitFullscreen(); this.innerHTML = '🔲 Pantalla Completa'; }">🔲 Pantalla Completa</button>""", height=40)

            try:
                ann_data = supabase.table("anuncios_urgentes").select("*").eq("is_active", True).execute().data
                active_ann = []
                for ann in ann_data:
                    if ann.get('prioridad') == 999: continue
                    try:
                        exp_dt = pd.to_datetime(ann['expiracion']).tz_localize(None)
                        if exp_dt > now_dt: active_ann.append(ann)
                    except: pass
                
                active_ann = sorted(active_ann, key=lambda x: x['prioridad'])
                if st.session_state.tv_profile == "Apoderados": titulo_panel, texto_vacio = "📰 Noticias y Comunicados", "No hay comunicados vigentes en este momento."
                else: titulo_panel, texto_vacio = "🚨 Avisos Urgentes y Alertas", "No hay avisos en este momento."

                html_anuncios = '<div class="announcements-container">'
                if not active_ann:
                    st.markdown(f"<div class='tv-sub-header'>{titulo_panel}</div>", unsafe_allow_html=True)
                    html_anuncios += f"<p style='color: #64748b; text-align:center; font-style:italic; margin-top: 10px;'>{texto_vacio}</p>"
                else:
                    ITEMS_POR_PAGINA_ANN = 3
                    total_paginas_ann = max(1, (len(active_ann) + ITEMS_POR_PAGINA_ANN - 1) // ITEMS_POR_PAGINA_ANN)
                    pagina_actual_ann = refresh_count % total_paginas_ann 
                    inicio_idx_ann, fin_idx_ann = pagina_actual_ann * ITEMS_POR_PAGINA_ANN, (pagina_actual_ann * ITEMS_POR_PAGINA_ANN) + ITEMS_POR_PAGINA_ANN
                    anuncios_a_mostrar = active_ann[inicio_idx_ann:fin_idx_ann]
                    
                    st.markdown(f"<div class='tv-sub-header'>{titulo_panel} (Pág. {pagina_actual_ann + 1}/{total_paginas_ann})</div>", unsafe_allow_html=True)

                    for i, ann in enumerate(anuncios_a_mostrar):
                        delay_ann = i * 0.15 
                        if st.session_state.tv_profile == "Apoderados": bg_color, border_color, title_color, desc_color, animacion_extra = "#f0f9ff", "#38bdf8", "#0369a1", "#334155", ""
                        else:
                            if ann['prioridad'] == 1: bg_color, border_color, title_color, desc_color, animacion_extra = "#fef2f2", "#ef4444", "#dc2626", "#334155", ", pulseAlert 2s infinite"
                            else: bg_color, border_color, title_color, desc_color, animacion_extra = "#fffbeb", "#f59e0b", "#d97706", "#334155", ""
                        html_anuncios += (f"<div class='announcement-card' style='border-left-color: {border_color}; background-color: {bg_color}; animation: cascadeIn 0.8s cubic-bezier(0.2, 0.8, 0.2, 1) forwards{animacion_extra}; animation-delay: {delay_ann}s; opacity: 0;'><div class='announcement-title' style='color: {title_color};'>{ann['titulo']}</div><div class='announcement-desc' style='color: {desc_color};'>{ann['descripcion']}</div></div>")
                html_anuncios += '</div>'
                st.markdown(html_anuncios, unsafe_allow_html=True)
            except Exception as e:
                pass

            # 🔔 GESTOR DE SONIDOS NORMALES
            try:
                ids_eventos_actuales = set([f"{e.get('titulo', '')}_{e.get('display_hora', '')}" for e in events_hoy_list])
                ids_avisos_actuales = set([str(a.get('id', '')) for a in active_ann])
                if "tv_inicializada" not in st.session_state:
                    st.session_state.tv_inicializada = True
                    st.session_state.memorias_eventos = ids_eventos_actuales
                    st.session_state.memorias_avisos = ids_avisos_actuales
                else:
                    nuevos_eventos = ids_eventos_actuales - st.session_state.memorias_eventos
                    nuevos_avisos = ids_avisos_actuales - st.session_state.memorias_avisos
                    if nuevos_avisos: st.audio("alarma.mp3", format="audio/mp3", autoplay=True)
                    elif nuevos_eventos: st.audio("evento.mp3", format="audio/mp3", autoplay=True)
                    if nuevos_avisos or nuevos_eventos: st.markdown("""<style>audio { display: none !important; }</style>""", unsafe_allow_html=True)
                    st.session_state.memorias_eventos = ids_eventos_actuales
                    st.session_state.memorias_avisos = ids_avisos_actuales
            except Exception as e:
                pass

        # Frenar todo al final de la TV para que no siga dibujando
        st.stop()


    # ==============================================================================
    # 💻 2. PANEL DE CONTROL INTERNO (MODO ADMINISTRADOR)
    # ==============================================================================
    # Este panel solo es visible si NO hiciste click en "Iniciar Pantalla"
    
    st.title("📺 Panel de Mensajería y Pantalla TV")
    st.markdown("Desde aquí puedes gestionar la pantalla pública del colegio, sincronizar calendarios y crear anuncios.")

    with st.container(border=True):
        col_launch1, col_launch2 = st.columns([1, 2])
        with col_launch1:
            st.subheader("🖥️ Proyección")
            st.write("Inicia el modo kiosco para proyectar en la TV.")
            if st.button("🚀 Iniciar Pantalla Pública", type="primary", use_container_width=True):
                st.session_state.ver_pantalla_tv = True
                st.rerun()
                
        with col_launch2:
            st.subheader("📅 Sincronización Google Calendar")
            st.write("Pega el enlace público (.ics) para sumar eventos automáticos.")
            url_cal = st.text_input("Enlace iCal (.ics)", value=st.session_state.get('url_calendario_tv', ''), label_visibility="collapsed")
            if st.button("Guardar Enlace y Sincronizar", type="secondary"):
                st.session_state['url_calendario_tv'] = url_cal
                try: obtener_eventos_google_calendar.clear() 
                except: pass
                st.success("✅ ¡Calendario sincronizado!")

    # --- SECCIÓN 2: ALERTA ROJA CENTRALIZADA (AHORA CON SUPABASE) ---
    st.divider()
    st.subheader("🚨 Mensaje Centralizado a Pantalla Completa")
    st.markdown("Usa esta función para **interrumpir la pantalla pública** con un aviso urgente.")

    # ====== 🔴 MAGIA: LA PANTALLA ROJA GIGANTE EN TU PROPIO PANEL ======
    if st.session_state.get("alerta_lanzada_panel", False):
        st.markdown("""
            <style>
            /* Tiñe todo el fondo de Streamlit de rojo emergencia */
            .stApp { background-color: rgba(220, 38, 38, 0.95) !important; background-image: none !important; }
            /* Oculta las barras laterales y menús para un efecto inmersivo */
            header { visibility: hidden; }
            [data-testid="stSidebar"] { display: none !important; }
            .texto-alerta { text-align: center; color: white; }
            </style>
        """, unsafe_allow_html=True)
        
        # Textos gigantes de alerta
        st.markdown("<h1 class='texto-alerta' style='font-size: 5rem; margin-top: 15vh; font-weight: 900;'>⚠️ AVISO IMPORTANTE ⚠️</h1>", unsafe_allow_html=True)
        st.markdown(f"<p class='texto-alerta' style='font-size: 3rem; margin-bottom: 50px;'>{st.session_state.get('mensaje_alerta_temporal', '')}</p>", unsafe_allow_html=True)
        
        # Botón para escapar de la pantalla roja sin recargar la página a la fuerza
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            if st.button("🛑 APAGAR ALERTA Y VOLVER AL PANEL NORMAL", use_container_width=True):
                st.session_state.alerta_lanzada_panel = False
                try:
                    # Apaga la alerta en la base de datos para que la TV del pasillo también vuelva a la normalidad
                    supabase.table("anuncios_urgentes").update({"is_active": False}).eq("prioridad", 999).execute()
                except:
                    pass
                st.rerun()
                
        # Reproducir sonido de alarma en tu panel
        st.audio("alarma.mp3", format="audio/mp3", autoplay=True)
        st.markdown("<style>audio { display: none !important; }</style>", unsafe_allow_html=True)
        
        st.stop() # Frena todo el resto del sistema para que solo se vea la alerta
    # ====================================================================

    with st.container(border=True):
        mensaje_alerta = st.text_area("Texto del mensaje:", placeholder="Ej: SE SUSPENDEN LAS CLASES DEL TURNO TARDE POR CORTE DE AGUA.", height=100)
        col_dur1, col_dur2 = st.columns(2)
        with col_dur1:
            tipo_duracion_alerta = st.radio("Definir duración por:", ["Minutos (Rápido)", "Fecha/Hora Exacta"], horizontal=True, key="rad_alerta")

        with col_dur2:
            import datetime as dt 
            from datetime import datetime as dt_datetime
            if tipo_duracion_alerta == "Minutos (Rápido)":
                minutos = st.number_input("¿Cuántos minutos durará en pantalla?", min_value=1, value=5, step=1)
                expiracion_alerta_roja = dt_datetime.now() + dt.timedelta(minutes=minutos)
                st.caption(f"Terminará aprox a las: {expiracion_alerta_roja.strftime('%H:%M:%S')}")
            else:
                f_alerta = st.date_input("Fecha límite:", key="f_alerta")
                h_alerta = st.time_input("Hora límite:", key="h_alerta")
                expiracion_alerta_roja = dt_datetime.combine(f_alerta, h_alerta)

        col_b1, col_b2 = st.columns(2)
        with col_b1:
            if st.button("🔴 LANZAR ALERTA EN PANTALLA AHORA", type="primary", use_container_width=True):
                if not mensaje_alerta.strip():
                    st.warning("Debes escribir un mensaje primero.")
                else:
                    try:
                        # Mandamos la alerta a Supabase (para que las TV de todo el colegio la reciban)
                        supabase.table("anuncios_urgentes").insert({
                            "titulo": "🚨 ALERTA ROJA", 
                            "descripcion": mensaje_alerta.strip(), 
                            "prioridad": 999, 
                            "expiracion": expiracion_alerta_roja.isoformat(), 
                            "is_active": True
                        }).execute()
                        
                        # === ACTIVAMOS LA PANTALLA ROJA EN TU PANEL ===
                        st.session_state.alerta_lanzada_panel = True
                        st.session_state.mensaje_alerta_temporal = mensaje_alerta.strip()
                        st.rerun() # Recargamos al instante para que la pantalla explote en rojo
                        
                    except Exception as e:
                        st.error(f"Error: {e}")

        with col_b2:
            if st.button("🛑 Cancelar Alerta / Limpiar Pantalla", use_container_width=True):
                try:
                    supabase.table("anuncios_urgentes").update({"is_active": False}).eq("prioridad", 999).execute()
                    st.info("Alerta cancelada. La pantalla de TV volverá a la normalidad en unos segundos.")
                except: 
                    pass
