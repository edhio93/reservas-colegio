import streamlit as st
import pandas as pd
from supabase import create_client, Client
import json
import plotly.express as px
from streamlit_option_menu import option_menu
from streamlit_sanitizer import Sanitizer
import hashlib
import time
import requests
from io import BytesIO
import datetime as dt

# --- NUEVA LIBRERÍA PARA WORD ---
try:
    from docx import Document
    from docx.shared import Pt, Inches
    from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
except ImportError:
    st.error("🚨 Falta la librería 'python-docx'. Instálala con: pip install python-docx")

st.set_page_config(layout="wide", page_title="Sistema de Reservas Enlaces")

# Instancia del sanitizador
html_sanitizer = Sanitizer()

# --- CONEXIÓN SUPABASE ---
try:
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    if not url or not key:
        raise ValueError("Credenciales vacías")
    supabase: Client = create_client(url, key)
except Exception as e:
    st.error(f"🚨 Error al conectar con Supabase: {e}")
    st.info("Revisa si has subido la app a Streamlit Cloud y has configurado los Secrets.")
    st.stop()

# Funciones de consulta a la base de datos
def run_query(table_name):
    try:
        query = supabase.table(table_name).select("*").execute()
        return query.data
    except Exception as e:
        st.error(f"Error consultando la tabla {table_name}: {e}")
        return []

# Cargar datos básicos
res_data = run_query("recursos")
# Filtramos recursos para que aparezcan solo los Activos en las reservas
if res_data:
    # Usamos DataFrame para filtrar fácilmente
    df_res_all = pd.DataFrame(res_data)
    if 'estado' in df_res_all.columns:
        res_data_activos = df_res_all[df_res_all['estado'] == 'Activo'].to_dict('records')
    else:
        # Si la columna no existe (antigua estructura), mostramos todos
        res_data_activos = res_data
else:
    res_data_activos = []

recursos_list = [r['Nombre'] for r in res_data_activos] if res_data_activos else []
profs_data = run_query("profesores")
profesores_list = [p['Nombre'] for p in profs_data] if profs_data else []
cursos_data = run_query("cursos")
cursos_list = [c['Nombre'] for c in cursos_data] if cursos_data else []

# Cargar reservas
reservas_query = supabase.table("reservas").select("*").execute()
reservas_data = reservas_query.data
if reservas_data:
    df = pd.DataFrame(reservas_data)
else:
    df = pd.DataFrame(columns=['id', 'Recurso', 'Profesor', 'Curso', 'Fecha', 'Hora inicio', 'Hora fin', 'Observaciones'])

# --- MENU LATERAL (ACTUALIZADO CON 'BAJA DE EQUIPOS') ---
with st.sidebar:
    st.image("https://images.vexels.com/content/135222/preview/university-building-simple-icon-135222.png", width=70) # Un logo simple
    page = option_menu("Menú", 
                       ["Dashboard", "Crear Reserva", "Semana", "Recursos", "Técnicos", "Baja de Equipos", "Administración"], 
                       icons=['speedometer2', 'calendar-plus', 'calendar-range', 'pc-display-horizontal', 'wrench-adjustable-circle', 'trash3', 'gear'], 
                       menu_icon="cast", default_index=0,
                       styles={
                           "container": {"padding": "0!important", "background-color": "#f0f2f6"},
                           "icon": {"color": "#4A90E2", "font-size": "1.1rem"}, 
                           "nav-link": {"font-size": "0.95rem", "text-align": "left", "margin":"3px", "--hover-color": "#e2e8f0"},
                           "nav-link-selected": {"background-color": "#4A90E2"},
                       })

    st.markdown("---")
    st.write("**Desarrollado para:**")
    st.write("Escuela Básica San Rafael")
    st.write("Dpto. de Enlaces")

# --- FUNCIONES AUXILIARES PARA EL DASHBOARD ---
def get_color_from_string(s):
    hash_val = int(hashlib.md5(s.encode('utf-8')).hexdigest(), 16)
    hue = hash_val % 360
    return f"hsl({hue}, 75%, 50%)"

# --- DASHBOARD ---
if page == "Dashboard":
    import tempfile
    import unicodedata
    import os
    st.header("📊 Panel de Supervisión General")

    COLOR_PRIMARIO = '#1E3A8A' 
    COLOR_SECUNDARIO = '#10B981' 

    def generar_pdf_profesional_con_graficos(df_datos):
        from fpdf import FPDF
        
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

            total_res = len(df_datos)
            df_datos['fecha_obj'] = pd.to_datetime(df_datos[C_FECHA]).dt.date
            hoy_res = len(df_datos[df_datos['fecha_obj'] == dt.date.today()])
            recurso_top = df_datos[C_REC].mode()[0] if not df_datos[C_REC].empty else "N/A"
            profesor_top = df_datos[C_PROF].mode()[0] if not df_datos[C_PROF].empty else "N/A"

            fecha_min = df_datos['fecha_obj'].min().strftime('%d/%m/%Y')
            fecha_max = df_datos['fecha_obj'].max().strftime('%d/%m/%Y')

            pdf.set_font('Arial', 'B', 16)
            pdf.set_text_color(30, 58, 138) 
            pdf.cell(0, 10, s('INFORME ESTADISTICO DE USO DE RECURSOS'), 0, 1, 'C')
            
            pdf.set_font('Arial', 'I', 11)
            pdf.set_text_color(100, 100, 100) 
            pdf.cell(0, 6, s(f'Periodo analizado: {fecha_min} al {fecha_max}'), 0, 1, 'C')
            pdf.ln(8)
            pdf.set_text_color(31, 41, 55) 

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

            pdf.section_title('2. Analisis Visual de Ocupacion')
            pdf.set_font('Arial', '', 11)
            pdf.multi_cell(0, 7, s('A continuacion se presentan los graficos detallados de uso de recursos y distribucion temporal de las reservas.'), 0, 'L')
            pdf.ln(5)

            # Usamos Matplotlib para el PDF
            import matplotlib.pyplot as plt

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
            
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_a:
                plt.savefig(tmp_a.name, format='png', dpi=150, bbox_inches='tight')
                img_path_a = tmp_a.name
            plt.close()
            
            current_y = pdf.get_y()
            pdf.image(img_path_a, x=25, y=current_y, w=160)
            os.remove(img_path_a)
            
            # Forzamos página 2 para pastel
            pdf.add_page() 
            
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
            
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_b:
                plt.savefig(tmp_b.name, format='png', dpi=150, bbox_inches='tight')
                img_path_b = tmp_b.name
            plt.close()
            
            current_y_b = pdf.get_y()
            pdf.image(img_path_b, x=55, y=current_y_b, w=100)
            os.remove(img_path_b)

        return pdf.output(dest='S').encode('latin-1')

    # BOTÓN DE DESCARGA PDF
    col_titulo, col_boton = st.columns([2, 1])
    with col_boton:
        if not df.empty:
            with st.spinner("Preparando reporte PDF..."):
                try:
                    pdf_data = generar_pdf_profesional_con_graficos(df)
                    st.download_button(
                        label="📄 Descargar Reporte PDF de Uso",
                        data=pdf_data,
                        file_name=f"Reporte_Gestión_Enlaces_{dt.date.today().strftime('%d_%m_%Y')}.pdf",
                        mime="application/pdf",
                        type="primary",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"⚠️ Error generando PDF: {e}")

    # MÉTRICAS Y GRÁFICOS EN PANTALLA (PLOTLY)
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

# --- CREAR RESERVA (Tu código original) ---
elif page == "Crear Reserva":
    st.header("➕ Nueva Reserva")
    
    if res_data_activos and profesores_list and cursos_list:
        with st.form("form_reserva", clear_on_submit=True):
            col1, col2 = st.columns(2)
            
            with col1:
                recurso = st.selectbox("Recurso a reservar", recursos_list, index=0)
                profesor = st.selectbox("Profesor", profesores_list, index=0)
                curso = st.selectbox("Curso / Nivel", cursos_list, index=0)
                fecha = st.date_input("Fecha", min_value=dt.date.today())

            with col2:
                base_horas = [
                    "08:00", "08:45", "09:30", "09:45", "10:30", "11:15", "11:30", "12:15",
                    "13:00", "14:00", "14:45", "15:30", "15:45", "16:30", "17:00", "17:30", "18:30"
                ]
                hora_i_str = st.selectbox("Hora inicio", base_horas, index=0)
                hora_f_str = st.selectbox("Hora fin", base_horas, index=2)
                
                observaciones = st.text_area("Observaciones (Ej. Taller, Reunión)", height=68)

            submit = st.form_submit_button("Confirmar Reserva", type="primary")

            if submit:
                try:
                    hora_i = dt.datetime.strptime(hora_i_str, "%H:%M").time()
                    hora_f = dt.datetime.strptime(hora_f_str, "%H:%M").time()
                    
                    # Sanitizar observaciones
                    obs_sanitizada = html_sanitizer.escape(observaciones).strip()

                    # Validaciones
                    if hora_f <= hora_i:
                        st.error("Error: La hora de fin debe ser posterior a la hora de inicio.")
                        st.stop()
                    
                    if fecha.weekday() >= 5: # 5=Sábado, 6=Domingo
                        st.warning("⚠️ Nota: Estás reservando en fin de semana.")

                    # Validación de Choques
                    if not df.empty:
                        choque = df[
                            (df['Recurso'] == recurso) & 
                            (df['Fecha'] == str(fecha)) & 
                            (
                                ((df['Hora inicio'].astype(str) <= str(hora_i)) & (df['Hora fin'].astype(str) > str(hora_i))) |
                                ((df['Hora inicio'].astype(str) < str(hora_f)) & (df['Hora fin'].astype(str) >= str(hora_f))) |
                                ((df['Hora inicio'].astype(str) >= str(hora_i)) & (df['Hora fin'].astype(str) < str(hora_f)))
                            )
                        ]
                        
                        if not choque.empty:
                            st.error(f"⚠️ El recurso {recurso} ya está reservado en ese horario.")
                            st.info("Revisa la pestaña 'Semana' para ver disponibilidad.")
                            st.stop()

                    data_res = {
                        "Recurso": recurso, "Profesor": profesor, "Curso": curso,
                        "Fecha": str(fecha), "Hora inicio": str(hora_i), "Hora fin": str(hora_f),
                        "Observaciones": obs_sanitizada
                    }
                    
                    supabase.table("reservas").insert(data_res).execute()
                    st.success("✅ ¡Reserva creada con éxito!")
                    st.balloons()
                    time.sleep(1)
                    st.rerun()

                except ValueError:
                    st.error("Error en el formato de hora.")
                except Exception as e:
                    st.error(f"Error al guardar: {e}")
    else:
        st.warning("⚠️ Requisito: Asegúrate de tener cargados Profesores, Cursos y Recursos Activos.")

# --- VISTA SEMANAL (Tu código original) ---
elif page == "Semana":
    st.header("🗓️ Vista Semanal")
    
    with st.container(border=True):
        st.write("🔍 **Filtros de Búsqueda Avanzados**")
        col_d, col_r, col_p, col_c = st.columns(4)
        
        selected_date = col_d.date_input("Semana del", value=dt.date.today(), format="DD/MM/YYYY")
        selected_recursos = col_r.multiselect("Filtrar Recursos", recursos_list, placeholder="Todos")
        selected_profesores = col_p.multiselect("Filtrar Profesores", profesores_list, placeholder="Todos")
        selected_cursos = col_c.multiselect("Filtrar Cursos", cursos_list, placeholder="Todos")

    start_of_week = selected_date - dt.timedelta(days=selected_date.weekday())
    week_days = [start_of_week + dt.timedelta(days=i) for i in range(5)]
    dias_es = {0: 'Lunes', 1: 'Martes', 2: 'Miércoles', 3: 'Jueves', 4: 'Viernes'}
    column_names = [f"{dias_es[d.weekday()]}<br><span style='font-size:0.8em; font-weight:normal;'>{d.strftime('%d/%m')}</span>" for d in week_days]
    
    # Horas base fijas y horas dinámicas de reservas
    base_horas = ["08:00 a 08:45", "08:45 a 09:30", "09:30 a 09:45", "09:45 a 10:30", "10:30 a 11:15", "11:15 a 11:30", "11:30 a 12:15", "12:15 a 13:00", "13:00 a 14:00", "14:00 a 14:45", "14:45 a 15:30", "15:30 a 15:45", "15:45 a 16:30", "16:30 a 17:00", "17:00 a 17:30", "17:30 a 18:30"]
    if not df.empty:
        df['bloque_hora'] = df['Hora inicio'].astype(str).str[:5] + " a " + df['Hora fin'].astype(str).str[:5]
        dynamic_horas = df['bloque_hora'].unique().tolist()
    else:
        dynamic_horas = []
    
    HORAS = sorted(list(set(base_horas + dynamic_horas)))
    schedule = pd.DataFrame(index=HORAS, columns=column_names).fillna('')

    if not df.empty:
        # Filtros
        mask = (pd.to_datetime(df['Fecha']).dt.date >= week_days[0]) & (pd.to_datetime(df['Fecha']).dt.date <= week_days[-1])
        if selected_recursos: mask &= df['Recurso'].isin(selected_recursos)
        if selected_profesores: mask &= df['Profesor'].isin(selected_profesores)
        if selected_cursos: mask &= df['Curso'].isin(selected_cursos)
        df_filtered = df[mask]

        for _, row in df_filtered.iterrows():
            f = pd.to_datetime(row['Fecha']).dt.date
            day_col = f"{dias_es[f.weekday()]}<br><span style='font-size:0.8em; font-weight:normal;'>{f.strftime('%d/%m')}</span>"
            bloque_actual = f"{str(row['Hora inicio'])[:5]} a {str(row['Hora fin'])[:5]}"
            
            if day_col in column_names:
                prof_color = get_color_from_string(str(row['Profesor']))
                observacion = str(row['Observaciones']) if pd.notna(row['Observaciones']) and str(row['Observaciones']).strip() != '' else ""
                icon = " 📌" if observacion else ""
                card_content = f"<div style='font-weight:bold; color:#1e293b; margin-bottom:4px;'>{row['Recurso']}{icon}</div><div style='color:#334155; margin-bottom:2px; font-size:0.95em;'>👨‍🏫 {row['Profesor']}</div><div style='color:#64748b; font-style:italic; font-size:0.9em;'>📚 {row['Curso']}</div>"
                safe_observacion = html_sanitizer.escape(observacion)
                card_html = f"<div style='background-color:#ffffff; padding:12px; margin-bottom:8px; border-radius:6px; border-left: 5px solid {prof_color}; box-shadow: 0 2px 5px rgba(0,0,0,0.08); font-size:0.85em; text-align:left; cursor:help;' title='{safe_observacion}'>{card_content}</div>"
                if schedule.at[bloque_actual, day_col] == '': schedule.at[bloque_actual, day_col] = card_html
                else: schedule.at[bloque_actual, day_col] += card_html

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

# --- RECURSOS Y QR (Tu código original INTÁCTO) ---
elif page == "Recursos":
    st.header("🖥️ Panel de Recursos")
    
    tab_reg, tab_gest_qr = st.tabs(["📝 Registro de Recursos", "📋 Gestión y Códigos QR"])
    
    with tab_reg:
        st.subheader("Registrar Nuevo Recurso")
        with st.form("form_nuevo_recurso", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
                nom_rec = st.text_input("Nombre del Recurso (Ej. Notebook 1, Sala 2)")
                cantidad_rec = st.number_input("Cantidad/Stock", min_value=1, value=1)
            with col2:
                obs_rec = st.text_input("Observaciones (Ej. Marca, Estado actual)")
                estado_inicial = st.selectbox("Estado Inicial", ["Activo", "Mantenimiento", "Inactivo"])
                
            submit_rec = st.form_submit_button("Registrar", type="primary")
            
            if submit_rec and nom_rec:
                # Sanitizar
                nom_s = html_sanitizer.escape(nom_rec).strip()
                obs_s = html_sanitizer.escape(obs_rec).strip()

                existing = supabase.table("recursos").select("Nombre").eq("Nombre", nom_s).execute().data
                if existing:
                    st.error(f"El recurso '{nom_s}' ya está registrado.")
                else:
                    supabase.table("recursos").insert({
                        "Nombre": nom_s, "Cantidad": cantidad_rec, "Observaciones": obs_s, "estado": estado_inicial
                    }).execute()
                    st.success(f"Recurso '{nom_s}' registrado.")
                    time.sleep(1)
                    st.rerun()

    with tab_gest_qr:
        # Código para generar QR y ver reportes
        st.subheader("📋 Gestión de Reportes Técnicos (Códigos QR)")
        st.write("Selecciona un recurso para generar su código QR único. Pégalo en el equipo físico para que los usuarios puedan reportar fallas.")
        
        if res_data_activos:
            col_sel, col_qr = st.columns([1, 2])
            with col_sel:
                # Usar selectbox con diccionario para obtener ID
                res_dict = {r['Nombre']: r['id'] for r in res_data_activos}
                res_qr_nom = st.selectbox("Recurso para QR", list(res_dict.keys()), index=0)
                res_qr_id = res_dict[res_qr_nom]
                
                # Generar URL del reporte público (¡Asegúrate de cambiar esta URL por la URL final de tu app publicada!)
                base_url = "https://escuela-san-rafael.streamlit.app/" # Ejemplo
                report_url = f"{base_url}?page=reporte&id={res_qr_id}"
                
                st.markdown(f"**URL del Reporte:**\n`{report_url}`")
                st.info("💡 Tip: Copia esta URL y úsala en un generador de QR externo o descarga el siguiente.")

            with col_qr:
                st.write("**Vista Previa del QR:**")
                # Usar API pública para generar QR rápido y mostrarlo
                qr_api_url = f"https://api.qrserver.com/v1/create-qr-code/?size=250x250&data={report_url}"
                try:
                    # Mostrar imagen directamente usando markdown HTML para centrar
                    st.markdown(f"<div style='text-align:center'><img src='{qr_api_url}' width='200'></div>", unsafe_allow_html=True)
                    
                    # Botón de descarga de la imagen QR
                    response_img = requests.get(qr_api_url)
                    if response_img.status_code == 200:
                        st.download_button(
                            label="⬇️ Descargar Imagen QR",
                            data=response_img.content,
                            file_name=f"QR_Enlaces_{res_qr_nom.replace(' ', '_')}.png",
                            mime="image/png",
                            use_container_width=True
                        )
                except Exception as e:
                    st.error(f"Error cargando QR: {e}")
        else:
            st.warning("No hay recursos activos para generar QR.")

# --- TÉCNICOS (Tu código con tickets, intácto) ---
elif page == "Técnicos":
    st.header("🛠️ Panel de Soporte Técnico (Tickets)")
    st.write("Gestiona los reportes de fallas ingresados vía QR. Atiende los tickets y lleva un historial de reparaciones.")

    mant_data = supabase.table("mantenimientos").select("*, recursos(Nombre)").order("fecha", desc=True).execute().data
    
    if mant_data:
        df_mant = pd.DataFrame(mant_data)
        # Extraer el nombre del recurso de la relación
        df_mant['Recurso'] = df_mant['recursos'].apply(lambda x: x['Nombre'] if isinstance(x, dict) and 'Nombre' in x else "Desconocido")
        
        # Prevenir errores si la columna notas_tecnico recién fue creada o está vacía
        if 'notas_tecnico' not in df_mant.columns:
            df_mant['notas_tecnico'] = ""
        else:
            df_mant['notas_tecnico'] = df_mant['notas_tecnico'].fillna("")
        
        # --- MÉTRICAS RÁPIDAS ---
        pendientes = len(df_mant[df_mant['estado'] == 'Reportado (Vía QR)'])
        en_revision = len(df_mant[df_mant['estado'] == 'En Revisión'])
        resueltos = len(df_mant[df_mant['estado'] == 'Resuelto'])
        
        c1, c2, c3 = st.columns(3)
        estilo_metrica = "background:white; border-radius:12px; padding: 15px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); border-left: 5px solid"
        
        with c1:
            st.markdown(f'<div style="{estilo_metrica} #EF4444;"><div style="color:#64748b; font-size:0.85em; font-weight:bold; letter-spacing:1px;">🔴 PENDIENTES</div><div style="font-size:2.2em; font-weight:900; color:#B91C1C; line-height:1.2;">{pendientes}</div></div>', unsafe_allow_html=True)
        with c2:
            st.markdown(f'<div style="{estilo_metrica} #F59E0B;"><div style="color:#64748b; font-size:0.85em; font-weight:bold; letter-spacing:1px;">🟡 EN REVISIÓN</div><div style="font-size:2.2em; font-weight:900; color:#D97706; line-height:1.2;">{en_revision}</div></div>', unsafe_allow_html=True)
        with c3:
            st.markdown(f'<div style="{estilo_metrica} #10B981;"><div style="color:#64748b; font-size:0.85em; font-weight:bold; letter-spacing:1px;">🟢 RESUELTOS</div><div style="font-size:2.2em; font-weight:900; color:#047857; line-height:1.2;">{resueltos}</div></div>', unsafe_allow_html=True)
            
        st.markdown("<br>", unsafe_allow_html=True)

        # --- TABLERO DE TAREAS (PESTAÑAS) ---
        t_pendientes, t_revision, t_resueltos = st.tabs(["🔴 Tickets Pendientes", "🟡 Tickets en Revisión", "🟢 Historial Resueltos"])
        
        def renderizar_tickets(df_filtrado, color_icon, estados_destino):
            if df_filtrado.empty:
                st.info("✨ No hay tickets en esta categoría actualmente.")
                return
            
            for _, row in df_filtrado.iterrows():
                with st.expander(f"{color_icon} Ticket #{row['id']} | {row['Recurso']} | Fecha: {row['fecha']}"):
                    st.markdown(f"**📝 Descripción de la Falla:**\n> {row['descripcion']}")
                    
                    if row['notas_tecnico']:
                        st.markdown(f"**🛠️ Historial/Notas Previas:**\n> {row['notas_tecnico']}")
                    
                    st.markdown("---")
                    col_a, col_b = st.columns([1, 2])
                    
                    with col_a:
                        nuevo_est = st.selectbox("Cambiar estado a:", estados_destino, key=f"est_{row['id']}")
                    with col_b:
                        nueva_nota = st.text_area("Agregar/Actualizar Notas de Reparación:", value=row['notas_tecnico'], key=f"not_{row['id']}", height=68)
                        
                    if st.button("💾 Guardar Cambios del Ticket", key=f"btn_{row['id']}", type="primary", use_container_width=True):
                        try:
                            supabase.table("mantenimientos").update({
                                "estado": nuevo_est,
                                "notas_tecnico": html_sanitizer.escape(nueva_nota).strip() # Sanitizar
                            }).eq("id", row['id']).execute()
                            
                            st.success(f"Ticket actualizado.")
                            time.sleep(1)
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error técnico al actualizar: {e}")

        with t_pendientes:
            renderizar_tickets(df_mant[df_mant['estado'] == 'Reportado (Vía QR)'], "🔴", ["En Revisión", "Resuelto", "Reportado (Vía QR)"])
        with t_revision:
            renderizar_tickets(df_mant[df_mant['estado'] == 'En Revisión'], "🟡", ["Resuelto", "En Revisión", "Reportado (Vía QR)"])
        with t_resueltos:
            renderizar_tickets(df_mant[df_mant['estado'] == 'Resuelto'], "🟢", ["Resuelto", "En Revisión", "Reportado (Vía QR)"])
    else:
        st.info("No hay reportes registrados.")

# ==============================================================================
# --- NUEVA SECCIÓN: BAJA DE EQUIPOS (CON INFORME WORD) ---
# ==============================================================================
elif page == "Baja de Equipos":
    st.header("🗑️ Baja de Equipos e Informes Técnicos")
    st.write("Selecciona un equipo inoperativo o antiguo para procesar su baja y generar automáticamente el informe Word reglamentario.")

    # 1. FUNCIÓN MAESTRA: GENERAR DOCX BASADO EN IMAGEN TEMPLATE
    def generar_docx_baja(datos):
        try:
            document = Document()
            
            # Estilos globales de fuente
            style = document.styles['Normal']
            font = style.font
            font.name = 'Arial'
            font.size = Pt(11)

            # --- CABECERA (Logo Placeholder y Título) ---
            # Para logo, se puede usar document.add_picture(path_logo, width=Inches(1))
            hdr = document.add_paragraph()
            hdr.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            r_hdr = hdr.add_run('[INSERTAR LOGO COLEGIO]\nDEPARTAMENTO DE ENLACES/INFORMÁTICA')
            r_hdr.font.size = Pt(10)
            r_hdr.font.color.rgb = docx.shared.RGBColor(100, 100, 100) # Gris

            titulo = document.add_paragraph()
            titulo.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            titulo.paragraph_format.space_before = Pt(12)
            titulo.paragraph_format.space_after = Pt(24)
            r_tit = titulo.add_run('INFORME TÉCNICO DE BAJA DE EQUIPO INFORMÁTICO')
            r_tit.bold = True
            r_tit.font.size = Pt(16)
            r_tit.font.color.rgb = docx.shared.RGBColor(30, 58, 138) # Azul institucional

            # --- SECCIÓN 1: IDENTIFICACIÓN DEL EQUIPO ---
            document.add_heading('1. IDENTIFICACIÓN DEL EQUIPO', level=1)
            
            table = document.add_table(rows=5, cols=2)
            table.style = 'Table Grid'
            
            # Formatear celdas y añadir datos
            def fill_cell(row, col, key, value, bold_key=True):
                cell = table.cell(row, col)
                cell.vertical_alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                p = cell.paragraphs[0]
                if bold_key:
                    run_k = p.add_run(f"{key}: ")
                    run_k.bold = True
                p.add_run(value)

            fill_cell(0, 0, 'Tipo de Equipo', datos['tipo'])
            fill_cell(1, 0, 'Marca / Modelo', datos['marca_modelo'])
            fill_cell(2, 0, 'N° de Serie / Inventario', datos['serie_inventario'])
            fill_cell(3, 0, 'Ubicación Habitual', datos['ubicacion'])
            fill_cell(4, 0, 'Fecha de Adquisición', datos['fecha_adq'])
            
            # La columna 2 se puede dejar vacía o usar para fotos (como en el original)
            cell_foto = table.cell(0, 1)
            cell_foto.merge(table.cell(4, 1)) # Combinar todas las filas
            p_foto = cell_foto.paragraphs[0]
            p_foto.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            p_foto.add_run('\n\n\n[ESPACIO PARA FOTO DEL EQUIPO]\n(Opcional)')
            
            document.add_paragraph().paragraph_format.space_after = Pt(12)

            # --- SECCIÓN 2: ESTADO TÉCNICO ACTUAL (Diagnosis) ---
            document.add_heading('2. ESTADO TÉCNICO ACTUAL / DIAGNÓSTICO', level=1)
            document.add_paragraph(datos['diagnosis'])
            document.add_paragraph().paragraph_format.space_after = Pt(12)

            # --- SECCIÓN 3: JUSTIFICACIÓN DE LA BAJA ---
            document.add_heading('3. JUSTIFICACIÓN DE LA BAJA', level=1)
            document.add_paragraph(datos['justificacion'])
            document.add_paragraph().paragraph_format.space_after = Pt(12)

            # --- SECCIÓN 4: RECOMENDACIÓN TÉCNICA ---
            document.add_heading('4. RECOMENDACIÓN TÉCNICA', level=1)
            document.add_paragraph(datos['recomendacion'])
            document.add_paragraph().paragraph_format.space_after = Pt(24)

            # --- PIE DE PÁGINA: FIRMAS Y FECHA ---
            fecha_p = document.add_paragraph()
            fecha_p.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
            fecha_p.add_run(f"San Rafael, {dt.date.today().strftime('%d de %B de %Y')}")
            fecha_p.paragraph_format.space_after = Pt(48)

            # Tabla de Firmas
            table_f = document.add_table(rows=1, cols=2)
            table_f.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            
            cell_f1 = table_f.cell(0, 0)
            p_f1 = cell_f1.paragraphs[0]
            p_f1.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            p_f1.add_run('__________________________\nFirma Técnico Responsable\n')
            p_f1.add_run(f"Nombre: {datos['tecnico_responsable']}").font.size = Pt(9)
            
            cell_f2 = table_f.cell(0, 1)
            p_f2 = cell_f2.paragraphs[0]
            p_f2.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            p_f2.add_run('__________________________\nFirma Dirección / Administración\n')
            p_f2.add_run('Nombre y Timbre').font.size = Pt(9)

            # Guardar en buffer
            docx_buf = BytesIO()
            document.save(docx_buf)
            docx_buf.seek(0)
            return docx_buf.read()

        except Exception as e:
            st.error(f"Error generando Word: {e}")
            return None

    # 2. PROCESO DE BAJA EN PANTALLA
    st.markdown("---")
    
    # Preparamos lista para selectbox (solo recursos activos)
    if res_data_activos:
        res_dict_baja = {r['Nombre']: r['id'] for r in res_data_activos}
        
        with st.container(border=True):
            st.subheader("🆕 Procesar Nueva Baja")
            
            col_sel, col_datos_m = st.columns([1, 2])
            
            with col_sel:
                recurso_baja_nom = st.selectbox("Equipo a dar de Baja:", list(res_dict_baja.keys()), key="sel_baja")
                recurso_baja_id = res_dict_baja[recurso_baja_nom]
                st.info(f"Seleccionado: {recurso_baja_nom}")

                # Estos campos son manuales porque no existen en tu DB actual
                marca_mod = st.text_input("Marca / Modelo", placeholder="Ej. Dell Vostro 3400")
                num_serie = st.text_input("Número de Serie / Inventario", placeholder="Ej. SN123456 / INV-001")
                ubicacion = st.text_input("Ubicación Habitual", placeholder="Ej. Sala de Profesores")
                fecha_adq = st.text_input("Fecha Adquisición (Aprox)", placeholder="Ej. 2018 o 15/05/2018")
            
            with col_datos_m:
                diagnosis = st.text_area("Diagnosis / Estado Técnico Actual", height=100, placeholder="Describe el estado técnico o daño detectado...")
                justificacion = st.text_area("Justificación de la Baja", height=100, placeholder="Por qué se debe dar de baja? Obsolescencia? Reparación muy cara?")
                recomendacion = st.text_area("Recomendación Técnica (Mínima)", height=100, placeholder="Baja definitiva? Reciclaje? Donación de partes?")
                tecnico = st.text_input("Nombre Técnico Responsable", placeholder="Ej. Juan Pérez")

            submit_baja = st.button("🚫 Procesar Baja y Generar Informe Word", type="primary", use_container_width=True)

            if submit_baja:
                if not diagnosis or not justificacion or not tecnico or not marca_mod:
                    st.warning("⚠️ Los campos Marca/Modelo, Diagnosis, Justificación y Técnico son obligatorios para el informe.")
                else:
                    with st.spinner("Procesando baja y creando historial..."):
                        # Sanitizar datos
                        datos_baja = {
                            "recurso_id": recurso_baja_id,
                            "recurso_nombre": html_sanitizer.escape(recurso_baja_nom).strip(),
                            "marca": html_sanitizer.escape(marca_mod).strip(),
                            "modelo": html_sanitizer.escape(marca_mod).strip(), # Usamos mismo input para marca/modelo
                            "serie": html_sanitizer.escape(num_serie).strip(),
                            "diagnosis": html_sanitizer.escape(diagnosis).strip(),
                            "justificacion": html_sanitizer.escape(justificacion).strip(),
                            "recomendacion": html_sanitizer.escape(recomendacion).strip(),
                            "tecnico_responsable": html_sanitizer.escape(tecnico).strip(),
                            "tipo": html_sanitizer.escape(recurso_baja_nom).strip() # Usamos nombre como tipo
                        }
                        # Datos extra para la tabla 'bajas_equipos'
                        datos_bd = datos_baja.copy()
                        del datos_bd['tipo'] # No está en BD

                        try:
                            # A. Insertar en tabla histórica 'bajas_equipos'
                            supabase.table("bajas_equipos").insert(datos_bd).execute()
                            
                            # B. Actualizar estado del equipo en 'recursos' a 'De Baja'
                            # Nota: Asegúrate que tu enum/estado en 'recursos' acepte 'De Baja'
                            supabase.table("recursos").update({"estado": "De Baja"}).eq("id", recurso_baja_id).execute()
                            
                            st.success(f"✅ ¡Se ha registrado la baja de '{recurso_baja_nom}'!")
                            
                            # C. GENERAR Y DESCARGAR EL WORD
                            docx_data = generar_docx_baja(datos_baja)
                            if docx_data:
                                st.download_button(
                                    label="⬇️ Descargar Informe Técnico Word (DOCX)",
                                    data=docx_data,
                                    file_name=f"Informe_Baja_{datos_baja['recurso_nombre'].replace(' ', '_')}_{dt.date.today()}.docx",
                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                    use_container_width=True
                                )
                                st.info("⬆️ Haz clic arriba para descargar el informe. Luego, actualiza la página.")
                        
                        except Exception as e:
                            st.error(f"Error en el proceso técnico: {e}")
                            if "bajas_equipos" in str(e).lower():
                                st.info("💡 Tip: Revisa si ejecutaste el SQL para crear la tabla 'bajas_equipos' en Supabase.")
    else:
        st.warning("No hay recursos activos para dar de baja.")

    # 3. HISTORIAL DE BAJAS
    st.markdown("---")
    st.subheader("📋 Historial de Equipos Dados de Baja")
    bajas_data = run_query("bajas_equipos")
    
    if bajas_data:
        df_bajas = pd.DataFrame(bajas_data)
        # Mostrar columnas relevantes
        df_display = df_bajas[['id', 'recurso_nombre', 'marca', 'serie', 'diagnosis', 'fecha_baja', 'tecnico_responsable']].copy()
        df_display.columns = ['Ticket ID', 'Recurso', 'Marca/Modelo', 'Serie', 'Diagnosis Breve', 'Fecha Baja', 'Técnico']
        st.dataframe(df_display, use_container_width=True, hide_index=True)
    else:
        st.info("No hay historial de bajas todavía.")

# --- ADMINISTRACIÓN (Tu código original) ---
elif page == "Administración":
    st.header("⚙️ Administración")
    
    tab_prof, tab_cur = st.tabs(["👨‍🏫 Profesores", "📚 Cursos"])
    
    with tab_prof:
        # Registro Prof
        with st.form("form_profesores", clear_on_submit=True):
            nom_prof_n = st.text_input("Nombre Completo del Profesor")
            submit_prof = st.form_submit_button("Registrar Profesor", type="primary")
            
            if submit_prof and nom_prof_n:
                nom_p_s = html_sanitizer.escape(nom_prof_n).strip()
                existing = supabase.table("profesores").select("Nombre").eq("Nombre", nom_p_s).execute().data
                if existing: st.error(f"El profesor '{nom_p_s}' ya existe.")
                else:
                    supabase.table("profesores").insert({"Nombre": nom_p_s}).execute()
                    st.success(f"Profesor '{nom_p_s}' registrado.")
                    time.sleep(1)
                    st.rerun()

        # Lista Prof
        st.markdown("---")
        if profs_data:
            df_prof = pd.DataFrame(profs_data)
            st.dataframe(df_prof[['id', 'Nombre']], use_container_width=True, hide_index=True)
        else: st.write("No hay profesores.")

    with tab_cur:
        # Registro Curso
        with st.form("form_cursos", clear_on_submit=True):
            nom_cur_n = st.text_input("Nombre del Curso (Ej. 1° Básico, 4° Medio)")
            submit_cur = st.form_submit_button("Registrar Curso", type="primary")
            
            if submit_cur and nom_cur_n:
                nom_c_s = html_sanitizer.escape(nom_cur_n).strip()
                existing = supabase.table("cursos").select("Nombre").eq("Nombre", nom_c_s).execute().data
                if existing: st.error(f"El curso '{nom_c_s}' ya existe.")
                else:
                    supabase.table("cursos").insert({"Nombre": nom_c_s}).execute()
                    st.success(f"Curso '{nom_c_s}' registrado.")
                    time.sleep(1)
                    st.rerun()

        # Lista Curso
        st.markdown("---")
        if cursos_data:
            df_cur = pd.DataFrame(cursos_data)
            st.dataframe(df_cur[['id', 'Nombre']], use_container_width=True, hide_index=True)
        else: st.write("No hay cursos.")
