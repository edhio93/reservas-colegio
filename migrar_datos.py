import pandas as pd
from supabase import create_client, Client, ClientOptions
import warnings
import time

warnings.filterwarnings('ignore')

# 1. CONFIGURACIÓN CON MÁXIMA PACIENCIA
URL_SUPABASE = "https://zxzpaubemwpgvswvwjh.supabase.co"
CLAVE_SUPABASE = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inp4enBhdWJlbXdwd2d2c3d2d2poIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc3Mzg1NzMzMiwiZXhwIjoyMDg5NDMzMzMyfQ.CGWbTQprQaAhYruqlIkmMAMhx7EzD9hJ8QnJ7wCBxto"

# Aumentamos el tiempo de espera a 60 segundos (el estándar son 10)
opciones = ClientOptions(
    postgrest_client_timeout=60,
    storage_client_timeout=60
)

supabase: Client = create_client(URL_SUPABASE, CLAVE_SUPABASE, options=opciones)

def subir_con_reintento(tabla, dato, max_intentos=3):
    """Función para reintentar si hay un timeout"""
    for intento in range(max_intentos):
        try:
            return supabase.table(tabla).upsert(dato, on_conflict="nombre" if tabla != "reservas" else None).execute()
        except Exception as e:
            if "timeout" in str(e).lower() and intento < max_intentos - 1:
                print(f"   ⚠️ Reintentando {tabla} (intento {intento+2})...")
                time.sleep(2)
                continue
            raise e

def migrar_datos():
    try:
        print("📊 Leyendo Excel...")
        xls = pd.ExcelFile("Recursos.xlsx")
        df_res = pd.read_excel(xls, "Reservas").fillna("")
        df_prof = pd.read_excel(xls, "Profesores").fillna("")
        df_cur = pd.read_excel(xls, "Cursos").fillna("")
        df_rec = pd.read_excel(xls, "Recursos").fillna("")

        print("👤 Subiendo Profesores...")
        for p in df_prof.iloc[:, 0].dropna().unique():
            if str(p).strip(): subir_con_reintento("profesores", {"nombre": str(p).strip()})

        print("🏫 Subiendo Cursos...")
        for c in df_cur.iloc[:, 0].dropna().unique():
            if str(c).strip(): subir_con_reintento("cursos", {"nombre": str(c).strip()})

        print("🛠️ Subiendo Recursos...")
        for r in df_rec.iloc[:, 0].dropna().unique():
            if str(r).strip(): subir_con_reintento("recursos", {"nombre": str(r).strip()})

        print("🔗 Sincronizando IDs...")
        map_prof = {p["nombre"]: p["id"] for p in supabase.table("profesores").select("id, nombre").execute().data}
        map_cur = {c["nombre"]: c["id"] for c in supabase.table("cursos").select("id, nombre").execute().data}
        map_rec = {r["nombre"]: r["id"] for r in supabase.table("recursos").select("id, nombre").execute().data}

        print("📅 Subiendo Reservas...")
        exito = 0
        for _, row in df_res.iterrows():
            p_n, c_n, r_n = str(row['Profesor']).strip(), str(row['Curso']).strip(), str(row['Recurso']).strip()
            if p_n in map_prof and c_n in map_cur and r_n in map_rec:
                try:
                    supabase.table("reservas").insert({
                        "fecha": str(row['Fecha'])[:10],
                        "hora_inicio": str(row['Hora inicio']),
                        "hora_fin": str(row['Hora fin']),
                        "profesor_id": map_prof[p_n],
                        "curso_id": map_cur[c_n],
                        "recurso_id": map_rec[r_n],
                        "observaciones": str(row.get('Observaciones', ''))
                    }).execute()
                    exito += 1
                except: continue

        print(f"\n✨ ¡MIGRACIÓN EXITOSA! {exito} reservas en la nube.")

    except Exception as e:
        print(f"❌ Error crítico: {e}")

if __name__ == "__main__":
    migrar_datos()