# Archivo: keep_alive.py
from playwright.sync_api import sync_playwright
import time

# REEMPLAZA ESTO CON LA URL REAL DE TU APP DE STREAMLIT
URL_STREAMLIT = "https://enlaces.streamlit.app/" 

def despertar_app():
    print(f"Iniciando el navegador fantasma para visitar: {URL_STREAMLIT}")
    
    with sync_playwright() as p:
        # Abrimos un navegador Chrome invisible
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        
        try:
            # Entramos a la página (le damos hasta 60 segundos para cargar si estaba dormida)
            page.goto(URL_STREAMLIT, timeout=60000)
            
            # ¡EL TRUCO MÁGICO! Esperamos 10 segundos en la página.
            # Esto da tiempo a que Streamlit cargue sus WebSockets y registre que es una visita "real".
            print("Página cargada. Esperando 10 segundos para registrar la actividad...")
            page.wait_for_timeout(10000) 
            
            print("¡Misión cumplida! La aplicación está despierta y activa.")
        except Exception as e:
            print(f"Hubo un error al visitar la página: {e}")
        finally:
            browser.close()

if __name__ == "__main__":
    despertar_app()
