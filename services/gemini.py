"""Cliente Gemini centralizado del Sistema CAV."""

import streamlit as st
import google.generativeai as genai


@st.cache_resource(show_spinner=False)
def get_gemini_model(api_key):
    genai.configure(api_key=api_key)
    return genai.GenerativeModel("gemini-2.5-flash")


def estado_gemini():
    """Comprueba que exista la clave y que el modelo pueda construirse."""
    try:
        api_key = st.secrets["GEMINI_API_KEY"]
        get_gemini_model(api_key)
        return True, "gemini-2.5-flash configurado"
    except KeyError:
        return False, "Falta GEMINI_API_KEY en Streamlit Secrets"
    except Exception as error:
        return False, str(error)


def probar_gemini():
    """Hace una llamada mínima para verificar conexión real con Gemini."""
    respuesta = consultar_gemini(
        "Responde únicamente con la palabra OK, sin puntuación adicional."
    )
    if str(respuesta).startswith("Error con la IA:"):
        return False, respuesta
    return True, str(respuesta).strip() or "Respuesta recibida"


def consultar_gemini(prompt):
    try:
        api_key = st.secrets["GEMINI_API_KEY"]
    except KeyError:
        return "Error con la IA: falta GEMINI_API_KEY en Streamlit Secrets."

    try:
        model = get_gemini_model(api_key)
        response = model.generate_content(prompt)
        return response.text
    except Exception as error:
        return f"Error con la IA: {error}"
