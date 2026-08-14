"""Cliente Gemini centralizado del Sistema CAV."""

import streamlit as st
import google.generativeai as genai


@st.cache_resource(show_spinner=False)
def get_gemini_model(api_key):
    genai.configure(api_key=api_key)
    return genai.GenerativeModel("gemini-2.5-flash")


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
