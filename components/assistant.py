"""Punto de entrada reservado para extraer el asistente en V24.x.

El asistente actual permanece en schedule_app.py durante V24.0 porque usa
contexto de varias páginas todavía monolíticas. Moverlo ahora aumentaría
innecesariamente el riesgo de regresiones.
"""
