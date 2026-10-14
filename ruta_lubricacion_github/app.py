# -*- coding: utf-8 -*-
"""
RUTA DE LUBRICACIÓN - Dashboard oscuro
Basado en el Excel maestro actualizado de Ruta de Lubricación.

La aplicación:
1. Lee HOJA 1 HORARIO para saber qué máquina corresponde a cada día/hora.
2. Lee las hojas de actividades y sus tiempos.
3. Relaciona la máquina del horario con el bloque de actividades correspondiente.
4. Permite marcar cada actividad individualmente.
5. Calcula el tiempo de las actividades seleccionadas.
6. Guarda el resultado en datos_ruta_lub.xlsx.
"""

from pathlib import Path
from datetime import datetime, date, timedelta
import re
import html

import pandas as pd
import streamlit as st
from openpyxl import load_workbook


# ============================================================
# CONFIGURACIÓN
# ============================================================

st.set_page_config(
    page_title="Ruta de Lubricación",
    page_icon="💧",
    layout="wide",
    initial_sidebar_state="collapsed",
)

BASE_DIR = Path(__file__).resolve().parent

# ============================================================
# ARCHIVO EXCEL MAESTRO
# ============================================================
# En tu repositorio el archivo aparece como:
#     ruta_lub_v2_oct_2026.xlsx
#
# Linux/GitHub distingue mayúsculas y minúsculas, por eso aquí
# se busca exactamente ese nombre y también algunas variantes.
CANDIDATOS_EXCEL = [
    # Excel nuevo adjuntado por el administrador (incluye el espacio antes de (2)).
    BASE_DIR / "RUTA LUB V2 OCT 2026 (2).xlsx",
    # Variante sin espacio que también se usó en versiones anteriores.
    BASE_DIR / "RUTA LUB V2 OCT 2026(2).xlsx",
    # Nombre canónico que se conservará en GitHub cuando se reemplace el Excel antiguo.
    BASE_DIR / "ruta_lub_v2_oct_2026.xlsx",
    BASE_DIR / "ruta_lub_v2_oct_2026.xlsm",
    BASE_DIR / "RUTA LUB V2 OCT 2026.xlsx",
    BASE_DIR / "RUTA LUB V2 OCT 2026(1).xlsx",
]

# También busca en la carpeta desde donde se ejecuta Streamlit.
CANDIDATOS_EXCEL += [
    Path.cwd() / "RUTA LUB V2 OCT 2026 (2).xlsx",
    Path.cwd() / "RUTA LUB V2 OCT 2026(2).xlsx",
    Path.cwd() / "ruta_lub_v2_oct_2026.xlsx",
    Path.cwd() / "ruta_lub_v2_oct_2026.xlsm",
    Path.cwd() / "RUTA LUB V2 OCT 2026.xlsx",
    Path.cwd() / "RUTA LUB V2 OCT 2026(1).xlsx",
]

ARCHIVO_RUTA = next((p for p in CANDIDATOS_EXCEL if p.exists()), None)

# Último intento: cualquier Excel de la carpeta que no sea el archivo
# de datos generado por la aplicación.
if ARCHIVO_RUTA is None:
    posibles = []
    for carpeta in {BASE_DIR, Path.cwd()}:
        if carpeta.exists():
            posibles.extend(carpeta.glob("*.xlsx"))
    posibles = [
        p for p in posibles
        if p.name.lower() != "datos_ruta_lub.xlsx"
        and not p.name.startswith("~$")
    ]
    if posibles:
        ARCHIVO_RUTA = posibles[0]

ARCHIVO_DATOS = BASE_DIR / "datos_ruta_lub.xlsx"


# ============================================================
# TEMA OSCURO
# ============================================================

st.markdown(
    """
<style>
#MainMenu, header, footer, [data-testid="stToolbar"] {
    visibility: hidden !important;
