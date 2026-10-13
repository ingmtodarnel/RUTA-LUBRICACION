# -*- coding: utf-8 -*-
"""
RUTA DE LUBRICACIÓN - Dashboard oscuro
Basado directamente en:
RUTA LUB V2 OCT 2026(1).xlsx

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
import io

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

CANDIDATOS_EXCEL = [
    BASE_DIR / "RUTA LUB V2 OCT 2026(2).xlsx",
    BASE_DIR / "ruta_lub_v2_oct_2026.xlsx",
    BASE_DIR / "ruta_lub_v2_oct_2026.xlsm",
    BASE_DIR / "RUTA LUB V2 OCT 2026(1).xlsx",
    BASE_DIR / "RUTA LUB V2 OCT 2026.xlsx",
]

CANDIDATOS_EXCEL += [
    Path.cwd() / "RUTA LUB V2 OCT 2026(2).xlsx",
    Path.cwd() / "ruta_lub_v2_oct_2026.xlsx",
    Path.cwd() / "RUTA LUB V2 OCT 2026(1).xlsx",
    Path.cwd() / "RUTA LUB V2 OCT 2026.xlsx",
]

ARCHIVO_RUTA = next(
    (p for p in CANDIDATOS_EXCEL if p.exists()),
    None,
)

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
}

.stApp {
    background:
        radial-gradient(circle at 10% 0%, rgba(0,140,255,.11), transparent 27%),
        radial-gradient(circle at 92% 10%, rgba(105,55,255,.08), transparent 25%),
        #03080f;
    color: #f5f8ff;
}

.main .block-container {
    max-width: 100% !important;
    padding: 0 !important;
}

p, span, label, div, input, textarea, button {
    font-family: "Segoe UI", Arial, sans-serif;
}

.tablet-header {
    background: linear-gradient(90deg, #061321, #081827);
    border-bottom: 2px solid #0796ff;
    padding: 17px 34px;
    margin-bottom: 20px;
    box-shadow: 0 8px 30px rgba(0,0,0,.35);
}

.header-title {
    color: #f8fbff;
    font-size: 30px;
    font-weight: 900;
}

.header-subtitle {
    color: #8dccff;
    font-size: 15px;
}

.header-meta {
    color: #f4f8ff;
    font-size: 15px;
    font-weight: 800;
    text-align: right;
}

.header-meta small {
    color: #85c9ff;
    font-weight: 500;
}

.page {
    padding: 0 30px 28px;
}

.welcome {
    background: linear-gradient(110deg,#081421,#06101a);
    border: 1px solid #17334c;
    border-radius: 15px;
    padding: 18px 24px;
    margin-bottom: 15px;
}

.welcome-title {
    color: #f7fbff;
    font-size: 29px;
    font-weight: 900;
}

.welcome-sub {
    color: #91cfff;
    font-size: 16px;
}

.panel {
    background: #07111b;
    border: 1px solid #17324b;
    border-radius: 14px;
    padding: 18px;
    margin-bottom: 15px;
}

.panel-title {
    color: #f5f9ff;
    font-size: 21px;
    font-weight: 900;
}

.panel-sub {
    color: #8eabc4;
    font-size: 13px;
}

.kpi {
    background: #07111b;
    border: 1px solid #17324b;
    border-radius: 14px;
    padding: 16px 18px;
    min-height: 120px;
}

.kpi-blue { border-color:#075ca7; }
.kpi-green { border-color:#087965; }
.kpi-orange { border-color:#8a4b0b; }
.kpi-purple { border-color:#5a2b9c; }

.kpi-label {
    color: #9bcbef;
    font-size: 14px;
    font-weight: 800;
}

.kpi-value {
    color: #ffffff;
    font-size: 30px;
    font-weight: 900;
    margin-top: 7px;
}

.kpi-small {
    color: #829bb2;
    font-size: 12px;
}

.menu-title {
    color: #f8fbff;
    font-size: 25px;
    font-weight: 900;
    margin: 20px 0 3px;
}

.menu-sub {
    color: #8ec8f5;
    margin-bottom: 12px;
}

.card {
    background: linear-gradient(145deg,#08131f,#050b12);
    border: 1px solid #183c59;
    border-radius: 14px;
    padding: 18px;
    min-height: 275px;
    text-align: center;
}

.card-blue { border-color:#087fe5; }
.card-green { border-color:#08a978; }
.card-orange { border-color:#ff7900; }
.card-purple { border-color:#7b42ec; }

.card-icon {
    width: 92px;
    height: 92px;
    margin: 0 auto 10px;
    border-radius: 50%;
    display:flex;
    align-items:center;
    justify-content:center;
    font-size:40px;
}

.blue { background:#087de9; }
.green { background:#08a978; }
.orange { background:#ff7900; }
.purple { background:#7139dc; }

.card-title {
    color:#fff;
    font-size:23px;
    font-weight:900;
}

.card-text {
    color:#a8c2d9;
    font-size:14
