
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
    font-size:14px;
    line-height:1.4;
    min-height:56px;
    margin:8px 0;
}

.stButton > button {
    min-height:44px !important;
    border-radius:10px !important;
    background:#0a1928 !important;
    border:1px solid #28506c !important;
    color:#f7fbff !important;
    font-weight:800 !important;
}

.stButton > button:hover {
    border-color:#078fff !important;
    background:#10283d !important;
}

button[kind="primary"] {
    background:linear-gradient(90deg,#087de9,#0b92f4) !important;
    border-color:#087de9 !important;
}

div[data-baseweb="select"] > div,
.stTextInput input,
.stTextArea textarea,
.stDateInput input,
.stNumberInput input {
    background:#07131f !important;
    color:#f5f9ff !important;
    border-color:#24445f !important;
}

[data-testid="stDataFrame"] {
    border:1px solid #183850 !important;
    border-radius:10px;
}

.activity {
    background:#08131e;
    border:1px solid #1b3b55;
    border-radius:12px;
    padding:12px 14px;
    margin:6px 0;
}

.activity:hover {
    border-color:#0b8ff2;
}

.activity-name {
    color:#eef7ff;
    font-size:14px;
    font-weight:700;
}

.activity-time {
    color:#68bfff;
    font-weight:900;
    white-space:nowrap;
}

.activity-special {
    color:#ffb45b;
    font-size:11px;
    margin-top:4px;
}

.footer {
    background:#07111b;
    border:1px solid #17324b;
    border-radius:14px;
    padding:15px;
    text-align:center;
    color:#79c8ff;
    font-weight:800;
    margin-top:18px;
}

.badge {
    display:inline-block;
    border-radius:999px;
    padding:5px 10px;
    background:#0b2032;
    color:#84caff;
    border:1px solid #1d4967;
    font-size:12px;
    font-weight:800;
}

.danger-note {
    background:#24150a;
    border:1px solid #73420d;
    color:#ffc77d;
    border-radius:10px;
    padding:10px 12px;
}

@media (max-width: 900px) {
    .page { padding:0 12px 20px; }
    .tablet-header { padding:14px 16px; }
    .header-title { font-size:23px; }
    .header-meta { display:none; }
}
</style>
""",
    unsafe_allow_html=True,
)


# ============================================================
# UTILIDADES
# ============================================================

def esc(x):
    return html.escape("" if x is None else str(x))


def normalizar_maquina(valor):
    if valor is None:
        return ""
    s = str(valor).strip().upper()
    s = s.replace(" ", "").replace("-", "")
    s = s.replace("_", "")
    return s


def extraer_numero_tf(s):
    m = re.search(r"TF\s*0*(\d+)", str(s).upper())
    return int(m.group(1)) if m else None


def maquina_equivalente(maquina):
    """
    Normaliza nombres del HORARIO para compararlos con las columnas
    de las hojas de actividades.
    """
    raw = str(maquina).strip().upper()
    n = normalizar_maquina(raw)

    aliases = {
        "EL-11": "EL11",
        "EL11": "EL11",
        "EL-14": "EL14",
        "EL14": "EL14",
        "EXTTF1": "EXT01",
        "EXTTF2": "EXT02",
        "EXTTF3": "EXT03",
        "TF04": "TF04",
        "TF05": "TF05",
        "TF07": "TF07",
        "TF10": "TF10",
        "TF11": "TF11",
        "TF12": "TF12",
        "TF13": "TF13",
        "TF14": "TF14",
        "TF15": "TF15",
        "TF16": "TF16",
        "TF17": "TF17",
        "TF20": "TF20",
        "TF22": "TF22+ML",
        "TF23": "TF23+ML",
        "TF24": "TF24",
        "TF26": "TF26",
        "TF31": "TF31",
        # TF32 fue retirado del cronograma por solicitud del usuario.
        "TF34": "TF34",
        "TF35": "TF35",
        "TF36": "TF36",
        "TF37": "TF37",
        "TF38": "TF38",
        "TF39": "TF39",
        "TF40": "TF40",
        "TF41": "TF41",
        "TF42": "TF42",
        "TF43": "TF43",
        "TF44": "TF44",
        "TF45": "TF45",
        "TF46": "TF46",
        "TF47": "TF47",
        "TF48": "TF48",
        "TF49": "TF49",
        "TF50": "TF50",
        "TF01": "TF01",
        "TF02": "TF02",
        "TF03": "TF03",
        # MP del horario corresponde a la columna Materias Primas del Excel.
        "MP": "MATERIASPRIMAS",
        "MATERIASPRIMAS": "MATERIASPRIMAS",
        # En el Excel actualizado S.T RIG y S.T ESP tienen columnas separadas;
        # no se deben mezclar sus actividades ni sus minutos.
        "S.TRIG": "STRIG",
        "S.TESP": "STESP",
        "STRIG": "STRIG",
        "STESP": "STESP",
        # El horario abrevia Materias Primas como MT PRIMAS.
        "MTPRIMAS": "MATERIASPRIMAS",
        "MATERIASPRIMAS": "MATERIASPRIMAS",
        # Encabezados de máquinas de ML Periferico que incluyen el día.
        "TF22SAB": "TF22+ML",
        "TF23SAB": "TF23+ML",
        "41VIERNES": "TF41",
        "42VIERNES": "TF42",
        # En HOJA 1 HORARIO las máquinas se llaman TAN; sus actividades
        # están en el segundo bloque de EXTRUSION con encabezados EE01, etc.
        # El Excel usa el número cero: EE01, EE02... (no la letra O).
        "TAN1": "EE01",
        "TAN2": "EE02",
        "TAN3": "EE03",
        "TAN4": "EE04",
        "TAN6": "EE06",
    }
    return aliases.get(raw.replace(" ", "").replace("-", ""), raw)


def normalizar_encabezado_maquina(valor):
    """Convierte encabezados del Excel en identificadores comparables al horario.

    Quita notas de encabezado como "(Viernes)" sin perder el nombre real de
    la máquina y normaliza abreviaturas usadas en ML Periferico.
    """
    if valor is None:
        return ""
    texto = str(valor).strip()
    # TF15 (Viernes) identifica TF15; el día de la ruta sigue gobernado por
    # HOJA 1 HORARIO, porque en el Excel el horario y esta nota no coinciden.
    texto = re.sub(r"\s*\([^)]*\)\s*$", "", texto)
    texto = texto.strip()
    return normalizar_maquina(maquina_equivalente(texto))


def actividad_especial_para_maquina(texto, maquina):
    """
    Solo aplica las restricciones explícitas que aparecen en el Excel.
    No inventa otras restricciones.
    """
    t = str(texto).upper()
    m = maquina_equivalente(maquina).upper()

    if "SOLO DAVIS" in t and m != "DAVIS":
        return False

    if "SOLO APLICA TF15" in t and m != "TF15":
        return False

    if "SOLO PARA TF41-42" in t:
        if m not in {"TF41", "TF42"}:
            return False

    return True


def frecuencia_permite(texto, maquina, fecha):
    """
    Respeta las frecuencias explícitas del texto:
    - Viernes
    - Sábado
    - Quincenal
    """
    t = str(texto).upper()
    dia = fecha.weekday()  # lunes=0 ... sábado=5

    if "VIERNES" in t and dia != 4:
        return False

    if "SÁBADO" in t or "SABADO" in t:
        if dia != 5:
            return False

    return True


# ============================================================
# LECTURA DEL EXCEL MAESTRO
# ============================================================

@st.cache_data(show_spinner=False)
def leer_excel_maestro(ruta, version_archivo):
    # version_archivo invalida la caché cuando el Excel cambia aunque conserve
    # el mismo nombre.
    wb = load_workbook(ruta, data_only=True)

    # ---------- HORARIO ----------
    ws = wb["HOJA 1 HORARIO"]

    horario = []

    # HOJA 1 HORARIO usa Excel 1-based:
    # Lunes=B/OT=C, Martes=F/OT=G, Miércoles=J/OT=K.
    # Jueves=B/OT=C, Viernes=F/OT=G, Sábado=J/OT=K.
    # La columna A contiene la hora solo en el bloque lunes-miércoles.
    grupos = [
        ("Lunes", 5, 16, 2, 3),
        ("Martes", 5, 16, 6, 7),
        ("Miércoles", 5, 16, 10, 11),
        ("Jueves", 22, 40, 2, 3),
        ("Viernes", 22, 40, 6, 7),
        ("Sábado", 22, 40, 10, 11),
    ]

    for dia, r1, r2, col_maquina, col_ot in grupos:
        for r in range(r1, r2 + 1):
            hora = ws.cell(r, 1).value
            maquina = ws.cell(r, col_maquina).value
            ot = ws.cell(r, col_ot).value

            if maquina is None:
                continue

            maquina = str(maquina).strip()

            if maquina in {"ALISTAMIENTO"}:
                continue

            # Correcciones solicitadas: TF32 se elimina del cronograma y
            # VOLCADOR es una actividad de Materias Primas, no una máquina.
            maquina_id = normalizar_maquina(maquina)
            if maquina_id in {"TF32", "VOLCADOR"}:
                continue

            if str(maquina).startswith("Nota:"):
                continue

            horario.append({
                "dia": dia,
                "hora": str(hora).strip() if hora else "",
                "maquina": maquina,
                "ot": "" if ot is None else str(ot),
            })

    # ---------- ACTIVIDADES ----------
    actividades = []

    def valor_no_vacio(valor):
        if valor is None:
            return False
        if isinstance(valor, str) and not valor.strip():
            return False
        return True

    def a_minutos(valor, fallback=None):
        """Convierte a minutos. Si la celda contiene una marca de texto, usa B como respaldo."""
        if valor_no_vacio(valor):
            try:
                return float(valor)
            except (TypeError, ValueError):
                pass
        if valor_no_vacio(fallback):
            try:
                return float(fallback)
            except (TypeError, ValueError):
                return None
        return None

    def leer_bloque(
        ws_name, header_row, start_row, end_row, start_col=3,
        dia_permitido=None,
    ):
        """Lee una tabla de actividades y sus máquinas.

        Regla de lectura del Excel:
        - Si el cuadro usa cantidades/minutos dentro de columnas de máquina,
          cada fila se asigna solo a las máquinas con una celda llena y usa ese
          valor como duración. Una celda vacía significa que esa fila no está
          asignada a esa máquina.
        - Si ninguna celda de máquina del cuadro contiene datos, el cuadro usa
          la columna B como tiempo común y las actividades se asignan a todos
          los encabezados de máquina del cuadro.
        """
        ws = wb[ws_name]
        headers = {}
        for c in range(start_col, ws.max_column + 1):
            value = ws.cell(header_row, c).value
            if value is None or not str(value).strip():
                continue
            maquina = normalizar_encabezado_maquina(value)
            if maquina:
                headers[c] = maquina

        if not headers:
            return

        # Detecta el formato del cuadro completo: valores por máquina o tiempo
        # compartido en la columna B. Esto evita asignar una tarea a una máquina
        # solo por aparecer en el encabezado cuando esa fila no tiene cantidad.
        usa_valores_por_maquina = any(
            valor_no_vacio(ws.cell(r, c).value)
            for r in range(start_row, end_row + 1)
            for c in headers
        )

        for r in range(start_row, end_row + 1):
            actividad = ws.cell(r, 1).value
            tiempo_compartido = ws.cell(r, 2).value
            if not valor_no_vacio(actividad):
                continue
            actividad_txt = str(actividad).strip()
            if not actividad_txt:
                continue

            if usa_valores_por_maquina:
                # Los minutos de cada columna definen la pertenencia y duración.
                for c, maquina in headers.items():
                    valor_maquina = ws.cell(r, c).value
                    if not valor_no_vacio(valor_maquina):
                        continue
                    minutos = a_minutos(valor_maquina, tiempo_compartido)
                    item = {
                        "hoja": ws_name,
                        "bloque": f"{ws_name} / bloque {header_row}",
                        "maquina_columna": maquina,
                        "actividad": actividad_txt,
                        "minutos": minutos,
                        "fila_excel": r,
                    }
                    if dia_permitido is not None:
                        item["dia_permitido"] = dia_permitido
                    actividades.append(item)
            else:
                # Este cuadro no tiene cantidades dentro de columnas de máquina:
                # aplica cada actividad a todas las máquinas del encabezado.
                minutos = a_minutos(tiempo_compartido)
                for maquina in headers.values():
                    item = {
                        "hoja": ws_name,
                        "bloque": f"{ws_name} / bloque {header_row}",
                        "maquina_columna": maquina,
                        "actividad": actividad_txt,
                        "minutos": minutos,
                        "fila_excel": r,
                    }
                    if dia_permitido is not None:
                        item["dia_permitido"] = dia_permitido
                    actividades.append(item)

    # Los cuadros sin cantidades por máquina usan el tiempo común de la columna B.
    leer_bloque("EXTRUSION", 4, 5, 20)
    leer_bloque("EXTRUSION", 24, 25, 39)
    leer_bloque("TF ESPUMADOS", 4, 5, 22)
    leer_bloque("TF ESPUMADOS", 26, 27, 45)
    leer_bloque("TF PRESION", 4, 5, 16)
    leer_bloque("TF PRESION", 20, 21, 34)

    # TF RIG tiene minutos dentro de las columnas de las máquinas; los espacios
    # vacíos de cada fila se respetan. El segundo cuadro indica revisión sábado.
    leer_bloque("TF RIG", 4, 5, 17)
    leer_bloque("TF RIG", 21, 22, 34, dia_permitido=5)

    # ML Periferico usa cantidades/minutos en las columnas de cada máquina.
    # Cada actividad y su tiempo se asignan según el valor de su columna. Las
    # columnas con Viernes se restringen al viernes; Materias Primas se usa
    # cuando MP aparece programado, y las demás columnas corresponden al sábado.
    ws = wb["ML Periferico"]
    headers = {}
    dias_por_maquina = {}
    for c in range(3, ws.max_column + 1):
        encabezado = ws.cell(3, c).value
        if not valor_no_vacio(encabezado):
            continue
        maquina = normalizar_encabezado_maquina(encabezado)
        if not maquina:
            continue
        headers[c] = maquina
        if maquina == "MATERIASPRIMAS":
            # La ruta puede programar MP en varios días; siempre toma las
            # actividades de la columna "Materias Primas".
            dias_por_maquina[maquina] = None
        else:
            dias_por_maquina[maquina] = 4 if "VIERNES" in str(encabezado).upper() else 5

    for r in range(4, ws.max_row + 1):
        actividad = ws.cell(r, 1).value
        tiempo_compartido = ws.cell(r, 2).value
        if not valor_no_vacio(actividad):
            continue
        actividad_txt = str(actividad).strip()
        if not actividad_txt:
            continue

        # Esta hoja sí contiene valores por máquina. Solo las columnas con
        # cantidad asignada reciben la actividad y el tiempo de esa celda.
        for c, maquina in headers.items():
            valor_maquina = ws.cell(r, c).value
            if not valor_no_vacio(valor_maquina):
                continue
            item = {
                "hoja": "ML Periferico",
                "bloque": f"ML Periferico / {maquina}",
                "maquina_columna": maquina,
                "actividad": actividad_txt,
                "minutos": a_minutos(valor_maquina, tiempo_compartido),
                "fila_excel": r,
                "dia_permitido": dias_por_maquina[maquina],
            }
            actividades.append(item)

    return horario, actividades


def cargar_maestro():
    """Carga el Excel maestro definido por el administrador.

    No se permite reemplazar/subir el horario maestro desde la interfaz,
    para que los usuarios de la aplicación no puedan cambiar la programación.
    """
    global ARCHIVO_RUTA

    if ARCHIVO_RUTA is None:
        st.error("No se encontró el Excel maestro de la ruta de lubricación.")
        st.warning(
            "La carga o sustitución del Excel maestro está deshabilitada para los usuarios. "
            "Un administrador debe colocar el archivo aprobado en la carpeta de la aplicación."
        )
        st.caption(
            f"Archivo esperado: RUTA LUB V2 OCT 2026.xlsx · Carpeta de la aplicación: {BASE_DIR}"
        )
        st.stop()

    version_archivo = ARCHIVO_RUTA.stat().st_mtime_ns if ARCHIVO_RUTA.exists() else 0
    return leer_excel_maestro(str(ARCHIVO_RUTA), version_archivo)


HORARIO, ACTIVIDADES = cargar_maestro()

# Diagnóstico discreto: permite comprobar en Streamlit qué Excel fue cargado.
if ARCHIVO_RUTA is not None:
    st.session_state["_excel_cargado"] = ARCHIVO_RUTA.name

# Diagnóstico visible para comprobar que la app está usando el Excel maestro.
with st.expander("🔎 Excel maestro cargado", expanded=False):
    st.write(f"**Archivo:** {ARCHIVO_RUTA.name if ARCHIVO_RUTA else 'No encontrado'}")
    if ARCHIVO_RUTA is not None:
        st.write(f"**Ruta:** `{ARCHIVO_RUTA}`")
        st.write(f"**Máquinas en HOJA 1 HORARIO:** {len(HORARIO)}")
        st.write(f"**Actividades leídas:** {len(ACTIVIDADES)}")
        st.write("**Hojas esperadas:** HOJA 1 HORARIO, EXTRUSION, TF ESPUMADOS, TF PRESION, TF RIG, ML Periferico")
        st.write("**Equivalencias TAN:** TAN1→EE01, TAN2→EE02, TAN3→EE03, TAN4→EE04, TAN6→EE06")



# ============================================================
# GUARDADO DE REGISTROS
# ============================================================

COLUMNAS_DATOS = [
    "fecha",
    "dia",
    "hora",
    "maquina",
    "hoja",
    "bloque",
    "fila_excel",
    "actividad",
    "minutos",
    "estado",
    "ejecutor",
    "ot",
    "recibio",
    "observacion",
    "fecha_hora_guardado",
]


def cargar_datos():
    if not ARCHIVO_DATOS.exists():
        return pd.DataFrame(columns=COLUMNAS_DATOS)

    try:
        df = pd.read_excel(ARCHIVO_DATOS)
    except Exception:
        return pd.DataFrame(columns=COLUMNAS_DATOS)

    for c in COLUMNAS_DATOS:
        if c not in df.columns:
            df[c] = ""

    return df[COLUMNAS_DATOS]


def guardar_datos(df):
    df.to_excel(ARCHIVO_DATOS, index=False)


if "datos" not in st.session_state:
    st.session_state.datos = cargar_datos()

if "pagina" not in st.session_state:
    st.session_state.pagina = "home"

if "fecha_trabajo" not in st.session_state:
    st.session_state.fecha_trabajo = date.today()

if "ejecutor" not in st.session_state:
    st.session_state.ejecutor = ""

if "maquina_seleccionada" not in st.session_state:
    st.session_state.maquina_seleccionada = None


def ir(pagina, maquina=None):
    st.session_state.pagina = pagina
    if maquina is not None:
        st.session_state.maquina_seleccionada = maquina
    st.rerun()


def upsert_actividad(registro):
    df = st.session_state.datos.copy()

    if df.empty:
        df = pd.DataFrame(columns=COLUMNAS_DATOS)

    mask = (
        df["fecha"].astype(str).eq(str(registro["fecha"]))
        & df["maquina"].astype(str).eq(str(registro["maquina"]))
        & df["actividad"].astype(str).eq(str(registro["actividad"]))
        & df["hoja"].astype(str).eq(str(registro["hoja"]))
    )

    if mask.any():
        idx = df.index[mask][0]
        for k, v in registro.items():
            df.at[idx, k] = v
    else:
        df = pd.concat([df, pd.DataFrame([registro])], ignore_index=True)

    st.session_state.datos = df


# ============================================================
# RELACIÓN MÁQUINA -> ACTIVIDADES
# ============================================================

def buscar_actividades(maquina, fecha):
    m = maquina_equivalente(maquina)
    resultado = []

    # MP se normaliza a MATERIASPRIMAS y recibe solamente las actividades
    # asignadas a esa columna del cuadro ML Periferico. VOLCADOR permanece
    # dentro de la descripción de una actividad de Materias Primas, nunca como
    # máquina independiente del horario.

    for a in ACTIVIDADES:
        dia_permitido = a.get("dia_permitido")
        if dia_permitido is not None and fecha.weekday() != dia_permitido:
            continue

        col = normalizar_maquina(a["maquina_columna"])
        mm = normalizar_maquina(m)

        # EL14 debe llevar exactamente la misma ruta base de EXTRUSION
        # que EL11. En esa hoja se toma como fuente la columna EL11, incluso
        # si en el futuro alguien añade otra columna EL14, para evitar sumar
        # actividades diferentes y romper la equivalencia solicitada.
        if a["hoja"] == "EXTRUSION" and mm == "EL14":
            coincide = col == "EL11"
        elif mm == "TF46" and col == "TF45":
            # TF46 utiliza las mismas actividades y tiempos definidos para TF45.
            coincide = True
        else:
            # En el resto de las hojas se usa la coincidencia normal.
            coincide = col == mm

        # El horario TF07 usa una columna TF06/07.
        if not coincide and mm == "TF07" and col == "TF06/07":
            coincide = True

        # TF22 y TF23 pueden aparecer como TF22/TF23 o con el sufijo ML.
        if not coincide and mm in {"TF22", "TF22ML", "TF22+ML"} and col == "TF22+ML":
            coincide = True
        if not coincide and mm in {"TF23", "TF23ML", "TF23+ML"} and col == "TF23+ML":
            coincide = True
        # Mantener compatibilidad con archivos antiguos que usaban un encabezado
        # compartido de Sistema de Transporte. En el Excel actual STRIG y STESP
        # tienen columnas separadas y ya se comparan de forma independiente.
        if not coincide and mm in {"STRIG", "STESP", "SISTTRANSPORTE"} and col == "SISTTRANSPORTE":
            coincide = True

        # TF20 aparece en ambos bloques de TF RIG. Se conserva el bloque
        # correspondiente a la fila de actividades sin fusionarlos.
        if not coincide and mm == "TF20" and col == "TF20":
            coincide = True

        if coincide:
            if actividad_especial_para_maquina(a["actividad"], maquina) and frecuencia_permite(
                a["actividad"], maquina, fecha
            ):
                resultado.append(a)

    # Evita duplicar exactamente la misma actividad proveniente de bloques
    # repetidos con el mismo nombre y tiempo.
    unicos = []
    vistos = set()
    for a in resultado:
        key = (a["hoja"], a["fila_excel"], a["actividad"], a["minutos"])
        if key not in vistos:
            vistos.add(key)
            unicos.append(a)

    return unicos


def obtener_ruta_del_dia(fecha):
    dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"]
    dia = dias[fecha.weekday()]
    return [x for x in HORARIO if x["dia"] == dia]


# ============================================================
# CABECERA
# ============================================================

def header():
    f = st.session_state.fecha_trabajo
    lunes = f - timedelta(days=f.weekday())
    domingo = lunes + timedelta(days=6)
    semana = f.isocalendar().week

    st.markdown(
        f"""
        <div class="tablet-header">
            <div style="display:flex;align-items:center;justify-content:space-between;gap:20px;">
                <div style="display:flex;align-items:center;gap:15px;">
                    <div style="font-size:48px;">💧</div>
                    <div>
                        <div class="header-title">Ruta de Lubricación</div>
                        <div class="header-subtitle">
                            Registro y control semanal de lubricación
                        </div>
                    </div>
                </div>
                <div class="header-meta">
                    📅 Semana {semana}<br>
                    <small>
                        {lunes.strftime('%d/%m/%Y')} - {domingo.strftime('%d/%m/%Y')}
                    </small>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ============================================================
# HOME
# ============================================================

def home():
    header()
    f = st.session_state.fecha_trabajo
    ruta = obtener_ruta_del_dia(f)
    datos = st.session_state.datos

    fecha_str = f.strftime("%Y-%m-%d")
    registros_hoy = datos[datos["fecha"].astype(str) == fecha_str]
    completadas = len(registros_hoy[registros_hoy["estado"] == "Completada"])

    total_hoy = 0
    for item in ruta:
        total_hoy += len(buscar_actividades(item["maquina"], f))

    pct = round(completadas / total_hoy * 100) if total_hoy else 0

    st.markdown('<div class="page">', unsafe_allow_html=True)

    st.markdown(
        """
        <div class="welcome">
            <div style="display:flex;justify-content:space-between;align-items:center;">
                <div>
                    <div class="welcome-title">⚙️ Bienvenido</div>
                    <div class="welcome-sub">
                        Selecciona una opción para comenzar
                    </div>
                </div>
                <div style="color:#0795ff;font-size:18px;font-weight:900;">
                    🏭 Ruta de Lubricación
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    k1, k2, k3, k4 = st.columns(4)

    with k1:
        st.markdown(
            f"""
            <div class="kpi kpi-blue">
                <div class="kpi-label">📋 Actividades de hoy</div>
                <div class="kpi-value">{total_hoy}</div>
                <div class="kpi-small">según el Excel maestro</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with k2:
        st.markdown(
            f"""
            <div class="kpi kpi-green">
                <div class="kpi-label">✓ Completadas</div>
                <div class="kpi-value">{completadas}</div>
                <div class="kpi-small">{pct}% de las actividades de hoy</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with k3:
        minutos = 0
        for item in ruta:
            for a in buscar_actividades(item["maquina"], f):
                if a["minutos"] is not None:
                    minutos += a["minutos"]

        st.markdown(
            f"""
            <div class="kpi kpi-orange">
                <div class="kpi-label">⏱ Tiempo programado</div>
                <div class="kpi-value">{int(minutos)} min</div>
                <div class="kpi-small">suma de tiempos del Excel</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with k4:
        st.markdown(
            f"""
            <div class="kpi kpi-purple">
                <div class="kpi-label">🏭 Máquinas</div>
                <div class="kpi-value">{len(ruta)}</div>
                <div class="kpi-small">programadas para {f.strftime('%d/%m/%Y')}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown(
        """
        <div class="menu-title">▦ &nbsp; Menú Principal</div>
        <div class="menu-sub">Accede a las diferentes secciones de la aplicación</div>
        """,
        unsafe_allow_html=True,
    )

    c1, c2, c3, c4 = st.columns(4)

    cards = [
        (c1, "card-blue", "blue", "📅", "Ruta Diaria",
         "Las máquinas y horarios salen directamente de HOJA 1 HORARIO.",
         "ABRIR RUTA", "ruta"),
        (c2, "card-green", "green", "✓", "Actividades",
         "Marca cada actividad de la máquina y conserva su tiempo del Excel.",
         "ABRIR ACTIVIDADES", "actividad"),
        (c3, "card-orange", "orange", "📊", "Cronograma",
         "Consulta máquinas, actividades y tiempos programados.",
         "ABRIR CRONOGRAMA", "cronograma"),
        (c4, "card-purple", "purple", "◷", "Historial",
         "Consulta lo que ya fue registrado y exporta los datos.",
         "ABRIR HISTORIAL", "historial"),
    ]

    for col, border, icon_cls, icon, title, desc, btn, page in cards:
        with col:
            st.markdown(
                f"""
                <div class="card {border}">
                    <div class="card-icon {icon_cls}">{icon}</div>
                    <div class="card-title">{title}</div>
                    <div class="card-text">{desc}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            if st.button(btn, use_container_width=True, type="primary", key=f"home_{page}"):
                ir(page)

    st.markdown('<div class="footer">💧 &nbsp; La lubricación es vida para tus equipos &nbsp; 💧</div>',
                unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    c1, c2 = st.columns([1, 2])
    with c1:
        nueva_fecha = st.date_input(
            "Fecha de trabajo",
            value=f,
            key="fecha_home",
        )
        if nueva_fecha != st.session_state.fecha_trabajo:
            st.session_state.fecha_trabajo = nueva_fecha
            st.rerun()

    with c2:
        st.text_input(
            "Lubricador / ejecutor",
            key="ejecutor",
            placeholder="Nombre de quien realiza la ruta",
        )

    st.markdown("</div>", unsafe_allow_html=True)


# ============================================================
# RUTA DIARIA
# ============================================================

def ruta_diaria():
    header()

    f = st.session_state.fecha_trabajo
    dia = ["Lunes","Martes","Miércoles","Jueves","Viernes","Sábado"][f.weekday()]
    ruta = obtener_ruta_del_dia(f)

    st.markdown('<div class="page">', unsafe_allow_html=True)

    if st.button("← INICIO", key="ruta_inicio"):
        ir("home")

    st.markdown(
        f"""
        <div class="panel">
            <div class="panel-title">📅 Ruta del {dia}</div>
            <div class="panel-sub">
                {f.strftime('%d/%m/%Y')} · Las máquinas vienen de HOJA 1 HORARIO.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if not ruta:
        st.warning("No hay máquinas programadas para este día en el Excel.")
        st.markdown("</div>", unsafe_allow_html=True)
        return

    for i, item in enumerate(ruta):
        maquina = item["maquina"]
        acts = buscar_actividades(maquina, f)
        minutos = sum(a["minutos"] or 0 for a in acts)

        c1, c2, c3, c4 = st.columns([1.5, 2.3, 1, 1])

        with c1:
            st.markdown(f"### {item['hora']}")

        with c2:
            st.markdown(f"**🏭 {maquina}**")
            st.caption(f"{len(acts)} actividades · {int(minutos)} min")

        with c3:
            st.markdown(f'<span class="badge">{item["ot"] or "SIN OT"}</span>',
                        unsafe_allow_html=True)

        with c4:
            if st.button("INICIAR", key=f"iniciar_{i}_{maquina}", type="primary"):
                ir("actividad", maquina)

        st.divider()

    st.markdown("</div>", unsafe_allow_html=True)


# ============================================================
# ACTIVIDADES DE UNA MÁQUINA
# ============================================================

def actividad_maquina():
    header()

    f = st.session_state.fecha_trabajo
    maquina = st.session_state.maquina_seleccionada

    if not maquina:
        ir("ruta")

    acts = buscar_actividades(maquina, f)

    st.markdown('<div class="page">', unsafe_allow_html=True)

    c1, c2 = st.columns([1, 5])
    with c1:
        if st.button("← RUTA", key="act_ruta"):
            ir("ruta")
    with c2:
        st.markdown(
            f"""
            <div class="panel">
                <div class="panel-title">🏭 {esc(maquina)}</div>
                <div class="panel-sub">
                    {f.strftime('%A %d/%m/%Y')} · {len(acts)} actividades encontradas en el Excel
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    if not acts:
        st.markdown(
            """
            <div class="danger-note">
                Esta máquina aparece en el horario, pero el Excel no contiene
                un bloque de actividades identificado para ella.
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.markdown("</div>", unsafe_allow_html=True)
        return

    # OT de la máquina en el horario.
    ot = ""
    for x in obtener_ruta_del_dia(f):
        if x["maquina"] == maquina:
            ot = x["ot"]
            break

    st.text_input("OT", value=ot, key=f"ot_{maquina}")
    st.text_input("Recibió", key=f"recibio_{maquina}")
    st.text_area("Observación general", key=f"obs_{maquina}")

    # Leer estados actuales.
    datos = st.session_state.datos
    fecha_str = f.strftime("%Y-%m-%d")

    completadas_previas = set()
    if not datos.empty:
        filtro = (
            datos["fecha"].astype(str).eq(fecha_str)
            & datos["maquina"].astype(str).eq(str(maquina))
            & datos["estado"].astype(str).eq("Completada")
        )
        for _, row in datos[filtro].iterrows():
            completadas_previas.add(
                (str(row["hoja"]), str(row["fila_excel"]), str(row["actividad"]))
            )

    # Marcar todas.
    key_all = f"all_{fecha_str}_{maquina}"
    marcar_todas = st.checkbox(
        "✓ Marcar todas las actividades",
        value=False,
        key=key_all,
    )

    seleccionadas = []

    for idx, a in enumerate(acts):
        special = ""
        texto = a["actividad"]

        if "SOLO DAVIS" in texto.upper():
            special = "Solo DAVIS"
        elif "SOLO APLICA TF15" in texto.upper():
            special = "Solo aplica TF15"
        elif "SOLO PARA TF41-42" in texto.upper():
            special = "Solo TF41-42"
        elif "VIERNES" in texto.upper():
            special = "Condición: viernes"

        key = f"act_{fecha_str}_{normalizar_maquina(maquina)}_{a['fila_excel']}_{idx}"

        default = marcar_todas or (
            str(a["hoja"]), str(a["fila_excel"]), str(a["actividad"])
        ) in completadas_previas

        c1, c2 = st.columns([6, 1])

        with c1:
            marcado = st.checkbox(
                texto,
                value=default,
                key=key,
            )

        with c2:
            if a["minutos"] is None:
                st.markdown('<span class="badge">sin tiempo</span>',
                            unsafe_allow_html=True)
            else:
                st.markdown(
                    f'<div class="activity-time">{int(a["minutos"])} min</div>',
                    unsafe_allow_html=True,
                )

        if special:
            st.caption(f"⚠️ {special}")

        if marcado:
            seleccionadas.append(a)

    total = sum(a["minutos"] or 0 for a in seleccionadas)

    st.markdown(
        f"""
        <div class="panel">
            <div class="panel-title">Resumen</div>
            <div class="panel-sub">
                ✓ {len(seleccionadas)} de {len(acts)} actividades seleccionadas
                &nbsp;&nbsp; | &nbsp;&nbsp;
                ⏱ {int(total)} minutos
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if st.button("💾 GUARDAR ACTIVIDADES", type="primary",
                 use_container_width=True, key=f"guardar_{fecha_str}_{maquina}"):

        ejecutor = st.session_state.get("ejecutor", "").strip()
        ot_val = st.session_state.get(f"ot_{maquina}", ot)
        recibio = st.session_state.get(f"recibio_{maquina}", "")
        obs = st.session_state.get(f"obs_{maquina}", "")

        seleccion_keys = {
            (str(a["hoja"]), str(a["fila_excel"]), str(a["actividad"]))
            for a in seleccionadas
        }

        # Guardamos cada actividad del bloque, completada o pendiente.
        for a in acts:
            k = (str(a["hoja"]), str(a["fila_excel"]), str(a["actividad"]))

            registro = {
                "fecha": fecha_str,
                "dia": ["Lunes","Martes","Miércoles","Jueves","Viernes","Sábado"][f.weekday()],
                "hora": next(
                    (x["hora"] for x in obtener_ruta_del_dia(f) if x["maquina"] == maquina),
                    ""
                ),
                "maquina": maquina,
                "hoja": a["hoja"],
                "bloque": a["bloque"],
                "fila_excel": a["fila_excel"],
                "actividad": a["actividad"],
                "minutos": a["minutos"],
                "estado": "Completada" if k in seleccion_keys else "Pendiente",
                "ejecutor": ejecutor,
                "ot": ot_val,
                "recibio": recibio,
                "observacion": obs,
                "fecha_hora_guardado": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }

            upsert_actividad(registro)

        guardar_datos(st.session_state.datos)

        st.success(
            f"Guardado: {len(seleccionadas)} actividades completadas · "
            f"{int(total)} minutos."
        )

    st.markdown("</div>", unsafe_allow_html=True)


# ============================================================
# CRONOGRAMA
# ============================================================

def cronograma():
    header()
    f = st.session_state.fecha_trabajo
    lunes = f - timedelta(days=f.weekday())
    sabado = lunes + timedelta(days=5)

    st.markdown('<div class="page">', unsafe_allow_html=True)

    if st.button("← INICIO", key="crono_inicio"):
        ir("home")

    st.markdown(
        f"""
        <div class="panel">
            <div style="display:flex;justify-content:space-between;align-items:center;gap:16px;flex-wrap:wrap;">
                <div>
                    <div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap;">
                        <div class="panel-title">📊 Cronograma semanal</div>
                        <span class="crono-readonly">🔒 SOLO LECTURA</span>
                    </div>
                    <div class="panel-sub">
                        Planificación protegida para consulta: no se puede editar ni reemplazar el horario desde esta pantalla.
                    </div>
                </div>
                <div style="text-align:right;">
                    <div class="badge">SEMANA {f.isocalendar().week}</div>
                    <div style="color:#8dccff;font-size:13px;margin-top:7px;">
                        {lunes.strftime('%d/%m/%Y')} – {sabado.strftime('%d/%m/%Y')}
                    </div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Estilos exclusivos de esta vista: no cambian el resto de la aplicación.
    st.markdown(
        """
        <style>
        .crono-kpi {
            background:linear-gradient(145deg,#0a1b2b,#07111b);
            border:1px solid #1a3a54;
            border-radius:15px;
            padding:16px 18px;
            min-height:124px;
            box-shadow:0 7px 18px rgba(0,0,0,.16);
        }
        .crono-kpi-label { color:#9ec8e8;font-size:12px;font-weight:800;letter-spacing:.5px;text-transform:uppercase; }
        .crono-kpi-value { color:#f5faff;font-size:29px;font-weight:900;margin-top:8px;line-height:1.15; }
        .crono-kpi-note { color:#7899b5;font-size:12px;margin-top:6px; }
        .crono-day-card {
            background:linear-gradient(150deg,#091827,#060e17);
            border:1px solid #1a3851;
            border-radius:15px;
            padding:16px;
            margin-bottom:13px;
            min-height:205px;
        }
        .crono-day-title { color:#f5f9ff;font-size:18px;font-weight:900; }
        .crono-day-date { color:#82bde8;font-size:12px;margin-top:3px; }
        .crono-day-time { color:#ffffff;font-size:24px;font-weight:900;margin-top:14px; }
        .crono-day-stats { color:#a4bed3;font-size:12px;margin-top:8px;line-height:1.8; }
        .crono-bar-track { height:7px;background:#14283a;border-radius:9px;overflow:hidden;margin-top:13px; }
        .crono-bar-fill { height:7px;background:linear-gradient(90deg,#087de9,#25c6f7);border-radius:9px; }
        .crono-section-title { color:#f5f9ff;font-size:18px;font-weight:900;margin:15px 0 10px; }
        .crono-readonly {
            display:inline-block;
            color:#bdebd5;
            background:#0b3028;
            border:1px solid #176c55;
            border-radius:999px;
            padding:5px 11px;
            font-size:12px;
            font-weight:900;
            letter-spacing:.4px;
            white-space:nowrap;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    nombres_dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"]
    resumen = []

    for desplazamiento, dia in enumerate(nombres_dias):
        fecha = lunes + timedelta(days=desplazamiento)
        ruta = obtener_ruta_del_dia(fecha)
        actividades_dia = 0
        minutos_dia = 0

        for item in ruta:
            acts = buscar_actividades(item["maquina"], fecha)
            actividades_dia += len(acts)
            minutos_dia += sum(a["minutos"] or 0 for a in acts)

        resumen.append({
            "dia": dia,
            "fecha": fecha,
            "ruta": ruta,
            "maquinas": len(ruta),
            "actividades": actividades_dia,
            "minutos": int(minutos_dia),
        })

    total_asignaciones = sum(d["maquinas"] for d in resumen)
    total_actividades = sum(d["actividades"] for d in resumen)
    total_minutos = sum(d["minutos"] for d in resumen)
    dias_programados = sum(1 for d in resumen if d["maquinas"] > 0)
    total_horas, minutos_restantes = divmod(total_minutos, 60)
    max_actividades_dia = max((d["actividades"] for d in resumen), default=0)

    st.markdown('<div class="crono-section-title">Resumen de la semana</div>', unsafe_allow_html=True)
    k1, k2, k3, k4 = st.columns(4)
    kpis = [
        (k1, "Asignaciones de máquina", f"{total_asignaciones}", "Suma de máquinas programadas por día"),
        (k2, "Actividades previstas", f"{total_actividades}", "Según las hojas de actividades del Excel"),
        (k3, "Tiempo estimado", f"{total_horas} h {minutos_restantes:02d} min", f"{total_minutos:,} minutos en total".replace(",", ".")),
        (k4, "Días con programación", f"{dias_programados} de 6", "Lunes a sábado"),
    ]
    for col, etiqueta, valor, nota in kpis:
        with col:
            st.markdown(
                f"""
                <div class="crono-kpi">
                    <div class="crono-kpi-label">{etiqueta}</div>
                    <div class="crono-kpi-value">{valor}</div>
                    <div class="crono-kpi-note">{nota}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    st.markdown('<div class="crono-section-title">Carga programada por día</div>', unsafe_allow_html=True)
    st.caption("La barra compara el número de actividades de cada día con el día más cargado de esta semana.")

    # Dos filas de tres tarjetas: lectura rápida en computador y tablet.
    for inicio in (0, 3):
        columnas = st.columns(3)
        for indice, dia_info in enumerate(resumen[inicio:inicio + 3]):
            with columnas[indice]:
                porcentaje = round(dia_info["actividades"] / max_actividades_dia * 100) if max_actividades_dia else 0
                horas_dia, mins_dia = divmod(dia_info["minutos"], 60)
                st.markdown(
                    f"""
                    <div class="crono-day-card">
                        <div class="crono-day-title">{dia_info['dia']}</div>
                        <div class="crono-day-date">{dia_info['fecha'].strftime('%d/%m/%Y')}</div>
                        <div class="crono-day-time">{horas_dia} h {mins_dia:02d} min</div>
                        <div class="crono-day-stats">
                            🏭 <b>{dia_info['maquinas']}</b> máquinas programadas<br>
                            📋 <b>{dia_info['actividades']}</b> actividades previstas<br>
                            ⏱️ <b>{dia_info['minutos']}</b> minutos estimados
                        </div>
                        <div class="crono-bar-track"><div class="crono-bar-fill" style="width:{porcentaje}%;"></div></div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

    st.markdown('<div class="crono-section-title">Detalle del cronograma</div>', unsafe_allow_html=True)
    st.caption("Abre cada día para consultar el horario y las actividades. Las tablas son de solo lectura y no guardan cambios en el cronograma.")

    for dia_info in resumen:
        etiqueta = (
            f"📅 {dia_info['dia']} · {dia_info['fecha'].strftime('%d/%m/%Y')}"
            f"  —  {dia_info['maquinas']} máquinas · {dia_info['actividades']} actividades · {dia_info['minutos']} min"
        )
        with st.expander(etiqueta, expanded=False):
            if not dia_info["ruta"]:
                st.info("No hay máquinas programadas para este día en el Excel.")
                continue

            filas_detalle = []
            for item in dia_info["ruta"]:
                acts = buscar_actividades(item["maquina"], dia_info["fecha"])
                minutos_maquina = int(sum(a["minutos"] or 0 for a in acts))
                horas_maquina, mins_maquina = divmod(minutos_maquina, 60)
                filas_detalle.append({
                    "Horario": item.get("hora", ""),
                    "Máquina": item["maquina"],
                    "OT": item.get("ot") or "Sin OT",
                    "Actividades": len(acts),
                    "Minutos": minutos_maquina,
                    "Duración estimada": f"{horas_maquina} h {mins_maquina:02d} min",
                })

            st.dataframe(
                pd.DataFrame(filas_detalle),
                use_container_width=True,
                hide_index=True,
            )

            maquinas_dia = [str(item["maquina"]) for item in dia_info["ruta"]]
            maquina_detalle = st.selectbox(
                "Consultar actividades de una máquina",
                maquinas_dia,
                key=f"crono_maquina_{dia_info['fecha'].isoformat()}",
            )
            actividades_maquina = buscar_actividades(maquina_detalle, dia_info["fecha"])
            if actividades_maquina:
                filas_actividades = []
                for actividad in actividades_maquina:
                    valor_minutos = actividad["minutos"]
                    filas_actividades.append({
                        "Actividad": actividad["actividad"],
                        "Tiempo (min)": int(valor_minutos) if valor_minutos is not None else "—",
                        "Hoja de Excel": actividad["hoja"],
                    })
                st.markdown(f"**Actividades previstas para {maquina_detalle}**")
                st.dataframe(
                    pd.DataFrame(filas_actividades),
                    use_container_width=True,
                    hide_index=True,
                )
            else:
                st.info(f"No se encontraron actividades vinculadas a {maquina_detalle} para este día.")

    st.markdown("</div>", unsafe_allow_html=True)


# ============================================================
# HISTORIAL
# ============================================================

def historial():
    header()

    st.markdown('<div class="page">', unsafe_allow_html=True)

    if st.button("← INICIO", key="hist_inicio"):
        ir("home")

    st.markdown(
        """
        <div class="panel">
            <div class="panel-title">◷ Historial</div>
            <div class="panel-sub">
                Registros guardados por actividad.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    df = st.session_state.datos.copy()

    if df.empty:
        st.info("Todavía no hay actividades guardadas.")
        st.markdown("</div>", unsafe_allow_html=True)
        return

    c1, c2, c3 = st.columns(3)

    with c1:
        fecha_filtro = st.date_input(
            "Fecha",
            value=st.session_state.fecha_trabajo,
            key="hist_fecha",
        )

    with c2:
        maquinas = ["Todas"] + sorted(df["maquina"].dropna().astype(str).unique().tolist())
        maq = st.selectbox("Máquina", maquinas, key="hist_maquina")

    with c3:
        estados = ["Todos", "Completada", "Pendiente"]
        estado = st.selectbox("Estado", estados, key="hist_estado")

    filtro = df.copy()
    filtro = filtro[filtro["fecha"].astype(str) == fecha_filtro.strftime("%Y-%m-%d")]

    if maq != "Todas":
        filtro = filtro[filtro["maquina"].astype(str) == maq]

    if estado != "Todos":
        filtro = filtro[filtro["estado"].astype(str) == estado]

    st.dataframe(
        filtro[
            [
                "fecha", "dia", "hora", "maquina", "actividad",
                "minutos", "estado", "ejecutor", "ot", "recibio", "observacion"
            ]
        ],
        use_container_width=True,
        hide_index=True,
    )

    excel_bytes = None
    import io
    buffer = io.BytesIO()
    filtro.to_excel(buffer, index=False)
    excel_bytes = buffer.getvalue()

    st.download_button(
        "⬇️ Descargar historial en Excel",
        data=excel_bytes,
        file_name=f"historial_lubricacion_{fecha_filtro.strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    st.markdown("</div>", unsafe_allow_html=True)


# ============================================================
# ENRUTAMIENTO
# ============================================================

if st.session_state.pagina == "home":
    home()
elif st.session_state.pagina == "ruta":
    ruta_diaria()
elif st.session_state.pagina == "actividad":
    actividad_maquina()
elif st.session_state.pagina == "cronograma":
    cronograma()
elif st.session_state.pagina == "historial":
    historial()
else:
    home()
