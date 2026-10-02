"""RUTA DE LUBRICACIÓN - App tipo tablet
Mantiene la lógica de la app original y adopta una interfaz tipo dashboard/tablet.
Ejecutar con: streamlit run app.py
"""

import re
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st

BASE_DIR = Path(__file__).parent
ARCHIVO_PLANTILLA = BASE_DIR / "ruta_lub_v2_oct_2026.xlsx"
ARCHIVO_DATOS = BASE_DIR / "datos_ruta_lub.xlsx"

DIAS_SEMANA = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"]


# ---------------------------------------------------------------------------
# PARSER DE LA PLANTILLA ORIGINAL
# ---------------------------------------------------------------------------

def normalizar(texto: str) -> str:
    texto = str(texto).upper()
    texto = re.sub(r"[ÁÀ]", "A", texto)
    texto = re.sub(r"[ÉÈ]", "E", texto)
    texto = re.sub(r"[ÍÌ]", "I", texto)
    texto = re.sub(r"[ÓÒ]", "O", texto)
    texto = re.sub(r"[ÚÙ]", "U", texto)
    texto = re.sub(r"[\s\-\_\.]+", "", texto)
    return texto


@st.cache_data(show_spinner=False)
def parse_horario(archivo) -> list:
    df = pd.read_excel(archivo, sheet_name=0, header=None)
    dias = []
    for fila_encabezado, rango in [(3, range(4, 17)), (20, range(21, 40))]:
        for col in (1, 5, 9):
            dia = str(df.iloc[fila_encabezado, col]).strip()
            if dia not in DIAS_SEMANA:
                continue
            maquinas = []
            for r in rango:
                valor = df.iloc[r, col]
                if pd.notna(valor) and str(valor).strip():
                    maquinas.append(str(valor).strip())
            dias.append({"dia": dia, "maquinas": maquinas})
    return dias


@st.cache_data(show_spinner=False)
def parse_bloques(archivo) -> list:
    xl = pd.ExcelFile(archivo)
    bloques = []
    for hoja in xl.sheet_names:
        if "HORARIO" in hoja.upper():
            continue

        df = pd.read_excel(archivo, sheet_name=hoja, header=None)
        n = len(df)
        i = 0
        titulo_pendiente = None

        while i < n:
            col0 = str(df.iloc[i, 0]).strip() if pd.notna(df.iloc[i, 0]) else ""

            if col0 == "ACTIVIDADES":
                titulo = titulo_pendiente or hoja
                titulo_pendiente = None

                for j in range(i, max(-1, i - 4), -1):
                    fila = df.iloc[j]
                    if any(str(v).strip() == "FECHA" for v in fila if pd.notna(v)):
                        candidatos = [
                            str(v).strip()
                            for v in fila.iloc[2:]
                            if pd.notna(v) and str(v).strip() != "FECHA"
                        ]
                        if candidatos:
                            titulo = candidatos[0]
                        break

                maquinas, vistos = [], set()
                for v in df.iloc[i, 2:]:
                    if pd.notna(v) and str(v).strip():
                        m = str(v).strip()
                        if m in vistos:
                            m = f"{m} (2)"
                        vistos.add(m)
                        maquinas.append(m)

                actividades = []
                k = i + 1

                while k < n:
                    fila = df.iloc[k]
                    c0 = fila[0]
                    c0s = str(c0).strip() if pd.notna(c0) else ""

                    if c0s in ("ACTIVIDADES", "EJECUTADO POR:"):
                        break
                    if any(str(v).strip() == "FECHA" for v in fila if pd.notna(v)):
                        break

                    if c0s and not re.match(r"^\d+\.?", c0s) and any(
                        pd.notna(v) and str(v).strip() for v in fila.iloc[2:]
                    ):
                        cand = [
                            str(v).strip()
                            for v in fila.iloc[2:]
                            if pd.notna(v) and str(v).strip()
                        ]
                        if cand:
                            titulo_pendiente = cand[0]
                        break

                    if c0s:
                        tiempo = fila[1] if pd.notna(fila[1]) else None
                        actividades.append({"actividad": c0s, "tiempo": tiempo})
                    k += 1

                if maquinas and actividades:
                    bloques.append({
                        "area": hoja,
                        "titulo": titulo,
                        "maquinas": maquinas,
                        "actividades": actividades,
                    })
                i = k
            else:
                i += 1

    return bloques


# ---------------------------------------------------------------------------
# CAPA DE DATOS ORIGINAL
# ---------------------------------------------------------------------------

def cargar_datos() -> tuple:
    if ARCHIVO_DATOS.exists():
        xl = pd.ExcelFile(ARCHIVO_DATOS)
        registros = (
            pd.read_excel(xl, "REGISTROS")
            if "REGISTROS" in xl.sheet_names
            else pd.DataFrame()
        )
        checks = (
            pd.read_excel(xl, "CHECKS")
            if "CHECKS" in xl.sheet_names
            else pd.DataFrame()
        )
    else:
        registros = pd.DataFrame(
            columns=[
                "fecha", "dia", "maquina", "ot",
                "ejecuto", "recibio", "observaciones"
            ]
        )
        checks = pd.DataFrame(
            columns=[
                "fecha", "area", "bloque", "maquina",
                "actividad", "tiempo", "ok", "ejecutado_por"
            ]
        )
    return registros, checks


def guardar_datos(registros: pd.DataFrame, checks: pd.DataFrame) -> None:
    with pd.ExcelWriter(ARCHIVO_DATOS, engine="openpyxl") as writer:
        registros.to_excel(writer, sheet_name="REGISTROS", index=False)
        checks.to_excel(writer, sheet_name="CHECKS", index=False)


def dia_a_ingles(dia: str) -> str:
    return {
        "Lunes": "Monday",
        "Martes": "Tuesday",
        "Miércoles": "Wednesday",
        "Jueves": "Thursday",
        "Viernes": "Friday",
        "Sábado": "Saturday",
    }.get(dia, "")


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    import io

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return buffer.getvalue()


# ---------------------------------------------------------------------------
# CONFIGURACIÓN / ESTILO TIPO TABLET
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="Ruta de Lubricación",
    page_icon="🛢️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
<style>
#MainMenu, header, footer, [data-testid="stToolbar"] {
    visibility: hidden !important;
}

.stApp {
    background-color: #F1F5F9;
    max-width: 100vw;
    overflow-x: hidden;
}

.main .block-container {
    padding: 0.35rem 0.5rem 1.5rem !important;
    max-width: 100% !important;
}

div[data-testid="stVerticalBlock"] {
    gap: 0.35rem !important;
}

.stButton > button {
    border-radius: 8px;
    font-weight: 700;
    font-size: 12px !important;
    min-height: 38px;
}

.stSelectbox label,
.stTextInput label,
.stDateInput label,
.stTextArea label {
    color: #475569 !important;
    font-size: 12px !important;
}

.tablet-header {
    background: linear-gradient(135deg, #0EA5E9 0%, #38BDF8 100%);
    color: white;
    padding: 12px 16px;
    border-radius: 0 0 16px 16px;
    text-align: center;
    font-size: 18px;
    font-weight: 700;
    margin: -0.35rem -0.5rem 12px;
    box-shadow: 0 4px 15px rgba(14,165,233,.25);
    position: sticky;
    top: 0;
    z-index: 100;
}

.home-screen {
    padding: 8px 5px;
    color: #0F172A;
}

.big-counter {
    font-size: 58px;
    font-weight: 900;
    color: #0EA5E9;
    line-height: 1;
    margin: 5px 0;
    text-align: center;
}

.counter-label {
    font-size: 16px;
    color: #475569;
    margin-bottom: 15px;
    text-align: center;
}

.dashboard-card {
    background: white;
    border: 1px solid #E2E8F0;
    border-radius: 14px;
    padding: 15px;
    box-shadow: 0 2px 10px rgba(15,23,42,.06);
    min-height: 110px;
}

.dashboard-card h3 {
    margin: 0 0 6px;
    font-size: 15px;
    color: #0F172A;
}

.dashboard-card p {
    margin: 0;
    font-size: 12px;
    color: #64748B;
    line-height: 1.45;
}

.kpi {
    background: white;
    border-radius: 12px;
    padding: 10px 16px;
    min-width: 100px;
    text-align: center;
    border: 1px solid #E2E8F0;
}

.kpi-value {
    font-size: 22px;
    font-weight: 800;
}

.kpi-label {
    font-size: 10px;
    color: #64748B;
}

.section-title {
    color: #0F172A;
    font-size: 16px;
    font-weight: 800;
    margin: 10px 0 5px;
}

.info-panel {
    background: white;
    border: 1px solid #E2E8F0;
    border-radius: 12px;
    padding: 12px;
    margin: 6px 0;
}

@media (max-width: 768px) {
    .tablet-header {
        font-size: 16px;
        padding: 10px 12px;
    }

    .big-counter {
        font-size: 46px;
    }

    .dashboard-card {
        min-height: 95px;
        padding: 12px;
    }
}
</style>
""",
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------------
# ESTADO DE LA APLICACIÓN
# ---------------------------------------------------------------------------

horario = parse_horario(ARCHIVO_PLANTILLA)
bloques = parse_bloques(ARCHIVO_PLANTILLA)
registros, checks = cargar_datos()

for key, value in {
    "pagina": "home",
    "fecha_trabajo": datetime.now().date(),
    "lubricador": "",
    "area_check": None,
    "bloque_check": 0,
}.items():
    if key not in st.session_state:
        st.session_state[key] = value


def ir(pagina):
    st.session_state.pagina = pagina
    st.rerun()


def nav_atras(pagina):
    c1, c2 = st.columns(2)

    with c1:
        if st.button(
            "← Volver",
            use_container_width=True,
            key=f"back_{pagina}",
        ):
            ir(pagina)

    with c2:
        if st.button(
            "⌂ Inicio",
            use_container_width=True,
            key=f"home_{pagina}",
        ):
            ir("home")


def header_tablet(titulo, badge=""):
    badge_html = (
        f"<span style='font-size:12px;opacity:.85'>{badge}</span>"
        if badge
        else ""
    )

    st.markdown(
        f"""
        <div class="tablet-header"
             style="display:flex;align-items:center;justify-content:space-between">
            <span>{titulo}</span>{badge_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def fecha_actual_str():
    return st.session_state.fecha_trabajo.strftime("%d/%m/%Y")


# ---------------------------------------------------------------------------
# PANTALLA INICIO
# ---------------------------------------------------------------------------

def pantalla_home():
    global registros, checks

    header_tablet("🛢️ Ruta de Lubricación", "Planta Madrid · 2026")

    f = st.session_state.fecha_trabajo
    f_str = f.strftime("%Y-%m-%d")

    hechos = len(registros[registros["fecha"] == f_str])
    checks_hoy = len(checks[checks["fecha"] == f_str])

    lunes = f - timedelta(days=f.weekday())
    total_hechas_semana = sum(
        len(registros[registros["fecha"] == (lunes + timedelta(days=i)).strftime("%Y-%m-%d")])
        for i in range(6)
    )

    st.markdown("<div class='home-screen'>", unsafe_allow_html=True)

    st.markdown(
        f"""
        <div style="text-align:left;font-size:20px;font-weight:700;
                    color:#475569;margin:4px 0 8px">
            📅 {f.strftime('%A %d/%m/%Y').capitalize()}
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        f"""
        <div class="big-counter">{hechos}</div>
        <div class="counter-label">registros realizados hoy</div>
        """,
        unsafe_allow_html=True,
    )

    c1, c2, c3 = st.columns(3)

    with c1:
        st.markdown(
            f"""
            <div class="kpi">
                <div class="kpi-value" style="color:#0EA5E9">{hechos}</div>
                <div class="kpi-label">Ruta hoy</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with c2:
        st.markdown(
            f"""
            <div class="kpi">
                <div class="kpi-value" style="color:#10B981">{checks_hoy}</div>
                <div class="kpi-label">Checklist hoy</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with c3:
        st.markdown(
            f"""
            <div class="kpi">
                <div class="kpi-value" style="color:#64748B">{total_hechas_semana}</div>
                <div class="kpi-label">Semana</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown(
        "<div class='section-title'>Selecciona una función</div>",
        unsafe_allow_html=True,
    )

    c1, c2 = st.columns(2)

    with c1:
        st.markdown(
            """
            <div class="dashboard-card">
                <h3>📅 Ruta diaria</h3>
                <p>Registra las máquinas lubricadas, OT, ejecutor,
                recibido y observaciones.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

        if st.button(
            "ABRIR RUTA DIARIA",
            use_container_width=True,
            type="primary",
            key="home_ruta",
        ):
            ir("horario")

    with c2:
        st.markdown(
            """
            <div class="dashboard-card">
                <h3>🔧 Checklists</h3>
                <p>Ejecuta las actividades por área, bloque y máquina.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

        if st.button(
            "ABRIR CHECKLISTS",
            use_container_width=True,
            type="primary",
            key="home_check",
        ):
            ir("checklists")

    c1, c2 = st.columns(2)

    with c1:
        st.markdown(
            """
            <div class="dashboard-card">
                <h3>📊 Cronograma</h3>
                <p>Consulta el avance de la semana y las actividades
                realizadas.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

        if st.button(
            "VER CRONOGRAMA",
            use_container_width=True,
            key="home_crono",
        ):
            ir("cronograma")

    with c2:
        st.markdown(
            """
            <div class="dashboard-card">
                <h3>🕓 Historial</h3>
                <p>Consulta y descarga los registros guardados por fecha.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

        if st.button(
            "VER HISTORIAL",
            use_container_width=True,
            key="home_hist",
        ):
            ir("historial")

    st.markdown("<div class='info-panel'>", unsafe_allow_html=True)
    st.markdown("**⚙️ Datos de trabajo**")

    c1, c2 = st.columns(2)

    with c1:
        st.session_state.fecha_trabajo = st.date_input(
            "Fecha de trabajo",
            value=f,
            key="fecha_home",
        )

    with c2:
        st.session_state.lubricador = st.text_input(
            "Lubricador / Ejecutor",
            value=st.session_state.lubricador,
            key="lub_home",
        )

    st.markdown("</div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# PANTALLA RUTA DIARIA
# ---------------------------------------------------------------------------

def pantalla_horario():
    global registros, checks

    header_tablet("📅 Registro de ruta diaria", fecha_actual_str())
    nav_atras("home")

    f = st.session_state.fecha_trabajo

    dia_sel = st.selectbox(
        "Día de la ruta",
        DIAS_SEMANA,
        index=min(f.weekday(), 5),
        key="dia_ruta",
    )

    info_dia = next(
        (d for d in horario if d["dia"] == dia_sel),
        None,
    )

    if not info_dia or not info_dia["maquinas"]:
        st.info(f"No hay máquinas programadas para el {dia_sel}.")
        return

    f_str = f.strftime("%Y-%m-%d")

    prev = registros[
        (registros["fecha"] == f_str)
        & (registros["dia"] == dia_sel)
    ]

    filas = []

    for m in info_dia["maquinas"]:
        fp = prev[
            prev["maquina"].apply(normalizar)
            == normalizar(m)
        ]

        filas.append(
            {
                "Máquina / Equipo": m,
                "OT": fp["ot"].iloc[0] if len(fp) else "",
                "Ejecutó": fp["ejecuto"].iloc[0] if len(fp) else "",
                "Recibió": fp["recibio"].iloc[0] if len(fp) else "",
                "Observaciones": (
                    fp["observaciones"].iloc[0]
                    if len(fp)
                    else ""
                ),
                "✓ Hecho": bool(len(fp)),
            }
        )

    st.markdown(
        f"""
        <div class="section-title">
            {dia_sel} · {len(filas)} máquinas programadas
        </div>
        """,
        unsafe_allow_html=True,
    )

    editado = st.data_editor(
        pd.DataFrame(filas),
        hide_index=True,
        use_container_width=True,
        column_config={
            "OT": st.column_config.TextColumn(
                "OT",
                width="small",
            ),
            "Ejecutó": st.column_config.TextColumn(
                "Ejecutó",
                width="small",
            ),
            "Recibió": st.column_config.TextColumn(
                "Recibió",
                width="small",
            ),
            "Observaciones": st.column_config.TextColumn(
                "Observaciones",
                width="large",
            ),
            "✓ Hecho": st.column_config.CheckboxColumn(
                "✓ Hecho",
            ),
        },
        key="editor_ruta",
    )

    if st.button(
        "💾 GUARDAR REGISTRO DEL DÍA",
        type="primary",
        use_container_width=True,
        key="save_ruta",
    ):
        registros = registros[
            ~(
                (registros["fecha"] == f_str)
                & (registros["dia"] == dia_sel)
            )
        ]

        nuevos = [
            {
                "fecha": f_str,
                "dia": dia_sel,
                "maquina": r["Máquina / Equipo"],
                "ot": r["OT"],
                "ejecuto": (
                    r["Ejecutó"]
                    or st.session_state.lubricador
                ),
                "recibio": r["Recibió"],
                "observaciones": r["Observaciones"],
            }
            for _, r in editado.iterrows()
            if r["✓ Hecho"]
        ]

        if nuevos:
            registros = pd.concat(
                [registros, pd.DataFrame(nuevos)],
                ignore_index=True,
            )

            guardar_datos(registros, checks)
            st.success(
                f"✅ Se guardaron {len(nuevos)} registros."
            )
            st.rerun()
        else:
            st.warning(
                "Marca al menos una máquina como '✓ Hecho'."
            )


# ---------------------------------------------------------------------------
# PANTALLA CHECKLISTS
# ---------------------------------------------------------------------------

def pantalla_checklists():
    global registros, checks

    header_tablet("🔧 Checklists", fecha_actual_str())
    nav_atras("home")

    areas = sorted({b["area"] for b in bloques})

    if not areas:
        st.info("No se encontraron bloques de checklist.")
        return

    area_default = (
        st.session_state.area_check
        if st.session_state.area_check in areas
        else areas[0]
    )

    area_sel = st.selectbox(
        "Área",
        areas,
        index=areas.index(area_default),
        key="area_check_select",
    )

    st.session_state.area_check = area_sel

    bloques_area = [
        b for b in bloques
        if b["area"] == area_sel
    ]

    idx = min(
        st.session_state.bloque_check,
        len(bloques_area) - 1,
    )

    bloque_sel = st.selectbox(
        "Bloque / Grupo",
        range(len(bloques_area)),
        index=idx,
        format_func=lambda i:
            f"{bloques_area[i]['titulo']} · "
            f"{len(bloques_area[i]['maquinas'])} máquinas",
        key="bloque_check_select",
    )

    st.session_state.bloque_check = bloque_sel
    bloque = bloques_area[bloque_sel]

    st.markdown(
        f"""
        <div class="section-title">
            {area_sel} — {bloque['titulo']}
        </div>
        """,
        unsafe_allow_html=True,
    )

    ejecutor = st.text_input(
        "Ejecutado por",
        value=st.session_state.lubricador,
        key="ejec_check",
    )

    f_str = st.session_state.fecha_trabajo.strftime(
        "%Y-%m-%d"
    )

    prev = checks[
        (checks["fecha"] == f_str)
        & (checks["area"] == area_sel)
        & (checks["bloque"] == bloque["titulo"])
    ]

    datos = {}

    for act in bloque["actividades"]:
        fila = []

        for m in bloque["maquinas"]:
            coinciden = prev[
                (prev["maquina"].apply(normalizar)
                 == normalizar(m))
                & (prev["actividad"] == act["actividad"])
            ]

            fila.append(
                bool(
                    len(coinciden)
                    and coinciden["ok"].iloc[0]
                )
            )

        etiqueta = (
            f"{act['actividad']}  [{act['tiempo']} min]"
            if pd.notna(act["tiempo"])
            else act["actividad"]
        )

        datos[etiqueta] = fila

    df_grid = pd.DataFrame(
        datos,
        index=bloque["maquinas"],
    ).T

    config = {
        "_index": st.column_config.Column(
            "Actividad",
            width="large",
        )
    }

    for m in bloque["maquinas"]:
        config[m] = st.column_config.CheckboxColumn(
            m,
            width="small",
        )

    grid = st.data_editor(
        df_grid.reset_index().rename(
            columns={"index": "Actividad"}
        ),
        hide_index=True,
        use_container_width=True,
        column_config=config,
        key="editor_check",
    )

    hechos = int(
        grid[bloque["maquinas"]].sum().sum()
    )

    total = (
        len(bloque["actividades"])
        * len(bloque["maquinas"])
    )

    st.progress(
        hechos / total if total else 0,
        text=f"Avance del bloque: {hechos}/{total}",
    )

    if st.button(
        "💾 GUARDAR CHECKLIST",
        type="primary",
        use_container_width=True,
        key="save_check",
    ):
        checks = checks[
            ~(
                (checks["fecha"] == f_str)
                & (checks["area"] == area_sel)
                & (checks["bloque"] == bloque["titulo"])
            )
        ]

        nuevos = []

        for _, fila in grid.iterrows():
            act_nombre = fila["Actividad"].split(
                "  ["
            )[0]

            tiempo = next(
                (
                    a["tiempo"]
                    for a in bloque["actividades"]
                    if a["actividad"] == act_nombre
                ),
                None,
            )

            for m in bloque["maquinas"]:
                if fila[m]:
                    nuevos.append(
                        {
                            "fecha": f_str,
                            "area": area_sel,
                            "bloque": bloque["titulo"],
                            "maquina": m,
                            "actividad": act_nombre,
                            "tiempo": tiempo,
                            "ok": True,
                            "ejecutado_por": ejecutor,
                        }
                    )

        if nuevos:
            checks = pd.concat(
                [checks, pd.DataFrame(nuevos)],
                ignore_index=True,
            )

            guardar_datos(registros, checks)
            st.success(
                f"✅ Checklist guardado: "
                f"{len(nuevos)} actividades."
            )
            st.rerun()
        else:
            st.warning(
                "No marcaste ninguna actividad."
            )


# ---------------------------------------------------------------------------
# PANTALLA CRONOGRAMA
# ---------------------------------------------------------------------------

def pantalla_cronograma():
    header_tablet(
        "📊 Cronograma semanal",
        fecha_actual_str(),
    )
    nav_atras("home")

    f = st.session_state.fecha_trabajo
    lunes = f - timedelta(days=f.weekday())

    fechas_semana = [
        lunes + timedelta(days=i)
        for i in range(6)
    ]

    hechos_dia = len(
        registros[
            registros["fecha"]
            == f.strftime("%Y-%m-%d")
        ]
    )

    checks_dia = len(
        checks[
            checks["fecha"]
            == f.strftime("%Y-%m-%d")
        ]
    )

    c1, c2, c3 = st.columns(3)

    c1.metric("Registros hoy", hechos_dia)
    c2.metric("Checklist hoy", checks_dia)
    c3.metric(
        "Semana",
        f"{lunes.strftime('%d/%m')} – "
        f"{(lunes + timedelta(days=5)).strftime('%d/%m')}",
    )

    filas_crono = []

    for d in horario:
        fila = {"Día": d["dia"]}

        for fecha in fechas_semana:
            if (
                fecha.strftime("%A")
                != dia_a_ingles(d["dia"])
            ):
                continue

            fs = fecha.strftime("%Y-%m-%d")
            regs_dia = registros[
                registros["fecha"] == fs
            ]

            hechas = sum(
                1
                for m in d["maquinas"]
                if len(
                    regs_dia[
                        regs_dia["maquina"].apply(
                            normalizar
                        )
                        == normalizar(m)
                    ]
                )
            )

            pct = (
                hechas / len(d["maquinas"])
                if d["maquinas"]
                else 0
            )

            fila[
                fecha.strftime("%d/%m")
            ] = (
                f"{'🟢' if pct == 1 else '🟡' if pct > 0 else '⚪'} "
                f"{hechas}/{len(d['maquinas'])}"
            )

        filas_crono.append(fila)

    st.dataframe(
        pd.DataFrame(filas_crono).set_index("Día"),
        use_container_width=True,
    )

    st.markdown(
        "<div class='section-title'>Actividades por área</div>",
        unsafe_allow_html=True,
    )

    if len(checks):
        resumen = (
            checks.groupby(
                ["fecha", "area"]
            )
            .size()
            .unstack(fill_value=0)
            .reindex(
                [
                    x.strftime("%Y-%m-%d")
                    for x in fechas_semana
                ],
                fill_value=0,
            )
        )

        resumen.index = [
            x.strftime("%a %d/%m")
            for x in fechas_semana
        ]

        st.dataframe(
            resumen,
            use_container_width=True,
        )
    else:
        st.info(
            "Aún no hay checklists registrados."
        )


# ---------------------------------------------------------------------------
# PANTALLA HISTORIAL
# ---------------------------------------------------------------------------

def pantalla_historial():
    header_tablet(
        "🕓 Historial de registros",
        fecha_actual_str(),
    )
    nav_atras("home")

    if not len(registros) and not len(checks):
        st.info(
            "Todavía no hay registros guardados."
        )
        return

    tipo = st.radio(
        "Ver historial de",
        ["Ruta (horario)", "Checklists"],
        horizontal=True,
        key="hist_tipo",
    )

    df_hist = (
        registros
        if tipo == "Ruta (horario)"
        else checks
    )

    if not len(df_hist):
        st.info(
            f"No hay registros de {tipo.lower()} todavía."
        )
        return

    c1, c2 = st.columns(2)

    fechas_hist = sorted(
        df_hist["fecha"].unique()
    )

    f_ini = c1.selectbox(
        "Desde",
        fechas_hist,
        index=0,
        key="hist_ini",
    )

    f_fin = c2.selectbox(
        "Hasta",
        fechas_hist,
        index=len(fechas_hist) - 1,
        key="hist_fin",
    )

    filtrado = df_hist[
        (df_hist["fecha"] >= f_ini)
        & (df_hist["fecha"] <= f_fin)
    ]

    st.dataframe(
        filtrado,
        use_container_width=True,
        hide_index=True,
    )

    st.download_button(
        "⬇️ DESCARGAR HISTORIAL (EXCEL)",
        data=to_excel_bytes(filtrado),
        file_name=(
            f"historial_"
            f"{tipo.split()[0].lower()}_"
            f"{f_ini}_{f_fin}.xlsx"
        ),
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True,
    )


# ---------------------------------------------------------------------------
# ENRUTADOR
# ---------------------------------------------------------------------------

PANTALLAS = {
    "home": pantalla_home,
    "horario": pantalla_horario,
    "checklists": pantalla_checklists,
    "cronograma": pantalla_cronograma,
    "historial": pantalla_historial,
}

PANTALLAS.get(
    st.session_state.pagina,
    pantalla_home,
)()
