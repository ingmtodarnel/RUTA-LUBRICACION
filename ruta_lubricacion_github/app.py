# -*- coding: utf-8 -*-
"""
RUTA DE LUBRICACIÓN - App de registro semanal
Lee la plantilla maestra (Excel) y guarda los registros en un Excel de datos.
Ejecutar con:  streamlit run app.py
"""
import re
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st

# ---------------------------------------------------------------------------
# CONFIGURACIÓN
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).parent
ARCHIVO_PLANTILLA = BASE_DIR / "ruta_lub_v2_oct_2026.xlsx"
ARCHIVO_DATOS = BASE_DIR / "datos_ruta_lub.xlsx"

st.set_page_config(page_title="Ruta de Lubricación", page_icon="🛢️", layout="wide")

DIAS_SEMANA = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"]


# ---------------------------------------------------------------------------
# PARSER DE LA PLANTILLA
# ---------------------------------------------------------------------------
def normalizar(texto: str) -> str:
    """Quita acentos, espacios y guiones para comparar nombres de máquinas."""
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
    """Extrae la ruta semanal: días y sus máquinas (HOJA 1 HORARIO)."""
    df = pd.read_excel(archivo, sheet_name=0, header=None)
    dias = []
    # Dos bloques en la hoja: filas 3 (encabezado) + 4-16 datos,
    # y fila 20 (encabezado) + 21-39 datos. Días en columnas 1, 5 y 9.
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
    """Extrae los bloques de checklist de cada hoja (área)."""
    xl = pd.ExcelFile(archivo)
    bloques = []
    for hoja in xl.sheetnames:
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
                # Título del bloque: fila más cercana hacia atrás que tenga FECHA
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

                # Máquinas: columnas 2 en adelante de la fila de encabezado
                maquinas, vistos = [], set()
                for v in df.iloc[i, 2:]:
                    if pd.notna(v) and str(v).strip():
                        m = str(v).strip()
                        if m in vistos:  # evitar nombres duplicados (ej. TF11 x2)
                            m = f"{m} (2)"
                        vistos.add(m)
                        maquinas.append(m)

                # Actividades: filas siguientes hasta encabezado de nuevo bloque.
                # Se saltan filas vacías (hay actividades sin número en medio).
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
                    # Fila-anotación con nombre de sub-grupo en columnas 2+
                    # (ej. "Estas revisar el sábado ... | TF RIG LYLE")
                    if c0s and not re.match(r"^\d+\.?", c0s) and any(
                        pd.notna(v) and str(v).strip() for v in fila.iloc[2:]
                    ):
                        cand = [str(v).strip() for v in fila.iloc[2:] if pd.notna(v) and str(v).strip()]
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
# CAPA DE DATOS (Excel de registros)
# ---------------------------------------------------------------------------
def cargar_datos() -> tuple:
    if ARCHIVO_DATOS.exists():
        xl = pd.ExcelFile(ARCHIVO_DATOS)
        registros = pd.read_excel(xl, "REGISTROS") if "REGISTROS" in xl.sheet_names else pd.DataFrame()
        checks = pd.read_excel(xl, "CHECKS") if "CHECKS" in xl.sheet_names else pd.DataFrame()
    else:
        registros = pd.DataFrame(columns=["fecha", "dia", "maquina", "ot", "ejecuto", "recibio", "observaciones"])
        checks = pd.DataFrame(columns=["fecha", "area", "bloque", "maquina", "actividad", "tiempo", "ok", "ejecutado_por"])
    return registros, checks


def guardar_datos(registros: pd.DataFrame, checks: pd.DataFrame) -> None:
    with pd.ExcelWriter(ARCHIVO_DATOS, engine="openpyxl") as writer:
        registros.to_excel(writer, sheet_name="REGISTROS", index=False)
        checks.to_excel(writer, sheet_name="CHECKS", index=False)




def dia_a_ingles(dia: str) -> str:
    """Convierte nombre de día en español al inglés para comparar con weekday()."""
    return {
        "Lunes": "Monday", "Martes": "Tuesday", "Miércoles": "Wednesday",
        "Jueves": "Thursday", "Viernes": "Friday", "Sábado": "Saturday",
    }.get(dia, "")


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    import io
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return buffer.getvalue()

# ---------------------------------------------------------------------------
# CARGA INICIAL
# ---------------------------------------------------------------------------
horario = parse_horario(ARCHIVO_PLANTILLA)
bloques = parse_bloques(ARCHIVO_PLANTILLA)
registros, checks = cargar_datos()

# ---------------------------------------------------------------------------
# BARRA LATERAL
# ---------------------------------------------------------------------------
st.sidebar.title("🛢️ Ruta de Lubricación")
st.sidebar.caption("Planta Madrid · 2026")
fecha_trabajo = st.sidebar.date_input("📆 Fecha de trabajo", datetime.now())
lubricador = st.sidebar.text_input("👤 Lubricador / Ejecutor", value="")
st.sidebar.divider()
semana_iso = fecha_trabajo.isocalendar()[1]
st.sidebar.metric("Semana del año", semana_iso)

tab_horario, tab_check, tab_crono, tab_hist = st.tabs(
    ["📅 Horario semanal", "🔧 Checklists", "📊 Cronograma", "🕓 Historial"]
)

# ---------------------------------------------------------------------------
# PESTAÑA 1: HORARIO SEMANAL
# ---------------------------------------------------------------------------
with tab_horario:
    st.subheader("Registro de ruta diaria")
    dia_sel = st.selectbox("Día de la ruta", DIAS_SEMANA,
                           index=min(fecha_trabajo.weekday(), 5))
    info_dia = next((d for d in horario if d["dia"] == dia_sel), None)

    if not info_dia or not info_dia["maquinas"]:
        st.info(f"No hay máquinas programadas para el {dia_sel}.")
    else:
        # Precargar registros existentes de esa fecha + día
        f_str = fecha_trabajo.strftime("%Y-%m-%d")
        prev = registros[(registros["fecha"] == f_str) & (registros["dia"] == dia_sel)]
        filas = []
        for m in info_dia["maquinas"]:
            fila_prev = prev[prev["maquina"].apply(normalizar) == normalizar(m)]
            filas.append({
                "Máquina / Equipo": m,
                "OT": fila_prev["ot"].iloc[0] if len(fila_prev) else "",
                "Ejecutó": fila_prev["ejecuto"].iloc[0] if len(fila_prev) else "",
                "Recibió": fila_prev["recibio"].iloc[0] if len(fila_prev) else "",
                "Observaciones": fila_prev["observaciones"].iloc[0] if len(fila_prev) else "",
                "✓ Hecho": bool(len(fila_prev)),
            })
        editado = st.data_editor(
            pd.DataFrame(filas),
            hide_index=True, use_container_width=True,
            column_config={
                "OT": st.column_config.TextColumn("OT", width="small"),
                "Ejecutó": st.column_config.TextColumn("Ejecutó", width="small"),
                "Recibió": st.column_config.TextColumn("Recibió", width="small"),
                "Observaciones": st.column_config.TextColumn("Observaciones", width="large"),
                "✓ Hecho": st.column_config.CheckboxColumn("✓ Hecho"),
            },
        )
        if st.button("💾 Guardar registro del día", type="primary"):
            # Reemplazar registros de esa fecha+día
            registros = registros[~((registros["fecha"] == f_str) & (registros["dia"] == dia_sel))]
            nuevos = [
                {
                    "fecha": f_str, "dia": dia_sel,
                    "maquina": r["Máquina / Equipo"],
                    "ot": r["OT"], "ejecuto": r["Ejecutó"] or lubricador,
                    "recibio": r["Recibió"], "observaciones": r["Observaciones"],
                }
                for _, r in editado.iterrows() if r["✓ Hecho"]
            ]
            if nuevos:
                registros = pd.concat([registros, pd.DataFrame(nuevos)], ignore_index=True)
                guardar_datos(registros, checks)
                st.success(f"✅ Se guardaron {len(nuevos)} registros del {dia_sel} {f_str}.")
            else:
                st.warning("Marca al menos una máquina como '✓ Hecho' para guardar.")

# ---------------------------------------------------------------------------
# PESTAÑA 2: CHECKLISTS
# ---------------------------------------------------------------------------
with tab_check:
    areas = sorted({b["area"] for b in bloques})
    area_sel = st.selectbox("Área", areas)
    bloques_area = [b for b in bloques if b["area"] == area_sel]
    bloque_sel = st.selectbox(
        "Bloque / Grupo",
        range(len(bloques_area)),
        format_func=lambda i: f"{bloques_area[i]['titulo']}  ({len(bloques_area[i]['maquinas'])} máquinas)",
    )
    bloque = bloques_area[bloque_sel]

    st.subheader(f"{area_sel} — {bloque['titulo']}")
    ejecutor = st.text_input("Ejecutado por", value=lubricador, key="ejec_check")
    st.caption("Marca ✓ en cada actividad realizada por máquina.")

    # Cuadrícula actividad x máquina
    f_str = fecha_trabajo.strftime("%Y-%m-%d")
    prev = checks[(checks["fecha"] == f_str) &
                  (checks["area"] == area_sel) &
                  (checks["bloque"] == bloque["titulo"])]
    datos = {}
    for act in bloque["actividades"]:
        fila = []
        for m in bloque["maquinas"]:
            coinciden = prev[
                (prev["maquina"].apply(normalizar) == normalizar(m)) &
                (prev["actividad"] == act["actividad"])
            ]
            fila.append(bool(len(coinciden) and coinciden["ok"].iloc[0]))
        datos[f"{act['actividad']}  [{act['tiempo']} min]" if pd.notna(act["tiempo"]) else act["actividad"]] = fila

    df_grid = pd.DataFrame(datos, index=bloque["maquinas"]).T
    config = {"_index": st.column_config.Column("Actividad", width="large")}
    for m in bloque["maquinas"]:
        config[m] = st.column_config.CheckboxColumn(m, width="small")
    grid = st.data_editor(
        df_grid.reset_index().rename(columns={"index": "Actividad"}),
        hide_index=True, use_container_width=True, column_config=config,
    )

    hechos = int(grid[bloque["maquinas"]].sum().sum())
    total = len(bloque["actividades"]) * len(bloque["maquinas"])
    st.progress(hechos / total if total else 0, text=f"Avance del bloque: {hechos}/{total}")

    if st.button("💾 Guardar checklist", type="primary"):
        checks = checks[~((checks["fecha"] == f_str) &
                          (checks["area"] == area_sel) &
                          (checks["bloque"] == bloque["titulo"]))]
        nuevos = []
        for _, fila in grid.iterrows():
            act_nombre = fila["Actividad"].split("  [")[0]
            tiempo = next((a["tiempo"] for a in bloque["actividades"] if a["actividad"] == act_nombre), None)
            for m in bloque["maquinas"]:
                if fila[m]:
                    nuevos.append({
                        "fecha": f_str, "area": area_sel, "bloque": bloque["titulo"],
                        "maquina": m, "actividad": act_nombre, "tiempo": tiempo,
                        "ok": True, "ejecutado_por": ejecutor,
                    })
        if nuevos:
            checks = pd.concat([checks, pd.DataFrame(nuevos)], ignore_index=True)
            guardar_datos(registros, checks)
            st.success(f"✅ Checklist guardado: {len(nuevos)} actividades marcadas.")
        else:
            st.warning("No marcaste ninguna actividad.")

# ---------------------------------------------------------------------------
# PESTAÑA 3: CRONOGRAMA
# ---------------------------------------------------------------------------
with tab_crono:
    st.subheader("Cronograma semanal de ejecución")
    # Semana (Lunes a Sábado) de la fecha seleccionada
    lunes = fecha_trabajo - timedelta(days=fecha_trabajo.weekday())
    fechas_semana = [(lunes + timedelta(days=i)) for i in range(6)]

    c1, c2, c3 = st.columns(3)
    f_str = fecha_trabajo.strftime("%Y-%m-%d")
    hechos_dia = len(registros[registros["fecha"] == f_str])
    checks_dia = len(checks[checks["fecha"] == f_str])
    c1.metric(f"Registros del {fecha_trabajo.strftime('%d/%m')}", hechos_dia)
    c2.metric(f"Actividades checklist del {fecha_trabajo.strftime('%d/%m')}", checks_dia)
    c3.metric("Semana", f"{lunes.strftime('%d/%m')} – {(lunes + timedelta(days=5)).strftime('%d/%m')}")

    st.markdown("**Avance por día (ruta del horario)**")
    filas_crono = []
    for d in horario:
        fila = {"Día": d["dia"]}
        for fecha in fechas_semana:
            if fecha.strftime("%A") != dia_a_ingles(d["dia"]):
                continue
            f = fecha.strftime("%Y-%m-%d")
            regs_dia = registros[registros["fecha"] == f]
            hechas = sum(
                1 for m in d["maquinas"]
                if len(regs_dia[regs_dia["maquina"].apply(normalizar) == normalizar(m)])
            )
            pct = hechas / len(d["maquinas"]) if d["maquinas"] else 0
            fila[fecha.strftime("%d/%m")] = f"{'🟢' if pct == 1 else '🟡' if pct > 0 else '⚪'} {hechas}/{len(d['maquinas'])}"
        filas_crono.append(fila)
    st.dataframe(pd.DataFrame(filas_crono).set_index("Día"), use_container_width=True)

    st.markdown("**Actividades por área (checklists)**")
    if len(checks):
        resumen = (
            checks.groupby(["fecha", "area"]).size().unstack(fill_value=0)
            .reindex([f.strftime("%Y-%m-%d") for f in fechas_semana], fill_value=0)
        )
        resumen.index = [f.strftime("%a %d/%m") for f in fechas_semana]
        st.dataframe(resumen, use_container_width=True)
    else:
        st.info("Aún no hay checklists registrados.")

# ---------------------------------------------------------------------------
# PESTAÑA 4: HISTORIAL
# ---------------------------------------------------------------------------
with tab_hist:
    st.subheader("Historial completo de registros")
    if not len(registros) and not len(checks):
        st.info("Todavía no hay registros guardados.")
    else:
        tipo = st.radio("Ver historial de", ["Ruta (horario)", "Checklists"], horizontal=True)
        df_hist = registros if tipo == "Ruta (horario)" else checks
        if len(df_hist):
            c1, c2 = st.columns(2)
            fechas_hist = sorted(df_hist["fecha"].unique())
            f_ini = c1.selectbox("Desde", fechas_hist, index=0)
            f_fin = c2.selectbox("Hasta", fechas_hist, index=len(fechas_hist) - 1)
            filtrado = df_hist[(df_hist["fecha"] >= f_ini) & (df_hist["fecha"] <= f_fin)]
            st.dataframe(filtrado, use_container_width=True, hide_index=True)
            st.download_button(
                "⬇️ Descargar historial filtrado (Excel)",
                data=to_excel_bytes(filtrado),
                file_name=f"historial_{tipo.split()[0].lower()}_{f_ini}_{f_fin}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.info(f"No hay registros de {tipo.lower()} todavía.")
