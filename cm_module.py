# -*- coding: utf-8 -*-
"""Módulo de Centros de Maestros — DFC/SEJ.
Piloto de toma de lista de asesores (ENEG, Autlán) y hogar de las futuras
funciones CM (solicitudes a nombre del asesor, reportes de asistencia).
Dependencias inyectadas vía configurar() — mismo patrón que tools.py del
agente de constancias. NO importa nada de app_incidencias (sin import circular).
"""
import streamlit as st
import pandas as pd
import gspread
import pytz
from datetime import datetime

_deps = {}

def configurar(get_client, error_amable, columnas_horario, subir_archivo_drive, carpeta_asistencia_cm):
    """Llamar UNA vez desde app_incidencias después de definir dependencias."""
    _deps["get_client"] = get_client
    _deps["error_amable"] = error_amable
    _deps["COLUMNAS_HORARIO"] = columnas_horario
    _deps["subir_archivo_drive"] = subir_archivo_drive
    _deps["carpeta_asistencia_cm"] = carpeta_asistencia_cm

def get_client():
    return _deps["get_client"]()

def _error_amable(e, contexto=""):
    return _deps["error_amable"](e, contexto)


# ═══════════════════════════════════════════════════════════════════
# PILOTO — TOMA DE LISTA DE ASESORES (Centros de Maestros)
# ═══════════════════════════════════════════════════════════════════
# Sin reloj checador para asesores: el responsable de cada Centro registra
# quién asistió, día por día. Empieza con 2 centros piloto (ENEG, Autlán).
#
# SECRETS NECESARIOS (mapea RFC del responsable -> nombre EXACTO del centro,
# igual a como aparece en la columna AREA de la tab Directorio):
#   [responsables_cm_piloto]
#   RFC_RESPONSABLE_ENEG   = "Centro de Maestros ENEG 1415"
#   RFC_RESPONSABLE_AUTLAN = "Centro de Maestros AUTLÁN"

HORARIOS_ASESORES_TAB = "Horarios_Asesores"
def _horarios_headers():
    return ["RFC_ASESOR", "NOMBRE", "CENTRO", "ACTIVO"] + \
        [c for par in _deps["COLUMNAS_HORARIO"].values() for c in par if "SAB" not in c and "DOM" not in c]

ASISTENCIA_ASESORES_TAB = "Asistencia_Asesores"
ASISTENCIA_ASESORES_HEADERS = ["FECHA", "CENTRO", "RFC_ASESOR", "NOMBRE_ASESOR",
                                "ESTADO", "MOTIVO", "REGISTRADO_POR", "FECHA_REGISTRO"]

def _ws_horarios_asesores():
    client = get_client()
    sh = client.open_by_key(st.secrets["sheet_checador_id"])
    try:
        return sh.worksheet(HORARIOS_ASESORES_TAB)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(HORARIOS_ASESORES_TAB, rows=2, cols=len(_horarios_headers()))
        ws.append_row(_horarios_headers())
        return ws

def _ws_asistencia_asesores():
    client = get_client()
    sh = client.open_by_key(st.secrets["sheet_checador_id"])
    try:
        return sh.worksheet(ASISTENCIA_ASESORES_TAB)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(ASISTENCIA_ASESORES_TAB, rows=2, cols=len(ASISTENCIA_ASESORES_HEADERS))
        ws.append_row(ASISTENCIA_ASESORES_HEADERS)
        return ws

@st.cache_data(ttl=120)
def _cargar_asesores_centro(centro: str):
    try:
        ws = _ws_horarios_asesores()
        data = ws.get_all_records(numericise_ignore=["all"])
        df = pd.DataFrame(data) if data else pd.DataFrame(columns=_horarios_headers())
        if df.empty:
            return df
        return df[df["CENTRO"].astype(str).str.strip() == centro.strip()]
    except Exception:
        return pd.DataFrame(columns=_horarios_headers())

def _centro_del_responsable(rfc_actual: str):
    try:
        mapa = dict(st.secrets.get("responsables_cm_piloto", {}))
        return mapa.get(rfc_actual)
    except Exception:
        return None

def registrar_asistencia_dia(fecha_iso: str, centro: str, registros: list[dict], registrado_por: str):
    """registros: [{RFC_ASESOR, NOMBRE, ESTADO, MOTIVO}, ...].
    Si ya existe un registro para esa FECHA+RFC_ASESOR, lo ACTUALIZA en vez
    de duplicar (localizado en vivo, mismo patrón que folios/emergencias:
    nunca por índice de fila cacheado)."""
    ws = _ws_asistencia_asesores()
    data = ws.get_all_records(numericise_ignore=["all"])
    idx_existente = {}
    for i, row in enumerate(data, start=2):
        clave = (str(row.get("FECHA", "")).strip(), str(row.get("RFC_ASESOR", "")).strip().upper())
        idx_existente[clave] = i
    ahora = datetime.now(pytz.timezone("America/Mexico_City")).strftime("%Y-%m-%d %H:%M")
    from gspread.cell import Cell
    actualizaciones, nuevas = [], []
    for r in registros:
        rfc_a = str(r.get("RFC_ASESOR", "")).strip().upper()
        clave = (fecha_iso, rfc_a)
        fila_vals = [fecha_iso, centro, rfc_a, r.get("NOMBRE", ""), r.get("ESTADO", ""),
                     r.get("MOTIVO", ""), registrado_por, ahora]
        if clave in idx_existente:
            fila_num = idx_existente[clave]
            for col, val in enumerate(fila_vals, start=1):
                actualizaciones.append(Cell(fila_num, col, val))
        else:
            nuevas.append(fila_vals)
    if actualizaciones:
        ws.update_cells(actualizaciones, value_input_option="USER_ENTERED")
    if nuevas:
        ws.append_rows(nuevas, value_input_option="USER_ENTERED")
    _cargar_asistencia_reciente.clear()

@st.cache_data(ttl=60)
def _cargar_asistencia_reciente(centro: str, dias: int = 7):
    try:
        ws = _ws_asistencia_asesores()
        data = ws.get_all_records(numericise_ignore=["all"])
        df = pd.DataFrame(data) if data else pd.DataFrame(columns=ASISTENCIA_ASESORES_HEADERS)
        if df.empty:
            return df
        df = df[df["CENTRO"].astype(str).str.strip() == centro.strip()]
        df["_F"] = pd.to_datetime(df["FECHA"], errors="coerce")
        return df.sort_values("_F", ascending=False).head(dias * 10)
    except Exception:
        return pd.DataFrame(columns=ASISTENCIA_ASESORES_HEADERS)


# ── Asistencia MENSUAL en PDF (documento firmado, complementa — no
# reemplaza — la toma de lista diaria digital de arriba) ──
ASISTENCIA_MENSUAL_TAB = "Asistencia_Mensual_CM"
ASISTENCIA_MENSUAL_HEADERS = ["ID", "CENTRO", "RESPONSABLE_RFC", "RESPONSABLE_NOMBRE",
                               "PERIODO", "URL_ARCHIVO", "FECHA_SUBIDA"]

def _ws_asistencia_mensual():
    client = get_client()
    sh = client.open_by_key(st.secrets["sheet_checador_id"])
    try:
        return sh.worksheet(ASISTENCIA_MENSUAL_TAB)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(ASISTENCIA_MENSUAL_TAB, rows=2, cols=len(ASISTENCIA_MENSUAL_HEADERS))
        ws.append_row(ASISTENCIA_MENSUAL_HEADERS)
        return ws

@st.cache_data(ttl=60)
def _cargar_asistencia_mensual(centro: str):
    try:
        ws = _ws_asistencia_mensual()
        data = ws.get_all_records(numericise_ignore=["all"])
        df = pd.DataFrame(data) if data else pd.DataFrame(columns=ASISTENCIA_MENSUAL_HEADERS)
        if df.empty:
            return df
        return df[df["CENTRO"].astype(str).str.strip() == centro.strip()]
    except Exception:
        return pd.DataFrame(columns=ASISTENCIA_MENSUAL_HEADERS)

def registrar_asistencia_mensual(centro: str, rfc: str, nombre: str, periodo: str, url: str):
    ws = _ws_asistencia_mensual()
    todas = ws.get_all_records(numericise_ignore=["all"])
    nuevo_id = len(todas) + 1
    ahora = datetime.now(pytz.timezone("America/Mexico_City")).strftime("%Y-%m-%d %H:%M")
    ws.append_row([nuevo_id, centro, rfc, nombre, periodo, url, ahora], value_input_option="USER_ENTERED")
    _cargar_asistencia_mensual.clear()

def _tab_asistencia_mensual(centro: str, rfc_actual: str):
    st.caption("Sube la lista de asistencia mensual firmada de tu Centro, en PDF. "
              "Se guarda en la unidad compartida, en una carpeta aparte de los "
              "anexos de incidencias.")
    hoy = datetime.now(pytz.timezone("America/Mexico_City")).date()
    periodo = st.text_input("Periodo que cubre esta lista", value=hoy.strftime("%B %Y"),
                            placeholder="Ej: Julio 2026")
    archivo = st.file_uploader("Lista de asistencia (PDF)", type=["pdf"])

    if archivo and st.button("📎 Subir lista de asistencia", type="primary", use_container_width=True):
        if not periodo.strip():
            st.warning("Indica a qué periodo corresponde la lista antes de subirla.")
        else:
            with st.spinner("Subiendo a la unidad compartida..."):
                nombre_arch = f"Asistencia_{centro}_{periodo}_{rfc_actual}.pdf".replace(" ", "_")
                url = _deps["subir_archivo_drive"](archivo, nombre_arch, _deps["carpeta_asistencia_cm"])
            # El bug clásico de esta función es tratar "ERROR: ..." como si
            # fuera un link válido. Aquí se checa explícito, no se repite.
            if url.startswith("ERROR:"):
                _error_amable(Exception(url), "al subir la lista de asistencia")
            else:
                registrar_asistencia_mensual(centro, rfc_actual,
                                             st.session_state.get("nombre", rfc_actual), periodo, url)
                st.success(f"Lista de {periodo} subida y registrada correctamente.")

    hist = _cargar_asistencia_mensual(centro)
    if not hist.empty:
        st.markdown("#### Listas ya subidas de este centro")
        for _, r in hist.sort_values("FECHA_SUBIDA", ascending=False).iterrows():
            st.markdown(f"📄 **{r['PERIODO']}** — subida por {r['RESPONSABLE_NOMBRE']} "
                       f"el {r['FECHA_SUBIDA']} · [Ver archivo]({r['URL_ARCHIVO']})")


def vista_toma_lista_cm():
    rfc_actual = str(st.session_state.get("rfc", "")).upper().strip()
    centro = _centro_del_responsable(rfc_actual)
    if st.session_state.get("rol") != "admin" and not centro:
        st.error("No tienes permiso para esta sección.")
        return
    if st.session_state.get("rol") == "admin" and not centro:
        centros_disp = list(dict(st.secrets.get("responsables_cm_piloto", {})).values())
        if not centros_disp:
            st.info("Aún no hay centros piloto configurados en secrets.")
            return
        centro = st.selectbox("Centro (vista admin)", centros_disp)

    st.markdown(f"## 📋 {centro}")
    tab_diaria, tab_mensual = st.tabs(["📋 Toma de lista diaria", "📎 Asistencia mensual (PDF)"])

    with tab_diaria:
        hoy = datetime.now(pytz.timezone("America/Mexico_City")).date()
        fecha_sel = st.date_input("Fecha", value=hoy, max_value=hoy)

        asesores = _cargar_asesores_centro(centro)
        if asesores.empty:
            st.warning(f"No hay asesores registrados para '{centro}' en la tab {HORARIOS_ASESORES_TAB}. "
                      "Agrega sus filas (RFC_ASESOR, NOMBRE, CENTRO) antes de tomar lista.")
        else:
            st.caption("Marca el estado de cada asesor. Si falta o justifica, agrega un motivo breve.")
            registros = []
            for _, a in asesores.iterrows():
                nombre_a = str(a["NOMBRE"]).strip()
                rfc_a = str(a["RFC_ASESOR"]).strip()
                c1, c2 = st.columns([2, 3])
                with c1:
                    estado = st.radio(nombre_a, ["Asistió", "Faltó", "Justificado"],
                                      key=f"est_{rfc_a}_{fecha_sel}", horizontal=True)
                with c2:
                    motivo = ""
                    if estado != "Asistió":
                        motivo = st.text_input("Motivo", key=f"mot_{rfc_a}_{fecha_sel}",
                                               placeholder="Breve motivo (opcional)")
                registros.append({"RFC_ASESOR": rfc_a, "NOMBRE": nombre_a,
                                  "ESTADO": estado.upper(), "MOTIVO": motivo})

            if st.button("💾 Guardar asistencia del día", type="primary", use_container_width=True):
                try:
                    registrar_asistencia_dia(fecha_sel.isoformat(), centro, registros,
                                             st.session_state.get("nombre", rfc_actual))
                    st.success(f"Asistencia del {fecha_sel.strftime('%d/%m/%Y')} guardada para {len(registros)} asesor(es).")
                except Exception as e:
                    _error_amable(e, "al guardar la asistencia")

            with st.expander("📅 Historial reciente de este centro"):
                hist = _cargar_asistencia_reciente(centro)
                if hist.empty:
                    st.caption("Sin registros previos.")
                else:
                    st.dataframe(hist[["FECHA","NOMBRE_ASESOR","ESTADO","MOTIVO","REGISTRADO_POR"]],
                                use_container_width=True, hide_index=True)

    with tab_mensual:
        _tab_asistencia_mensual(centro, rfc_actual)
