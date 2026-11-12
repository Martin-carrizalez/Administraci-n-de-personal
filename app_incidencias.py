import streamlit as st
import os
# Import protegido: si nomina_module falta o truena, la app sigue funcionando
# (solo se deshabilita esa pestaña). Un módulo secundario no debe tumbar el login.
try:
    from nomina_module import render_pendientes_nomina
except Exception as _e_nomina:
    render_pendientes_nomina = None
    _ERROR_NOMINA = str(_e_nomina)
# Mismo import protegido: si el módulo falta, solo se apaga su botón.
try:
    from pendientes_module import render_pendientes_rh
except Exception as _e_prh:
    render_pendientes_rh = None
    _ERROR_PENDIENTES_RH = str(_e_prh)
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import re
import unicodedata
from datetime import datetime, date, timedelta
from io import BytesIO
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
import qrcode
from datetime import timezone
import pytz
from reportlab.platypus import Image as RLImage

# ─────────────────────────────────────────────
# FLAGS DE FUNCIONALIDAD
# ─────────────────────────────────────────────
HABILITAR_CUMPLEANOS = False  # Cambiar a True cuando Ángel lo autorice

# ─────────────────────────────────────────────
# CONSTANTES
# ─────────────────────────────────────────────
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

TIPO_LABELS = {
    "ECO": "Día económico",
    "PSE": "Pase de salida sin retorno",
    "PSR": "Pase de salida con retorno",
    "PEN": "Pase de entrada",
    "COM": "Comisión",
    "CHO": "Cambio de horario",
    "RGU": "Reposición de guardias",
}
if HABILITAR_CUMPLEANOS:
    TIPO_LABELS["CUM"] = "Día de cumpleaños"

DIAS_SEMANA = ["LUN", "MAR", "MIE", "JUE", "VIE"]

def formato_jefe_pdf(jefe_raw: str) -> str:
    """Convierte el valor de la columna JEFE_INMEDIATO del Sheet Usuarios
    (ej. 'Maricela Esquivel Domínguez — Dir. Desarrollo Académico')
    al formato del PDF con salto de línea entre nombre y cargo.
    El catálogo de jefes vive en el Sheet, no en el código."""
    jefe_raw = str(jefe_raw or "").strip()
    if not jefe_raw:
        return ""
    # Separa nombre y cargo por el guion largo (—) o el guion normal (-) con espacios
    for sep in (" — ", " – ", " - "):
        if sep in jefe_raw:
            nombre, cargo = jefe_raw.split(sep, 1)
            return f"{nombre.strip()}<br/>{cargo.strip()}"
    return jefe_raw

def _dato_institucional(clave: str, default: str = "") -> str:
    """Dato institucional (nombre/cargo de la directora, etc.) leído de
    secrets — NUNCA hardcodeado en el código que se sube a GitHub."""
    try:
        return str(st.secrets.get(clave, default)).strip()
    except Exception:
        return default

# En secrets: directora_nombre, directora_cargo
DIRECTORA_NOMBRE = _dato_institucional("directora_nombre", "(directora_nombre no configurado en secrets)")
DIRECTORA_CARGO  = _dato_institucional("directora_cargo", "(directora_cargo no configurado en secrets)")
NOMBRE_DIRECTORA = f"{DIRECTORA_NOMBRE}<br/>{DIRECTORA_CARGO}"
def _drive_folder(clave: str) -> str:
    """ID de carpeta de Drive leído de secrets — nunca hardcodeado en el
    código que se sube a GitHub. Si falta, devuelve '' (quien lo use debe
    mostrar aviso claro, no tronar la app entera)."""
    try:
        return str(st.secrets.get(clave, "")).strip()
    except Exception:
        return ""

# IDs reales van en secrets: drive_anexos_folder_id, drive_asistencia_cm_folder_id
DRIVE_ANEXOS_FOLDER = _drive_folder("drive_anexos_folder_id")
DRIVE_ASISTENCIA_CM_FOLDER = _drive_folder("drive_asistencia_cm_folder_id")

COLUMNAS_HORARIO = {
    "LUN": ("ENTRADA_LUN", "SALIDA_LUN"),
    "MAR": ("ENTRADA_MAR", "SALIDA_MAR"),
    "MIE": ("ENTRADA_MIE", "SALIDA_MIE"),
    "JUE": ("ENTRADA_JUE", "SALIDA_JUE"),
    "VIE": ("ENTRADA_VIE", "SALIDA_VIE"),
    "SAB": ("ENTRADA_SAB", "SALIDA_SAB"),
    "DOM": ("ENTRADA_DOM", "SALIDA_DOM"),
}

COLS_INCIDENCIAS = [
    "ID", "FOLIO", "RFC", "NOMBRE", "TIPO", "FECHA_SOLICITUD",
    "FECHA_INICIO", "FECHA_FIN", "DIAS", "HORAS_PASE", "HORA_RETORNO",
    "MOTIVO", "TIENE_ANEXO", "LINK_ANEXO", "ESTADO", "AUTORIZADO_POR",
    "FECHA_AUTORIZACION", "OBSERVACIONES"
]

# ─────────────────────────────────────────────
# CONEXIÓN GOOGLE SHEETS
# ─────────────────────────────────────────────
@st.cache_resource
def get_client():
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=SCOPES
    )
    return gspread.authorize(creds)

@st.cache_data(ttl=300)
def cargar_empleados():
    client = get_client()
    sh = client.open_by_key(st.secrets["sheet_economicos_id"])
    ws = sh.worksheet("Empleados")
    data = ws.get_all_records(numericise_ignore=["all"])
    return pd.DataFrame(data) if data else pd.DataFrame(columns=[
        "ID", "RFC", "CURP", "PATERNO", "MATERNO", "NOMBRE",
        "PLAZA", "PUESTO", "BASE/INTERINO", "QNA FIN", "C.C.T.",
        "CENTRO DE TRABAJO", "DIAS DISPONIBLES", "DIAS TOTALES"
    ])

COL_CORREO_USR = "Correo electrónico institucional"


@st.cache_data(ttl=300)
def cargar_usuarios():
    """Quién puede iniciar sesión.

    FASE 3 — Se compone del padrón MÁS la tab Usuarios. La tab solo cubría al
    personal de oficinas centrales: 73 de los 76 asesores de Centros de
    Maestros no podían entrar, y sin sesión no pueden pedir un pase ni
    registrar su asistencia. El padrón sí trae su correo institucional.

    La tab Usuarios manda sobre el padrón cuando ambos tienen a la persona:
    ahí es donde RH corrige un correo o un jefe sin esperar al padrón.
    """
    client = get_client()
    sh = client.open_by_key(st.secrets["sheet_checador_id"])
    try:
        ws = sh.worksheet("Usuarios")
        data = ws.get_all_records(numericise_ignore=["all"])
        tab = pd.DataFrame(data) if data else pd.DataFrame(columns=["RFC", COL_CORREO_USR])
    except Exception:
        tab = pd.DataFrame(columns=["RFC", COL_CORREO_USR])

    padron = cargar_padron()
    if padron.empty or "RFC" not in padron.columns or "CORREO" not in padron.columns:
        st.session_state["_fuente_usuarios"] = "tab_Usuarios"
        return tab

    desde_padron = padron[padron["CORREO"].astype(str).str.strip() != ""].copy()
    # Excluir comisionados: cobran en DFC pero trabajan físicamente en otra
    # área (PRESTADO_A...). No necesitan acceso a esta app — no están aquí
    # para checar, pedir pases de salida, ni tomar lista.
    # Excluir comisionados: cobran en DFC pero trabajan físicamente en otra
    # área. La etiqueta real en el Sheet es COMISIONADO_FUERA (confirmado
    # por Ángel) — se deja también PRESTADO_A por si algún registro viejo
    # usa esa otra convención, pero la que manda es COMISIONADO_FUERA.
    if "PRESTADO" in desde_padron.columns:
        _valor = desde_padron["PRESTADO"].astype(str).str.strip().str.upper()
        _es_comisionado = _valor.str.startswith("PRESTADO_A") | _valor.str.contains("COMISIONADO_FUERA")
        desde_padron = desde_padron[~_es_comisionado]
    if desde_padron.empty:
        st.session_state["_fuente_usuarios"] = "tab_Usuarios"
        return tab

    desde_padron[COL_CORREO_USR] = desde_padron["CORREO"]
    for col in ("JEFE_INMEDIATO", "DIRECTOR_GENERAL"):
        if col not in desde_padron.columns:
            desde_padron[col] = ""
    desde_padron = desde_padron[["RFC", COL_CORREO_USR,
                                 "JEFE_INMEDIATO", "DIRECTOR_GENERAL"]]

    ya = set(tab["RFC"].astype(str).str.upper().str.strip()) if "RFC" in tab.columns else set()
    nuevos = desde_padron[~desde_padron["RFC"].astype(str).str.upper().str.strip().isin(ya)]
    if nuevos.empty:
        st.session_state["_fuente_usuarios"] = "tab_Usuarios"
        return tab

    st.session_state["_fuente_usuarios"] = f"tab_Usuarios + padrón (+{len(nuevos)})"
    return pd.concat([tab, nuevos], ignore_index=True).fillna("")

@st.cache_data(ttl=300)
def cargar_solicitudes_eco():
    client = get_client()
    sh = client.open_by_key(st.secrets["sheet_economicos_id"])
    ws = sh.worksheet("Solicitudes")
    data = ws.get_all_records(numericise_ignore=["all"])
    return pd.DataFrame(data) if data else pd.DataFrame()

@st.cache_data(ttl=300)
def cargar_horarios():
    """Horarios vigentes del personal.

    Compone la tab `empleados` MÁS el padrón. La tab solo cubre a quienes
    marcan en el reloj de piso 5 y 6; el personal de Centros de Maestros no
    aparece ahí y su consulta de horario salía vacía aunque el padrón sí lo
    tuviera. El padrón usa los mismos nombres de columna, así que se anexa
    tal cual.

    La tab manda cuando ambos tienen a la persona: es donde queda un cambio
    de horario autorizado antes de que llegue al padrón.
    """
    client = get_client()
    sh = client.open_by_key(st.secrets["sheet_checador_id"])
    try:
        ws = sh.worksheet("empleados")
        data = ws.get_all_records(numericise_ignore=["all"])
        tab = pd.DataFrame(data) if data else pd.DataFrame()
    except Exception:
        tab = pd.DataFrame()

    padron = cargar_padron()
    cols_h = [c for par in COLUMNAS_HORARIO.values() for c in par]
    if padron.empty or "RFC" not in padron.columns:
        return tab
    if not any(c in padron.columns for c in cols_h):
        return tab

    # Solo quien tenga al menos un horario capturado: una fila sin horas no
    # aporta nada y ensuciaría la consulta.
    presentes = [c for c in cols_h if c in padron.columns]
    con_horario = padron[
        padron[presentes].apply(
            lambda r: any(str(v).strip() for v in r), axis=1)].copy()
    if con_horario.empty:
        return tab

    for c in cols_h:
        if c not in con_horario.columns:
            con_horario[c] = ""
    if "NOMBRE" not in con_horario.columns and "NOMBRE_COMPLETO" in con_horario.columns:
        con_horario["NOMBRE"] = con_horario["NOMBRE_COMPLETO"]
    con_horario["ACTIVO"] = con_horario.get("ACTIVO_CHECADOR", "SI")
    con_horario = con_horario[["RFC", "NOMBRE", "ACTIVO"] + cols_h]

    ya = set(tab["RFC"].astype(str).str.upper().str.strip()) if "RFC" in tab.columns else set()
    nuevos = con_horario[~con_horario["RFC"].astype(str).str.upper().str.strip().isin(ya)]
    if nuevos.empty:
        return tab
    if tab.empty:
        return nuevos.reset_index(drop=True)
    return pd.concat([tab, nuevos], ignore_index=True).fillna("")

COLS_NOMINA = ["ID", "NOMBRE_COMPLETO", "CORREO",
               "JEFE_INMEDIATO", "CORREO_JEFE", "CC_FIJO", "CCT_NOMINA"]


@st.cache_data(ttl=300)
def cargar_directorio_nomina():
    """Personal que va en nómina de la DFC.

    FASE 2 — Se deriva del padrón filtrando EN_NOMINA_DFC == "SI", en lugar de
    leerse de una tab capturada a mano. Ese cambio es el que resuelve el
    problema de que la lista se quedara corta: ya no puede desincronizarse del
    padrón porque ES el padrón filtrado.

    Requiere que el padrón traiga JEFE_INMEDIATO, CORREO_JEFE y CC_FIJO. Si
    esas columnas están vacías o no existen, se usa la tab antigua: es
    preferible la lista de siempre a una sin jefes ni copias.
    """
    def _tab_antigua():
        client = get_client()
        sh = client.open_by_key(st.secrets["sheet_checador_id"])
        try:
            ws = sh.worksheet("Directorio_Nomina")
        except gspread.WorksheetNotFound:
            return pd.DataFrame(columns=COLS_NOMINA)
        data = ws.get_all_records(numericise_ignore=["all"])
        return pd.DataFrame(data) if data else pd.DataFrame(columns=COLS_NOMINA)

    padron = cargar_padron()
    c_nom = _col(padron, *ALIAS_PADRON["NOMBRE"]) if not padron.empty else None
    tiene_jefes = (c_nom is not None
                   and "JEFE_INMEDIATO" in padron.columns
                   and padron["JEFE_INMEDIATO"].astype(str).str.strip().ne("").any())

    if not (c_nom and "EN_NOMINA_DFC" in padron.columns and tiene_jefes):
        st.session_state["_fuente_nomina"] = "tab_Directorio_Nomina"
        return _tab_antigua()

    df = padron[padron["EN_NOMINA_DFC"].astype(str).str.upper().str.strip() == "SI"].copy()
    df["NOMBRE_COMPLETO"] = df[c_nom]
    # El padrón no tiene columna ID: sin esto, TODOS quedaban con ID vacío y el
    # módulo de nómina confundía a cada empleado con el primero de la lista.
    # El RFC es la llave única del padrón; si falta, se usa el nombre.
    _c_rfc = _col(padron, "RFC")
    if "ID" not in df.columns or df["ID"].astype(str).str.strip().eq("").all():
        _rfc = df[_c_rfc].astype(str).str.strip().str.upper() if _c_rfc else ""
        df["ID"] = [r if r else f"NOM:{n}" for r, n in
                    zip(_rfc if _c_rfc else [""] * len(df), df["NOMBRE_COMPLETO"].astype(str))]
    # CCT_NOMINA dice en qué centro(s) cobra cada quien ("14ADG1075P" o
    # "14ADG1075P / 14FMP0001B"). Sin esto, la precarga metía a TODOS en las
    # dos nóminas, aunque cobraran en una sola.
    _c_cct = _col(padron, "CCT_NOMINA")
    if _c_cct and "CCT_NOMINA" not in df.columns:
        df["CCT_NOMINA"] = padron.loc[df.index, _c_cct].astype(str)
    for col in COLS_NOMINA:
        if col not in df.columns:
            df[col] = ""
    df = df[COLS_NOMINA].fillna("")

    # CC_FIJO es una lista fija de correos en copia y no se deriva de ningún
    # dato del padrón. Si viene vacía, se toma de la tab antigua: perder las
    # copias sería una regresión silenciosa en los correos de nómina.
    if df["CC_FIJO"].astype(str).str.strip().eq("").all():
        try:
            previo = _tab_antigua()
            if not previo.empty and "CC_FIJO" in previo.columns:
                cc = [c for c in previo["CC_FIJO"].astype(str) if c.strip()]
                for i, valor in enumerate(cc):
                    if i < len(df):
                        df.iat[i, df.columns.get_loc("CC_FIJO")] = valor
                st.session_state["_fuente_nomina"] = "padron (CC_FIJO de la tab antigua)"
                return df
        except Exception:
            pass

    st.session_state["_fuente_nomina"] = "padron"
    return df

@st.cache_data(ttl=300)
def cargar_incidencias():
    client = get_client()
    sh = client.open_by_key(st.secrets["sheet_checador_id"])
    ws = sh.worksheet("Incidencias")
    data = ws.get_all_records(numericise_ignore=["all"])
    return pd.DataFrame(data) if data else pd.DataFrame(columns=COLS_INCIDENCIAS)

@st.cache_data(ttl=300)
def cargar_historial_horarios():
    client = get_client()
    sh = client.open_by_key(st.secrets["sheet_checador_id"])
    ws = sh.worksheet("HISTORIAL_HORARIOS")
    data = ws.get_all_records(numericise_ignore=["all"])
    return pd.DataFrame(data) if data else pd.DataFrame(columns=[
        "RFC","NOMBRE","FECHA_INICIO","FECHA_FIN",
        "ENTRADA_LUN","SALIDA_LUN","ENTRADA_MAR","SALIDA_MAR",
        "ENTRADA_MIE","SALIDA_MIE","ENTRADA_JUE","SALIDA_JUE",
        "ENTRADA_VIE","SALIDA_VIE"
    ])

def _leer_hoja_con_reintento(ws, intentos=3, espera=5):
    """Lee get_all_records reintentando si Google devuelve 429 (cuota excedida)."""
    import time as _time
    for i in range(intentos):
        try:
            return ws.get_all_records()
        except gspread.exceptions.APIError as e:
            if "429" in str(e) and i < intentos - 1:
                _time.sleep(espera)
                continue
            raise
    return []

def cargar_festivos():
    """Lee la tab festivos con el MISMO método que el conteo de faltas
    (load_fest_ch), que sí funciona. Sin cache para evitar servir vacíos."""
    import time as _time
    for intento in range(3):
        try:
            client = get_client()
            ws = client.open_by_key(st.secrets["sheet_checador_id"]).worksheet("festivos")
            data = ws.get_all_records()
            return pd.DataFrame(data).fillna("") if data else pd.DataFrame()
        except gspread.exceptions.APIError as e:
            if "429" in str(e) and intento < 2:
                _time.sleep(5)
                continue
            raise
    return pd.DataFrame()

ASISTENCIA_TAB = "Asistencia_Mes"
ASISTENCIA_HEADERS = ["RFC", "NOMBRE", "PERIODO", "FALTAS", "JUSTIFICADAS",
                      "NO_JUSTIFICADAS", "RETARDOS", "DIAS_FALTA", "AVISAR", "FECHA_PROCESO"]

def _ws_asistencia():
    client = get_client()
    sh = client.open_by_key(st.secrets["sheet_checador_id"])
    try:
        return sh.worksheet(ASISTENCIA_TAB)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(ASISTENCIA_TAB, rows=2, cols=len(ASISTENCIA_HEADERS))
        ws.append_row(ASISTENCIA_HEADERS)
        return ws

def guardar_asistencia_mes(filas: list[dict], periodo: str):
    """Reemplaza el contenido de Asistencia_Mes con el último reporte procesado.
    SIN clear() previo: se escribe encima y el sobrante se limpia solo si la
    escritura tuvo éxito. Si Google falla (429), los datos anteriores quedan
    intactos en vez de perder la tab completa."""
    ws = _ws_asistencia()
    rows = [ASISTENCIA_HEADERS]
    ahora = datetime.now(pytz.timezone("America/Mexico_City")).strftime("%Y-%m-%d %H:%M")
    for f in filas:
        rows.append([
            f.get("RFC",""), f.get("NOMBRE",""), periodo,
            f.get("FALTAS",0), f.get("JUSTIFICADAS",0), f.get("NO_JUSTIFICADAS",0),
            f.get("RETARDOS",0), f.get("DIAS_FALTA",""), f.get("AVISAR",""), ahora
        ])
    ws.update(rows, value_input_option="USER_ENTERED")
    # Éxito: limpiar filas viejas que hayan quedado debajo de la tabla nueva
    try:
        if ws.row_count > len(rows):
            ws.batch_clear([f"A{len(rows) + 1}:J{ws.row_count}"])
    except Exception:
        pass
    cargar_asistencia_mes.clear()

@st.cache_data(ttl=300)
def cargar_asistencia_mes():
    try:
        ws = _ws_asistencia()
        data = ws.get_all_records(numericise_ignore=["all"])
        return pd.DataFrame(data) if data else pd.DataFrame(columns=ASISTENCIA_HEADERS)
    except Exception:
        return pd.DataFrame(columns=ASISTENCIA_HEADERS)

# ── Observaciones de faltas: histórico permanente por RFC + fecha ──
OBS_TAB = "Observaciones_Faltas"
OBS_HEADERS = ["RFC", "NOMBRE", "FECHA_FALTA", "OBSERVACION", "FECHA_REGISTRO"]

def _ws_observaciones():
    client = get_client()
    sh = client.open_by_key(st.secrets["sheet_checador_id"])
    try:
        return sh.worksheet(OBS_TAB)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(OBS_TAB, rows=2, cols=len(OBS_HEADERS))
        ws.append_row(OBS_HEADERS)
        return ws

@st.cache_data(ttl=300)
def cargar_observaciones():
    try:
        ws = _ws_observaciones()
        data = ws.get_all_records(numericise_ignore=["all"])
        return pd.DataFrame(data) if data else pd.DataFrame(columns=OBS_HEADERS)
    except Exception:
        return pd.DataFrame(columns=OBS_HEADERS)

def guardar_observaciones(nuevas: list[dict]):
    """Agrega/actualiza observaciones por RFC+fecha sin borrar el histórico.
    nuevas: [{RFC, NOMBRE, FECHA_FALTA, OBSERVACION}, ...]"""
    if not nuevas:
        return
    ws = _ws_observaciones()
    existentes = ws.get_all_records(numericise_ignore=["all"])
    # índice por (RFC, FECHA_FALTA) → número de fila (2-based)
    idx = {}
    for i, row in enumerate(existentes, start=2):
        idx[(str(row.get("RFC","")).upper(), str(row.get("FECHA_FALTA","")))] = i
    ahora = datetime.now(pytz.timezone("America/Mexico_City")).strftime("%Y-%m-%d %H:%M")
    from gspread.cell import Cell
    actualizaciones, nuevas_filas = [], []
    for n in nuevas:
        rfc_n = str(n.get("RFC","")).upper()
        fecha_n = str(n.get("FECHA_FALTA",""))
        obs = str(n.get("OBSERVACION","")).strip()
        if not obs:
            continue
        clave = (rfc_n, fecha_n)
        if clave in idx:
            fila = idx[clave]
            actualizaciones.append(Cell(fila, 4, obs))           # col OBSERVACION
            actualizaciones.append(Cell(fila, 5, ahora))         # col FECHA_REGISTRO
        else:
            nuevas_filas.append([rfc_n, n.get("NOMBRE",""), fecha_n, obs, ahora])
    if actualizaciones:
        ws.update_cells(actualizaciones, value_input_option="USER_ENTERED")
    if nuevas_filas:
        ws.append_rows(nuevas_filas, value_input_option="USER_ENTERED")
    cargar_observaciones.clear()


# ── Justificaciones de Dirección: faltas justificadas con Vo.Bo. de la titular ──
# Se guardan en el Sheet y se aplican ANTES de generar el reporte, para que
# Excel y PDF salgan ya corregidos (adiós edición manual del Excel).
JUSTIF_DIR_TAB = "Justificaciones_Direccion"
JUSTIF_DIR_HEADERS = ["RFC", "NOMBRE", "FECHA", "MOTIVO", "REGISTRADO"]

def _ws_justif_dir():
    client = get_client()
    sh = client.open_by_key(st.secrets["sheet_checador_id"])
    try:
        return sh.worksheet(JUSTIF_DIR_TAB)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(JUSTIF_DIR_TAB, rows=2, cols=len(JUSTIF_DIR_HEADERS))
        ws.append_row(JUSTIF_DIR_HEADERS)
        return ws

@st.cache_data(ttl=60)
def cargar_justif_direccion():
    try:
        ws = _ws_justif_dir()
        data = ws.get_all_records(numericise_ignore=["all"])
        return pd.DataFrame(data) if data else pd.DataFrame(columns=JUSTIF_DIR_HEADERS)
    except Exception:
        return pd.DataFrame(columns=JUSTIF_DIR_HEADERS)

def guardar_justif_direccion(nuevas: list[dict]):
    """nuevas: [{RFC, NOMBRE, FECHA (YYYY-MM-DD), MOTIVO}, ...].
    No duplica: si ya existe RFC+FECHA, actualiza el motivo."""
    if not nuevas:
        return
    ws = _ws_justif_dir()
    existentes = ws.get_all_records(numericise_ignore=["all"])
    idx = {}
    for i, row in enumerate(existentes, start=2):
        idx[(str(row.get("RFC","")).upper().strip(), str(row.get("FECHA","")).strip())] = i
    ahora = datetime.now(pytz.utc).astimezone(pytz.timezone("America/Mexico_City")).strftime("%Y-%m-%d %H:%M")
    from gspread.cell import Cell
    actualizaciones, filas_nuevas = [], []
    for n in nuevas:
        rfc_n = str(n.get("RFC","")).upper().strip()
        fecha_n = str(n.get("FECHA","")).strip()
        if not rfc_n or not fecha_n:
            continue
        clave = (rfc_n, fecha_n)
        if clave in idx:
            fila = idx[clave]
            actualizaciones.append(Cell(fila, 4, str(n.get("MOTIVO",""))))
            actualizaciones.append(Cell(fila, 5, ahora))
        else:
            filas_nuevas.append([rfc_n, n.get("NOMBRE",""), fecha_n, str(n.get("MOTIVO","")), ahora])
    if actualizaciones:
        ws.update_cells(actualizaciones, value_input_option="USER_ENTERED")
    if filas_nuevas:
        ws.append_rows(filas_nuevas, value_input_option="USER_ENTERED")
    cargar_justif_direccion.clear()

def eliminar_justif_direccion(rfc: str, fecha: str):
    """Quita una justificación (para deshacer errores). Localiza la fila EN VIVO."""
    try:
        ws = _ws_justif_dir()
        data = ws.get_all_records(numericise_ignore=["all"])
        for i, row in enumerate(data, start=2):
            if (str(row.get("RFC","")).upper().strip() == str(rfc).upper().strip()
                    and str(row.get("FECHA","")).strip() == str(fecha).strip()):
                ws.delete_rows(i)
                break
        cargar_justif_direccion.clear()
    except Exception as e:
        st.error(f"No se pudo eliminar la justificación: {e}")



# ─────────────────────────────────────────────
# UTILIDADES
# ─────────────────────────────────────────────
def generar_folio(tipo: str, incidencias_df: pd.DataFrame) -> str:
    anio = datetime.now().year
    prefijo = f"{tipo}-{anio}-"
    # Para ECO buscar en Solicitudes
    if tipo == "ECO":
        try:
            sol = cargar_solicitudes_eco()
            col_folio = next((c for c in sol.columns if "FOLIO" in c.upper()), None)
            if col_folio:
                existentes = sol[sol[col_folio].astype(str).str.startswith(prefijo)][col_folio].tolist()
                nums = []
                for f in existentes:
                    partes = f.split("-")
                    if len(partes) == 3 and partes[2].isdigit():
                        nums.append(int(partes[2]))
                return f"{prefijo}{str(max(nums) + 1 if nums else 1).zfill(4)}"
        except Exception:
            pass
        return f"{prefijo}0001"
    if incidencias_df.empty:
        return f"{prefijo}0001"
    existentes = incidencias_df[
        incidencias_df["FOLIO"].astype(str).str.startswith(prefijo)
    ]["FOLIO"].tolist()
    if not existentes:
        return f"{prefijo}0001"
    nums = []
    for f in existentes:
        partes = f.split("-")
        if len(partes) == 3 and partes[2].isdigit():
            nums.append(int(partes[2]))
    return f"{prefijo}{str(max(nums) + 1 if nums else 1).zfill(4)}"

def festivos_en_periodo(festivos_df: pd.DataFrame, ini: date, fin: date) -> list:
    """Devuelve [(fecha, descripcion), ...] de los festivos que caen dentro
    del período [ini, fin], ordenados por fecha. Expande rangos.
    Usa parseo flexible (pd.to_datetime) porque el Sheet puede devolver
    las fechas en distintos formatos (YYYY-MM-DD, DD/MM/YYYY, serial, etc.)."""
    out = []
    if festivos_df is None or not hasattr(festivos_df, "empty") or festivos_df.empty:
        return out
    # Normalizar ini/fin a date puro (la app puede pasarlos como datetime/Timestamp
    # con hora, lo que rompe la comparación date <= datetime).
    if hasattr(ini, "date") and not isinstance(ini, date):
        ini = ini.date()
    elif isinstance(ini, datetime):
        ini = ini.date()
    if hasattr(fin, "date") and not isinstance(fin, date):
        fin = fin.date()
    elif isinstance(fin, datetime):
        fin = fin.date()
    col_desc = next((c for c in festivos_df.columns if "DESCRIP" in str(c).upper()), None)
    col_fi   = next((c for c in festivos_df.columns if "INICIO" in str(c).upper()), None)
    col_ff   = next((c for c in festivos_df.columns if "FIN" in str(c).upper()), None)
    if not col_fi:
        return out
    def _parse_fecha(v):
        """Parsea fecha desde texto (ISO, DD/MM), datetime, o serial de Sheets."""
        if v is None or v == "":
            return None
        # Número serial de Google Sheets (días desde 1899-12-30)
        try:
            if isinstance(v, (int, float)) or (isinstance(v, str) and v.replace(".","",1).isdigit()):
                serial = float(v)
                if serial > 30000:  # rango razonable de fechas modernas
                    return (datetime(1899,12,30) + timedelta(days=int(serial))).date()
        except Exception:
            pass
        ts = pd.to_datetime(v, errors="coerce")
        return ts.date() if not pd.isna(ts) else None

    for _, row in festivos_df.iterrows():
        try:
            fi = _parse_fecha(row.get(col_fi, ""))
            ff = _parse_fecha(row.get(col_ff, "") if col_ff else row.get(col_fi, ""))
            if fi is None:
                continue
            if ff is None:
                ff = fi
            desc = str(row[col_desc]).strip() if col_desc else ""
            d = fi
            while d <= ff:
                if ini <= d <= fin:
                    out.append((d, desc))
                d += timedelta(days=1)
        except Exception:
            pass
    return sorted(out, key=lambda x: x[0])

def festivos_a_set(festivos_df: pd.DataFrame) -> set:
    """Expande los rangos de la tab festivos a un set de fechas individuales."""
    festivos_set = set()
    for _, row in festivos_df.iterrows():
        try:
            fi = datetime.strptime(str(row["FECHA_INICIO"]), "%Y-%m-%d").date()
            ff = datetime.strptime(str(row["FECHA_FIN"]), "%Y-%m-%d").date()
            d = fi
            while d <= ff:
                festivos_set.add(d)
                d += timedelta(days=1)
        except Exception:
            pass
    return festivos_set

def dias_habiles_entre(fecha_inicio: date, fecha_fin: date, festivos_df: pd.DataFrame) -> int:
    festivos_set = festivos_a_set(festivos_df)
    total = 0
    d = fecha_inicio
    while d <= fecha_fin:
        if d.weekday() < 5 and d not in festivos_set:
            total += 1
        d += timedelta(days=1)
    return total

def dias_economicos_usados(rfc: str, solicitudes_df: pd.DataFrame) -> int:
    df = solicitudes_df[solicitudes_df["RFC"].astype(str).str.upper() == rfc.upper()]
    # Solo cuentan las APROBADAS: la columna tiene nombre de admin.
    # Las rechazadas escriben "RECHAZADO — ..." en la misma columna y NO deben descontar días.
    aprob = df["Aprobado Por"].astype(str).str.strip()
    df_aprobados = df[(aprob != "") & (~aprob.str.upper().str.startswith("RECHAZADO"))]
    total = 0
    for _, row in df_aprobados.iterrows():
        try:
            total += int(row["Dias Solicitados"])
        except Exception:
            pass
    return total

def saldo_economicos_rfc(rfc: str):
    """(disponibles, totales) con la MISMA fórmula que vista_empleado, para
    que coordinador y empleado vean el mismo número. None si el RFC no está
    en la tab Empleados (no se inventa un saldo)."""
    try:
        emp = cargar_empleados()
        fila = emp[emp["RFC"].astype(str).str.upper().str.strip() == str(rfc).upper().strip()]
        if fila.empty:
            return None
        totales = int(fila["DIAS TOTALES"].iloc[0])
        return totales - dias_economicos_usados(rfc, cargar_solicitudes_eco()), totales
    except Exception:
        return None

def actualizar_dias_disponibles_sheet(rfc_objetivo: str = None):
    """Recalcula DIAS DISPONIBLES (= DIAS TOTALES - usados aprobados) y lo escribe
    en la hoja Empleados del Sheet de económicos.
    Si rfc_objetivo se indica, solo actualiza esa fila; si no, actualiza todas.
    Devuelve (actualizados, errores)."""
    from gspread.cell import Cell
    client = get_client()
    sh = client.open_by_key(st.secrets["sheet_economicos_id"])
    ws = sh.worksheet("Empleados")
    registros = ws.get_all_records(numericise_ignore=["all"])
    if not registros:
        return 0, []
    headers = list(registros[0].keys())
    try:
        col_disp = headers.index("DIAS DISPONIBLES") + 1
    except ValueError:
        return 0, ["No existe la columna 'DIAS DISPONIBLES' en la hoja Empleados"]
    col_rfc = next((h for h in headers if h.upper() == "RFC"), "RFC")
    col_tot = next((h for h in headers if "DIAS TOTALES" in h.upper()), "DIAS TOTALES")

    solicitudes = cargar_solicitudes_eco()
    celdas, errores = [], []
    for i, row in enumerate(registros, start=2):  # fila 1 = headers
        rfc = str(row.get(col_rfc, "")).upper().strip()
        if not rfc:
            continue
        if rfc_objetivo and rfc != rfc_objetivo.upper().strip():
            continue
        try:
            total = int(row.get(col_tot, 0) or 0)
        except Exception:
            total = 0
        usados = dias_economicos_usados(rfc, solicitudes)
        disponibles = total - usados
        celdas.append(Cell(i, col_disp, disponibles))
    if celdas:
        ws.update_cells(celdas, value_input_option="USER_ENTERED")
    cargar_empleados.clear()
    return len(celdas), errores

def horas_pases_mes(rfc: str, incidencias_df: pd.DataFrame) -> float:
    """Suma las horas de pases autorizados del mes en curso usando FECHA_INICIO."""
    import pytz
    tz_mx = pytz.timezone("America/Mexico_City")
    ahora = datetime.now(pytz.utc).astimezone(tz_mx)
    df = incidencias_df[
        (incidencias_df["RFC"].astype(str).str.upper() == rfc.upper()) &
        (incidencias_df["TIPO"].isin(["PSE","PSR","PEN"])) &
        (_estado_norm(incidencias_df["ESTADO"]).str.startswith("AUTORIZADO"))
    ].copy()
    if df.empty:
        return 0.0
    df["FECHA_DT"] = pd.to_datetime(df["FECHA_INICIO"], errors="coerce")
    df = df[(df["FECHA_DT"].dt.year == ahora.year) & (df["FECHA_DT"].dt.month == ahora.month)]
    total = 0.0
    for _, row in df.iterrows():
        try:
            val = str(row.get("HORAS_PASE", 0) or 0).replace(",", ".")
            total += float(val)
        except Exception:
            pass
    return round(total, 2)

def rfc_a_fecha_nacimiento(rfc: str):
    """Extrae fecha de nacimiento del RFC (posiciones 4-9: AAMMDD)."""
    try:
        anio = int(rfc[4:6])
        mes  = int(rfc[6:8])
        dia  = int(rfc[8:10])
        pivote = datetime.now().year % 100
        anio_completo = 2000 + anio if anio <= pivote else 1900 + anio
        return date(anio_completo, mes, dia)
    except Exception:
        return None

def cumpleanos_laboral(rfc: str) -> date:
    """Retorna el día hábil de cumpleaños (lunes si cae sábado, viernes si domingo)."""
    fn = rfc_a_fecha_nacimiento(rfc)
    if not fn:
        return None
    hoy = date.today()
    cumple = date(hoy.year, fn.month, fn.day)
    if cumple < hoy:
        cumple = date(hoy.year + 1, fn.month, fn.day)
    if cumple.weekday() == 5:  # sábado → lunes
        cumple += timedelta(days=2)
    elif cumple.weekday() == 6:  # domingo → viernes
        cumple -= timedelta(days=2)
    return cumple

def rfc_oculto(rfc: str) -> str:
    """Muestra solo los primeros 4 caracteres del RFC."""
    return rfc[:4] + "****" + rfc[-3:] if len(rfc) >= 7 else rfc

# ─────────────────────────────────────────────
# GENERACIÓN DE COMPROBANTE PDF
# ─────────────────────────────────────────────
def generar_comprobante_pdf(datos: dict) -> bytes:
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter,
                            leftMargin=1.5*cm, rightMargin=1.5*cm,
                            topMargin=1.2*cm, bottomMargin=1.2*cm)
    styles = getSampleStyleSheet()

    estilo_titulo   = ParagraphStyle("t", parent=styles["Normal"], fontSize=13, fontName="Helvetica-Bold", alignment=TA_CENTER, spaceAfter=4)
    estilo_sub      = ParagraphStyle("s", parent=styles["Normal"], fontSize=10, fontName="Helvetica", alignment=TA_CENTER, spaceAfter=2)
    estilo_folio    = ParagraphStyle("f", parent=styles["Normal"], fontSize=16, fontName="Helvetica-Bold", alignment=TA_CENTER, textColor=colors.HexColor("#002F6C"), spaceAfter=6)
    estilo_aviso    = ParagraphStyle("a", parent=styles["Normal"], fontSize=8, fontName="Helvetica-Oblique", alignment=TA_CENTER, textColor=colors.HexColor("#888888"))

    elementos = []
    # Logo SEJ
    import os
    logo_path = "logos_gris.png"
    if os.path.exists(logo_path):
        logo_rl = RLImage(logo_path, width=6*cm, height=1.5*cm)
        tabla_header = Table([[logo_rl, Paragraph(
            "<b>Dirección de Formación Continua</b><br/>Área de Recursos Humanos · SEJ",
            ParagraphStyle("hdr", parent=styles["Normal"], fontSize=9, fontName="Helvetica", alignment=TA_CENTER)
        )]], colWidths=[7*cm, 10*cm])
        tabla_header.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"MIDDLE"),("ALIGN",(1,0),(1,0),"CENTER")]))
        elementos.append(tabla_header)
    else:
        elementos.append(Paragraph("Secretaría de Educación Jalisco", estilo_titulo))
        elementos.append(Paragraph("Dirección de Formación Continua · Área de Recursos Humanos", estilo_sub))
    elementos.append(Spacer(1, 0.1*cm))
    elementos.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#002F6C")))
    elementos.append(Spacer(1, 0.1*cm))
    elementos.append(Paragraph("COMPROBANTE DE CAPTURA DE INCIDENCIA", estilo_titulo))
    elementos.append(Spacer(1, 0.1*cm))
    elementos.append(Paragraph(f"FOLIO: {datos['folio']}", estilo_folio))
    elementos.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#CCCCCC")))
    elementos.append(Spacer(1, 0.15*cm))

    tabla_datos = [
        ["Nombre completo:", datos["nombre"]],
        ["Filiación:", rfc_oculto(datos["rfc"])],
        ["Tipo de incidencia:", datos.get("subtipo_label", datos["tipo_label"])],
        ["Fecha de solicitud:", datos["fecha_solicitud"]],
    ]
    if datos["tipo"] == "CHO":
        tabla_datos.insert(4, ["Fecha de aplicación:", datos["fecha_inicio"]])
    else:
        tabla_datos.insert(4, ["Fecha inicio:", datos["fecha_inicio"]])
        tabla_datos.insert(5, ["Fecha fin:", datos["fecha_fin"]])
    horas_val = datos.get("horas_pase", 0)
    tabla_datos += [
        ["Días solicitados:", "N/A" if datos["tipo"] in ["PSE", "CHO"] else str(datos["dias"])],
        ["Horas de pase:", (f"{horas_val}h" if horas_val else "No registradas") if datos["tipo"] == "PSE" else "N/A"],
        ["Motivo / Descripción:", Paragraph(str(datos["motivo"]), ParagraphStyle("mot", parent=styles["Normal"], fontSize=9, fontName="Helvetica"))],
        ["Documento anexo:", "Sí, se presentará en RH" if datos["tiene_anexo"] else "No aplica"],
    ]

    t = Table(tabla_datos, colWidths=[5*cm, 11*cm])
    t.setStyle(TableStyle([
        ("FONTNAME",    (0, 0), (0, -1), "Helvetica-Bold"),
        ("FONTNAME",    (1, 0), (1, -1), "Helvetica"),
        ("FONTSIZE",    (0, 0), (-1, -1), 9),
        ("VALIGN",      (0, 0), (-1, -1), "TOP"),
        ("ROWBACKGROUNDS", (0, 0), (-1, -1), [colors.HexColor("#F4F6F9"), colors.white]),
        ("GRID",        (0, 0), (-1, -1), 0.3, colors.HexColor("#DDDDDD")),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING",(0, 0), (-1, -1), 8),
        ("TOPPADDING",  (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 3),
    ]))
    elementos.append(t)
    elementos.append(Spacer(1, 0.1*cm))
    elementos.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#CCCCCC")))
    elementos.append(Spacer(1, 0.1*cm))

    color_estado = colors.HexColor("#F59E0B")
    et = Table([[Paragraph("<b>ESTADO: PENDIENTE DE AUTORIZACIÓN</b>",
                           ParagraphStyle("e", parent=styles["Normal"], fontSize=10,
                                          fontName="Helvetica-Bold", textColor=color_estado,
                                          alignment=TA_CENTER))]], colWidths=[16*cm])
    et.setStyle(TableStyle([
        ("BOX", (0,0), (-1,-1), 1, color_estado),
        ("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#FFFBEB")),
        ("TOPPADDING", (0,0), (-1,-1), 6),
        ("BOTTOMPADDING", (0,0), (-1,-1), 6),
    ]))
    elementos.append(et)
    elementos.append(Spacer(1, 0.1*cm))
    # ── Bloque de firmas ─────────────────────────
    elementos.append(Paragraph("<b>REQUISITO OBLIGATORIO: SECCIÓN DE FIRMAS</b>",
        ParagraphStyle("tf", parent=styles["Normal"], fontSize=9, fontName="Helvetica-Bold", spaceAfter=6)))

    estilo_firma = ParagraphStyle("firmas", parent=styles["Normal"], fontSize=8, fontName="Helvetica", alignment=TA_CENTER)
    nombre_interesado = datos.get("nombre", "Servidor(a) Público(a)")
    jefe_pdf_texto    = datos.get("jefe_inmediato", "Nombre y Firma")
    # Director(a) General: viene del Sheet (columna DIRECTOR_GENERAL); si no, usa la constante
    director_pdf      = datos.get("director_general") or NOMBRE_DIRECTORA

    # ¿El jefe inmediato es la misma persona que el/la director(a) general?
    def _solo_nombre(txt):
        return str(txt or "").split("<br/>")[0].split("—")[0].split(" - ")[0].strip().lower()
    jefe_es_director = _solo_nombre(jefe_pdf_texto) and _solo_nombre(jefe_pdf_texto) == _solo_nombre(director_pdf)

    firma_interesado  = Paragraph(f"<br/><br/>___________________________<br/><b>Firma del Interesado</b><br/>{nombre_interesado}", estilo_firma)
    firma_jefe        = Paragraph(f"<br/><br/>___________________________<br/><b>Autoriza Jefe(a) Inmediato</b><br/>{jefe_pdf_texto}", estilo_firma)
    firma_vob         = Paragraph(f"<br/><br/>___________________________<br/><b>Vo.Bo. Titular del Área</b><br/>{director_pdf}", estilo_firma)

    if datos["tipo"] in ["ECO", "CHO", "RGU"]:
        if jefe_es_director:
            # La titular es a su vez la jefa inmediata: solo se pone Vo.Bo. del titular.
            t_firmas = Table([[firma_interesado, firma_vob]], colWidths=[8*cm, 8*cm])
        else:
            t_firmas = Table([[firma_interesado, firma_jefe, firma_vob]], colWidths=[5.3*cm, 5.3*cm, 5.4*cm])
    else:
        t_firmas = Table([[firma_interesado, firma_jefe]], colWidths=[8*cm, 8*cm])

    t_firmas.setStyle(TableStyle([
        ("VALIGN",        (0,0), (-1,-1), "TOP"),
        ("ALIGN",         (0,0), (-1,-1), "CENTER"),
        ("BOTTOMPADDING", (0,0), (-1,-1), 10),
    ]))
    elementos.append(t_firmas)
    elementos.append(Spacer(1, 0.1*cm))
    elementos.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#CCCCCC")))
    elementos.append(Spacer(1, 0.1*cm))

    # ── Instrucciones finales ─────────────────────
    elementos.append(Paragraph("<b>INSTRUCCIONES DE ENTREGA:</b>",
        ParagraphStyle("tit_i", parent=styles["Normal"], fontSize=8, fontName="Helvetica-Bold", spaceAfter=3)))
    for inst in [
        "1. Recaba las firmas físicas obligatorias que se muestran arriba.",
        "2. Adjunta la documentación de soporte original si aplica (cita IMSS, oficio de comisión, etc.).",
        "3. Entrega el expediente completo en el Área de Recursos Humanos (Administración de Personal DFC).",
        "4. Resguarda tu copia digital. El folio es tu número de seguimiento oficial.",
    ]:
        elementos.append(Paragraph(inst, ParagraphStyle("i", parent=styles["Normal"],
                                                         fontSize=8, fontName="Helvetica",
                                                         leftIndent=10, spaceAfter=3)))

    # ── QR de seguridad ──────────────────────────
    url_validacion = f"https://gestion-personal-dfc.streamlit.app/?validar_folio={datos['folio']}"
    qr = qrcode.QRCode(version=1, box_size=10, border=1)
    qr.add_data(url_validacion)
    qr.make(fit=True)
    img_qr = qr.make_image(fill_color="black", back_color="white")
    qr_buffer = BytesIO()
    img_qr.save(qr_buffer, format="PNG")
    qr_buffer.seek(0)
    qr_flowable = RLImage(qr_buffer, width=2.5*cm, height=2.5*cm)
    tabla_qr = Table([
        [qr_flowable, Paragraph(
            "<b>ESCÁNER DE SEGURIDAD DIGITAL</b><br/>Escanea este QR para verificar la autenticidad de este documento en tiempo real. Cualquier alteración anula el comprobante.",
            ParagraphStyle("txt_qr", parent=styles["Normal"], fontSize=7, textColor=colors.HexColor("#555555"))
        )]
    ], colWidths=[3*cm, 13*cm])
    tabla_qr.setStyle(TableStyle([("VALIGN", (0,0), (-1,-1), "MIDDLE")]))
    elementos.append(tabla_qr)
    elementos.append(Spacer(1, 0.15*cm))
    elementos.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#CCCCCC")))
    elementos.append(Spacer(1, 0.1*cm))
    elementos.append(Paragraph(f"Generado el {datos['fecha_solicitud']} · Sistema ARI — DFC · SEJ", estilo_aviso))
    doc.build(elementos)
    buffer.seek(0)
    return buffer.read()

# ─────────────────────────────────────────────
# ESCRITURA EN SHEETS
# ─────────────────────────────────────────────
def subir_archivo_drive(archivo, nombre_archivo: str, folder_id: str) -> str:
    """Sube un archivo a Google Drive y retorna el link de visualización,
    o un string que empieza con 'ERROR:' si falla. Quien la llame DEBE
    checar ese prefijo antes de tratar el resultado como link válido."""
    try:
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaIoBaseUpload
        from google.oauth2.service_account import Credentials
        import io

        creds = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=["https://www.googleapis.com/auth/drive"]
        )
        service = build("drive", "v3", credentials=creds)
        media = MediaIoBaseUpload(io.BytesIO(archivo.read()), mimetype=archivo.type)
        file_metadata = {"name": nombre_archivo, "parents": [folder_id]}
        archivo_drive = service.files().create(
            body=file_metadata, media_body=media, fields="id", supportsAllDrives=True
        ).execute()
        file_id = archivo_drive.get("id")
        # NOTA: sin permiso "anyone" — el acceso se hereda de la membresía
        # de la unidad compartida (solo RH y la service account).
        return f"https://drive.google.com/file/d/{file_id}/view"
    except Exception as e:
        return f"ERROR: {e}"

def subir_anexo_drive(archivo, folio: str, rfc: str) -> str:
    """Sube un archivo a Google Drive y retorna el link de visualización."""
    ext = archivo.name.split(".")[-1].lower()
    return subir_archivo_drive(archivo, f"{folio}_{rfc}.{ext}", DRIVE_ANEXOS_FOLDER)

# ═══════════════════════════════════════════════════════════════════
# BITÁCORA DE SOLICITUDES — trazabilidad de TODO intento de envío
# ═══════════════════════════════════════════════════════════════════
# Se registra el intento haya terminado bien o mal. Sin esto, una solicitud
# que falla al escribirse no deja absolutamente ningún rastro: ni el empleado
# ni RH pueden demostrar después que se intentó enviar.
BITACORA_SOL_TAB = "Bitacora_Solicitudes"
# Orden pensado para leerse de izquierda a derecha: cuándo, qué pasó, a cuál
# folio, de quién es, de qué estado a cuál, y quién lo hizo.
BITACORA_SOL_HEADERS = [
    "FECHA_HORA", "EVENTO", "FOLIO", "TIPO", "NOMBRE", "RFC",
    "ESTADO_ANTERIOR", "ESTADO_NUEVO", "USUARIO_SESION", "ROL",
    "RESULTADO", "DETALLE", "FECHA_INICIO", "FECHA_FIN",
]


def _ws_bitacora_sol():
    """Hoja de bitácora, creándola si no existe.

    Si la hoja ya existía con menos columnas (la primera versión solo
    registraba altas), se le AGREGAN las que falten en vez de romperse: así
    los renglones viejos se conservan y los nuevos quedan alineados.
    """
    sh = get_client().open_by_key(st.secrets["sheet_checador_id"])
    try:
        ws = sh.worksheet(BITACORA_SOL_TAB)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(BITACORA_SOL_TAB, rows=2, cols=len(BITACORA_SOL_HEADERS))
        ws.append_row(BITACORA_SOL_HEADERS, value_input_option="USER_ENTERED")
        return ws
    try:
        actuales = [str(h).strip().upper() for h in ws.row_values(1)]
        faltan = [h for h in BITACORA_SOL_HEADERS if h not in actuales]
        if faltan:
            if ws.col_count < len(actuales) + len(faltan):
                ws.add_cols(len(actuales) + len(faltan) - ws.col_count)
            ws.update(
                [actuales + faltan],
                f"A1:{chr(64 + len(actuales) + len(faltan)) if len(actuales)+len(faltan) <= 26 else 'AZ'}1",
                value_input_option="USER_ENTERED")
    except Exception:
        pass  # si no se pudo ampliar, se escribe con lo que haya
    return ws


def _volcar_bitacora_sol():
    """Vacía la cola al Sheet. Cada evento viaja como diccionario y aquí se
    acomoda según los encabezados REALES de la hoja: si alguien le cambia el
    orden a las columnas, el dato sigue cayendo donde va.

    Si Google está caído la cola se queda en sesión y se reintenta en el
    siguiente rerun: un 429 no se lleva también el registro de auditoría.
    """
    cola = st.session_state.get("_bitacora_cola") or []
    if not cola:
        return
    try:
        ws = _ws_bitacora_sol()
        enc = [str(h).strip().upper() for h in ws.row_values(1)] or BITACORA_SOL_HEADERS
        filas = []
        for ev in cola:
            if isinstance(ev, dict):
                filas.append([str(ev.get(c, "")) for c in enc])
            else:
                filas.append(list(ev))  # cola de la versión anterior (lista)
        ws.append_rows(filas, value_input_option="USER_ENTERED",
                       insert_data_option="INSERT_ROWS")
        st.session_state["_bitacora_cola"] = []
    except Exception:
        pass  # se reintenta solo; nunca debe tumbar la pantalla


def registrar_bitacora_evento(evento: str, folio: str = "", tipo: str = "",
                              nombre: str = "", rfc: str = "",
                              estado_anterior: str = "", estado_nuevo: str = "",
                              resultado: str = "OK", detalle: str = "",
                              fecha_inicio: str = "", fecha_fin: str = ""):
    """Anota UN evento del ciclo de vida de una solicitud.

    EVENTO es ALTA (la mandó el empleado), AUTORIZACION, RECHAZO o
    CAMBIO_HORARIO. Con el alta y la resolución en la misma hoja, cada folio
    tiene su rastro completo: si un folio tiene ALTA pero nunca una
    resolución y ya no está en pendientes, se perdió — y la bitácora lo
    demuestra.

    NUNCA lanza excepción: la bitácora no puede ser el motivo de que una
    solicitud o una autorización fallen.
    """
    try:
        ahora = datetime.now(pytz.utc).astimezone(
            pytz.timezone("America/Mexico_City")).strftime("%Y-%m-%d %H:%M:%S")
        st.session_state.setdefault("_bitacora_cola", []).append({
            "FECHA_HORA": ahora,
            "EVENTO": evento,
            "FOLIO": str(folio),
            "TIPO": str(tipo),
            "NOMBRE": str(nombre),
            "RFC": str(rfc),
            "ESTADO_ANTERIOR": str(estado_anterior),
            "ESTADO_NUEVO": str(estado_nuevo),
            # Quién estaba dentro al hacerlo, que no siempre es el titular
            # de la solicitud (un admin puede capturar o resolver por otro).
            "USUARIO_SESION": str(st.session_state.get("nombre", "")),
            "ROL": str(st.session_state.get("rol", "")),
            "RESULTADO": resultado,
            "DETALLE": texto_seguro(str(detalle)[:300]),
            "FECHA_INICIO": str(fecha_inicio),
            "FECHA_FIN": str(fecha_fin),
        })
        _volcar_bitacora_sol()
    except Exception:
        pass


def registrar_bitacora_sol(datos: dict, resultado: str, detalle: str = ""):
    """Alta de una solicitud. Envoltura sobre registrar_bitacora_evento para
    no cambiar los puntos de llamada que ya existían."""
    registrar_bitacora_evento(
        "ALTA",
        folio=datos.get("folio", ""), tipo=datos.get("tipo", ""),
        nombre=datos.get("nombre", ""), rfc=datos.get("rfc", ""),
        # "CREADA" y no "PENDIENTE": la bitácora es un histórico, y un renglón
        # que dice PENDIENTE se lee como si la solicitud siguiera esperando
        # aunque ya esté resuelta. Solo las filas de resolución hablan de
        # PENDIENTE, y ahí sí describe de dónde venía.
        estado_anterior="", estado_nuevo="CREADA" if resultado == "GUARDADA" else "",
        resultado=resultado, detalle=detalle,
        fecha_inicio=datos.get("fecha_inicio", ""), fecha_fin=datos.get("fecha_fin", ""))


def texto_seguro(valor) -> str:
    """Neutraliza el texto que escribe una persona antes de mandarlo al Sheet.

    La app escribe con value_input_option="USER_ENTERED", así que Google
    INTERPRETA lo que recibe. Un motivo que empiece con "=", "+", "-" o "@"
    deja de ser texto y se convierte en una FÓRMULA VIVA dentro de la hoja.
    Y una fórmula puede leer otras celdas y mandarlas fuera —por ejemplo con
    IMPORTXML hacia un sitio ajeno— en el momento en que alguien de RH abre
    el Sheet, sin que se vea nada raro en pantalla.

    Anteponer un apóstrofo obliga a Google a tratarlo como texto. El apóstrofo
    no se ve en la celda ni se copia al leerla, así que el motivo se sigue
    mostrando igual en los reportes y en el PDF.

    Solo se toca lo que empieza con esos cuatro caracteres: "Cita médica" o
    "Junta a las 10:00" pasan intactos.
    """
    txt = "" if valor is None else str(valor)
    return "'" + txt if txt[:1] in ("=", "+", "-", "@") else txt


def _estado_norm(serie) -> pd.Series:
    """ESTADO en mayúsculas y sin espacios sobrantes.

    Las celdas del Sheet se editan a mano en la oficina, así que el ESTADO
    llega con variantes: "PENDIENTE " con espacio, "Pendiente", o la celda
    vaciada por accidente. Comparar con == contra el texto crudo hacía que la
    solicitud desapareciera de la pantalla del admin aunque siguiera en el
    Sheet, sin autorizar ni rechazar — mientras el empleado SÍ la seguía
    viendo, porque su filtro usaba contains.
    """
    return serie.astype(str).str.upper().str.strip()


def _mascara_pendiente(df: pd.DataFrame) -> pd.Series:
    """Filas realmente pendientes: dice PENDIENTE en cualquier variante, o la
    celda ESTADO quedó vacía SIN que nadie haya autorizado ni rechazado (una
    solicitud sin estado nunca debe perderse de vista)."""
    est = _estado_norm(df["ESTADO"])
    if "AUTORIZADO_POR" in df.columns:
        aut = df["AUTORIZADO_POR"].astype(str).str.strip()
    else:
        aut = pd.Series([""] * len(df), index=df.index)
    return est.str.startswith("PENDIENTE") | (est.eq("") & aut.eq(""))


def _hash_envio(clave: str) -> str:
    import hashlib
    return hashlib.md5(clave.encode()).hexdigest()


def _envio_duplicado(clave: str) -> bool:
    """Candado anti doble clic: True solo si esta MISMA solicitud ya se guardó
    CON ÉXITO en los últimos 2 minutos de esta sesión.

    BUG QUE ESTO ARREGLA (solicitudes que "desaparecían" sin registro):
    antes el candado se cerraba al ENTRAR, sin saber si el guardado iba a
    funcionar. Si append_row tronaba (429 de cuota de Google, red), pasaba
    esto:
      1. intento 1 → candado cerrado → append_row falla → "espera 15 s y
         vuelve a intentar";
      2. el empleado reintenta a los 15 s, como se le indicó;
      3. el candado respondía "esta solicitud ya se envió hace un momento";
      4. el empleado se iba creyendo que quedó registrada.
    Nunca se escribió nada. Los dos mensajes se contradecían y entre ambos
    simulaban un registro exitoso.

    Ahora solo bloquea un envío CONFIRMADO: un intento fallido deja el candado
    abierto para que el reintento sí entre.
    """
    import time as _time
    h = _hash_envio(clave)
    ult = st.session_state.get("_ultimo_envio")
    # ult = (hash, timestamp, confirmado). Las tuplas viejas de 2 elementos se
    # tratan como confirmadas para no cambiar el comportamiento a media sesión.
    if ult and ult[0] == h and (_time.time() - ult[1]) < 120:
        confirmado = ult[2] if len(ult) > 2 else True
        if confirmado:
            return True
    st.session_state["_ultimo_envio"] = (h, _time.time(), False)
    return False


def _confirmar_envio(clave: str):
    """Cierra el candado. Se llama SOLO después de que la fila ya quedó escrita
    en el Sheet; a partir de aquí un segundo clic sí es un duplicado real."""
    import time as _time
    st.session_state["_ultimo_envio"] = (_hash_envio(clave), _time.time(), True)

def _error_amable(e: Exception, contexto: str = ""):
    """Traduce errores de cuota de Google a un mensaje humano en vez de traceback."""
    if "429" in str(e) or "Quota" in str(e) or "quota" in str(e):
        # Antes decía "Tu información no se perdió", y era falso: si el fallo
        # ocurre al escribir, la solicitud NO quedó registrada. Ese mensaje era
        # la mitad de la ilusión de solicitudes "desaparecidas".
        st.error("⏳ El sistema está ocupado en este momento (muchas personas usándolo a la vez). "
                 "**No se guardó nada todavía.** Espera 15 segundos y vuelve a intentar.")
    else:
        st.error(f"Error {contexto}: {e}")

def guardar_dia_economico(datos: dict):
    """Guarda día económico en tab Solicitudes del Sheet económicos — igual que app3."""
    # Candado: la misma persona + mismas fechas en <2 min = doble clic
    if _envio_duplicado(f"ECO|{datos.get('rfc','')}|{datos.get('fecha_inicio','')}|{datos.get('fecha_fin','')}"):
        # Solo llega aquí si el envío anterior SÍ quedó guardado.
        st.warning("Esta solicitud **ya quedó registrada** hace un momento; no se "
                   "duplicó. Búscala en 'Mis solicitudes'.")
        return
    client = get_client()
    sh = client.open_by_key(st.secrets["sheet_economicos_id"])
    ws = sh.worksheet("Solicitudes")
    todas = ws.get_all_records(numericise_ignore=["all"])
    nuevo_id = len(todas) + 1
    # FOLIO recalculado EN VIVO sobre lo recién leído (no sobre cache):
    # cierra la ventana de folios duplicados por concurrencia.
    try:
        anio_f = datetime.now().year
        pref = f"ECO-{anio_f}-"
        nums = []
        for t in todas:
            fol = str(t.get("FOLIO", ""))
            if fol.startswith(pref):
                pp = fol.split("-")
                if len(pp) == 3 and pp[2].isdigit():
                    nums.append(int(pp[2]))
        datos["folio"] = f"{pref}{str(max(nums) + 1 if nums else 1).zfill(4)}"
    except Exception:
        pass  # conserva el folio previamente generado
    emp_id = ""
    try:
        empleados = cargar_empleados()
        match = empleados[empleados["RFC"].astype(str).str.upper() == datos["rfc"].upper()]
        if not match.empty:
            emp_id = match.iloc[0].get("ID", "")
    except Exception:
        pass
    link_anexo = ""
    if datos.get("tiene_anexo") and datos.get("archivo_anexo"):
        try:
            link_anexo = subir_anexo_drive(datos["archivo_anexo"], datos["folio"], datos["rfc"])
        except Exception as e:
            st.warning(f"No se pudo subir el anexo: {e}")
    fila = [
        nuevo_id,                    # ID
        emp_id,                      # EmpleadoID
        datos["rfc"],                 # RFC
        datos["nombre"],              # Nombre Completo
        "economico",                 # Tipo Permiso
        datos["fecha_inicio"],        # Fecha Inicio
        datos["fecha_fin"],           # Fecha Fin
        datos["dias"],                # Dias Solicitados
        texto_seguro(datos["motivo"]),   # Motivo (saneado: ver texto_seguro)
        datos["fecha_solicitud"],     # Fecha Registro
        "",                          # Aprobado Por
        datos["nombre"],              # Registrado Por
        datos["folio"],               # FOLIO
        link_anexo,                  # LINK_ANEXO
    ]
    # insert_data_option="INSERT_ROWS" es OBLIGATORIO aquí.
    # Sin él, gspread deja que la API de Google use su modo por defecto
    # (OVERWRITE): Google detecta dónde "termina la tabla" y escribe a partir
    # de ahí. Si la detección termina antes del final real de los datos —pasa
    # cuando una fila quedó en blanco porque alguien borró su contenido con
    # Supr en vez de eliminar la fila— el append ESCRIBE ENCIMA de las filas
    # existentes y se las lleva, sin error y sin rastro. INSERT_ROWS fuerza a
    # insertar renglones nuevos: nunca pisa lo que ya estaba.
    ws.append_row(fila, value_input_option="USER_ENTERED",
                  insert_data_option="INSERT_ROWS")
    # La fila ya está en el Sheet: ahora sí se cierra el candado.
    _confirmar_envio(f"ECO|{datos.get('rfc','')}|{datos.get('fecha_inicio','')}|{datos.get('fecha_fin','')}")
    cargar_solicitudes_eco.clear()
    return True

def guardar_incidencia(datos: dict):
    # Candado: la misma persona + tipo + fechas en <2 min = doble clic
    if _envio_duplicado(f"{datos.get('tipo','')}|{datos.get('rfc','')}|{datos.get('fecha_inicio','')}|{datos.get('fecha_fin','')}"):
        # Solo llega aquí si el envío anterior SÍ quedó guardado.
        st.warning("Esta solicitud **ya quedó registrada** hace un momento; no se "
                   "duplicó. Búscala en 'Mis solicitudes'.")
        return False
    client = get_client()
    sh = client.open_by_key(st.secrets["sheet_checador_id"])
    ws = sh.worksheet("Incidencias")
    todas = ws.get_all_records(numericise_ignore=["all"])
    # FOLIO recalculado EN VIVO sobre lo recién leído (cierra duplicados por concurrencia)
    try:
        anio_f = datetime.now().year
        pref = f"{datos.get('tipo','')}-{anio_f}-"
        nums = []
        for t in todas:
            fol = str(t.get("FOLIO", ""))
            if fol.startswith(pref):
                pp = fol.split("-")
                if len(pp) == 3 and pp[2].isdigit():
                    nums.append(int(pp[2]))
        datos["folio"] = f"{pref}{str(max(nums) + 1 if nums else 1).zfill(4)}"
    except Exception:
        pass
    fila = [
        len(todas) + 1,
        datos["folio"],
        datos["rfc"],
        datos["nombre"],
        datos["tipo"],
        datos["fecha_solicitud"],
        datos["fecha_inicio"],
        datos["fecha_fin"],
        datos["dias"],
        datos.get("horas_pase", ""),
        datos.get("hora_retorno", ""),
        texto_seguro(datos["motivo"]),   # saneado: ver texto_seguro
        "SÍ" if datos["tiene_anexo"] else "NO",
        datos.get("link_anexo", ""),
        "PENDIENTE",
        "", "", "",
    ]
    # insert_data_option="INSERT_ROWS" es OBLIGATORIO aquí.
    # Sin él, gspread deja que la API de Google use su modo por defecto
    # (OVERWRITE): Google detecta dónde "termina la tabla" y escribe a partir
    # de ahí. Si la detección termina antes del final real de los datos —pasa
    # cuando una fila quedó en blanco porque alguien borró su contenido con
    # Supr en vez de eliminar la fila— el append ESCRIBE ENCIMA de las filas
    # existentes y se las lleva, sin error y sin rastro. INSERT_ROWS fuerza a
    # insertar renglones nuevos: nunca pisa lo que ya estaba.
    ws.append_row(fila, value_input_option="USER_ENTERED",
                  insert_data_option="INSERT_ROWS")
    # La fila ya está en el Sheet: ahora sí se cierra el candado.
    _confirmar_envio(f"{datos.get('tipo','')}|{datos.get('rfc','')}|{datos.get('fecha_inicio','')}|{datos.get('fecha_fin','')}")
    cargar_incidencias.clear()
    return True

def autorizar_incidencia(folio: str, obs: str = ""):
    try:
        import pytz
        tz_mx = pytz.timezone("America/Mexico_City")
        ahora = datetime.now(pytz.utc).astimezone(tz_mx).strftime("%Y-%m-%d %H:%M")
        client = get_client()
        sh = client.open_by_key(st.secrets["sheet_checador_id"])
        ws = sh.worksheet("Incidencias")
        headers = [h.upper().strip() for h in ws.row_values(1)]
        data = ws.get_all_records(numericise_ignore=["all"])
        if "ESTADO" not in headers:
            st.error("No se encontró columna ESTADO en el Sheet.")
            return
        from gspread.cell import Cell
        for i, row in enumerate(data, start=2):
            if str(row.get("FOLIO", "")) == folio:
                celdas = [
                    Cell(i, headers.index("ESTADO") + 1,             "AUTORIZADO"),
                    Cell(i, headers.index("AUTORIZADO_POR") + 1,     st.session_state.get("nombre", "admin")),
                    Cell(i, headers.index("FECHA_AUTORIZACION") + 1, ahora),
                ]
                if obs and "OBSERVACIONES" in headers:
                    celdas.append(Cell(i, headers.index("OBSERVACIONES") + 1, texto_seguro(obs)))
                ws.update_cells(celdas, value_input_option="USER_ENTERED")
                # El estado ANTERIOR se toma del renglón tal como estaba antes
                # de sobrescribirlo: es la mitad que faltaba para poder auditar
                # una solicitud que desaparece del panel.
                registrar_bitacora_evento(
                    "AUTORIZACION", folio=folio, tipo=row.get("TIPO", ""),
                    nombre=row.get("NOMBRE", ""), rfc=row.get("RFC", ""),
                    estado_anterior=row.get("ESTADO", ""), estado_nuevo="AUTORIZADO",
                    detalle=obs, fecha_inicio=row.get("FECHA_INICIO", ""),
                    fecha_fin=row.get("FECHA_FIN", ""))
                break
        else:
            # No se encontró el folio: eso también se asienta, porque es
            # justo la huella de una solicitud que ya no está en el Sheet.
            registrar_bitacora_evento("AUTORIZACION", folio=folio,
                                      resultado="NO_ENCONTRADA",
                                      detalle="el folio no está en la hoja Incidencias")
        cargar_incidencias.clear()
    except gspread.exceptions.APIError:
        st.error("⏳ Google Sheets está saturado temporalmente. Espera 5 segundos e intenta de nuevo.")
    except Exception as e:
        _error_amable(e, "al autorizar")

def rechazar_incidencia(folio: str, obs: str):
    client = get_client()
    sh = client.open_by_key(st.secrets["sheet_checador_id"])
    ws = sh.worksheet("Incidencias")
    headers = ws.row_values(1)
    data = ws.get_all_records(numericise_ignore=["all"])
    from gspread.cell import Cell
    ahora = datetime.now(pytz.utc).astimezone(pytz.timezone("America/Mexico_City")).strftime("%Y-%m-%d %H:%M")
    for i, row in enumerate(data, start=2):
        if str(row.get("FOLIO", "")) == folio:
            ws.update_cells([
                Cell(i, headers.index("ESTADO") + 1,             "RECHAZADO"),
                Cell(i, headers.index("AUTORIZADO_POR") + 1,     st.session_state.get("nombre", "admin")),
                Cell(i, headers.index("FECHA_AUTORIZACION") + 1, ahora),
                Cell(i, headers.index("OBSERVACIONES") + 1,      texto_seguro(obs)),
            ], value_input_option="USER_ENTERED")
            registrar_bitacora_evento(
                "RECHAZO", folio=folio, tipo=row.get("TIPO", ""),
                nombre=row.get("NOMBRE", ""), rfc=row.get("RFC", ""),
                estado_anterior=row.get("ESTADO", ""), estado_nuevo="RECHAZADO",
                detalle=obs, fecha_inicio=row.get("FECHA_INICIO", ""),
                fecha_fin=row.get("FECHA_FIN", ""))
            break
    else:
        registrar_bitacora_evento("RECHAZO", folio=folio, resultado="NO_ENCONTRADA",
                                  detalle="el folio no está en la hoja Incidencias")
    cargar_incidencias.clear()

def _localizar_fila_eco(ws, folio: str = "", rfc: str = "", fecha_inicio: str = "", fecha_registro: str = ""):
    """Relee el Sheet EN VIVO y localiza la fila de la solicitud.
    Prioridad: FOLIO exacto; si no hay folio, RFC + Fecha Inicio + Fecha Registro
    (para solicitudes viejas sin folio).

    Devuelve (num_fila, headers_norm, fila_norm). La TERCERA pieza es nueva:
    la fila ya estaba leída aquí, pero antes se tiraba, y por eso la bitácora
    de los días económicos salía sin NOMBRE y sin FECHA_FIN — se vio en
    producción, con renglones que solo traían el RFC. Viene con las llaves en
    MAYÚSCULAS ("NOMBRE COMPLETO", "FECHA FIN") y los valores ya recortados.

    NUNCA usar índices de DataFrames cacheados: si alguien insertó/borró filas
    en el Sheet, se aprobaría la solicitud equivocada."""
    headers = [h.upper().strip() for h in ws.row_values(1)]
    data = ws.get_all_records(numericise_ignore=["all"])
    folio = str(folio or "").strip()
    for i, row in enumerate(data, start=2):
        row_norm = {str(k).upper().strip(): str(v).strip() for k, v in row.items()}
        if folio and row_norm.get("FOLIO", "") == folio:
            return i, headers, row_norm
        if not folio and rfc:
            if (row_norm.get("RFC", "").upper() == rfc.upper().strip()
                    and row_norm.get("FECHA INICIO", "") == str(fecha_inicio).strip()
                    and row_norm.get("FECHA REGISTRO", "") == str(fecha_registro).strip()):
                return i, headers, row_norm
    return None, headers, {}

def _datos_eco_para_bitacora(fila_eco: dict, rfc: str, fecha_inicio: str):
    """(nombre, fecha_inicio, fecha_fin) de una solicitud de día económico.

    La hoja de económicos usa otros nombres de columna que la de incidencias
    ("Nombre Completo" en vez de "NOMBRE", "Fecha Fin" en vez de "FECHA_FIN"),
    así que se buscan por su nombre real y se cae a lo que haya si falta.
    """
    f = fila_eco or {}
    nombre = (f.get("NOMBRE COMPLETO") or f.get("NOMBRE") or "").strip()
    fi = (f.get("FECHA INICIO") or fecha_inicio or "").strip()
    # Si no hay fecha fin, un día económico de un solo día tiene fin = inicio.
    ff = (f.get("FECHA FIN") or fi or "").strip()
    return nombre, fi, ff


def rechazar_dia_economico(obs: str, folio: str = "", rfc: str = "", fecha_inicio: str = "", fecha_registro: str = ""):
    try:
        client = get_client()
        sh = client.open_by_key(st.secrets["sheet_economicos_id"])
        ws = sh.worksheet("Solicitudes")
        fila, headers, _fila_eco = _localizar_fila_eco(ws, folio, rfc, fecha_inicio, fecha_registro)
        if fila is None:
            registrar_bitacora_evento("RECHAZO", folio=folio, tipo="ECO", rfc=rfc,
                                      resultado="NO_ENCONTRADA",
                                      detalle="la solicitud no está en la hoja Solicitudes",
                                      fecha_inicio=fecha_inicio, fecha_fin=fecha_inicio)
            st.error(f"No encontré la solicitud {folio or rfc} en el Sheet. Recarga la página e intenta de nuevo.")
            return
        col_aprobado = headers.index("APROBADO POR") + 1 if "APROBADO POR" in headers else None
        if col_aprobado:
            ws.update_cell(fila, col_aprobado, texto_seguro(f"RECHAZADO — {obs}"))
        # ESTADO_ANTERIOR es PENDIENTE por construcción: la lista de días
        # económicos pendientes filtra justamente las que tienen "Aprobado Por"
        # vacío, así que solo desde ahí se puede llegar a rechazar una.
        _nom_eco, _fi_eco, _ff_eco = _datos_eco_para_bitacora(_fila_eco, rfc, fecha_inicio)
        registrar_bitacora_evento(
            "RECHAZO", folio=folio, tipo="ECO", nombre=_nom_eco, rfc=rfc,
            estado_anterior="PENDIENTE", estado_nuevo="RECHAZADO",
            detalle=obs, fecha_inicio=_fi_eco, fecha_fin=_ff_eco)
        cargar_solicitudes_eco.clear()
        st.warning("Solicitud rechazada.")
    except Exception as e:
        _error_amable(e, "al rechazar")

def aprobar_dia_economico(nombre_admin: str, folio: str = "", rfc: str = "", fecha_inicio: str = "", fecha_registro: str = ""):
    """Escribe el nombre del admin en Aprobado Por, localizando la fila por FOLIO en vivo."""
    try:
        import pytz
        from gspread.cell import Cell
        tz_mx = pytz.timezone("America/Mexico_City")
        ahora = datetime.now(pytz.utc).astimezone(tz_mx).strftime("%Y-%m-%d %H:%M")
        client = get_client()
        sh = client.open_by_key(st.secrets["sheet_economicos_id"])
        ws = sh.worksheet("Solicitudes")
        fila, headers, _fila_eco = _localizar_fila_eco(ws, folio, rfc, fecha_inicio, fecha_registro)
        if fila is None:
            registrar_bitacora_evento("AUTORIZACION", folio=folio, tipo="ECO", rfc=rfc,
                                      resultado="NO_ENCONTRADA",
                                      detalle="la solicitud no está en la hoja Solicitudes",
                                      fecha_inicio=fecha_inicio, fecha_fin=fecha_inicio)
            st.error(f"No encontré la solicitud {folio or rfc} en el Sheet. Recarga la página e intenta de nuevo.")
            return
        celdas = []
        if "APROBADO POR" in headers:
            celdas.append(Cell(fila, headers.index("APROBADO POR") + 1, nombre_admin))
        col_fecha_aut = next((j + 1 for j, h in enumerate(headers) if "FECHA" in h and "AUTORIZACION" in h), None)
        if col_fecha_aut:
            celdas.append(Cell(fila, col_fecha_aut, ahora))
        if celdas:
            ws.update_cells(celdas, value_input_option="USER_ENTERED")
        _nom_eco, _fi_eco, _ff_eco = _datos_eco_para_bitacora(_fila_eco, rfc, fecha_inicio)
        # El DETALLE ya no repite quién aprobó: eso vive en USUARIO_SESION.
        registrar_bitacora_evento(
            "AUTORIZACION", folio=folio, tipo="ECO", nombre=_nom_eco, rfc=rfc,
            estado_anterior="PENDIENTE", estado_nuevo="AUTORIZADO",
            fecha_inicio=_fi_eco, fecha_fin=_ff_eco)
        cargar_solicitudes_eco.clear()
        # Actualizar el saldo de DIAS DISPONIBLES de ese empleado en el Sheet Empleados
        try:
            rfc_sol = rfc
            if not rfc_sol and "RFC" in headers:
                rfc_sol = ws.cell(fila, headers.index("RFC") + 1).value
            if rfc_sol:
                actualizar_dias_disponibles_sheet(rfc_sol)
        except Exception:
            pass
        st.success("Día económico autorizado correctamente.")
    except Exception as e:
        _error_amable(e, "al aprobar")

def _registrar_historial_horario(sh, rfc, nombre, row_anterior, horario_nuevo, fecha_aplicacion):
    """Deja en HISTORIAL_HORARIOS dos filas: la que cierra el horario anterior
    (FECHA_FIN = día previo al cambio) y la del nuevo (FECHA_INICIO = fecha de
    aplicación, FECHA_FIN vacía porque es el vigente).

    Se sacó del cuerpo de autorizar_cambio_horario para que el personal que
    solo vive en el Padrón Maestro (Centros de Maestros, Dirección Académica)
    también deje rastro: antes solo se escribía dentro del bucle de la tab
    'empleados', así que para ellos el cambio no quedaba registrado.

    row_anterior es el renglón previo como dict; sirve igual el de 'empleados'
    que el del padrón porque ambos usan los mismos nombres de columna.
    """
    try:
        ws_hist = sh.worksheet("HISTORIAL_HORARIOS")
    except Exception as e:
        st.warning(f"No se pudo guardar historial: {e}")
        return
    try:
        fecha_cierre = (datetime.strptime(fecha_aplicacion, "%Y-%m-%d")
                        - timedelta(days=1)).strftime("%Y-%m-%d")
    except Exception:
        fecha_cierre = ""
    try:
        # Fila 1: horario ANTERIOR, cerrado el día previo al cambio
        ws_hist.append_row([
            str(rfc), str(nombre), "", fecha_cierre,
            str(row_anterior.get("ENTRADA_LUN", "")), str(row_anterior.get("SALIDA_LUN", "")),
            str(row_anterior.get("ENTRADA_MAR", "")), str(row_anterior.get("SALIDA_MAR", "")),
            str(row_anterior.get("ENTRADA_MIE", "")), str(row_anterior.get("SALIDA_MIE", "")),
            str(row_anterior.get("ENTRADA_JUE", "")), str(row_anterior.get("SALIDA_JUE", "")),
            str(row_anterior.get("ENTRADA_VIE", "")), str(row_anterior.get("SALIDA_VIE", "")),
        ], value_input_option="USER_ENTERED", insert_data_option="INSERT_ROWS")
        # Fila 2: horario NUEVO, vigente desde la fecha de aplicación
        ws_hist.append_row([
            str(rfc), str(nombre), str(fecha_aplicacion), "",
            horario_nuevo.get("LUN", {}).get("entrada", ""), horario_nuevo.get("LUN", {}).get("salida", ""),
            horario_nuevo.get("MAR", {}).get("entrada", ""), horario_nuevo.get("MAR", {}).get("salida", ""),
            horario_nuevo.get("MIE", {}).get("entrada", ""), horario_nuevo.get("MIE", {}).get("salida", ""),
            horario_nuevo.get("JUE", {}).get("entrada", ""), horario_nuevo.get("JUE", {}).get("salida", ""),
            horario_nuevo.get("VIE", {}).get("entrada", ""), horario_nuevo.get("VIE", {}).get("salida", ""),
        ], value_input_option="USER_ENTERED", insert_data_option="INSERT_ROWS")
    except Exception as e:
        st.warning(f"No se pudo guardar historial: {e}")


def _escribir_horario_padron(emp_id: str, horario_nuevo: dict):
    """Escribe el horario autorizado en el PADRÓN MAESTRO (tab Padron).

    Esto faltaba por completo: el cambio de horario solo se guardaba en la tab
    'empleados' del Sheet del checador. Quien no marca en el reloj de piso 5 y
    6 —Centros de Maestros, Dirección Académica— no está en esa tab, así que su
    horario autorizado no se escribía en ningún lado aunque la pantalla dijera
    'autorizado'. El padrón es la fuente única, así que aquí siempre se escribe.

    Devuelve (encontrado, row_anterior): row_anterior es el renglón del padrón
    ANTES de sobrescribirlo, para poder archivarlo en HISTORIAL_HORARIOS.
    """
    try:
        sh_p = get_client().open_by_key(st.secrets["sheet_padron_id"])
        ws_p = sh_p.worksheet(PADRON_TAB)
        headers_p = ws_p.row_values(1)
        registros_p = ws_p.get_all_records(numericise_ignore=["all"])
    except Exception as e:
        st.warning(f"No se pudo abrir el Padrón Maestro para escribir el horario: {e}")
        return False, {}

    # El padrón puede no tener columnas de horario todavía: sin ellas no hay
    # dónde escribir y hay que decirlo, no fallar en silencio.
    cols_h = [c for par in COLUMNAS_HORARIO.values() for c in par]
    if not any(c in headers_p for c in cols_h):
        st.warning("El Padrón Maestro no tiene columnas de horario "
                   "(ENTRADA_LUN, SALIDA_LUN, …): el horario no se pudo escribir ahí.")
        return False, {}

    from gspread.cell import Cell
    objetivo = str(emp_id).upper().strip()
    for i, row in enumerate(registros_p, start=2):
        rfc_fila = str(row.get("RFC", "")).upper().strip()
        nom_fila = str(row.get("NOMBRE", row.get("NOMBRE_COMPLETO", ""))).upper().strip()
        if rfc_fila != objetivo and nom_fila != objetivo:
            continue
        celdas = []
        for dia, (col_e, col_s) in COLUMNAS_HORARIO.items():
            if dia not in horario_nuevo:
                continue
            if col_e in headers_p:
                celdas.append(Cell(i, headers_p.index(col_e) + 1, horario_nuevo[dia]["entrada"]))
            if col_s in headers_p:
                celdas.append(Cell(i, headers_p.index(col_s) + 1, horario_nuevo[dia]["salida"]))
        if celdas:
            ws_p.update_cells(celdas, value_input_option="USER_ENTERED")
        return True, row
    return False, {}


def autorizar_cambio_horario(emp_id: str, horario_nuevo: dict, folio: str, fecha_inicio_cho: str = None):
    """Devuelve (encontro_empleado, encontro_incidencia, encontro_padron).

    Antes esta función no devolvía nada y el llamador mostraba éxito sin
    verificar: si el RFC/NOMBRE no calzaba con ninguna fila de 'empleados'
    (típico en personal de Centros de Maestros, que no vive en esa tab) o el
    FOLIO no se encontraba en Incidencias, el cambio no se aplicaba en ningún
    lado pero la pantalla igual decía 'autorizado'.

    Ahora escribe en DOS lugares, porque cada uno lo lee alguien distinto:
      · Padrón Maestro (tab Padron) — fuente única; es el único sitio donde
        vive el personal de Centros de Maestros y Dirección Académica.
      · tab 'empleados' del Sheet del checador — de ahí lee checador_module
        para calcular retardos; si solo escribiéramos el padrón, el reloj
        seguiría midiendo contra el horario viejo.
    """
    client = get_client()
    sh = client.open_by_key(st.secrets["sheet_checador_id"])
    ws_emp = sh.worksheet("empleados")
    registros = ws_emp.get_all_records(numericise_ignore=["all"])
    headers   = ws_emp.row_values(1)
    fecha_hoy = datetime.now().strftime("%Y-%m-%d")
    fecha_aplicacion = fecha_inicio_cho if fecha_inicio_cho else fecha_hoy
    encontro_empleado = False

    for i, row in enumerate(registros, start=2):
        if str(row.get("RFC", "")).upper().strip() == str(emp_id).upper().strip() or str(row.get("NOMBRE", "")).upper().strip() == str(emp_id).upper().strip():
            encontro_empleado = True
            # Guardar horario ANTERIOR en HISTORIAL_HORARIOS antes de sobrescribir
            _registrar_historial_horario(sh, row.get("RFC", emp_id), row.get("NOMBRE", ""),
                                         row, horario_nuevo, fecha_aplicacion)

            # Sobrescribir horario actual en tab empleados
            from gspread.cell import Cell
            celdas_horario = []
            for dia, (col_e, col_s) in COLUMNAS_HORARIO.items():
                if dia in horario_nuevo:
                    if col_e in headers:
                        celdas_horario.append(Cell(i, headers.index(col_e) + 1, horario_nuevo[dia]["entrada"]))
                    if col_s in headers:
                        celdas_horario.append(Cell(i, headers.index(col_s) + 1, horario_nuevo[dia]["salida"]))
            if celdas_horario:
                ws_emp.update_cells(celdas_horario, value_input_option="USER_ENTERED")
            break

    # ── Padrón Maestro: se escribe SIEMPRE, esté o no en 'empleados' ──
    encontro_padron, _fila_padron = _escribir_horario_padron(emp_id, horario_nuevo)
    # Si no estaba en 'empleados', el historial no se escribió arriba: hacerlo
    # aquí con el renglón del padrón, para que el cambio no quede sin rastro.
    if encontro_padron and not encontro_empleado:
        _registrar_historial_horario(
            sh, _fila_padron.get("RFC", emp_id),
            _fila_padron.get("NOMBRE", _fila_padron.get("NOMBRE_COMPLETO", "")),
            _fila_padron, horario_nuevo, fecha_aplicacion)

    ws_inc = sh.worksheet("Incidencias")
    inc_h  = ws_inc.row_values(1)
    inc_d  = ws_inc.get_all_records(numericise_ignore=["all"])
    from gspread.cell import Cell
    ahora_inc = datetime.now(pytz.utc).astimezone(pytz.timezone("America/Mexico_City")).strftime("%Y-%m-%d %H:%M")
    encontro_incidencia = False
    for i, row in enumerate(inc_d, start=2):
        if str(row.get("FOLIO", "")) == folio:
            encontro_incidencia = True
            ws_inc.update_cells([
                Cell(i, inc_h.index("ESTADO") + 1,             "AUTORIZADO"),
                Cell(i, inc_h.index("AUTORIZADO_POR") + 1,     st.session_state.get("nombre", "admin")),
                Cell(i, inc_h.index("FECHA_AUTORIZACION") + 1, ahora_inc),
            ], value_input_option="USER_ENTERED")
            break
    # Sin esto el horario nuevo tardaba hasta 5 min (ttl del cache) en verse,
    # y parecía que no se había guardado.
    # Dónde quedó escrito el horario importa para auditar después: no es lo
    # mismo "autorizado y aplicado" que "autorizado pero no se escribió".
    _destinos = []
    if encontro_padron:   _destinos.append("Padrón Maestro")
    if encontro_empleado: _destinos.append("tab empleados")
    registrar_bitacora_evento(
        "CAMBIO_HORARIO", folio=folio, tipo="CHO", rfc=emp_id,
        estado_anterior="PENDIENTE",
        estado_nuevo="AUTORIZADO" if encontro_incidencia else "",
        resultado="OK" if (encontro_padron or encontro_empleado) else "NO_APLICADO",
        detalle=("horario escrito en: " + ", ".join(_destinos)) if _destinos
                else "no se encontró a la persona ni en el padrón ni en la tab empleados")

    cargar_incidencias.clear()
    cargar_horarios.clear()
    cargar_padron.clear()
    return encontro_empleado, encontro_incidencia, encontro_padron

# ─────────────────────────────────────────────
# AUTENTICACIÓN
# ─────────────────────────────────────────────
def login():
    col_izq, col_centro, col_der = st.columns([1.1, 1.8, 1.1])
    with col_centro:
        with st.container(border=True):
            col_l, col_img, col_r = st.columns([1, 2, 1])
            with col_img:
                st.image("dfc_logo.png", use_container_width=True)
            st.markdown("<h3 style='text-align:center;margin-bottom:0;font-size:22px'>🪪 Ingreso al Sistema</h3>", unsafe_allow_html=True)
            st.markdown("<p style='text-align:center;color:gray;font-size:13px'>Dirección de Formación Continua · Personal</p>", unsafe_allow_html=True)
            st.divider()
            correo    = st.text_input("Correo electrónico institucional", placeholder="nombre@jalisco.gob.mx")
            rfc_input = st.text_input("RFC (Contraseña)", type="password", placeholder="XXXX000000XXX")
            st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
            entrar = st.button("Iniciar Sesión", type="primary", use_container_width=True)

    if entrar:
        if not correo or not rfc_input:
            st.error("Ingresa tu correo y RFC.")
            return

        admin_correo   = st.secrets.get("admin_correo", "")
        admin_rfc      = st.secrets.get("admin_rfc", "")
        # Contraseña propia del admin (secret admin_password). Si existe, el RFC
        # deja de servir como contraseña de admin: quien conozca el RFC no entra.
        # Si no existe, se conserva el comportamiento anterior (correo + RFC).
        admin_password = str(st.secrets.get("admin_password", ""))
        es_correo_admin = bool(admin_correo) and correo.strip().lower() == admin_correo.strip().lower()
        if admin_password:
            import hmac as _hmac
            clave_ok = _hmac.compare_digest(rfc_input.strip(), admin_password)  # respeta mayúsculas
        else:
            clave_ok = rfc_input.upper() == admin_rfc.upper()
        if es_correo_admin and clave_ok:
            st.session_state["rol"]    = "admin"
            st.session_state["correo"] = correo
            # Siempre el RFC real, nunca la contraseña: el resto de la app
            # usa este valor para permisos y registros.
            st.session_state["rfc"]    = admin_rfc.upper()
            st.session_state["nombre"] = "Administrador RH"
            st.rerun()
            return

        # Con contraseña de admin activa, el RFC del admin NO puede entrar por la
        # ruta de empleado (si no, quien conozca tu RFC entraría como "Ángel").
        # Mensaje genérico a propósito: no revela que ese RFC es el del admin.
        if admin_password and admin_rfc and rfc_input.strip().upper() == admin_rfc.upper():
            st.error("RFC no encontrado. Verifica tus datos.")
            return

        usuarios  = cargar_usuarios()
        match_usr = usuarios[usuarios["RFC"].astype(str).str.upper() == rfc_input.upper()]
        if match_usr.empty:
            st.error("RFC no encontrado. Verifica tus datos.")
            return

        usr_row           = match_usr.iloc[0]
        correo_registrado = str(usr_row.get("Correo electrónico institucional", "")).strip().lower()
        if not correo_registrado:
            # SEGURIDAD: sin correo registrado NO se permite el acceso.
            # Antes, un correo vacío en el Sheet dejaba entrar con solo el RFC.
            st.error("Tu RFC no tiene correo registrado en el sistema. Contacta a RH para completar tu registro.")
            return
        if correo.strip().lower() != correo_registrado:
            st.error("El correo no coincide con el RFC registrado.")
            return

        empleados = cargar_empleados()
        match_emp = empleados[empleados["RFC"].astype(str).str.upper() == rfc_input.upper()]
        if not match_emp.empty:
            emp_row         = match_emp.iloc[0]
            nombre_completo = f"{emp_row.get('PATERNO','')} {emp_row.get('MATERNO','')} {emp_row.get('NOMBRE','')}".strip()
            emp_dict        = emp_row.to_dict()
        else:
            # Buscar nombre en tab empleados del reloj checador por RFC
            try:
                horarios_tmp = cargar_horarios()
                match_hor = horarios_tmp[horarios_tmp["RFC"].astype(str).str.upper().str.strip() == rfc_input.upper().strip()]
                if not match_hor.empty:
                    nombre_completo = match_hor.iloc[0].get("NOMBRE", correo.split("@")[0].replace(".", " ").title())
                else:
                    nombre_completo = correo.split("@")[0].replace(".", " ").title()
            except Exception:
                nombre_completo = correo.split("@")[0].replace(".", " ").title()
            emp_dict = {}

        jefe_auto = str(usr_row.get("JEFE_INMEDIATO", "")).strip()
        jefe_pdf_auto = formato_jefe_pdf(jefe_auto)
        director_auto = formato_jefe_pdf(str(usr_row.get("DIRECTOR_GENERAL", "")).strip())
        st.session_state["director_general"] = director_auto
        st.session_state["rol"]             = "empleado"
        st.session_state["correo"]          = correo
        st.session_state["rfc"]             = rfc_input.upper()
        st.session_state["nombre"]          = nombre_completo
        st.session_state["empleado_row"]    = emp_dict
        st.session_state["jefe_inmediato"]  = jefe_pdf_auto
        st.rerun()

# ─────────────────────────────────────────────
# VISTA EMPLEADO
# ─────────────────────────────────────────────
def mapear_emojis_estado(estado: str) -> str:
    est = str(estado).upper()
    if "AUTORIZADO" in est or "✅" in est: return "✅ AUTORIZADO"
    if "RECHAZADO"  in est or "🔴" in est: return "🔴 RECHAZADO"
    return "🟡 PENDIENTE"

# ── TAB RELOJ CHECADOR (solo admin) ──────────────
# ── Reloj checador: extraído a checador_module.py (ver motor v2.2) ──
try:
    import checador_module as _checador_mod
    _ERROR_CHECADOR = ""
except Exception as _e_chk:
    _checador_mod = None
    _ERROR_CHECADOR = str(_e_chk)

# ── Minutario de oficios: módulo aparte (aún no entregado) ──
try:
    import oficios_module as _oficios_mod
    _ERROR_OFICIOS = ""
except Exception as _e_of:
    _oficios_mod = None
    _ERROR_OFICIOS = str(_e_of)

def render_checador():
    if _checador_mod is None:
        st.error(f"El módulo del checador no está disponible: {_ERROR_CHECADOR}")
        return
    _checador_mod.render_checador({
        "cargar_festivos": cargar_festivos,
        "festivos_en_periodo": festivos_en_periodo,
        "get_client": get_client,
        "cargar_justif_direccion": cargar_justif_direccion,
        "guardar_justif_direccion": guardar_justif_direccion,
        "eliminar_justif_direccion": eliminar_justif_direccion,
        "cargar_observaciones": cargar_observaciones,
        "guardar_observaciones": guardar_observaciones,
        "guardar_asistencia_mes": guardar_asistencia_mes,
        "DIRECTORA_NOMBRE": DIRECTORA_NOMBRE,
        "DIRECTORA_CARGO": DIRECTORA_CARGO,
    })

def render_oficios():
    if _oficios_mod is None:
        st.error(f"Módulo de oficios no disponible: {_ERROR_OFICIOS}")
        return
    _oficios_mod.render_oficios({
        "get_client": get_client,
        "rfcs_autorizados": _rfcs_secret("rfcs_oficios"),
    })




def _selector_fechas(prefijo: str, etq_ini: str, etq_fin: str):
    """Rango continuo o días sueltos, igual que días económicos, para los
    tipos de incidencia que abarcan días (comisión, reposición). Devuelve
    (fechas_lista, fi, ff, dias_hab, es_suelto) o None si falta capturar.
    En días sueltos, quien llama agrega "Fechas: ..." al motivo: el
    checador justifica SOLO esos días, no el rango inicio→fin."""
    festivos = cargar_festivos()
    festivos_set = festivos_a_set(festivos)
    modo = st.radio("Selección de fechas", ["Rango continuo", "Días sueltos"],
                    horizontal=True, key=f"{prefijo}_modo")
    if modo == "Rango continuo":
        c1, c2 = st.columns(2)
        with c1:
            fi = st.date_input(etq_ini, value=None, format="DD/MM/YYYY", key=f"{prefijo}_fi")
        with c2:
            ff = st.date_input(etq_fin, value=None, format="DD/MM/YYYY", key=f"{prefijo}_ff")
        if not fi or not ff:
            st.info("Selecciona las fechas de tu solicitud para continuar.")
            return None
        if ff < fi:
            st.error("La fecha fin no puede ser anterior a la fecha inicio.")
            return None
        dias_hab = dias_habiles_entre(fi, ff, festivos)
        st.caption(f"Días hábiles: **{dias_hab}**")
        return [], fi, ff, dias_hab, False
    # Días sueltos: 10 hábiles hacia atrás (comisiones que se registran
    # después) y 60 hacia adelante.
    NOMBRES_DIA = ["Lun", "Mar", "Mié", "Jue", "Vie"]
    MESES = ["", "Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
    atras, d = [], date.today() - timedelta(days=1)
    while len(atras) < 10:
        if d.weekday() < 5 and d not in festivos_set:
            atras.append(d)
        d -= timedelta(days=1)
    adelante, d = [], date.today()
    while len(adelante) < 60:
        if d.weekday() < 5 and d not in festivos_set:
            adelante.append(d)
        d += timedelta(days=1)
    st.caption("Marca cada día que aplica.")
    fechas, cols = [], st.columns(5)
    for i, dia in enumerate(sorted(atras) + adelante):
        if cols[i % 5].checkbox(f"{NOMBRES_DIA[dia.weekday()]} {dia.day}/{MESES[dia.month]}",
                                key=f"{prefijo}_dia_{dia}"):
            fechas.append(dia)
    if not fechas:
        st.info("Selecciona al menos un día.")
        return None
    st.caption(f"Días seleccionados: **{len(fechas)}**")
    return sorted(fechas), min(fechas), max(fechas), len(fechas), True


def vista_empleado():
    rfc      = st.session_state["rfc"]
    nombre   = st.session_state["nombre"]
    emp_data = st.session_state.get("empleado_row", {})

    st.markdown(f'### Sesión iniciada como: {nombre}')
    st.caption('Filiación: ' + rfc_oculto(rfc) + ' · ' + emp_data.get('PUESTO', ''))

    # Alerta de falta no justificada (solo si RH marcó avisar a este empleado)
    try:
        _asis = cargar_asistencia_mes()
        if not _asis.empty and "RFC" in _asis.columns:
            _mi = _asis[_asis["RFC"].astype(str).str.upper() == str(rfc).upper()]
            if not _mi.empty:
                _f = _mi.iloc[0]
                _nojust = int(_f.get("NO_JUSTIFICADAS", 0) or 0)
                if str(_f.get("AVISAR","")).upper() == "SI" and _nojust > 0:
                    _dias = _f.get("DIAS_FALTA","")
                    st.error(
                        f"⚠️ Tienes {_nojust} falta(s) sin justificar"
                        + (f" ({_dias})" if _dias else "")
                        + ". Acude a Recursos Humanos o presenta tu justificante."
                    )
    except Exception:
        pass

    # Alerta días económicos vencen 31 dic
    hoy      = date.today()
    fin_anio = date(hoy.year, 12, 31)
    dias_para_vencer = (fin_anio - hoy).days
    if dias_para_vencer <= 60:
        st.warning(f"⚠️ Tus días económicos disponibles vencen el 31 de diciembre ({dias_para_vencer} días restantes).")

    # ── Carga de datos ───────────────────────────
    solicitudes = cargar_solicitudes_eco()
    empleados   = cargar_empleados()
    incidencias = cargar_incidencias()
    horarios_df = cargar_horarios()

    emp_row          = empleados[empleados["RFC"].astype(str).str.upper() == rfc]
    dias_totales     = int(emp_row["DIAS TOTALES"].iloc[0])     if not emp_row.empty else 0
    dias_usados      = dias_economicos_usados(rfc, solicitudes)
    dias_disponibles = dias_totales - dias_usados
    horas_pases      = horas_pases_mes(rfc, incidencias)

    # ── Tablero de saldos ────────────────────────
    col1, _col2 = st.columns(2)
    col1.metric("📅 Días económicos disponibles", dias_disponibles, delta=f"-{dias_usados} ejercidos")

    st.divider()

    # ── Horario actual ───────────────────────────
    with st.expander("🕐 Mi horario registrado", expanded=False):
        emp_hor = horarios_df[horarios_df["RFC"].astype(str).str.upper().str.strip() == rfc.upper().strip()]
        if not emp_hor.empty:
            row_hor = emp_hor.iloc[0]
            resumen = []
            for dia, (col_e, col_s) in COLUMNAS_HORARIO.items():
                ent = row_hor.get(col_e, "")
                sal = row_hor.get(col_s, "")
                if ent:
                    resumen.append({"Día": dia, "Entrada": ent, "Salida": sal})
            if resumen:
                st.dataframe(pd.DataFrame(resumen), use_container_width=True, hide_index=True)
            else:
                st.caption("Sin horario registrado.")
        else:
            st.caption("No se encontró tu horario en el sistema.")

    # ── Mis solicitudes ──────────────────────────
    st.markdown("#### 📋 Mis solicitudes este mes")
    import pytz
    tz_mx  = pytz.timezone("America/Mexico_City")
    ahora  = datetime.now(pytz.utc).astimezone(tz_mx)
    # Solo incidencias del mes actual
    mis_inc_all = incidencias[incidencias["RFC"].astype(str).str.upper() == rfc].copy()
    if not mis_inc_all.empty:
        mis_inc_all["FECHA_DT"] = pd.to_datetime(mis_inc_all["FECHA_INICIO"], errors="coerce")
        mis_inc = mis_inc_all[(mis_inc_all["FECHA_DT"].dt.year == ahora.year) & (mis_inc_all["FECHA_DT"].dt.month == ahora.month)].copy()
    else:
        mis_inc = mis_inc_all
    # Solo días económicos del mes actual
    sol_hist_all = solicitudes[solicitudes["RFC"].astype(str).str.upper().str.strip() == rfc.upper().strip()].copy() if "RFC" in solicitudes.columns else pd.DataFrame()
    if not sol_hist_all.empty:
        col_fi_eco = next((c for c in sol_hist_all.columns if "Inicio" in c or "INICIO" in c), None)
        if col_fi_eco:
            sol_hist_all["FECHA_DT"] = pd.to_datetime(sol_hist_all[col_fi_eco], errors="coerce")
            sol_hist = sol_hist_all[
                (sol_hist_all["FECHA_DT"].dt.year == ahora.year) &
                (sol_hist_all["FECHA_DT"].dt.month == ahora.month)
            ].copy()
            # Mes sin solicitudes = tabla vacía. NO mostrar el histórico como
            # respaldo: hacía que en julio aparecieran solicitudes de junio.
        else:
            sol_hist = sol_hist_all
    else:
        sol_hist = sol_hist_all

    if not sol_hist.empty:
        sol_hist = sol_hist.rename(columns={
            "Tipo Permiso":     "TIPO",
            "Fecha Inicio":     "FECHA_INICIO",
            "Fecha Fin":        "FECHA_FIN",
            "Dias Solicitados": "DIAS",
            "Aprobado Por":     "AUTORIZADO_POR",
            "Fecha Registro":   "FECHA_SOLICITUD",
        })
        col_folio_sol = next((c for c in sol_hist.columns if "FOLIO" in c.upper()), None)
        if col_folio_sol and col_folio_sol != "FOLIO":
            sol_hist["FOLIO"] = sol_hist[col_folio_sol]
        elif "FOLIO" not in sol_hist.columns:
            sol_hist["FOLIO"] = "HISTÓRICO"
        sol_hist["HORAS_PASE"]         = ""
        sol_hist["ESTADO"]             = sol_hist["AUTORIZADO_POR"].apply(
            lambda x: "🔴 RECHAZADO" if str(x).strip().upper().startswith("RECHAZADO") else ("✅ AUTORIZADO" if str(x).strip() != "" else "🟡 PENDIENTE")
        )
        # Observaciones: texto después del — en Aprobado Por para rechazos
        sol_hist["OBSERVACIONES"] = sol_hist["AUTORIZADO_POR"].apply(
            lambda x: str(x).split("—",1)[1].strip() if str(x).upper().startswith("RECHAZADO") and "—" in str(x) else ""
        )
    if not mis_inc.empty:
        mis_inc["ESTADO"] = mis_inc["ESTADO"].apply(mapear_emojis_estado)

    cols_mostrar = ["FOLIO", "TIPO", "FECHA_INICIO", "FECHA_FIN", "DIAS", "HORAS_PASE", "ESTADO", "FECHA_AUTORIZACION", "OBSERVACIONES"]
    frames = []
    if not sol_hist.empty:
        frames.append(sol_hist[[c for c in cols_mostrar if c in sol_hist.columns]])
    if not mis_inc.empty:
        frames.append(mis_inc[[c for c in cols_mostrar if c in mis_inc.columns]])

    if not frames:
        st.info("No tienes solicitudes este mes.")
    else:
        df_consolidado = pd.concat(frames, ignore_index=True)
        st.dataframe(df_consolidado, use_container_width=True, hide_index=True)

        # ── Alertas de pendientes y rechazos ──────────
        # Pendientes
        col_estado = "ESTADO" if "ESTADO" in df_consolidado.columns else None
        if col_estado:
            pendientes_emp = df_consolidado[df_consolidado[col_estado].astype(str).str.contains("PENDIENTE", case=False)]
            if not pendientes_emp.empty:
                folios_pend = ", ".join(pendientes_emp["FOLIO"].astype(str).tolist())
                st.warning(f"⏳ Tienes {len(pendientes_emp)} solicitud(es) pendiente(s): **{folios_pend}**\n\n"
                           "Para que tu solicitud sea procesada debes **entregar el comprobante PDF firmado** "
                           "en el Área de Recursos Humanos (Administración de Personal DFC). "
                           "Si ya lo entregaste, espera a que el área lo valide.")
    st.divider()

    # ── Firmas de nómina pendientes (lo que RH registró en el Sheet) ──
    try:
        _pn = cargar_pendientes_nomina_sheet()
        if not _pn.empty and "NOMBRE" in _pn.columns:
            if "ESTADO" in _pn.columns:
                _pn = _pn[_pn["ESTADO"].astype(str).str.upper().str.strip() != "FIRMADO"]
            _tok_yo = _tokens_persona(st.session_state.get("nombre", ""))
            _mios = [r for _, r in _pn.iterrows()
                     if _mismo_nombre(_tok_yo, _tokens_persona(r.get("NOMBRE", "")))]
            if _mios:
                st.markdown("#### 💰 Mis firmas de nómina pendientes")
                st.warning(f"Tienes **{len(_mios)}** registro(s) de nómina sin firmar. "
                           "La Dirección de Pagos permite un rezago máximo de 2 quincenas; "
                           "acude a la Dirección a firmar.")
                for _r in _mios:
                    st.markdown(f"• **{_r.get('CONCEPTO','')}** · nómina {_r.get('NOMINA','')}")
                st.divider()
    except Exception:
        pass  # informativo: nunca debe impedir que el empleado use su vista

    # ── Faltas informativas ──────────────────────
    asis = cargar_asistencia_mes()
    mi_asis = pd.DataFrame()
    if not asis.empty and "RFC" in asis.columns:
        mi_asis = asis[asis["RFC"].astype(str).str.upper() == str(rfc).upper()]

    # Título honesto: el contenido es del ÚLTIMO reporte procesado del checador,
    # que casi siempre es del mes anterior. Nunca decir "este mes" si no lo es.
    _titulo_faltas = "📊 Mis faltas del último período procesado"
    try:
        if not mi_asis.empty:
            _per = str(mi_asis.iloc[0].get("PERIODO", ""))
            _f_per = pd.to_datetime(_per.split("~")[0].strip(), errors="coerce")
            if not pd.isna(_f_per):
                _meses_es = {1:"enero",2:"febrero",3:"marzo",4:"abril",5:"mayo",6:"junio",
                             7:"julio",8:"agosto",9:"septiembre",10:"octubre",11:"noviembre",12:"diciembre"}
                _hoy_mx = datetime.now(pytz.timezone("America/Mexico_City"))
                if _f_per.year == _hoy_mx.year and _f_per.month == _hoy_mx.month:
                    _titulo_faltas = "📊 Mis faltas este mes"
                else:
                    _titulo_faltas = f"📊 Mis faltas — reporte de {_meses_es[_f_per.month]} {_f_per.year}"
    except Exception:
        pass
    with st.expander(_titulo_faltas, expanded=False):
        st.info("ℹ️ Este contador es informativo. Las faltas no justificadas las gestiona Recursos Humanos.")
        if mi_asis.empty:
            c1, c2, c3 = st.columns(3)
            c1.metric("Faltas totales",        0)
            c2.metric("Faltas justificadas",   0)
            c3.metric("Faltas no justificadas",0)
            st.caption("Aún no hay un reporte de asistencia procesado para ti.")
        else:
            fila = mi_asis.iloc[0]
            faltas_tot = int(fila.get("FALTAS", 0) or 0)
            justis     = int(fila.get("JUSTIFICADAS", 0) or 0)
            nojust     = int(fila.get("NO_JUSTIFICADAS", 0) or 0)
            c1, c2, c3 = st.columns(3)
            c1.metric("Faltas totales",        faltas_tot)
            c2.metric("Faltas justificadas",   justis)
            c3.metric("Faltas no justificadas",nojust)
            st.caption(f"Período: {fila.get('PERIODO','')}")
            if fila.get("DIAS_FALTA"):
                st.caption(f"Días que faltaste: {fila.get('DIAS_FALTA')}")


    st.divider()
    st.markdown("### Nueva solicitud")

    jefe_guardado = st.session_state.get("jefe_inmediato", "")
    if jefe_guardado:
        jefe_pdf = jefe_guardado
        st.caption(f"👤 Jefe inmediato: {jefe_guardado.split(chr(60))[0]}")
    else:
        try:
            usuarios_df = cargar_usuarios()
            jefes_unicos = sorted({
                str(j).strip() for j in usuarios_df.get("JEFE_INMEDIATO", [])
                if str(j).strip()
            })
        except Exception:
            jefes_unicos = []
        if jefes_unicos:
            jefe_sel = st.selectbox("👤 Jefe inmediato que autoriza", options=jefes_unicos)
            jefe_pdf = formato_jefe_pdf(jefe_sel)
        else:
            jefe_manual = st.text_input("👤 Jefe inmediato que autoriza (nombre — cargo)")
            jefe_pdf = formato_jefe_pdf(jefe_manual)

    tipo = st.selectbox(
        "Tipo de incidencia",
        options=list(TIPO_LABELS.keys()),
        format_func=lambda x: TIPO_LABELS[x],
    )

    # ── DÍA ECONÓMICO ──────────────────────────
    if tipo == "ECO":
        st.info(f"Tienes **{dias_disponibles}** día(s) económico(s) disponibles. Vencen el 31 de diciembre.")
        if dias_disponibles <= 0:
            st.warning("No tienes días económicos disponibles.")
            return
        # Modo de selección: rango o días sueltos
        modo = st.radio("Selección de fechas", ["Rango continuo", "Días sueltos"], horizontal=True)
        festivos = cargar_festivos()
        festivos_set = festivos_a_set(festivos)

        if modo == "Rango continuo":
            col1, col2 = st.columns(2)
            with col1:
                fi = st.date_input("Fecha inicio", value=None, format="DD/MM/YYYY")
            with col2:
                ff = st.date_input("Fecha fin",    value=None, format="DD/MM/YYYY")
            # Las fechas nacen VACÍAS a propósito. Cuando venían precargadas
            # con hoy, el formulario era válido desde que cargaba y bastaba
            # un clic distraído para generar un folio con la fecha de hoy.
            if not fi or not ff:
                st.info("Selecciona las fechas de tu solicitud para continuar.")
                return
            if ff < fi:
                st.error("La fecha fin no puede ser anterior a la fecha inicio.")
                return
            dias_hab = dias_habiles_entre(fi, ff, festivos)
            st.caption(f"Días hábiles solicitados: **{dias_hab}**")
            if dias_hab > dias_disponibles:
                st.error(f"Excedes tus días disponibles ({dias_disponibles}).")
                return
            fechas_seleccionadas = []
            d = fi
            while d <= ff:
                if d.weekday() < 5 and d not in festivos_set:
                    fechas_seleccionadas.append(d)
                d += timedelta(days=1)
            fi_final, ff_final = fi, ff
        else:
            st.caption("Selecciona cada día económico que necesitas tomar.")
            # Mostrar checkboxes para los próximos 60 días hábiles
            from datetime import timedelta as _td
            dias_opciones = []
            d = date.today()
            while len(dias_opciones) < 60:
                if d.weekday() < 5 and d not in festivos_set:
                    dias_opciones.append(d)
                d += _td(days=1)
            NOMBRES_DIA = ["Lun","Mar","Mié","Jue","Vie"]
            MESES = ["","Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"]
            fechas_seleccionadas = []
            cols_sel = st.columns(5)
            for idx, dia in enumerate(dias_opciones):
                label = f"{NOMBRES_DIA[dia.weekday()]} {dia.day}/{MESES[dia.month]}"
                if cols_sel[idx % 5].checkbox(label, key=f"eco_dia_{dia}"):
                    fechas_seleccionadas.append(dia)
            dias_hab = len(fechas_seleccionadas)
            if dias_hab > 0:
                st.caption(f"Días seleccionados: **{dias_hab}**")
            if dias_hab > dias_disponibles:
                st.error(f"Excedes tus días disponibles ({dias_disponibles}).")
                return
            fi_final = min(fechas_seleccionadas) if fechas_seleccionadas else date.today()
            ff_final = max(fechas_seleccionadas) if fechas_seleccionadas else date.today()

        motivo = st.text_area("Motivo (opcional)", max_chars=300)
        archivo_eco = st.file_uploader("Adjuntar documento de soporte (opcional)", type=["pdf","png","jpg","jpeg"], key="eco_anexo")
        tiene_anexo_eco = archivo_eco is not None

        # Construir motivo con fechas exactas
        if fechas_seleccionadas:
            fechas_str = ", ".join([f.strftime("%d/%m/%Y") for f in sorted(fechas_seleccionadas)])
            motivo_final = f"{motivo} | Fechas: {fechas_str}".strip(" |")
        else:
            motivo_final = motivo

        if dias_hab == 0:
            st.warning("Selecciona al menos un día.")
            return
        enviar_solicitud(rfc, nombre, tipo, fi_final, ff_final, dias_hab, 0.0, motivo_final, tiene_anexo_eco, incidencias, archivo_eco, jefe_inmediato=jefe_pdf)

    # ── PASES ────────────────────────────────────
    elif tipo in ["PSE", "PSR", "PEN"]:
        # El tipo ya define el subtipo
        subtipo = TIPO_LABELS[tipo]
        fecha   = st.date_input("Fecha del pase", value=None, format="DD/MM/YYYY")
        if not fecha:
            st.info("Selecciona la fecha del pase para continuar.")
            return

        col1, col2 = st.columns(2)
        hora_salida = hora_entrada = hora_retorno = None
        with col1:
            if tipo in ["PSE", "PSR"]:
                hora_salida = st.text_input("Hora de salida (HH:MM)", placeholder="08:37", max_chars=5)
                hora_salida = hora_salida.strip() if hora_salida else None
        with col2:
            if tipo == "PEN":
                hora_entrada = st.text_input("Hora de entrada (HH:MM)", placeholder="08:37", max_chars=5)
                hora_entrada = hora_entrada.strip() if hora_entrada else None
            elif tipo == "PSR":
                hora_retorno = st.text_input("Hora estimada de retorno (HH:MM)", placeholder="10:30", max_chars=5)
                hora_retorno = hora_retorno.strip() if hora_retorno else None

        # Calcular horas del pase
        horas_pase = 0.0
        if hora_salida and hora_retorno:
            try:
                sal  = datetime.combine(fecha, datetime.strptime(hora_salida.strip(),  "%H:%M").time())
                ret  = datetime.combine(fecha, datetime.strptime(hora_retorno.strip(), "%H:%M").time())
                diff = (ret - sal).total_seconds() / 3600
                horas_pase = round(max(diff, 0), 2)
                st.caption(f"Horas de ausencia: **{horas_pase}h**")
            except:
                st.warning("Formato inválido. Usa HH:MM (ej: 08:37)")
        elif hora_salida:
            try:
                horarios_df_pse = cargar_horarios()
                emp_hor_pse = horarios_df_pse[horarios_df_pse["RFC"].astype(str).str.upper().str.strip() == rfc.upper().strip()]
                nd_pse = ["LUN","MAR","MIE","JUE","VIE","SAB","DOM"][fecha.weekday()]
                salida_prog = str(emp_hor_pse.iloc[0].get(f"SALIDA_{nd_pse}","")).strip() if not emp_hor_pse.empty else ""
                if salida_prog:
                    sal_dt = datetime.combine(fecha, datetime.strptime(hora_salida.strip(), "%H:%M").time())
                    fin_dt = datetime.combine(fecha, datetime.strptime(salida_prog, "%H:%M").time())
                    horas_pase = round(max((fin_dt - sal_dt).total_seconds() / 3600, 0), 2)
                    st.caption(f"Horas de ausencia (hasta fin de jornada {salida_prog}): **{horas_pase}h**")
                else:
                    horas_pase = 0.0
            except:
                horas_pase = 0.0
                st.warning("No se pudo calcular automáticamente.")
        elif hora_entrada:
            # Pase de entrada: horas ausentes = desde hora programada hasta hora real de llegada
            try:
                horarios_df_pse = cargar_horarios()
                emp_hor_pse = horarios_df_pse[horarios_df_pse["RFC"].astype(str).str.upper().str.strip() == rfc.upper().strip()]
                nd_pse = ["LUN","MAR","MIE","JUE","VIE","SAB","DOM"][fecha.weekday()]
                entrada_prog = str(emp_hor_pse.iloc[0].get(f"ENTRADA_{nd_pse}","")).strip() if not emp_hor_pse.empty else ""
                if entrada_prog:
                    prog_dt  = datetime.combine(fecha, datetime.strptime(entrada_prog, "%H:%M").time())
                    real_dt  = datetime.combine(fecha, datetime.strptime(hora_entrada.strip(), "%H:%M").time())
                    horas_pase = round(max((real_dt - prog_dt).total_seconds() / 3600, 0), 2)
                    st.caption(f"Horas de ausencia (desde {entrada_prog} hasta {hora_entrada}): **{horas_pase}h**")
                else:
                    horas_pase = 0.0
            except:
                horas_pase = 0.0
                st.warning("No se pudo calcular automáticamente.")


        motivo = st.text_area("Motivo", max_chars=300)
        archivo_anexo = st.file_uploader("Adjuntar justificante (opcional)", type=["pdf","png","jpg","jpeg"])
        tiene_anexo   = archivo_anexo is not None

        # ── La hora es OBLIGATORIA en un pase ────────────────────────────
        # Antes se podía enviar sin ella: el pase quedaba con 0 horas y no se
        # descontaba nada. Pasó de verdad (PSE-2026-0078 y PEN-2026-0068, que
        # llevan la hora escrita dentro del texto del motivo y 0 en HORAS).
        # Como la hora es lo que define cuánto se ausenta la persona, un pase
        # sin hora no es un pase: aquí se detiene antes de llegar al botón.
        _faltantes = []
        if tipo in ("PSE", "PSR") and not hora_salida:
            _faltantes.append("hora de salida")
        if tipo == "PSR" and not hora_retorno:
            _faltantes.append("hora de retorno")
        if tipo == "PEN" and not hora_entrada:
            _faltantes.append("hora de entrada")

        # El formato también: "10:45 am" o "1045" dan 0 horas en silencio.
        # Se exigen DOS dígitos en los minutos: "9:5" lo aceptaba strptime
        # como 9:05, pero quien lo escribe bien puede haber querido 9:50.
        # Una hora ambigua en un pase se traduce en horas de ausencia mal
        # contadas, así que se pide que la escriba completa.
        def _hora_valida(h):
            txt = str(h).strip()
            if not re.fullmatch(r"\d{1,2}:\d{2}", txt):
                return False
            try:
                datetime.strptime(txt, "%H:%M")
                return True
            except Exception:
                return False
        _malformadas = [etq for etq, val in (("salida", hora_salida),
                                             ("entrada", hora_entrada),
                                             ("retorno", hora_retorno))
                        if val and not _hora_valida(val)]

        if _faltantes:
            st.warning("⏰ Falta la **" + "** y la **".join(_faltantes) +
                       "**. Captúrala en su campo (formato 24 horas, por ejemplo "
                       "**14:30**); no la escribas en el motivo, porque entonces "
                       "el sistema no puede calcular tus horas de ausencia.")
            return
        if _malformadas:
            st.error("⏰ La hora de " + " y de ".join(_malformadas) +
                     " no tiene el formato correcto. Usa 24 horas con dos puntos: "
                     "**08:30**, **14:05**, **16:00**.")
            return
        if horas_pase <= 0:
            st.error("⏰ El cálculo dio **0 horas de ausencia**. Revisa que las "
                     "horas sean correctas; si tu horario de ese día no está en "
                     "el sistema, avisa a Recursos Humanos antes de continuar.")
            return

        detalle = subtipo
        if hora_salida:  detalle += f" | Salida: {hora_salida}"
        if hora_entrada: detalle += f" | Entrada: {hora_entrada}"
        if hora_retorno: detalle += f" | Retorno: {hora_retorno}"
        motivo_completo = (detalle + "\n" + motivo).strip()
        enviar_solicitud(rfc, nombre, tipo, fecha, fecha, 0, horas_pase, motivo_completo, tiene_anexo, incidencias, archivo_anexo, subtipo_label=subtipo, hora_retorno=hora_retorno or "", jefe_inmediato=jefe_pdf)

    # ── COMISIÓN ────────────────────────────────
    elif tipo == "COM":
        _sel = _selector_fechas("com", "Fecha inicio comisión", "Fecha fin comisión")
        if _sel is None:
            return
        _fechas_com, fi, ff, dias_hab, _suelto_com = _sel
        motivo      = st.text_area("Motivo de la comisión", max_chars=300)
        if _suelto_com:
            motivo = (f"{motivo} | Fechas: " if motivo else "Fechas: ") + \
                     ", ".join(f.strftime("%d/%m/%Y") for f in _fechas_com)
        archivo_anexo = st.file_uploader("Adjuntar constancia/oficio (opcional)", type=["pdf","png","jpg","jpeg"])
        tiene_anexo   = archivo_anexo is not None
        if not tiene_anexo:
            st.caption("Recuerda que sin anexo tu solicitud puede quedar sin soporte documental.")
        enviar_solicitud(rfc, nombre, tipo, fi, ff, dias_hab, 0.0,
                         motivo, tiene_anexo, incidencias, archivo_anexo, jefe_inmediato=jefe_pdf)
    # ── REPOSICIÓN DE GUARDIAS ──────────────────
    elif tipo == "RGU":
        st.info("Captura los días que repones a cambio de guardia(s) previamente realizada(s).")
        _sel = _selector_fechas("rgu", "Fecha inicio reposición", "Fecha fin reposición")
        if _sel is None:
            return
        _fechas_rgu, fi, ff, dias_hab, _suelto_rgu = _sel
        fecha_guardia = st.text_input(
            "Fecha(s) de la guardia que repones",
            placeholder="27/04/2026",
            help="Indica la fecha de la guardia realizada que compensas con estos días."
        )
        motivo = st.text_area("Observaciones (opcional)", max_chars=300)
        archivo_anexo = st.file_uploader("Adjuntar formato de guardia / soporte (opcional)", type=["pdf","png","jpg","jpeg"])
        tiene_anexo   = archivo_anexo is not None
        motivo_rgu = f"Reposición de guardia | Guardia repuesta: {fecha_guardia}".strip()
        if motivo:
            motivo_rgu += f" | {motivo}"
        if _suelto_rgu:  # al FINAL: el checador lee todo lo que sigue a "Fechas:"
            motivo_rgu += " | Fechas: " + ", ".join(f.strftime("%d/%m/%Y") for f in _fechas_rgu)
        if dias_hab == 0:
            st.warning("Selecciona al menos un día hábil para reponer.")
            return
        if not fecha_guardia.strip():
            st.warning("Indica la fecha de la guardia que estás reponiendo.")
            return
        enviar_solicitud(rfc, nombre, tipo, fi, ff, dias_hab, 0.0,
                         motivo_rgu, tiene_anexo, incidencias, archivo_anexo, jefe_inmediato=jefe_pdf)
    # ── CAMBIO DE HORARIO ───────────────────────
    elif tipo == "CHO":
        fecha_inicio_cho = st.date_input("¿A partir de qué fecha aplica el cambio?", value=None, format="DD/MM/YYYY")
        if not fecha_inicio_cho:
            st.info("Selecciona la fecha en que aplica el cambio para continuar.")
            return
        motivo      = st.text_area("Motivo del cambio de horario", max_chars=300)
        # Antes solo preguntaba "¿Traerás documento?" sin dejar subirlo.
        archivo_anexo = st.file_uploader("Adjuntar documento de soporte (oficio, etc.) (opcional)",
                                         type=["pdf","png","jpg","jpeg"], key="cho_anexo")
        tiene_anexo = archivo_anexo is not None
        st.markdown("**Horario solicitado:**")
        st.caption("Deja en blanco los días que no labora.")
        horario_solicitado = {}
        cols_dias = st.columns(7)
        for idx, dia in enumerate(DIAS_SEMANA):
            with cols_dias[idx]:
                st.caption(dia)
                ne = st.text_input("Entrada", value="", key=f"cho_e_{dia}", max_chars=5,
                                   placeholder="08:00")
                ns = st.text_input("Salida",  value="", key=f"cho_s_{dia}", max_chars=5,
                                   placeholder="16:00")
                horario_solicitado[dia] = {"entrada": ne.strip(), "salida": ns.strip()}
        # Serializar horario para guardarlo en MOTIVO
        horario_str = " | ".join(
            f"{dia} {v['entrada']}-{v['salida']}"
            for dia, v in horario_solicitado.items()
            if v["entrada"]
        )
        motivo_completo = f"Horario solicitado: {horario_str} | Motivo: {motivo}".strip()
        enviar_solicitud(rfc, nombre, tipo, fecha_inicio_cho, fecha_inicio_cho, 0, 0.0,
                         motivo_completo, tiene_anexo, incidencias, archivo_anexo, jefe_inmediato=jefe_pdf)

    # ── CUMPLEAÑOS (oculto hasta autorización) ──
    if HABILITAR_CUMPLEANOS and tipo == "CUM":
        cumple = cumpleanos_laboral(rfc)
        if cumple:
            st.info(f"Tu día de cumpleaños hábil es: **{cumple.strftime('%d/%m/%Y')}**")
            fi = cumple
            ff = cumple
            motivo = "Día de cumpleaños"
            st.caption("Solo el día hábil correspondiente a tu fecha de nacimiento.")
            enviar_solicitud(rfc, nombre, tipo, fi, ff, 1, 0.0, motivo, False, incidencias, jefe_inmediato=jefe_pdf)
        else:
            st.error("No se pudo calcular tu fecha de cumpleaños desde el RFC.")


def enviar_solicitud(rfc, nombre, tipo, fi, ff, dias, horas_pase, motivo, tiene_anexo, incidencias_df, archivo_anexo=None, subtipo_label=None, hora_retorno="", jefe_inmediato=""):
    # Identidad de la solicitud: misma persona, mismo tipo, mismas fechas.
    clave    = f"{tipo}_{rfc}_{fi}_{ff}"
    key_btn  = f"btn_registrar_{clave}"
    key_ok   = f"registrado_{clave}"

    # ── Si ya se registró en esta sesión, NO se vuelve a dibujar el botón ──
    # Antes el éxito se mostraba solo en el run del clic: al descargar el PDF
    # o al mover cualquier campo, el mensaje desaparecía y el botón volvía a
    # verse disponible. De ahí venían los folios dobles.
    ya = st.session_state.get(key_ok)
    if ya:
        st.success(f"✅ Esta solicitud ya quedó registrada con folio **{ya['folio']}**")
        st.download_button(
            label="⬇️ Descargar comprobante PDF",
            data=ya["pdf"],
            file_name=f"Comprobante_{ya['folio']}.pdf",
            mime="application/pdf",
            key=f"dl_{clave}",
        )
        st.info("Imprime tu comprobante y preséntalo en RH junto con el documento físico original.")
        st.caption("Si necesitas otra solicitud, cambia las fechas o vuelve al menú principal.")
        return

    # Resumen ANTES del botón: la persona ve exactamente qué fechas va a
    # registrar. Sin esto, un clic distraído generaba un folio con la fecha
    # equivocada y había que rechazarlo a mano.
    try:
        _et = TIPO_LABELS.get(tipo, tipo)
        if fi == ff:
            _cuando = f"el **{fi.strftime('%d/%m/%Y')}**"
        else:
            _cuando = f"del **{fi.strftime('%d/%m/%Y')}** al **{ff.strftime('%d/%m/%Y')}**"
        st.warning(f"Vas a registrar **{_et}** {_cuando}. Revisa las fechas antes de continuar.")
    except Exception:
        pass

    if st.button("Registrar solicitud", type="primary", key=key_btn):
        # Relectura EN VIVO: si la sesión se recicló (Streamlit Cloud tira la
        # sesión tras unos minutos, o el usuario recargó con F5), los candados
        # de session_state se perdieron y el Sheet es la única defensa real.
        try:
            incidencias_df = cargar_incidencias()
        except Exception:
            pass

        # Validar duplicado: mismo RFC, mismo tipo y misma fecha de inicio.
        # La comparación va normalizada a fecha: Google Sheets devuelve
        # "15/01/2026" mientras que str(fi) da "2026-01-15", así que la
        # comparación de texto plana NUNCA detectaba el duplicado.
        if not incidencias_df.empty and "FECHA_INICIO" in incidencias_df.columns:
            _fechas = pd.to_datetime(incidencias_df["FECHA_INICIO"],
                                     errors="coerce", dayfirst=True).dt.date
            _mismo_dia = (_fechas == fi) | (
                incidencias_df["FECHA_INICIO"].astype(str).str.strip() == str(fi)
            )
            dup = incidencias_df[
                (incidencias_df["RFC"].astype(str).str.strip().str.upper() == str(rfc).strip().upper()) &
                (incidencias_df["TIPO"].astype(str).str.strip() == str(tipo).strip()) &
                _mismo_dia
            ]
            if not dup.empty:
                folio_dup = dup.iloc[0]["FOLIO"]
                tipo_label_dup = TIPO_LABELS.get(tipo, tipo)
                st.error(f"⚠️ Ya tienes una solicitud **{folio_dup}** de **{tipo_label_dup}** registrada para el **{fi.strftime('%d/%m/%Y')}**. No puedes registrar dos solicitudes del mismo tipo en la misma fecha.")
                return

        # Validar duplicado en ECO
        if tipo == "ECO":
            try:
                sol_eco = cargar_solicitudes_eco()
                if not sol_eco.empty:
                    col_rfc_eco = next((c for c in sol_eco.columns if c.upper() == "RFC"), None)
                    col_fi_eco  = next((c for c in sol_eco.columns if "INICIO" in c.upper()), None)
                    if col_rfc_eco and col_fi_eco:
                        sol_eco["_FDT"] = pd.to_datetime(sol_eco[col_fi_eco], errors="coerce", dayfirst=True)
                        dup_eco = sol_eco[
                            (sol_eco[col_rfc_eco].astype(str).str.strip().str.upper() == str(rfc).strip().upper()) &
                            (sol_eco["_FDT"].dt.date == fi)
                        ]
                        if not dup_eco.empty:
                            col_f = next((c for c in sol_eco.columns if "FOLIO" in c.upper()), None)
                            folio_dup = str(dup_eco.iloc[0].get(col_f, "ECO")) if col_f else "ECO"
                            st.error(f"⚠️ Ya tienes una solicitud **{folio_dup}** de **Día económico** registrada para el **{fi.strftime('%d/%m/%Y')}**. No puedes registrar dos días económicos para la misma fecha.")
                            return
            except Exception:
                pass

        folio     = generar_folio(tipo, incidencias_df)
        import pytz
        tz_mx = pytz.timezone("America/Mexico_City")
        fecha_sol = datetime.now(pytz.utc).astimezone(tz_mx).strftime("%Y-%m-%d %H:%M")
        link_anexo = ""
        if archivo_anexo is not None:
            with st.spinner("Subiendo anexo a Drive..."):
                link_anexo = subir_anexo_drive(archivo_anexo, folio, rfc)
            # Si Drive falla, la función devuelve "ERROR: ...". Antes ese texto
            # se guardaba en LINK_ANEXO como si fuera un link válido.
            if link_anexo.startswith("ERROR:"):
                st.warning("La solicitud se registrará, pero el anexo NO se pudo subir "
                           "a Drive. Entrégalo en físico en RH o vuelve a intentarlo después.")
                link_anexo = ""
                archivo_anexo = None
        datos = {
            "folio":           folio,
            "rfc":             rfc,
            "nombre":          nombre,
            "tipo":            tipo,
            "tipo_label":      TIPO_LABELS[tipo],
            "fecha_solicitud": fecha_sol,
            "fecha_inicio":    str(fi),
            "fecha_fin":       str(ff),
            "dias":            dias,
            "horas_pase":      horas_pase,
            "motivo":          motivo,
            "tiene_anexo":     tiene_anexo or archivo_anexo is not None,
            "link_anexo":      link_anexo,
            "subtipo_label":   subtipo_label if subtipo_label else TIPO_LABELS[tipo],
            "jefe_inmediato":   jefe_inmediato,
            "director_general": st.session_state.get("director_general", ""),
            "hora_retorno":    hora_retorno,
        }
        _guardado_ok = False
        with st.spinner("Registrando solicitud..."):
            try:
                if tipo == "ECO":
                    _guardado_ok = guardar_dia_economico(datos)
                else:
                    _guardado_ok = guardar_incidencia(datos)
            except Exception as _e_save:
                # El fallo queda asentado ANTES de avisar al empleado: aunque
                # la solicitud no se haya escrito, el intento sí deja rastro.
                registrar_bitacora_sol(datos, "ERROR", f"{type(_e_save).__name__}: {_e_save}")
                _error_amable(_e_save, "al registrar la solicitud")
            else:
                registrar_bitacora_sol(
                    datos,
                    "GUARDADA" if _guardado_ok else "NO_GUARDADA",
                    "" if _guardado_ok else "el candado anti-duplicado bloqueó el envío")

        if _guardado_ok:
            folio = datos["folio"]  # puede haberse recalculado en vivo al guardar
            pdf_bytes = generar_comprobante_pdf(datos)
            # La marca se pone SOLO si el guardado salió bien. Antes se ponía
            # antes de guardar, así que un error de cuota de Google dejaba al
            # usuario bloqueado sin haber registrado nada.
            st.session_state[key_ok] = {"folio": folio, "pdf": pdf_bytes}
            st.rerun()


# ─────────────────────────────────────────────
# VISTA ADMIN
# ─────────────────────────────────────────────
def vista_admin():
    st.markdown("### Panel de administración · Incidencias")
    # Si algún intento quedó en cola porque Google estaba caído, se vuelca aquí.
    _volcar_bitacora_sol()

    incidencias = cargar_incidencias()
    horarios_df = cargar_horarios()

    tab1, tab2, tab3, tab4, tab5 = st.tabs(["Pendientes", "Historial completo", "Reporte mensual", "🕐 Reloj Checador", "📋 Nómina"])

    with tab1:
        with st.expander("🔄 Sincronizar saldos de días económicos al Sheet"):
            st.caption("Recalcula DIAS DISPONIBLES (= totales − usados aprobados) y lo escribe "
                       "en la hoja Empleados para todos. Útil para dejar el Sheet al día.")
            if st.button("Actualizar todos los saldos ahora"):
                with st.spinner("Recalculando y escribiendo en el Sheet..."):
                    try:
                        n, errs = actualizar_dias_disponibles_sheet()
                        if errs:
                            st.error(" · ".join(errs))
                        else:
                            st.success(f"✅ {n} saldo(s) actualizado(s) en la hoja Empleados.")
                    except Exception as e:
                        st.error(f"No se pudo actualizar: {e}")

        solicitudes_eco = cargar_solicitudes_eco()
        col_aprobado = next((c for c in solicitudes_eco.columns if "Aprobado" in c), None)
        col_folio    = next((c for c in solicitudes_eco.columns if "FOLIO" in c.upper() or "Folio" in c), None)
        eco_pend = pd.DataFrame()
        if col_aprobado and not solicitudes_eco.empty:
            eco_pend = solicitudes_eco[solicitudes_eco[col_aprobado].astype(str).str.strip() == ""]

        # Antes: incidencias["ESTADO"] == "PENDIENTE" — comparación cruda. Un
        # espacio o una minúscula en la celda del Sheet sacaba la solicitud de
        # esta lista sin autorizarla ni rechazarla.
        pendientes = incidencias[_mascara_pendiente(incidencias)]
        # Las que solo entran por anomalía de la celda se señalan para revisar.
        _raras = pendientes[~_estado_norm(pendientes["ESTADO"]).eq("PENDIENTE")]
        if not _raras.empty:
            st.warning("⚠️ " + str(len(_raras)) + " solicitud(es) con el campo ESTADO alterado "
                       "en el Sheet (espacios, minúsculas o celda vacía): " +
                       ", ".join(_raras["FOLIO"].astype(str).tolist()[:8]) +
                       ". Se muestran aquí para que no se pierdan; corrige el ESTADO "
                       "a PENDIENTE en el Sheet.")
        total_pend = len(pendientes) + len(eco_pend)

        if total_pend == 0:
            st.success("No hay solicitudes pendientes.")
        else:
            st.caption(f"{total_pend} solicitud(es) pendiente(s)")

            # Días económicos primero
            for idx, row in eco_pend.iterrows():
                folio_eco  = str(row.get(col_folio, "")) if col_folio else f"ECO-{idx}"
                nombre_eco = str(row.get("Nombre Completo", row.get("NOMBRE", "")))
                with st.expander(f"**{folio_eco}** · Día económico · {nombre_eco}"):
                    col1, col2 = st.columns(2)
                    col1.write(f"**RFC:** {str(row.get('RFC',''))}")
                    col1.write(f"**Fecha inicio:** {str(row.get('Fecha Inicio',''))}")
                    col1.write(f"**Fecha fin:** {str(row.get('Fecha Fin',''))}")
                    col1.write(f"**Días:** {str(row.get('Dias Solicitados',''))}")
                    col2.write(f"**Motivo:** {str(row.get('Motivo',''))}")
                    col2.write(f"**Registrado:** {str(row.get('Fecha Registro',''))}")
                    obs_eco = st.text_input("Observaciones (para rechazo)", key=f"obs_eco_{idx}")
                    col_a, col_r = st.columns(2)
                    _rfc_eco   = str(row.get("RFC", ""))
                    _fi_eco    = str(row.get("Fecha Inicio", ""))
                    _freg_eco  = str(row.get("Fecha Registro", ""))
                    _folio_real = str(row.get(col_folio, "")) if col_folio else ""
                    with col_a:
                        if st.button("✅ Aprobar", key=f"eco_{idx}", type="primary"):
                            aprobar_dia_economico(st.session_state.get("nombre", "admin"),
                                                  folio=_folio_real, rfc=_rfc_eco,
                                                  fecha_inicio=_fi_eco, fecha_registro=_freg_eco)
                            st.rerun()
                    with col_r:
                        if st.button("❌ Rechazar", key=f"rec_eco_{idx}"):
                            if not obs_eco:
                                st.error("Escribe una observación para rechazar.")
                            else:
                                rechazar_dia_economico(obs_eco,
                                                       folio=_folio_real, rfc=_rfc_eco,
                                                       fecha_inicio=_fi_eco, fecha_registro=_freg_eco)
                                st.rerun()
            for _, row in pendientes.iterrows():
                with st.expander(f"**{row['FOLIO']}** · {TIPO_LABELS.get(row['TIPO'], row['TIPO'])} · {row['NOMBRE']}"):
                    col1, col2 = st.columns(2)
                    col1.write('**Filiación:** ' + rfc_oculto(str(row['RFC'])))
                    col1.write(f"**Fechas:** {row['FECHA_INICIO']} → {row['FECHA_FIN']}")
                    col1.write(f"**Días:** {row['DIAS']}")
                    col1.write(f"**Horas pase:** {row.get('HORAS_PASE','')}")
                    col2.write(f"**Motivo:** {row['MOTIVO']}")
                    col2.write(f"**Anexo:** {row['TIENE_ANEXO']}")
                    col2.write(f"**Registrado:** {row['FECHA_SOLICITUD']}")

                    obs = st.text_input("Observaciones", key=f"obs_{row['FOLIO']}")
                    col_a, col_r = st.columns(2)
                    with col_a:
                        if st.button("✅ Autorizar", key=f"aut_{row['FOLIO']}", type="primary"):
                            if row["TIPO"] == "CHO":
                                st.warning("Para cambio de horario usa el editor de horario abajo.")
                            else:
                                autorizar_incidencia(row["FOLIO"], obs)
                                st.success(f"Folio {row['FOLIO']} autorizado.")
                                st.rerun()
                    with col_r:
                        if st.button("❌ Rechazar", key=f"rec_{row['FOLIO']}"):
                            if not obs:
                                st.error("Escribe una observación para rechazar.")
                            else:
                                rechazar_incidencia(row["FOLIO"], obs)
                                st.warning(f"Folio {row['FOLIO']} rechazado.")
                                st.rerun()

                    if row["TIPO"] == "CHO":
                        st.divider()
                        st.markdown("**Ajustar horario en el sistema:**")
                        emp_hor = horarios_df[horarios_df["RFC"].astype(str).str.upper().str.strip() == str(row["RFC"]).upper().strip()]

                        # Parsear horario solicitado del MOTIVO
                        motivo_cho = str(row.get("MOTIVO", ""))
                        horario_parsed = {}
                        for parte in motivo_cho.split(" | "):
                            for dia in DIAS_SEMANA:
                                if parte.strip().startswith(dia + " "):
                                    try:
                                        horas = parte.strip()[len(dia)+1:]
                                        e, s = horas.split("-")
                                        horario_parsed[dia] = {"entrada": e.strip(), "salida": s.strip()}
                                    except Exception:
                                        pass

                        nuevo_horario = {}
                        cols_dias = st.columns(5)
                        for idx, dia in enumerate(DIAS_SEMANA):
                            col_e, col_s = COLUMNAS_HORARIO[dia]
                            # Prellenar con horario solicitado, si no con el actual del Sheet
                            if dia in horario_parsed:
                                default_e = horario_parsed[dia]["entrada"]
                                default_s = horario_parsed[dia]["salida"]
                            elif not emp_hor.empty:
                                default_e = str(emp_hor.iloc[0].get(col_e, ""))
                                default_s = str(emp_hor.iloc[0].get(col_s, ""))
                            else:
                                default_e = ""
                                default_s = ""
                            with cols_dias[idx]:
                                st.caption(dia)
                                ne = st.text_input("Entrada", value=default_e, key=f"ne_{row['FOLIO']}_{dia}", max_chars=5)
                                ns = st.text_input("Salida",  value=default_s, key=f"ns_{row['FOLIO']}_{dia}", max_chars=5)
                                nuevo_horario[dia] = {"entrada": ne, "salida": ns}
                        if st.button("💾 Guardar horario y autorizar", key=f"hor_{row['FOLIO']}", type="primary"):
                            _enc_emp, _enc_inc, _enc_pad = autorizar_cambio_horario(
                                row["RFC"], nuevo_horario, row["FOLIO"])
                            # El RFC se enmascara igual que en el resto de la
                            # pantalla: no tiene por qué aparecer completo aquí.
                            _rfc_msg = rfc_oculto(str(row["RFC"]))
                            if not _enc_pad and not _enc_emp:
                                st.error(f"⚠️ No se encontró a {_rfc_msg} ni en el Padrón Maestro ni "
                                         "en la tab 'empleados': el horario NO se guardó en ningún "
                                         "lado. Captura primero a esta persona en el padrón.")
                            elif not _enc_pad:
                                st.warning("Horario actualizado en la tab 'empleados' (reloj checador), "
                                           f"pero {_rfc_msg} no está en el Padrón Maestro. Captúralo ahí: "
                                           "el padrón es la fuente única.")
                            elif not _enc_emp:
                                st.success("Horario actualizado en el **Padrón Maestro**.")
                                st.info("No está en la tab 'empleados' del reloj checador — normal si "
                                        "no marca en piso 5 y 6.")
                            else:
                                st.success("Horario actualizado en el **Padrón Maestro** y en la tab "
                                           "'empleados' (reloj checador).")
                            if not _enc_inc:
                                st.warning(f"No se encontró el folio {row['FOLIO']} en Incidencias "
                                           "para marcarlo autorizado.")
                            elif _enc_pad or _enc_emp:
                                st.rerun()

    with tab2:
        sol_eco_hist = cargar_solicitudes_eco()
        # Normalizar ECO para unir con incidencias
        df_eco_hist = pd.DataFrame()
        if not sol_eco_hist.empty:
            df_eco_hist = sol_eco_hist.rename(columns={
                "Tipo Permiso":"TIPO","Fecha Inicio":"FECHA_INICIO","Fecha Fin":"FECHA_FIN",
                "Dias Solicitados":"DIAS","Aprobado Por":"AUTORIZADO_POR",
                "Nombre Completo":"NOMBRE","Fecha Registro":"FECHA_SOLICITUD"
            }).copy()
            col_f = next((c for c in df_eco_hist.columns if "FOLIO" in c.upper()), None)
            if col_f and col_f != "FOLIO": df_eco_hist["FOLIO"] = df_eco_hist[col_f]
            df_eco_hist["TIPO"] = "ECO"
            df_eco_hist["ESTADO"] = df_eco_hist["AUTORIZADO_POR"].apply(
                lambda x: "🔴 RECHAZADO" if str(x).strip().upper().startswith("RECHAZADO") else ("✅ AUTORIZADO" if str(x).strip() else "🟡 PENDIENTE")
            )
            df_eco_hist["ID"] = ""
            df_eco_hist["HORAS_PASE"] = ""

        df_hist_completo = pd.concat([incidencias, df_eco_hist[[c for c in incidencias.columns if c in df_eco_hist.columns]]], ignore_index=True) if not df_eco_hist.empty else incidencias.copy()

        filtro_tipo   = st.selectbox("Filtrar por tipo", ["TODOS"] + list(TIPO_LABELS.keys()))
        filtro_estado = st.selectbox("Filtrar por estado", ["TODOS", "PENDIENTE", "AUTORIZADO", "RECHAZADO"])
        df_hist = df_hist_completo.copy()
        if filtro_tipo   != "TODOS": df_hist = df_hist[df_hist["TIPO"] == filtro_tipo]
        if filtro_estado != "TODOS": df_hist = df_hist[df_hist["ESTADO"].astype(str).str.contains(filtro_estado, case=False)]
        st.dataframe(df_hist, use_container_width=True, hide_index=True)

        # ── Export Excel padrón mensual ──────────────
        import pytz as _pytz
        _tz = _pytz.timezone("America/Mexico_City")
        _ahora = datetime.now(_pytz.utc).astimezone(_tz)
        mes_sel = st.selectbox("📅 Mes para exportar padrón", 
            options=list(range(1,13)),
            format_func=lambda m: ["","Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"][m],
            index=_ahora.month-1)

        if st.button("📊 Generar padrón Excel del mes", type="primary"):
            import io as _io
            from openpyxl import Workbook as _WB
            from openpyxl.styles import Font as _Font, PatternFill as _Fill, Alignment as _Align
            sol_eco = cargar_solicitudes_eco()
            # Filtrar incidencias del mes
            df_mes = incidencias.copy()
            df_mes["_FDT"] = pd.to_datetime(df_mes["FECHA_INICIO"], errors="coerce")
            df_mes = df_mes[(df_mes["_FDT"].dt.year == _ahora.year) & (df_mes["_FDT"].dt.month == mes_sel)]
            # Filtrar ECO del mes
            df_eco = sol_eco.copy() if not sol_eco.empty else pd.DataFrame()
            if not df_eco.empty:
                col_fi = next((c for c in df_eco.columns if "Inicio" in c or "INICIO" in c), None)
                if col_fi:
                    df_eco["_FDT"] = pd.to_datetime(df_eco[col_fi], errors="coerce")
                    df_eco = df_eco[(df_eco["_FDT"].dt.year == _ahora.year) & (df_eco["_FDT"].dt.month == mes_sel)]
                    df_eco = df_eco.rename(columns={
                        "Tipo Permiso":"TIPO","Fecha Inicio":"FECHA_INICIO","Fecha Fin":"FECHA_FIN",
                        "Dias Solicitados":"DIAS","Aprobado Por":"AUTORIZADO_POR",
                        "Nombre Completo":"NOMBRE","Fecha Registro":"FECHA_SOLICITUD"
                    })
                    col_f = next((c for c in df_eco.columns if "FOLIO" in c.upper()), None)
                    if col_f: df_eco["FOLIO"] = df_eco[col_f]
            # Unir
            cols_exp = ["FOLIO","NOMBRE","RFC","TIPO","FECHA_INICIO","FECHA_FIN","DIAS","HORAS_PASE","MOTIVO","ESTADO","AUTORIZADO_POR","FECHA_AUTORIZACION"]
            frames_exp = []
            if not df_mes.empty: frames_exp.append(df_mes[[c for c in cols_exp if c in df_mes.columns]])
            if not df_eco.empty: frames_exp.append(df_eco[[c for c in cols_exp if c in df_eco.columns]])
            if frames_exp:
                df_export = pd.concat(frames_exp, ignore_index=True)
                df_export["TIPO"] = df_export["TIPO"].apply(lambda x: "ECO" if str(x).lower() in ["economico","económico"] else x)
                df_export["TIPO"] = df_export["TIPO"].map(TIPO_LABELS).fillna(df_export["TIPO"])
                df_export = df_export.sort_values("NOMBRE")
                wb = _WB()
                ws = wb.active
                ws.title = "Padrón Solicitudes"
                mes_nombre = ["","Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"][mes_sel]
                ws.merge_cells("A1:L1")
                ws["A1"] = f"PADRÓN DE SOLICITUDES — {mes_nombre.upper()} {_ahora.year} · DFC · SEJ"
                ws["A1"].font = _Font(bold=True, size=12, color="002F6C")
                ws["A1"].alignment = _Align(horizontal="center")
                hdrs = ["Folio","Nombre","RFC","Tipo","Fecha Inicio","Fecha Fin","Días","Horas Pase","Motivo","Estado","Autorizado Por","Fecha Autorización"]
                ws.append(hdrs)
                for cell in ws[2]: cell.font = _Font(bold=True, color="FFFFFF"); cell.fill = _Fill("solid", fgColor="002F6C")
                for _, row in df_export.iterrows():
                    ws.append([str(row.get(c,"")) for c in ["FOLIO","NOMBRE","RFC","TIPO","FECHA_INICIO","FECHA_FIN","DIAS","HORAS_PASE","MOTIVO","ESTADO","AUTORIZADO_POR","FECHA_AUTORIZACION"]])
                ws.column_dimensions["A"].width = 18
                ws.column_dimensions["B"].width = 35
                ws.column_dimensions["I"].width = 40
                buf_xl = _io.BytesIO()
                wb.save(buf_xl)
                st.download_button(f"⬇️ Descargar padrón {mes_nombre} {_ahora.year}",
                    data=buf_xl.getvalue(),
                    file_name=f"Padron_Solicitudes_{mes_nombre}_{_ahora.year}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary")
            else:
                st.info(f"No hay solicitudes registradas para ese mes.")

        st.divider()
        st.markdown("#### 🔍 Regenerar PDF por folio")
        folio_buscar = st.text_input("Ingresa el folio", placeholder="ECO-2026-0001 / PSE-2026-0001", key="folio_regen")
        if folio_buscar:
            folio_buscar = folio_buscar.strip().upper()
            # Buscar en Incidencias
            match_inc = incidencias[incidencias["FOLIO"].astype(str).str.upper() == folio_buscar]
            # Buscar en Solicitudes ECO
            sol_eco = cargar_solicitudes_eco()
            col_f = next((c for c in sol_eco.columns if "FOLIO" in c.upper()), None)
            match_eco = sol_eco[sol_eco[col_f].astype(str).str.upper() == folio_buscar] if col_f else pd.DataFrame()
            if not match_inc.empty:
                row = match_inc.iloc[0]
                # Detectar subtipo real desde el motivo para registros viejos PSE
                motivo_str = str(row["MOTIVO"])
                tipo_real  = row["TIPO"]
                if tipo_real == "PSE":
                    if "Pase de entrada" in motivo_str:
                        tipo_real = "PEN"
                    elif "sin retorno" in motivo_str.lower() or "Salida" in motivo_str:
                        tipo_real = "PSE"
                    elif "Retorno" in motivo_str:
                        tipo_real = "PSR"
                # Buscar jefe inmediato del empleado en tab Usuarios
                jefe_pdf_regen = ""
                try:
                    usuarios_df = cargar_usuarios()
                    match_usr = usuarios_df[usuarios_df["RFC"].astype(str).str.upper().str.strip() == str(row["RFC"]).upper().strip()]
                    if not match_usr.empty:
                        jefe_col = match_usr.iloc[0].get("JEFE_INMEDIATO", "")
                        jefe_pdf_regen = formato_jefe_pdf(str(jefe_col).strip())
                except: pass
                datos_pdf = {
                    "folio":          row["FOLIO"],
                    "rfc":            row["RFC"],
                    "nombre":         row["NOMBRE"],
                    "tipo":           tipo_real,
                    "tipo_label":     TIPO_LABELS.get(tipo_real, tipo_real),
                    "subtipo_label":  TIPO_LABELS.get(tipo_real, tipo_real),
                    "fecha_solicitud":str(row["FECHA_SOLICITUD"]),
                    "fecha_inicio":   str(row["FECHA_INICIO"]),
                    "fecha_fin":      str(row["FECHA_FIN"]),
                    "dias":           row["DIAS"],
                    "horas_pase":     row.get("HORAS_PASE", 0),
                    "motivo":         motivo_str,
                    "tiene_anexo":    row.get("TIENE_ANEXO", "NO") == "SÍ",
                    "link_anexo":     row.get("LINK_ANEXO", ""),
                    "estado":         str(row["ESTADO"]),
                    "jefe_inmediato": jefe_pdf_regen,
                }
                st.success(f"Folio encontrado: {folio_buscar}")
                pdf_bytes = generar_comprobante_pdf(datos_pdf)
                st.download_button("⬇️ Descargar PDF", data=pdf_bytes,
                    file_name=f"Comprobante_{folio_buscar}.pdf", mime="application/pdf", type="primary")
            elif not match_eco.empty:
                row = match_eco.iloc[0]
                datos_pdf = {
                    "folio":          str(row.get(col_f, folio_buscar)),
                    "rfc":            str(row.get("RFC", "")),
                    "nombre":         str(row.get("Nombre Completo", "")),
                    "tipo":           "ECO",
                    "tipo_label":     "Día económico",
                    "subtipo_label":  "Día económico",
                    "fecha_solicitud":str(row.get("Fecha Registro", "")),
                    "fecha_inicio":   str(row.get("Fecha Inicio", "")),
                    "fecha_fin":      str(row.get("Fecha Fin", "")),
                    "dias":           row.get("Dias Solicitados", 0),
                    "horas_pase":     0,
                    "motivo":         str(row.get("Motivo", "")),
                    "tiene_anexo":    False,
                    "link_anexo":     str(row.get("LINK_ANEXO", "")),
                    "estado":         "AUTORIZADO" if str(row.get("Aprobado Por","")).strip() else "PENDIENTE",
                    "jefe_inmediato": str(row.get("Aprobado Por", "")),
                }
                st.success(f"Folio encontrado: {folio_buscar}")
                pdf_bytes = generar_comprobante_pdf(datos_pdf)
                st.download_button("⬇️ Descargar PDF", data=pdf_bytes,
                    file_name=f"Comprobante_{folio_buscar}.pdf", mime="application/pdf", type="primary")
            else:
                st.error(f"Folio {folio_buscar} no encontrado.")


    with tab4:
        render_checador()

    with tab5:
        if render_pendientes_nomina is None:
            st.error(f"El módulo de nómina no está disponible: {_ERROR_NOMINA}")
        else:
            render_pendientes_nomina(cargar_directorio_nomina, get_client,
                                     quincena_actual=quincena_pago_actual)

    with tab3:
        # ── Alerta días económicos por agotarse ──────
        st.markdown("#### ⚠️ Empleados con 5 días económicos o menos")
        try:
            import pytz as _ptz
            _tz   = _ptz.timezone("America/Mexico_City")
            _anio = datetime.now(_ptz.utc).astimezone(_tz).year
            sol_eco = cargar_solicitudes_eco()
            if sol_eco.empty:
                st.info("No hay solicitudes registradas.")
            else:
                col_rfc  = next((c for c in sol_eco.columns if c.upper() == "RFC"), None)
                col_nom  = next((c for c in sol_eco.columns if "NOMBRE" in c.upper()), None)
                col_dias = next((c for c in sol_eco.columns if "DIAS" in c.upper() or "SOLICIT" in c.upper()), None)
                col_apr  = next((c for c in sol_eco.columns if "APROBADO" in c.upper()), None)
                col_fi   = next((c for c in sol_eco.columns if "INICIO" in c.upper()), None)
                if col_rfc and col_dias and col_fi:
                    sol_eco["_FDT"] = pd.to_datetime(sol_eco[col_fi], errors="coerce")
                    sol_anio = sol_eco[sol_eco["_FDT"].dt.year == _anio].copy()
                    # Excluir rechazados
                    if col_apr:
                        sol_anio = sol_anio[~sol_anio[col_apr].astype(str).str.strip().str.upper().str.startswith("RECHAZADO")]
                    sol_anio["_DIAS"] = pd.to_numeric(sol_anio[col_dias], errors="coerce").fillna(0).astype(int)
                    resumen = sol_anio.groupby(col_rfc).agg(
                        Usados=("_DIAS", "sum")
                    ).reset_index()
                    resumen.rename(columns={col_rfc: "RFC"}, inplace=True)
                    if col_nom:
                        nombres = sol_anio.groupby(col_rfc)[col_nom].first().reset_index()
                        nombres.rename(columns={col_rfc:"RFC", col_nom:"Nombre"}, inplace=True)
                        resumen = resumen.merge(nombres, on="RFC", how="left")
                    resumen["Disponibles"] = (9 - resumen["Usados"]).clip(lower=0)
                    alertas = resumen[resumen["Disponibles"] <= 5].sort_values("Disponibles")
                    cols_show = ["Nombre","Usados","Disponibles"] if col_nom else ["RFC","Usados","Disponibles"]
                    if alertas.empty:
                        st.success("Ningún empleado tiene 5 días o menos disponibles.")
                    else:
                        st.dataframe(alertas[[c for c in cols_show if c in alertas.columns]], use_container_width=True, hide_index=True)
                    st.warning("No se encontraron las columnas necesarias en Solicitudes.")
        except Exception as e:
            st.warning(f"No se pudo calcular: {e}")
        st.divider()
        st.markdown("Reporte consolidado para la directora.")
        anio_rep = st.number_input("Año",  value=datetime.now().year, step=1)
        mes_rep  = st.selectbox("Mes", range(1, 13), index=datetime.now().month - 1,
                                format_func=lambda m: ["Enero","Febrero","Marzo","Abril","Mayo","Junio",
                                                        "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"][m-1])
        if st.button("Generar reporte", type="primary"):
            df_rep = incidencias.copy()
            df_rep["FECHA_DT"] = pd.to_datetime(df_rep["FECHA_SOLICITUD"], errors="coerce")
            df_rep = df_rep[(df_rep["FECHA_DT"].dt.year == anio_rep) & (df_rep["FECHA_DT"].dt.month == mes_rep)]
            if df_rep.empty:
                st.info("No hay incidencias en ese período.")
            else:
                mes_nombre = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto",
                              "Septiembre","Octubre","Noviembre","Diciembre"][mes_rep-1]
                st.markdown(f"**Resumen {mes_nombre} {anio_rep}**")
                st.dataframe(df_rep.groupby(["TIPO","ESTADO"]).size().reset_index(name="Total"),
                             use_container_width=True, hide_index=True)
                detalle = df_rep.groupby(["NOMBRE","TIPO"]).agg(
                    Total=("FOLIO","count"),
                    Dias=("DIAS","sum"),
                    Horas_pase=("HORAS_PASE","sum"),
                    Autorizadas=("ESTADO", lambda x: _estado_norm(x).str.startswith("AUTORIZADO").sum()),
                    Rechazadas=("ESTADO",  lambda x: _estado_norm(x).str.startswith("RECHAZADO").sum()),
                ).reset_index()
                st.dataframe(detalle, use_container_width=True, hide_index=True)
                st.download_button("⬇️ Descargar CSV",
                                   data=df_rep.to_csv(index=False).encode("utf-8"),
                                   file_name=f"Reporte_{mes_nombre}_{anio_rep}.csv",
                                   mime="text/csv")

# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
PAGOS = [
    # OJO: esta lista NO guarda el año — quincena_pago_actual arma la fecha
    # con date(año_actual, mes, día). O sea que es el calendario de UN año, y
    # al entrar 2027 estas mismas filas se leen como fechas de 2027.
    # Las fechas ya son las de 2027 (Q-01 el 14 y Q-02 el 28, que en 2026 fue
    # el 29). Los CONCEPTOS siguen siendo los de 2026: cuando salga el
    # calendario oficial de 2027 hay que cambiarlos, no solo aquí sino en todo
    # PAGOS. En ari_module.py no hace falta: allá la fecha lleva el año.
    ("Enero", [
        ("14 Ene", "Q-01", [("Estímulo puntualidad y asistencia 2ª parte", "🟣🔵"), ("Prima Dominical", "🟣🔵")]),
        ("28 Ene", "Q-02", [("Compensación Nacional Única 1ª Parte", "🟡🟢")]),
    ]),
    ("Febrero", [
        ("12 Feb", "Q-03", [("Sueldo ordinario", "")]),
        ("26 Feb", "Q-04", [("Sueldo ordinario", "")]),
    ]),
    ("Marzo", [
        ("12 Mar", "Q-05", [("Sueldo ordinario", "")]),
        ("26 Mar", "Q-06", [("1ª Parte Aguinaldo", "🟢🔵")]),
    ]),
    ("Abril", [
        ("14 Abr", "Q-07", [("Sueldo ordinario", "")]),
        ("29 Abr", "Q-08", [("Sueldo ordinario", "")]),
    ]),
    ("Mayo", [
        ("14 May", "Q-09", [
            ("1ª Parte Aguinaldo", "🟡"),
            ("Gratificación Día del Maestro", "🟡"),
            ("1ª Parte Aguinaldo", "🟣"),
            ("Reconocimiento Docentes Nivel Superior", "🟣"),
            ("Ayuda para Libros", "🟣"),
        ]),
        ("28 May", "Q-10", [("Sueldo ordinario", "")]),
    ]),
    ("Junio", [
        ("12 Jun", "Q-11", [("Estímulo puntualidad y asistencia 1ª parte", "🟣🔵")]),
        ("29 Jun", "Q-12", [("Reconocimiento a Directores", "🟡")]),
    ]),
    ("Julio", [
        ("14 Jul", "Q-13", [("Sueldo ordinario", "")]),
        ("30 Jul", "Q-14", [("Gratificación por el trabajo", "🟢")]),
    ]),
    ("Agosto", [
        ("13 Ago", "Q-15", [("Organización Escolar", "🟡"), ("Ayuda para gastos escolares", "🟢")]),
        ("28 Ago", "Q-16", [("Compensación Nacional Única 2ª Parte", "🟡🟢"), ("Medida Económica Única", "🟣🔵")]),
    ]),
    ("Septiembre", [
        ("14 Sep", "Q-17", [
            ("Estímulo a la Actividad Docente", "🟡"),
            ("Estímulo a Directores", "🟡"),
            ("Gratificación Única 1ª Parte", "🔵"),
        ]),
        ("29 Sep", "Q-18", [
            ("Gratificación Fortalecimiento Académico", "🟡"),
            ("Bono Extraordinario superación académica 1ª Parte", "🟣"),
        ]),
    ]),
    ("Octubre", [
        ("14 Oct", "Q-19", [("Sueldo ordinario", "")]),
        ("29 Oct", "Q-20", [("Fortalecimiento CC según ajustes salariales", "🟡"), ("Fortalecimiento CT según ajustes salariales", "🟢")]),
    ]),
    ("Noviembre", [
        ("12 Nov", "Q-21", [("Sueldo ordinario", "")]),
        ("27 Nov", "Q-22", [("Bono anual 24 días inicial", "🟡"), ("Apoyo a la integración educativa especial", "🟡")]),
    ]),
    ("Diciembre", [
        # Aguinaldo y prima vacacional (32) los cobran los CUATRO subsistemas.
        ("14 Dic", "Q-23", [
            ("Aguinaldo", "🟡🟣🟢🔵"),
            ("32 Prima vacacional", "🟡🟣🟢🔵"),
        ]),
        # Los G y la CF los cobran docentes Y apoyo de nivel superior.
        # BE es solo de docentes; GU y BN solo de apoyo.
        # El G9 se marca en esta quincena pero el depósito cae hasta enero:
        # va dicho en el nombre para que nadie lo reclame en diciembre.
        ("17 Dic", "Q-24", [
            ("G4 Prima dominical", "🟣🔵"),
            ("G5 Ajuste de calendario", "🟣🔵"),
            ("G6 Días económicos no disfrutados", "🟣🔵"),
            ("G7 Días de descanso obligatorio", "🟣🔵"),
            ("G9 Estímulo puntualidad y asistencia (se paga hasta enero)", "🟣🔵"),
            ("CF Compensación de fin de año", "🟣🔵"),
            ("BE Bono extraordinario de superación académica", "🟣"),
            ("GU Gratificación única", "🔵"),
            ("BN Bono anual", "🔵"),
        ]),
    ]),
]

def quincena_pago_actual(anio: int = None):
    """(numero_quincena, fecha_de_pago) de la quincena vigente, leída del
    MISMO calendario PAGOS que muestra la app. Vigente = último pago que ya
    ocurrió (o es hoy). None si aún no hay pagos en el año."""
    import pytz
    _MES_ABREV = {"Ene":1,"Feb":2,"Mar":3,"Abr":4,"May":5,"Jun":6,
                  "Jul":7,"Ago":8,"Sep":9,"Oct":10,"Nov":11,"Dic":12}
    hoy = datetime.now(pytz.timezone("America/Mexico_City")).date()
    anio = anio or hoy.year
    todas = []
    for _mes, quincenas in PAGOS:
        for fecha_str, qna, _ in quincenas:
            try:
                dia, mes_abr = fecha_str.split()
                f = date(anio, _MES_ABREV[mes_abr], int(dia))
            except Exception:
                continue
            m = re.search(r"(\d+)", str(qna))
            if m:
                todas.append((f, int(m.group(1))))
    todas.sort()
    actual = None
    for f, num in todas:
        if f <= hoy:
            actual = (num, f)
        else:
            break
    return actual


def vista_calendario():
    st.markdown("## 📅 Calendario de Pagos y Prestaciones 2026")
    st.markdown("""
    <div style='display:flex;gap:16px;flex-wrap:wrap;margin-bottom:12px;font-size:12px'>
        <span>🟡 Docentes Básica</span>
        <span>🟣 Docentes Nivel Superior</span>
        <span>🟢 Personal de Apoyo Básica</span>
        <span>🔵 Personal de Apoyo No Docente Nivel Superior</span>
    </div>""", unsafe_allow_html=True)


    import pytz
    from datetime import datetime, date
    tz_mx = pytz.timezone("America/Mexico_City")
    hoy = datetime.now(tz_mx).date()
    MESES_NUM = {"Enero":1,"Febrero":2,"Marzo":3,"Abril":4,"Mayo":5,"Junio":6,
                 "Julio":7,"Agosto":8,"Septiembre":9,"Octubre":10,"Noviembre":11,"Diciembre":12}
    MES_ABREV = {"Ene":1,"Feb":2,"Mar":3,"Abr":4,"May":5,"Jun":6,
                 "Jul":7,"Ago":8,"Sep":9,"Oct":10,"Nov":11,"Dic":12}

    def _fecha_real(fecha_str, anio=2026):
        try:
            dia_str, mes_abr = fecha_str.split()
            return date(anio, MES_ABREV[mes_abr], int(dia_str))
        except Exception:
            return None

    # Quincena ACTUAL de verdad: la última fecha de pago que ya pasó (o es
    # hoy). Antes solo se comparaba el MES, así que toda la tarjeta de
    # septiembre se resaltaba sin decir si iba la del 14 o la del 29.
    _todas = []
    for _mes, _quincenas in PAGOS:
        for _fecha, _qna, _ in _quincenas:
            _f = _fecha_real(_fecha)
            if _f:
                _todas.append((_f, _mes, _qna))
    _todas.sort()
    quincena_actual = None
    for _f, _mes, _qna in _todas:
        if _f <= hoy:
            quincena_actual = (_mes, _qna)
        else:
            break

    def render_mes(col, mes, quincenas, es_actual):
        borde = "#F97316" if es_actual else "var(--color-border-tertiary)"
        fondo = "#FFF7ED" if es_actual else "var(--color-background-primary)"
        # Construido en UN solo bloque HTML: antes el <div> se abría en un
        # st.markdown y el contenido en otros, y Streamlit los renderizaba
        # como piezas separadas — ese era el recuadro vacío flotando aparte.
        partes = [f"<div style='border:2px solid {borde};border-radius:12px;"
                 f"padding:12px;margin-bottom:12px;background:{fondo}'>"]
        partes.append(f"<p style='margin:0 0 8px 0'><b>{'🟠 ' if es_actual else ''}{mes.upper()}</b></p>")
        for fecha, qna, conceptos in quincenas:
            es_esta_quincena = quincena_actual == (mes, qna)
            estilo_fila = ("background:#FFEDD5;border-radius:8px;padding:6px 8px;"
                          "margin-bottom:6px;border-left:3px solid #F97316;") if es_esta_quincena else "margin-bottom:6px;"
            partes.append(f"<div style='{estilo_fila}'>")
            etiqueta = " · 📍 Estamos aquí" if es_esta_quincena else ""
            partes.append(f"<p style='margin:0'>📅 <b>{fecha}</b>"
                         f"{'  ·  ' + qna if qna else ''}{etiqueta}</p>")
            for concepto, cats in conceptos:
                texto = f"{cats}  {concepto}" if cats else concepto
                partes.append(f"<p style='margin:0;font-size:0.8em;color:gray'>{texto}</p>")
            partes.append("</div>")
        partes.append("</div>")
        col.markdown("".join(partes), unsafe_allow_html=True)

    for fila in range(0, len(PAGOS), 3):
        grupo = PAGOS[fila:fila+3]
        cols = st.columns(3)
        for j, (mes, quincenas) in enumerate(grupo):
            es_actual = MESES_NUM.get(mes, 0) == hoy.month
            render_mes(cols[j], mes, quincenas, es_actual)

    st.divider()
    try:
        with open("1780634863696_calendario_pagos_est_2026.ics", "rb") as f_ics:
            ics_bytes = f_ics.read()
        st.download_button(
            "⬇️ Descargar calendario (.ics)",
            data=ics_bytes,
            file_name="calendario_pagos_2026.ics",
            mime="text/calendar",
            use_container_width=True,
            type="primary"
        )
    except:
        st.info("Archivo .ics no disponible.")

    st.markdown("### 📲 ¿Cómo agregar este calendario a tu dispositivo?")
    with st.expander("📱 iPhone / iPad"):
        st.markdown("""
1. Descarga el archivo **.ics** con el botón de arriba
2. Abre la app **Archivos** en tu iPhone
3. Toca el archivo descargado
4. Aparecerá una pantalla para **Agregar todos** los eventos — confirma
5. Los pagos aparecerán en tu app **Calendario**
        """)
    with st.expander("🤖 Android / Google Calendar"):
        st.markdown("""
1. Descarga el archivo **.ics**
2. Ve a [calendar.google.com](https://calendar.google.com) desde tu navegador
3. En la barra lateral toca **Otros calendarios** → **+** → **Importar**
4. Selecciona el archivo **.ics** descargado y confirma
        """)
    with st.expander("💻 Outlook (PC o Mac)"):
        st.markdown("""
1. Descarga el archivo **.ics**
2. Abre **Outlook** → **Archivo** → **Abrir y exportar** → **Importar o exportar**
3. Selecciona **Importar un archivo iCalendar (.ics)**
4. Busca el archivo descargado y confirma
        """)


def vista_nomina():
    st.markdown("## 💰 Mis comprobantes de nómina")
    st.caption("Selecciona el portal según tu tipo de plaza.")

    col1, col2 = st.columns(2)

    with col1:
        with st.container(border=True):
            st.markdown("### 🏛️ Personal Estatal")
            st.markdown("""
**Paso 1** — Haz clic en el botón de abajo  
**Paso 2** — Se abrirá el portal del Gobierno de Jalisco en nueva pestaña  
**Paso 3** — Inicia sesión con tu **RFC** y contraseña institucional  
**Paso 4** — Selecciona el período a consultar
            """)
            st.info("⚠️ Si no carga, cópialo en tu navegador: **miscomprobantesnomina.jalisco.gob.mx/login**")
            st.link_button("🔗 Portal Estatal", "https://miscomprobantesnomina.jalisco.gob.mx/login", type="primary", use_container_width=True)

    with col2:
        with st.container(border=True):
            st.markdown("### Personal Federalizado")
            st.markdown("""
**Paso 1** — Haz clic en el botón de abajo  
**Paso 2** — Se abrirá el portal de la SEP Federal en nueva pestaña  
**Paso 3** — Inicia sesión con tus credenciales federales  
**Paso 4** — Selecciona el período a consultar
            """)
            st.info("⚠️ Si no carga, accede desde **miportal.fone.sep.gob.mx**")
            st.link_button("🔗 Portal Federal SEP", "https://miportal.fone.sep.gob.mx", type="primary", use_container_width=True)


def vista_directorio():
    import pytz

    @st.cache_data(ttl=300)
    def cargar_directorio():
        # FASE 1 migración: lee del Padrón Maestro con fallback automático
        # a la tab Directorio (mismas columnas que la vista espera).
        return _directorio_unificado()

    st.markdown("## 📞 Directorio interno DFC 2026")
    st.caption("Toca un área para ver su equipo · Busca por nombre, extensión o correo")

    df = cargar_directorio()
    aviso_fuente_personal()
    if df.empty:
        st.warning("No se pudo cargar el directorio.")
        return

    # ── Separar Centros de Maestros del personal DFC ──
    # Las filas cuya AREA contiene "Centro de Maestros" van a su propia
    # sección (solo visible para RFCs autorizados en secrets); la vista
    # DFC de siempre queda intacta y sin mezclarse.
    _mask_cm = df["AREA"].astype(str).str.contains("Centro de Maestros", case=False, na=False)
    df_cm = df[_mask_cm].copy()
    df = df[~_mask_cm].copy()

    # Contactos externos (Subsecretaría, etc.): solo admin y los RFC del secret
    # rfcs_horarios_todos. Se agregan aquí y no en _directorio_unificado
    # para que no lleguen al buscador de contacto de emergencia.
    _rfc_ext = str(st.session_state.get("rfc", "")).upper().strip()
    if (st.session_state.get("rol") == "admin"
            or _rfc_ext in _rfcs_secret("rfcs_horarios_todos")):
        _ext, _motivo_ext = cargar_directorio_externo()
        if not _ext.empty:
            df = pd.concat([df, _ext], ignore_index=True).fillna("")
        elif st.session_state.get("rol") == "admin":
            # Solo el admin ve el diagnóstico: antes esto fallaba en silencio.
            st.caption(f"ℹ️ Directorio externo sin mostrar: {_motivo_ext}.")
    if st.session_state.get("rol") != "admin":
        mask_ocultar = (
            df["AREA"].astype(str).str.contains("Comisionado a otra|Oficinas centrales", case=False, na=False) | 
            df["DEPARTAMENTO"].astype(str).str.contains("Comisionado a otra|Oficinas centrales", case=False, na=False)
        )
        df = df[~mask_ocultar].copy()

    busq = st.text_input("", placeholder="🔍 Busca por nombre, extensión o correo...", label_visibility="collapsed")

    COLORES = {
        "Dirección General":        ("#E1F5EE", "#0F6E56"),
        "Dir. Gestión y Evaluación": ("#EEEDFE", "#534AB7"),
        "Dir. Desarrollo Académico": ("#E6F1FB", "#185FA5"),
    }

    def color(area):
        for k, v in COLORES.items():
            if k in area:
                return v
        return ("#F4F6F9", "#444")

    _rfc_dir = str(st.session_state.get("rfc", "")).upper().strip()
    _ve_horarios_todos = (st.session_state.get("rol") == "admin"
                          or _rfc_dir in _rfcs_secret("rfcs_horarios_todos"))

    def tarjeta(row):
        bg, tc = color(row["AREA"])
        ini = (str(row["NOMBRE"])[0] + (str(row["NOMBRE"]).split()[1][0] if len(str(row["NOMBRE"]).split()) > 1 else "")).upper()
        ext_html = f"📞 **{row['EXTENSION']}**" if row["EXTENSION"] else ""
        email_html = row["CORREO"] if row["CORREO"] else ""
        # Departamento en su propio renglón, como etiqueta de color y en
        # mayúsculas (antes iba en gris chico y con asteriscos literales:
        # la cursiva de Markdown no funciona dentro de HTML).
        _dep = str(row.get("DEPARTAMENTO", "") or "").strip()
        dept_html = (f"<div style='margin-top:3px'><span style='background:{bg};color:{tc};"
                     f"font-size:11px;font-weight:700;letter-spacing:.3px;padding:2px 8px;"
                     f"border-radius:10px'>{_dep.upper()}</span></div>") if _dep else ""
        # Celular: SOLO contactos externos (oficinas fuera de la DFC). Si el
        # padrón tuviera celulares personales del personal, no se exponen aquí.
        _cel = str(row.get("CELULAR", "") or "").strip() if str(row.get("_EXTERNO", "")) == "SI" else ""

        col1, col2 = st.columns([3,1])
        with col1:
            st.markdown(
                f"<div style='display:flex;align-items:center;gap:10px;padding:4px 0'>"
                f"<div style='width:34px;height:34px;border-radius:50%;background:{bg};color:{tc};"
                f"display:flex;align-items:center;justify-content:center;font-weight:600;font-size:12px;flex-shrink:0'>{ini}</div>"
                f"<div><div style='font-size:14px;font-weight:500'>{row['NOMBRE']}</div>"
                f"<div style='font-size:12px;color:gray'>{row['AREA']}</div>{dept_html}</div></div>",
                unsafe_allow_html=True
            )
        with col2:
            if ext_html: st.markdown(ext_html)
            if _cel: st.markdown(f"📱 **{_cel}**")
            if email_html: st.caption(row["CORREO"])
        # Horario: SOLO admin lo ve aquí (las 10 áreas, no Centros de
        # Maestros). Las 5 secretarias de rfcs_directorio_cm siguen viendo
        # horario únicamente en la sección de Centros de Maestros.
        # Horarios de las 10 áreas: admin + RFCs del secret rfcs_horarios_todos
        # (Pichardo y Nadia). En secrets, nunca en el código.
        if _ve_horarios_todos:
            _horario_partes = []
            for _dia, (_col_e, _col_s) in COLUMNAS_HORARIO.items():
                _ent = str(row.get(_col_e, "")).strip()
                _sal = str(row.get(_col_s, "")).strip()
                if _ent and _sal and _ent.lower() != "nan" and _sal.lower() != "nan":
                    _ent = ":".join(_ent.split(":")[:2])
                    _sal = ":".join(_sal.split(":")[:2])
                    _horario_partes.append(f"{_dia} {_ent}-{_sal}")
            if _horario_partes:
                st.caption(f"🕒 {' | '.join(_horario_partes)}")
        st.divider()

    if busq:
        q = busq.lower().strip()
        resultados = df[
            df["NOMBRE"].str.lower().str.contains(q) |
            df["EXTENSION"].astype(str).str.contains(q) |
            df["CORREO"].str.lower().str.contains(q)
        ]
        if resultados.empty:
            st.info("Sin resultados.")
        else:
            st.caption(f"{len(resultados)} resultado(s)")
            for _, row in resultados.iterrows():
                tarjeta(row)
    else:
        # Agrupar por DEPARTAMENTO como área visible (7 secciones)
        ORDEN_AREAS = [
            "Despacho",
            "Recursos Humanos",
            "Viáticos",
            "Recursos Materiales",
            "Licitaciones y Presupuestos",
            "Comunicación Social",
            "Dir. Desarrollo Académico",
            "Base de Datos y Diseño Instruccional",
            #"Comisionado a otra área",
        ]
        ICONOS = {
            "Despacho":                    "🏢",
            "Recursos Humanos":            "👥",
            "Viáticos":                    "🧾",
            "Recursos Materiales":         "📦",
            "Licitaciones y Presupuestos": "📋",
            "Comunicación Social":         "📢",
            "Dir. Desarrollo Académico":   "👩‍🏫",
            "Base de Datos y Diseño Instruccional": "💾",
            "Comisionado a otra área":     "✈️",
            "Asesor":                      "🎓",
            "Responsable":                 "⭐",
            "Administrativo":              "🗂️",
        }
        # El padrón guarda todo en MAYÚSCULAS sin acentos, así que la búsqueda
        # de icono se hace sobre una clave normalizada. Con la comparación
        # literal, "VIATICOS" no encontraba "Viáticos" y todo caía al icono
        # genérico.
        def _clave_icono(t):
            import unicodedata as _ud
            t = _ud.normalize("NFKD", str(t or ""))
            return "".join(c for c in t if not _ud.combining(c)).upper().strip()

        ICONOS_NORM = {_clave_icono(k): v for k, v in ICONOS.items()}
        ORDEN_NORM = [_clave_icono(a) for a in ORDEN_AREAS]

        # Normalizar: si DEPARTAMENTO está vacío usar AREA como dept
        # Externos se agrupan por su oficina (AREA); su DEPARTAMENTO es el cargo.
        df["DEPT_VISTA"] = df.apply(
            lambda r: r["AREA"] if str(r.get("_EXTERNO", "")) == "SI"
            else (r["DEPARTAMENTO"] if r["DEPARTAMENTO"] else r["AREA"]), axis=1)

        depts_en_datos = df["DEPT_VISTA"].unique().tolist()
        _pos = {d: (ORDEN_NORM.index(_clave_icono(d))
                    if _clave_icono(d) in ORDEN_NORM else len(ORDEN_NORM))
                for d in depts_en_datos}
        orden_final = sorted(depts_en_datos, key=lambda d: (_pos[d], d))

        for dept in orden_final:
            personas = df[df["DEPT_VISTA"] == dept]
            _es_ext = (personas.get("_EXTERNO", pd.Series(dtype=str)).astype(str) == "SI").any()
            icono = ICONOS_NORM.get(_clave_icono(dept), "🏛️" if _es_ext else "🏢")
            bg, tc = color(df[df["DEPT_VISTA"]==dept]["AREA"].iloc[0])
            with st.expander(f"{icono} {dept}  ·  {len(personas)} personas"):
                for _, row in personas.iterrows():
                    tarjeta(row)

    # ── 🏫 Directorio de Centros de Maestros ──────────────────────────
    # Solo visible para el admin y los RFCs listados en secrets:
    #   rfcs_directorio_cm = ["RFC1", "RFC2", ...]
    # (en secrets, NO en el código: los RFC son dato personal)
    try:
        _rfcs_cm = [str(r).upper().strip() for r in st.secrets.get("rfcs_directorio_cm", [])]
    except Exception:
        _rfcs_cm = []
    _rfc_actual = str(st.session_state.get("rfc", "")).upper().strip()
    _permiso_cm = (st.session_state.get("rol") == "admin") or (_rfc_actual and _rfc_actual in _rfcs_cm) or _ve_horarios_todos

    if _permiso_cm:
        st.markdown("---")
        st.markdown("## 🏫 Centros de Maestros")
        if df_cm.empty:
            st.info("Aún no hay personal de Centros de Maestros en la tab Directorio. "
                    "Agrega filas con AREA = 'Centro de Maestros <nombre>' y "
                    "DEPARTAMENTO = 'Responsable' o 'Asesor'.")
        else:
            busq_cm = st.text_input("", placeholder="🔍 Busca por nombre o centro...",
                                    label_visibility="collapsed", key="busq_cm")
            _df_v = df_cm
            if busq_cm:
                _q = busq_cm.lower().strip()
                _df_v = df_cm[df_cm["NOMBRE"].str.lower().str.contains(_q, na=False) |
                              df_cm["AREA"].str.lower().str.contains(_q, na=False)]
            if _df_v.empty:
                st.info("Sin resultados.")
            else:
                for _centro in sorted(_df_v["AREA"].unique()):
                    _gente = _df_v[_df_v["AREA"] == _centro]
                    # Responsable primero, con corona
                    _es_resp = _gente["DEPARTAMENTO"].astype(str).str.upper().str.contains("RESPONSABLE", na=False)
                    _gente = pd.concat([_gente[_es_resp], _gente[~_es_resp]])
                    _etq = str(_centro).replace("Centro de Maestros", "").strip()
                    with st.expander(f"🏫 {_etq}  ·  {len(_gente)} persona(s)"):
                        for _, _r in _gente.iterrows():
                            _rol = str(_r.get("DEPARTAMENTO", "")).strip()
                            _icon = "🧑‍💼" if "RESPONSABLE" in _rol.upper() else "•"
                            _extra = f" · 📞 {_r['EXTENSION']}" if str(_r.get("EXTENSION", "")).strip() else ""
                            _correo = f" · {_r['CORREO']}" if str(_r.get("CORREO", "")).strip() else ""
                            st.markdown(f"{_icon} **{_r['NOMBRE']}** — {_rol or 'Asesor'}{_extra}{_correo}")
                            # Horario semanal: se había perdido en una edición
                            # anterior. split(":")[:2] quita los segundos sin
                            # importar el valor de los minutos (":00:00"→":00"
                            # solo servía cuando los minutos ya eran 00).
                            _horario_partes = []
                            for _dia, (_col_e, _col_s) in COLUMNAS_HORARIO.items():
                                _ent = str(_r.get(_col_e, "")).strip()
                                _sal = str(_r.get(_col_s, "")).strip()
                                if _ent and _sal and _ent.lower() != "nan" and _sal.lower() != "nan":
                                    _ent = ":".join(_ent.split(":")[:2])
                                    _sal = ":".join(_sal.split(":")[:2])
                                    _horario_partes.append(f"{_dia} {_ent}-{_sal}")
                            if _horario_partes:
                                st.caption(f"🕒 {' | '.join(_horario_partes)}")


# ═══════════════════════════════════════════════════════════════════
# PROTOCOLO DE EMERGENCIA — Contacto de emergencia (acceso directo + log)
# ═══════════════════════════════════════════════════════════════════
# Reutiliza la misma lista de RFCs autorizados que el directorio de
# Centros de Maestros (secrets: rfcs_directorio_cm). No hay aprobación
# en tiempo real: el control de acceso YA ES esa lista, decidida de
# antemano. Cada consulta se registra sola en la bitácora.
#
# EN TU GOOGLE SHEET: agrega 2 columnas nuevas a la tab "Directorio":
#   CONTACTO_EMERGENCIA_NOMBRE   |   CONTACTO_EMERGENCIA_TEL

EMERGENCIA_TAB = "Bitacora_Emergencias"
EMERGENCIA_HEADERS = ["ID", "CONSULTOR_RFC", "CONSULTOR_NOMBRE", "COMPAÑERO_CONSULTADO", "FECHA_HORA"]

# ═══════════════════════════════════════════════════════════════════
# PADRÓN MAESTRO — Fuente única de verdad del personal (Fase 1 de migración)
# ═══════════════════════════════════════════════════════════════════
# Secrets necesario:  sheet_padron_id = "1F_CHd4p0n_5RGMS6NXYRXS_KUNM9UNkjdzL1LcZaCYU"
# Y compartir ese spreadsheet con el client_email de la service account (Editor).
#
# RED DE SEGURIDAD: si el padrón no está accesible (secret faltante, permiso
# faltante, tab renombrada), los loaders CAEN AUTOMÁTICAMENTE a la tab
# Directorio de siempre — la app nunca se queda sin directorio por la migración.

PADRON_TAB = "Padron"

# URL pública de la app: es la que se codifica en el QR de asistencia.
URL_APP_ASISTENCIA = st.secrets.get(
    "url_app", "https://gestion-personal-dfc.streamlit.app")

# El padrón se armó con nombres propios (NOMBRE, ADSCRIPCION_REAL) y la app
# nació esperando otros (NOMBRE_COMPLETO, UBICACION_FISICA). Se aceptan ambos
# para que un renombrado en el Sheet no vuelva a dejar la migración muda.
ALIAS_PADRON = {
    "NOMBRE":      ("NOMBRE", "NOMBRE_COMPLETO"),
    "AREA":        ("ADSCRIPCION_REAL", "UBICACION_FISICA", "AREA"),
    "DEPARTAMENTO":("ROL", "DEPARTAMENTO"),
}


def _col(df, *candidatas):
    """Primera columna existente de entre varias equivalentes."""
    return next((c for c in candidatas if c in df.columns), None)


@st.cache_data(ttl=300)
def cargar_padron():
    """Lee el PADRON_MAESTRO_DFC completo. Devuelve DataFrame vacío si no está
    accesible; cada consumidor decide su propio fallback.

    NO filtra por ACTIVO: en el padrón esa columna es ACTIVO_CHECADOR y solo
    indica quién marca en el reloj de piso 5 y 6. Filtrar con ella sacaría del
    directorio a todo el personal de Centros de Maestros, que sí labora.
    """
    try:
        client = get_client()
        sh = client.open_by_key(st.secrets["sheet_padron_id"])
        ws = sh.worksheet(PADRON_TAB)
        data = ws.get_all_records(numericise_ignore=["all"])
        return pd.DataFrame(data).fillna("")
    except Exception:
        return pd.DataFrame()


def padron_disponible() -> bool:
    """True si el padrón se puede leer y trae la columna de nombre."""
    p = cargar_padron()
    return (not p.empty) and _col(p, *ALIAS_PADRON["NOMBRE"]) is not None


@st.cache_data(ttl=600)
def _cargar_directorio_externo_cached() -> pd.DataFrame:
    # Si falla LANZA la excepción: Streamlit no guarda en caché los errores,
    # así que al crear la pestaña se lee en el siguiente intento, no 10 min después.
    sh = get_client().open_by_key(st.secrets["sheet_padron_id"])
    # get_all_values en vez de get_all_records: este último truena si la fila 1
    # tiene 2+ celdas vacías (columnas sobrantes con formato o datos sueltos).
    # Aquí se toman solo las columnas con encabezado y se ignoran las demás.
    valores = sh.worksheet("Directorio_Externo").get_all_values()
    if not valores:
        return pd.DataFrame(columns=["NOMBRE", "AREA", "DEPARTAMENTO", "EXTENSION", "CORREO"])
    enc = [str(h).strip().upper() for h in valores[0]]
    utiles = [i for i, h in enumerate(enc) if h]
    filas = [[(f[i] if i < len(f) else "") for i in utiles] for f in valores[1:]]
    df = pd.DataFrame(filas, columns=[enc[i] for i in utiles]).fillna("")
    df = df.loc[:, ~df.columns.duplicated()]  # si un nombre se repite, gana el primero
    for col in ("NOMBRE", "AREA", "DEPARTAMENTO", "EXTENSION", "CORREO", "CELULAR"):
        if col not in df.columns:
            df[col] = ""
    df = df[df["NOMBRE"].astype(str).str.strip() != ""].copy()
    df["_EXTERNO"] = "SI"  # para agruparlos por su oficina, no por cargo
    return df


def cargar_directorio_externo():
    """Contactos de oficinas FUERA de la DFC (p. ej. Despacho de la
    Subsecretaría de Atención al Magisterio), en la pestaña
    'Directorio_Externo' del Sheet del padrón. No se mezclan con el padrón:
    sin login, sin contacto de emergencia, sin nómina.
    Devuelve (df, motivo): motivo explica por qué vino vacío."""
    try:
        df = _cargar_directorio_externo_cached()
        return df, ("" if not df.empty else "la pestaña existe pero no tiene filas con NOMBRE")
    except Exception as e:
        if type(e).__name__ == "WorksheetNotFound":
            return pd.DataFrame(), ("no existe una pestaña llamada exactamente "
                                    "'Directorio_Externo' en el Sheet del padrón")
        return pd.DataFrame(), f"no se pudo leer ({type(e).__name__}: {e})"


def _directorio_unificado():
    """Directorio derivado del padrón, o de la tab Directorio si no hay padrón.
    Devuelve siempre las columnas que las pantallas ya esperan."""
    padron = cargar_padron()
    c_nom = _col(padron, *ALIAS_PADRON["NOMBRE"]) if not padron.empty else None

    if c_nom:
        df = padron.copy()
        df["NOMBRE"] = df[c_nom]
        c_area = _col(df, *ALIAS_PADRON["AREA"])
        df["AREA"] = df[c_area] if c_area else ""
        c_dep = _col(df, *ALIAS_PADRON["DEPARTAMENTO"])
        df["DEPARTAMENTO"] = df[c_dep] if c_dep else ""
        for col in ("EXTENSION", "CORREO",
                    "CONTACTO_EMERGENCIA_NOMBRE", "CONTACTO_EMERGENCIA_TEL"):
            if col not in df.columns:
                df[col] = ""
        # Un directorio implica jerarquía. La tab antigua la llevaba implícita
        # en el orden de las filas —el primero de cada área era el jefe— y eso
        # se perdía al ordenar el padrón por nombre. Ahora es explícita:
        # ORDEN_JERARQUIA (1 jefe, 2 personal, 3 administrativo).
        if "ORDEN_JERARQUIA" in df.columns:
            df["_ord"] = pd.to_numeric(df["ORDEN_JERARQUIA"], errors="coerce").fillna(2)
            # En los Centros de Maestros la columna DEPARTAMENTO guarda el rol
            # (Asesor, Responsable), no un departamento: agrupar por ella los
            # desordenaría. Solo se usa como agrupador en oficinas centrales.
            _roles_cm = ("Asesor", "Responsable", "Administrativo")
            df["_grupo"] = df["DEPARTAMENTO"].where(
                ~df["DEPARTAMENTO"].isin(_roles_cm), "")
            df = (df.sort_values(["AREA", "_grupo", "_ord", "NOMBRE"])
                    .drop(columns=["_ord", "_grupo"]))
        st.session_state["_fuente_directorio"] = "padron"
        return df.reset_index(drop=True).fillna("")

    st.session_state["_fuente_directorio"] = "tab_directorio (padrón no accesible)"
    client = get_client()
    sh = client.open_by_key(st.secrets["sheet_checador_id"])
    ws = sh.worksheet("Directorio")
    data = ws.get_all_records(numericise_ignore=["all"])
    return pd.DataFrame(data).fillna("")


def aviso_fuente_personal():
    """Deja ver de dónde salieron los datos. Sin esto la migración es muda: el
    fallback hace que la app se vea idéntica aunque el padrón no se lea."""
    fuente = st.session_state.get("_fuente_directorio", "")
    nom = st.session_state.get("_fuente_nomina", "")
    if fuente == "padron":
        detalle = " · Nómina derivada del padrón." if nom.startswith("padron") else ""
        st.caption(f"Datos de personal: **Padrón Maestro**.{detalle}")
    elif fuente:
        st.warning("El Padrón Maestro no está accesible: se están usando las "
                   "tabs antiguas. Revisa el secret `sheet_padron_id` y que el "
                   "Sheet esté compartido con la cuenta de servicio.")


@st.cache_data(ttl=300)
def _cargar_directorio_completo():
    """Loader a nivel módulo para el protocolo de emergencia y el directorio.
    FASE 1: lee del Padrón Maestro, con fallback automático a la tab Directorio."""
    return _directorio_unificado()

def _ws_emergencias():
    client = get_client()
    sh = client.open_by_key(st.secrets["sheet_checador_id"])
    try:
        return sh.worksheet(EMERGENCIA_TAB)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(EMERGENCIA_TAB, rows=2, cols=len(EMERGENCIA_HEADERS))
        ws.append_row(EMERGENCIA_HEADERS)
        return ws

def registrar_consulta_emergencia(rfc: str, nombre: str, companero: str):
    """Escribe el log EN EL MOMENTO en que se revela el contacto.
    Nunca debe bloquear una emergencia real por un fallo de red."""
    try:
        ws = _ws_emergencias()
        todas = ws.get_all_records(numericise_ignore=["all"])
        nuevo_id = len(todas) + 1
        ahora = datetime.now(pytz.timezone("America/Mexico_City")).strftime("%Y-%m-%d %H:%M")
        ws.append_row([nuevo_id, rfc, nombre, companero, ahora], value_input_option="USER_ENTERED")
    except Exception:
        pass

def _puede_pendientes_rh() -> bool:
    """Quién entra a Pendientes RH DFC: el admin y sus asesores.

    Reusa rfcs_oficios, que son exactamente esas tres personas (el admin y
    sus dos asesores del área de RH). Tener dos listas con la misma gente
    garantiza que un día se desfasen y alguien pierda o gane acceso sin que
    nadie lo note.

    Cuidado con el día en que eso deje de ser cierto: rfcs_oficios dice
    "quién puede emitir oficios". Si mañana se le da esa facultad a alguien
    más, entraría aquí de pasada, sin que nadie lo haya decidido. Cuando pase
    —o antes, si se prefiere— se agrega el secret rfcs_pendientes_rh con su
    propia lista y esta función lo prefiere automáticamente, sin tocar código.
    """
    if st.session_state.get("rol") == "admin":
        return True
    rfc = str(st.session_state.get("rfc", "")).upper().strip()
    if not rfc:
        return False
    propios = _rfcs_secret("rfcs_pendientes_rh")
    return rfc in (propios if propios else _rfcs_secret("rfcs_oficios"))


def _rfcs_secret(clave: str) -> list[str]:
    try:
        return [str(r).upper().strip() for r in st.secrets.get(clave, [])]
    except Exception:
        return []

def _normalizar_busqueda(texto: str) -> str:
    """Mayúsculas y sin acentos, para comparar sin importar cómo esté
    capturado el nombre en el Sheet (con/sin acentos, mayús/minús).
    'Ángel' y 'angel' deben coincidir; antes no lo hacían."""
    import unicodedata
    t = str(texto or "").upper().strip()
    t = unicodedata.normalize("NFKD", t)
    return "".join(c for c in t if not unicodedata.combining(c))

def _tokens_persona(txt: str) -> frozenset:
    """Nombre en piezas comparables: sin acentos, sin puntuación y con Z→S
    (CARRIZALEZ/CARRIZALES). Se comparan como conjunto para que no importe si
    el minutario escribió 'Garcia Tirado Susana' y el padrón otro orden."""
    t = unicodedata.normalize("NFKD", str(txt or "").upper())
    t = "".join(c for c in t if not unicodedata.combining(c)).replace("Z", "S")
    t = re.sub(r"[^A-ZÑ ]", " ", t)
    return frozenset(p for p in t.split() if len(p) > 1)


def _mismo_nombre(a: frozenset, b: frozenset) -> bool:
    """Coinciden si uno contiene al otro (segundo nombre de más) y comparten
    al menos 2 piezas: así 'GARCIA TIRADO' solo no basta para dar un falso
    positivo con otro García."""
    if not a or not b:
        return False
    return len(a & b) >= 2 and (a <= b or b <= a)


def pendientes_acuse_de_centro(centro: str) -> list:
    """Comisiones del personal de ese Centro que aún no tienen acuse firmado.
    Se leen del minutario: asunto con 'comisión' y estado distinto de ACUSE
    (CANCELADO se ignora). Devuelve [{nombre, oficios: [...]}]."""
    if _oficios_mod is None or _aqr_mod is None or not centro:
        return []
    try:
        df = _oficios_mod.cargar_oficios(get_client)
    except Exception:
        return []
    if df.empty or "DIRIGIDO_A" not in df.columns:
        return []
    _asunto = df.get("ASUNTO", pd.Series([""] * len(df))).astype(str)
    _asunto_norm = _asunto.apply(lambda x: "".join(
        c for c in unicodedata.normalize("NFKD", x.upper()) if not unicodedata.combining(c)))
    _estado = df.get("ESTADO", pd.Series([""] * len(df))).astype(str).str.upper().str.strip()
    df = df[_asunto_norm.str.contains("COMISION", na=False)
            & ~_estado.isin(["ACUSE", "CANCELADO", "HISTORICO"])]
    if df.empty:
        return []
    try:
        gente = _aqr_mod.asesores_de(centro)
    except Exception:
        return []
    if gente.empty:
        return []
    filas = [(str(r["NOMBRE"]), _tokens_persona(r["NOMBRE"])) for _, r in gente.iterrows()]
    por_persona = {}
    for _, o in df.iterrows():
        tok_of = _tokens_persona(o.get("DIRIGIDO_A", ""))
        for nombre, tok_p in filas:
            if _mismo_nombre(tok_p, tok_of):
                por_persona.setdefault(nombre, []).append({
                    "id": str(o.get("ID_OFICIO", "")),
                    "fecha": str(o.get("FECHA_OFICIO", "") or o.get("FECHA_SOLICITUD", "")),
                    "asunto": str(o.get("ASUNTO", "")),
                    "estado": str(o.get("ESTADO", "")).upper(),
                })
                break
    return [{"nombre": n, "oficios": o} for n, o in
            sorted(por_persona.items(), key=lambda kv: -len(kv[1]))]


@st.cache_data(ttl=300)
def cargar_pendientes_nomina_sheet() -> pd.DataFrame:
    """Pendientes de firma de nómina que RH dejó registrados en la tab
    Pendientes_Nomina. Vacío si la tab aún no existe."""
    try:
        sh = get_client().open_by_key(st.secrets["sheet_checador_id"])
        valores = sh.worksheet("Pendientes_Nomina").get_all_values()
        if len(valores) < 2:
            return pd.DataFrame()
        enc = [str(h).strip().upper() for h in valores[0]]
        utiles = [i for i, h in enumerate(enc) if h]
        filas = [[(f[i] if i < len(f) else "") for i in utiles] for f in valores[1:]]
        return pd.DataFrame(filas, columns=[enc[i] for i in utiles]).fillna("")
    except Exception:
        return pd.DataFrame()


def pendientes_nomina_de_centro(centro: str) -> list:
    """Firmas de nómina pendientes del personal de ese Centro.
    Devuelve [{nombre, conceptos: [(nomina, concepto)]}]."""
    df = cargar_pendientes_nomina_sheet()
    if df.empty or "NOMBRE" not in df.columns or _aqr_mod is None or not centro:
        return []
    if "ESTADO" in df.columns:
        df = df[df["ESTADO"].astype(str).str.upper().str.strip() != "FIRMADO"]
    if df.empty:
        return []
    try:
        gente = _aqr_mod.asesores_de(centro)
    except Exception:
        return []
    if gente.empty:
        return []
    filas = [(str(r["NOMBRE"]), _tokens_persona(r["NOMBRE"])) for _, r in gente.iterrows()]
    por_persona = {}
    for _, p in df.iterrows():
        tok = _tokens_persona(p.get("NOMBRE", ""))
        for nombre, tok_p in filas:
            if _mismo_nombre(tok_p, tok):
                por_persona.setdefault(nombre, []).append(
                    (str(p.get("NOMINA", "")), str(p.get("CONCEPTO", ""))))
                break
    return [{"nombre": n, "conceptos": c} for n, c in
            sorted(por_persona.items(), key=lambda kv: -len(kv[1]))]


def vista_pendientes_cm():
    """Pendientes de entregar acuse, para el coordinador del Centro."""
    rfc_actual = str(st.session_state.get("rfc", "")).upper().strip()
    centro = ""
    if _aqr_mod is not None:
        try:
            centro = _aqr_mod.centro_del_responsable(rfc_actual)
        except Exception:
            centro = ""
    if st.session_state.get("rol") == "admin" and not centro:
        try:
            opciones = _aqr_mod.centros_disponibles() if _aqr_mod else []
        except Exception:
            opciones = []
        if not opciones:
            st.info("El padrón no tiene personal de Centros de Maestros.")
            return
        centro = st.selectbox("Centro (vista admin)", opciones, key="pend_centro")
    if not centro:
        st.error("No tienes permiso para esta sección.")
        return

    st.markdown(f"## 📌 Pendientes — {centro}")
    datos = pendientes_acuse_de_centro(centro)
    total = sum(len(d["oficios"]) for d in datos)
    nom = pendientes_nomina_de_centro(centro)
    total_nom = sum(len(d["conceptos"]) for d in nom)

    if not total and not total_nom:
        st.success("✅ Sin pendientes: tu Centro está al corriente en acuses y en firmas de nómina.")
        return

    # Encabezado con los totales y, abajo, el desglose plegado: un Centro con
    # muchos pendientes no satura la pantalla.
    c1, c2, c3 = st.columns(3)
    c1.metric("Acuses por entregar", total)
    c2.metric("Firmas de nómina", total_nom)
    c3.metric("Personas", len({d["nombre"] for d in datos} | {d["nombre"] for d in nom}))

    if total:
        st.markdown("### 📄 Acuses de comisión por entregar")
        st.caption("Entrega en RH el acuse firmado de cada comisión. "
                   "Toca un nombre para ver sus oficios.")
        for d in datos:
            with st.expander(f"{d['nombre']} — {len(d['oficios'])} pendiente(s)"):
                for o in d["oficios"]:
                    st.markdown(f"• **{o['id']}** · {o['fecha']} — {o['asunto']}")

    if total_nom:
        st.markdown("### 💰 Firmas de nómina pendientes")
        st.caption("Deben presentarse en la Dirección a firmar. La Dirección de Pagos "
                   "permite un rezago máximo de 2 quincenas.")
        for d in nom:
            with st.expander(f"{d['nombre']} — {len(d['conceptos'])} firma(s)"):
                for nomina, concepto in d["conceptos"]:
                    st.markdown(f"• **{concepto}** · nómina {nomina}")


def _es_responsable_cm(rfc_actual: str) -> bool:
    """Responsable de un Centro de Maestros según el padrón (vía
    asistencia_qr), sin depender de la lista piloto de cm_module."""
    try:
        return bool(_aqr_mod.centro_del_responsable(rfc_actual)) if _aqr_mod is not None else False
    except Exception:
        return False

def vista_contacto_emergencia():
    rfc_actual = str(st.session_state.get("rfc", "")).upper().strip()
    autorizados = _rfcs_secret("rfcs_directorio_cm")
    # Los responsables de Centros de Maestros TAMBIÉN necesitan esto: son
    # quienes en la práctica atienden emergencias de sus propios asesores.
    if (st.session_state.get("rol") != "admin"
            and rfc_actual not in autorizados
            and not _es_responsable_cm(rfc_actual)):
        st.error("No tienes permiso para esta sección.")
        return

    st.markdown("## 🚨 Contacto de Emergencia")

    # ── Disclaimer con aceptación explícita: se pide en CADA sesión ──
    st.error(
        "**AVISO IMPORTANTE — Uso restringido**\n\n"
        "Esta sección contiene datos personales confidenciales (contacto de "
        "emergencia) y debe usarse ÚNICAMENTE en caso de una emergencia real.\n\n"
        "Cada consulta queda registrada de forma permanente con tu nombre, RFC, "
        "la persona consultada y la fecha/hora exacta.\n\n"
        "El uso de esta información para fines distintos a una emergencia "
        "constituye una falta administrativa y, en su caso, puede derivar en "
        "responsabilidad legal conforme a la normatividad aplicable en materia "
        "de protección de datos personales."
    )
    acepto = st.checkbox(
        "Entiendo el aviso anterior, confirmo que se trata de una emergencia real "
        "y acepto que esta consulta quedará registrada con mis datos."
    )
    if not acepto:
        st.info("Marca la casilla anterior para continuar.")
        return

    try:
        df = _cargar_directorio_completo()
    except Exception:
        df = pd.DataFrame()
    if df.empty or "NOMBRE" not in df.columns:
        st.info("No se pudo cargar el directorio.")
        return

    # Mostrar SIEMPRE en mayúsculas, sin importar cómo esté capturado en el
    # Sheet (mezcla de mayúsculas/minúsculas rompía la vista y la búsqueda).
    df = df.copy()
    df["NOMBRE"] = df["NOMBRE"].astype(str).str.upper().str.strip()
    df["_NOMBRE_BUSQUEDA"] = df["NOMBRE"].apply(_normalizar_busqueda)

    busq = st.text_input("🔍 Busca por nombre", placeholder="Escribe nombre o apellido, con o sin acentos...")
    if busq:
        q = _normalizar_busqueda(busq)
        opciones = df[df["_NOMBRE_BUSQUEDA"].str.contains(q, na=False)]["NOMBRE"].tolist()
        st.caption(f"{len(opciones)} resultado(s)")
    else:
        opciones = df["NOMBRE"].tolist()

    companero = st.selectbox("Selecciona al compañero", [""] + sorted(set(opciones)))

    if companero and st.button("🚨 VER CONTACTO DE EMERGENCIA", type="primary", use_container_width=True):
        fila = df[df["NOMBRE"] == companero]
        if fila.empty:
            st.error("No se encontró a esa persona en el directorio.")
        else:
            r = fila.iloc[0]
            c_nombre = str(r.get("CONTACTO_EMERGENCIA_NOMBRE", "")).strip()
            c_tel = str(r.get("CONTACTO_EMERGENCIA_TEL", "")).strip()
            registrar_consulta_emergencia(rfc_actual, st.session_state.get("nombre", ""), companero)
            if c_nombre or c_tel:
                nombre_txt = c_nombre or "(sin nombre registrado)"
                tel_txt = c_tel or "(sin teléfono registrado)"
                st.success(f"**Contacto de {companero}:**\n\n👤 {nombre_txt}\n\n📞 {tel_txt}")
            else:
                st.warning(f"{companero} no tiene contacto de emergencia registrado en el sistema. "
                          "Contacta directamente a RH.")
            _tz_emer = pytz.timezone("America/Mexico_City")
            _sello = datetime.now(_tz_emer).strftime("%d/%m/%Y %H:%M")
            _nombre_ses = st.session_state.get("nombre", "")
            st.caption(f"Consulta registrada — {_nombre_ses} · {rfc_actual} · {_sello}")


# ── Centros de Maestros: extraído a cm_module.py ──
try:
    import cm_module as _cm_module
    _ERROR_CM = ""
except Exception as _e_cm:
    _cm_module = None
    _ERROR_CM = str(_e_cm)

if _cm_module is not None:
    _cm_module.configurar(get_client=get_client, error_amable=_error_amable,
                          columnas_horario=COLUMNAS_HORARIO,
                          subir_archivo_drive=subir_archivo_drive,
                          carpeta_asistencia_cm=DRIVE_ASISTENCIA_CM_FOLDER)

# ── Asistencia por QR rotativo: asistencia_qr.py ──
# Reemplaza la toma de lista manual (retirada del sidebar y del enrutador).
# La tab vieja en Sheets no se borró: sigue ahí como histórico, solo ya
# nadie puede llegar a ella desde la app.
try:
    import asistencia_qr as _aqr_mod
    _ERROR_AQR = ""
except Exception as _e_aqr:
    _aqr_mod = None
    _ERROR_AQR = str(_e_aqr)

if _aqr_mod is not None:
    try:
        _aqr_mod.configurar(
            get_client=get_client,
            error_amable=_error_amable,
            cargar_padron=cargar_padron,
            sheet_asistencia_id=st.secrets.get("sheet_asistencia_cm_id",
                                               st.secrets.get("sheet_checador_id", "")),
            subir_archivo_drive=subir_archivo_drive,
            carpeta_asistencia_cm=DRIVE_ASISTENCIA_CM_FOLDER,
            saldo_economicos=saldo_economicos_rfc,
        )
    except Exception as _e_cfg:
        _ERROR_AQR = f"no se pudo configurar: {_e_cfg}"


def _puede_qr(rfc_actual: str) -> bool:
    """Responsable de un Centro de Maestros, o admin."""
    if _aqr_mod is None:
        return False
    if st.session_state.get("rol") == "admin":
        return True
    try:
        return bool(_aqr_mod.centro_del_responsable(rfc_actual))
    except Exception:
        return False


def vista_qr_pantalla():
    if _aqr_mod is None:
        st.error(f"El módulo de asistencia QR no está disponible: {_ERROR_AQR}")
        return
    _aqr_mod.vista_pantalla(URL_APP_ASISTENCIA)


def vista_mi_asistencia():
    if _aqr_mod is None:
        st.error(f"El módulo de asistencia QR no está disponible: {_ERROR_AQR}")
        return
    _aqr_mod.vista_mi_asistencia()


def _es_personal_cm(rfc_actual: str) -> bool:
    """Cualquiera adscrito a un Centro de Maestros, no solo el Responsable."""
    if _aqr_mod is None or not rfc_actual:
        return False
    try:
        cm = _aqr_mod._padron_cm()
        if cm.empty:
            return False
        return bool((cm["RFC"].astype(str).str.upper() == rfc_actual.upper()).any())
    except Exception:
        return False


def vista_qr_coordinador():
    if _aqr_mod is None:
        st.error(f"El módulo de asistencia QR no está disponible: {_ERROR_AQR}")
        return
    _aqr_mod.vista_coordinador()


def _centro_del_responsable(rfc_actual: str):
    if _cm_module is None:
        return None
    return _cm_module._centro_del_responsable(rfc_actual)

def vista_toma_lista_cm():
    if _cm_module is None:
        st.error(f"El módulo de Centros de Maestros no está disponible: {_ERROR_CM}")
        return
    _cm_module.vista_toma_lista_cm()


def vista_ari():
    """ARI embebida: chat de dudas de RH con contexto del usuario autenticado."""
    try:
        import ari_module
    except ImportError:
        st.error("El módulo de ARI (ari_module.py) no está en el proyecto.")
        return
    if "GEMINI_API_KEY" not in st.secrets:
        st.warning("Falta GEMINI_API_KEY en los secrets para habilitar a ARI.")
        return

    contexto = ""
    if st.session_state.get("rol") == "empleado":
        rfc    = st.session_state.get("rfc", "")
        nombre = st.session_state.get("nombre", "")
        # Minimización de datos hacia la API externa: solo el nombre de pila.
        # Convención del Sheet: APELLIDO_P APELLIDO_M NOMBRE(S).
        _tk = str(nombre).split()
        nombre_pila = " ".join(_tk[2:]) if len(_tk) >= 3 else (_tk[-1] if _tk else "")
        partes = [f"Nombre: {nombre_pila}"]
        try:
            empleados   = cargar_empleados()
            solicitudes = cargar_solicitudes_eco()
            incidencias = cargar_incidencias()
            emp_row = empleados[empleados["RFC"].astype(str).str.upper() == rfc]
            if not emp_row.empty:
                dias_totales = int(emp_row["DIAS TOTALES"].iloc[0])
                usados       = dias_economicos_usados(rfc, solicitudes)
                partes.append(f"Días económicos: {dias_totales - usados} disponibles de {dias_totales} este año")
            if not incidencias.empty and "RFC" in incidencias.columns:
                propias = incidencias[incidencias["RFC"].astype(str).str.upper() == rfc]
                # Mismo criterio que el panel del admin: si las dos pantallas
                # filtran distinto, el empleado ve pendientes que RH no ve.
                pend = propias[_mascara_pendiente(propias)]
                if not pend.empty:
                    folios = ", ".join(pend["FOLIO"].astype(str).tolist()[:5])
                    partes.append(f"Solicitudes pendientes de autorización: {folios}")
        except Exception:
            pass
        contexto = "\n".join(partes)
    elif st.session_state.get("rol") == "admin":
        contexto = "El usuario actual es el administrador de RH de la DFC."

    # deps_pase enciende la captura de pases por chat. Si se omite, ARI
    # funciona igual que antes: solo responde dudas.
    #
    # ARI no escribe en ningún Sheet: arma los datos y llama a
    # enviar_solicitud(), la MISMA función del formulario, para que todo pase
    # por una sola puerta (folio en vivo, candado anti duplicado,
    # INSERT_ROWS, bitácora y comprobante PDF incluidos).
    _rfc_ari = str(st.session_state.get("rfc", "")).strip()
    _nom_ari = str(st.session_state.get("nombre", "")).strip()
    if _rfc_ari and _nom_ari:
        ari_module.render_ari(contexto, deps_pase={
            "enviar_solicitud":   enviar_solicitud,
            "cargar_horarios":    cargar_horarios,
            "cargar_incidencias": cargar_incidencias,
            "rfc":                _rfc_ari,
            "nombre":             _nom_ari,
            "jefe_inmediato":     st.session_state.get("jefe_inmediato", ""),
        })
    else:
        ari_module.render_ari(contexto)


def main():
    st.set_page_config(page_title="Incidencias DFC · RH", page_icon="📋", layout="wide")

    # ── Interceptor QR de validación ─────────────
    if "validar_folio" in st.query_params:
        folio_a_buscar = st.query_params["validar_folio"]
        st.markdown("## 🔍 Verificación de Autenticidad")
        st.caption("Módulo de Control Interno · Dirección de Formación Continua · SEJ")
        st.divider()
        incidencias   = cargar_incidencias()
        solicitudes   = cargar_solicitudes_eco()
        # Buscar en Incidencias (PSE, COM, CHO)
        match = incidencias[incidencias["FOLIO"].astype(str) == folio_a_buscar]
        # Si no está, buscar en Solicitudes (ECO)
        if match.empty and not solicitudes.empty and "FOLIO" in solicitudes.columns:
            match_eco = solicitudes[solicitudes["FOLIO"].astype(str) == folio_a_buscar]
            encontrado_en = "solicitudes"
        else:
            match_eco = pd.DataFrame()
            encontrado_en = "incidencias"

        registro = match if not match.empty else match_eco

        if not registro.empty:
            row = registro.iloc[0]
            st.success("✅ DOCUMENTO AUTÉNTICO Y REGISTRADO EN SISTEMA")
            with st.container(border=True):
                st.write(f"**Folio Oficial:** {folio_a_buscar}")
                nombre     = row.get("NOMBRE") or row.get("Nombre Completo", "")
                rfc_reg    = row.get("RFC", "")
                tipo_raw   = row.get("TIPO") or row.get("Tipo Permiso", "")
                tipo_reg   = "ECO" if str(tipo_raw).lower() in ["economico", "económico"] else tipo_raw
                fi_reg     = row.get("FECHA_INICIO") or row.get("Fecha Inicio", "")
                ff_reg     = row.get("FECHA_FIN") or row.get("Fecha Fin", "")
                aprobado   = row.get("AUTORIZADO_POR") or row.get("Aprobado Por", "")
                estado_reg = row.get("ESTADO") or ("AUTORIZADO" if str(aprobado).strip() != "" else "PENDIENTE")
                motivo_reg = row.get("MOTIVO") or row.get("Motivo", "")
                st.write(f"**Servidor Público:** {nombre}")
                st.write(f"**Filiación:** {rfc_oculto(str(rfc_reg))}")
                st.write(f"**Incidencia:** {TIPO_LABELS.get(tipo_reg, tipo_reg)}")
                st.write(f"**Periodo:** {fi_reg} al {ff_reg}")
                st.write(f"**Estado:** {estado_reg}")
                # PRIVACIDAD: el motivo NO se muestra en el validador público.
                # Los folios son enumerables en la URL y el motivo puede contener
                # información médica/familiar. Para validar autenticidad basta
                # folio + tipo + periodo + estado. El motivo se consulta con sesión.
        else:
            st.error("🚨 ALERTA: DOCUMENTO NO ENCONTRADO O ALTERADO")
            st.warning("Este folio no existe en los registros oficiales de la DFC. El formato impreso podría ser falso o modificado de manera ilícita.")
        if st.button("Ir al inicio de sesión"):
            st.query_params.clear()
            st.rerun()
        return

    # ── Escaneo del QR de asistencia ──
    # Solo se procesa CON sesión iniciada. Sin ella se deja pasar al login: el
    # parámetro sobrevive al rerun, así que en cuanto entra se registra sin
    # tener que volver a escanear. Cortar aquí dejaba al asesor atrapado en
    # una pantalla que le pedía sesión y no se la ofrecía.
    if "asistencia" in st.query_params and "rol" in st.session_state:
        if _aqr_mod is None:
            st.error(f"El módulo de asistencia no está disponible: {_ERROR_AQR}")
        else:
            _aqr_mod.procesar_escaneo(st.query_params["asistencia"])
        if st.button("Continuar al portal", key="btn_post_asistencia"):
            st.query_params.clear()
            st.rerun()
        st.stop()

    if "validar_oficio" in st.query_params and _oficios_mod is not None:
        _oficios_mod.render_validacion_oficio(get_client, st.query_params["validar_oficio"])
        return

    if "rol" not in st.session_state:
        st.markdown("<style>[data-testid='stSidebar']{display:none}</style>", unsafe_allow_html=True)
        if "asistencia" in st.query_params:
            st.info("Inicia sesión para registrar tu asistencia. El código que "
                    "escaneaste sigue vigente y se aplicará en cuanto entres.")
        login()
        return

    with st.sidebar:
        st.image("dfc_logo.png", use_container_width=True)
        st.markdown("**Portal de Gestión de Incidencias**")
        st.caption('👤 Servidor(a) Público(a): ' + st.session_state['nombre'])
        st.caption('🏷️ Nivel de acceso: ' + ('Administrador RH' if st.session_state['rol'] == 'admin' else 'Personal de la Dirección'))
        st.divider()
        if st.session_state.get("rol") == "admin":
            if st.button("🔄 Limpiar caché"):
                st.cache_data.clear()
                st.success("Caché limpiado.")
        if st.button("📅 Calendario de pagos"):
            st.session_state["vista"] = "calendario"
            st.rerun()
        if st.button("💰 Mis comprobantes de nómina"):
            st.session_state["vista"] = "nomina"
            st.rerun()
        if st.button("📞 Directorio DFC"):
            st.session_state["vista"] = "directorio"
            st.rerun()
        _rfc_sb = str(st.session_state.get("rfc", "")).upper().strip()
        _autoriz_emer = _rfcs_secret("rfcs_directorio_cm")
        if (st.session_state.get("rol") == "admin" or _rfc_sb in _autoriz_emer
                or _es_responsable_cm(_rfc_sb)):
            if st.button("🚨 Contacto de Emergencia", key="btn_emer_sidebar"):
                st.session_state["vista"] = "emergencia"
                st.rerun()
            # Colorea SOLO este botón (busca por su texto exacto, no depende
            # de clases internas de Streamlit que cambian entre versiones).
            import streamlit.components.v1 as components
            components.html("""
                <script>
                function _colorearBotonEmergencia() {
                    try {
                        const botones = window.parent.document.querySelectorAll('button');
                        botones.forEach(function(b) {
                            if (b.innerText && b.innerText.indexOf("Contacto de Emergencia") !== -1) {
                                b.style.backgroundColor = "#FF6B00";
                                b.style.color = "#FFFFFF";
                                b.style.fontWeight = "bold";
                                b.style.border = "none";
                            }
                        });
                    } catch (e) {}
                }
                _colorearBotonEmergencia();
                if (!window.parent._emerObserverSet) {
                    window.parent._emerObserverSet = true;
                    new MutationObserver(_colorearBotonEmergencia)
                        .observe(window.parent.document.body, {childList: true, subtree: true});
                }
                </script>
            """, height=0)

        # Toma de lista manual RETIRADA: el sistema de QR (abajo) la reemplaza
        # por decisión explícita. Antes de esto, Andrade Zamorano (único en
        # el secret piloto viejo) veía AMBOS sistemas a la vez y por eso
        # seguía usando la lista manual por costumbre — los demás coordinadores
        # nunca vieron ese botón, de ahí la inconsistencia.
        if st.session_state.get("rol") == "admin" or _es_responsable_cm(_rfc_sb):
            _n_pend = 0
            try:
                _c_pend = _aqr_mod.centro_del_responsable(_rfc_sb) if _aqr_mod else ""
                if _c_pend:
                    _n_pend = (sum(len(d["oficios"]) for d in pendientes_acuse_de_centro(_c_pend))
                               + sum(len(d["conceptos"]) for d in pendientes_nomina_de_centro(_c_pend)))
            except Exception:
                _n_pend = 0
            if st.button(f"📌 Pendientes ({_n_pend})" if _n_pend else "📌 Pendientes",
                         key="btn_pend_sidebar"):
                st.session_state["vista"] = "pendientes_cm"
                st.rerun()
            # Azul, con el mismo método del botón naranja de emergencia: se
            # busca por su texto, porque las clases internas de Streamlit
            # cambian entre versiones.
            import streamlit.components.v1 as _comp_pend
            _comp_pend.html("""
                <script>
                function _colorearBotonPendientes() {
                    try {
                        window.parent.document.querySelectorAll('button').forEach(function(b) {
                            if (b.innerText && b.innerText.indexOf("Pendientes") !== -1) {
                                b.style.backgroundColor = "#1F4E9C";
                                b.style.color = "#FFFFFF";
                                b.style.fontWeight = "bold";
                                b.style.border = "none";
                            }
                        });
                    } catch (e) {}
                }
                _colorearBotonPendientes();
                if (!window.parent._pendObserverSet) {
                    window.parent._pendObserverSet = true;
                    new MutationObserver(_colorearBotonPendientes)
                        .observe(window.parent.document.body, {childList: true, subtree: true});
                }
                </script>
            """, height=0)

        if _es_personal_cm(_rfc_sb):
            if st.button("📍 Mi asistencia", key="btn_mi_asistencia"):
                st.session_state["vista"] = "mi_asistencia"
                st.rerun()
        if _puede_qr(_rfc_sb):
            if st.button("📱 Pantalla de asistencia (QR)", key="btn_qr_pantalla"):
                st.session_state["vista"] = "qr_pantalla"
                st.rerun()
            if st.button("🤝 Asistencia: registro y consulta", key="btn_qr_coord"):
                st.session_state["vista"] = "qr_coordinador"
                st.rerun()
        if st.session_state.get("rol") == "admin" or _rfc_sb in _rfcs_secret("rfcs_oficios"):
            if st.button("📄 Minutario de oficios", key="btn_of_sidebar"):
                st.session_state["vista"] = "oficios"
                st.rerun()
        if st.button("🤖 Pregúntale a ARI la IA de RH"):
            st.session_state["vista"] = "ari"
            st.rerun()
        # Pendientes RH DFC: SOLO el admin y sus dos asesores (rfcs_oficios).
        # El detalle de por qué ese secret y no otro está en
        # _puede_pendientes_rh.
        if render_pendientes_rh is not None and _puede_pendientes_rh():
            if st.button("📋 Pendientes RH DFC", key="btn_pend_rh_sidebar",
                         use_container_width=True):
                st.session_state["vista"] = "pendientes_rh"
                st.rerun()
        if st.button("🏠 Inicio"):
            st.session_state["vista"] = "inicio"
            st.rerun()
        if st.button("Cerrar sesión"):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()

    vista = st.session_state.get("vista", "inicio")
    if vista == "ari":
        vista_ari()
    elif vista == "directorio":
        vista_directorio()
    elif vista == "emergencia":
        vista_contacto_emergencia()
    elif vista == "pendientes_cm":
        vista_pendientes_cm()
    elif vista == "pendientes_rh":
        # Se vuelve a revisar aquí, no solo al pintar el botón: esconder un
        # botón no es un permiso, y la vista queda guardada en la sesión.
        if not _puede_pendientes_rh():
            st.error("No tienes acceso a esta sección.")
        elif render_pendientes_rh is None:
            st.error("El módulo de Pendientes RH no está disponible: "
                     + globals().get("_ERROR_PENDIENTES_RH", ""))
        else:
            render_pendientes_rh(get_client)
    elif vista == "mi_asistencia":
        vista_mi_asistencia()
    elif vista == "qr_pantalla":
        vista_qr_pantalla()
    elif vista == "qr_coordinador":
        vista_qr_coordinador()
    elif vista == "calendario":
        vista_calendario()
    elif vista == "nomina":
        vista_nomina()
    elif vista == "oficios":
        render_oficios()
    elif st.session_state["rol"] == "admin":
        vista_admin()
    else:
        vista_empleado()

if __name__ == "__main__":
    main()