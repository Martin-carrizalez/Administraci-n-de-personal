# -*- coding: utf-8 -*-
"""asistencia_qr.py — Asistencia de Centros de Maestros por QR rotativo.

Sustituye la toma de lista manual. La diferencia de fondo: la hora la pone el
servidor, no la persona. Quien llega a las 10:01 queda registrado a las 10:01.

Cómo cierra los dos huecos de un QR impreso:

  · El código CAMBIA cada 30 segundos y va firmado con un secreto del
    servidor. Fotografiarlo no sirve: a los 60 segundos ya no vale, y no se
    puede fabricar uno sin el secreto.
  · El QR NO identifica a la persona, solo al centro y al momento. La
    identidad la pone la sesión del asesor. Por eso nadie puede registrar a
    un compañero ausente: tendría que tener su sesión.

Lo que este diseño NO resuelve, y conviene tenerlo claro: un asesor puede
escanear y retirarse. Prueba que estuvo ahí a esa hora, no que se quedó.

Dependencias inyectadas vía configurar(), sin importar nada de
app_incidencias (evita el import circular).
"""

import hashlib
import hmac
import re
import unicodedata
from datetime import datetime, timedelta

import pandas as pd
import pytz
import streamlit as st

_deps = {}

TZ = pytz.timezone("America/Mexico_City")

# Ventana de validez del código. 30 s de rotación y 60 s de tolerancia: cabe
# el tiempo de enfocar y que cargue la página, sin dar margen para reenviar
# el código a alguien que no está en el centro.
SEGUNDOS_VENTANA = 30
VENTANAS_TOLERADAS = 2

TAB_ASISTENCIA = "Asistencia_QR"
COLUMNAS_ASISTENCIA = [
    "FECHA", "CENTRO", "RFC", "NOMBRE",
    "HORA_ENTRADA", "HORA_SALIDA",
    "METODO_ENTRADA", "METODO_SALIDA",
    "REGISTRADO_POR", "MOTIVO",
    "TS_ENTRADA", "TS_SALIDA",
]

DIAS = ["LUN", "MAR", "MIE", "JUE", "VIE"]
TOLERANCIA_MIN = 10


def configurar(get_client, error_amable, cargar_padron, sheet_asistencia_id):
    """Llamar UNA vez desde app_incidencias, después de definir dependencias."""
    _deps["get_client"] = get_client
    _deps["error_amable"] = error_amable
    _deps["cargar_padron"] = cargar_padron
    _deps["sheet_id"] = sheet_asistencia_id


def _client():
    return _deps["get_client"]()


def _error(e, ctx=""):
    return _deps["error_amable"](e, ctx)


def _ahora():
    return datetime.now(TZ)


def _norm(s):
    s = unicodedata.normalize("NFKD", str(s or ""))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s.upper()).strip()


# ─────────────────────────────────────────────
# CÓDIGO ROTATIVO FIRMADO
# ─────────────────────────────────────────────
def _secreto() -> bytes:
    """Secreto del servidor. Sin él no se puede fabricar un código válido."""
    s = st.secrets.get("secreto_asistencia", "")
    if not s:
        raise RuntimeError(
            "Falta `secreto_asistencia` en secrets. Genera una cadena larga "
            "y aleatoria; de ella depende que el código no se pueda falsificar.")
    return str(s).encode()


def ventana_actual(momento=None) -> int:
    """Número de ventana de 30 segundos. Es lo que hace que el código caduque."""
    m = momento or _ahora()
    return int(m.timestamp()) // SEGUNDOS_VENTANA


def generar_codigo(centro: str, ventana: int = None) -> str:
    """Código del centro para una ventana de tiempo. No lleva ningún dato de
    persona: solo dice 'este centro, este momento'."""
    v = ventana if ventana is not None else ventana_actual()
    mensaje = f"{_norm(centro)}|{v}".encode()
    firma = hmac.new(_secreto(), mensaje, hashlib.sha256).hexdigest()[:16]
    return f"{v}.{firma}"


def validar_codigo(codigo: str, centro: str) -> tuple[bool, str]:
    """Comprueba firma y vigencia. Devuelve (válido, motivo del rechazo)."""
    partes = str(codigo or "").split(".")
    if len(partes) != 2 or not partes[0].isdigit():
        return False, "El código no tiene el formato esperado."
    v = int(partes[0])
    actual = ventana_actual()
    if v > actual:
        return False, "El código viene de un momento futuro."
    if actual - v > VENTANAS_TOLERADAS:
        segundos = (actual - v) * SEGUNDOS_VENTANA
        return False, (f"El código caducó hace {segundos} segundos. "
                       "Vuelve a escanear el de la pantalla.")
    # compare_digest evita filtrar información por el tiempo de comparación
    if not hmac.compare_digest(generar_codigo(centro, v), f"{v}.{partes[1]}"):
        return False, "El código no corresponde a este centro."
    return True, ""


# ─────────────────────────────────────────────
# PERSONAL DEL CENTRO (desde el PADRÓN)
# ─────────────────────────────────────────────
@st.cache_data(ttl=300)
def _padron_cm() -> pd.DataFrame:
    """Personal adscrito a Centros de Maestros, con su horario.

    Sale del Padrón Maestro y no de una tab propia: mantener otra lista de
    personal sería volver a duplicar lo que el padrón vino a unificar.
    """
    try:
        pad = _deps["cargar_padron"]()
    except Exception:
        return pd.DataFrame()
    if pad.empty or "ADSCRIPCION_REAL" not in pad.columns:
        return pd.DataFrame()
    m = pad["ADSCRIPCION_REAL"].astype(str).str.upper().str.startswith("CENTRO DE MAESTROS")
    return pad[m].copy()


def centros_disponibles() -> list:
    df = _padron_cm()
    if df.empty:
        return []
    return sorted(df["ADSCRIPCION_REAL"].astype(str).unique())


def asesores_de(centro: str) -> pd.DataFrame:
    df = _padron_cm()
    if df.empty:
        return df
    return df[df["ADSCRIPCION_REAL"].astype(str).str.strip() == str(centro).strip()]


def centro_del_responsable(rfc: str) -> str:
    """Centro que coordina esta persona, si es Responsable de alguno."""
    df = _padron_cm()
    if df.empty or not rfc:
        return ""
    hit = df[(df["RFC"].astype(str).str.upper() == str(rfc).upper())
             & (df["ROL"].astype(str).str.upper() == "RESPONSABLE")]
    return str(hit.iloc[0]["ADSCRIPCION_REAL"]) if not hit.empty else ""


def horario_del_dia(fila, momento=None) -> tuple:
    """(entrada, salida) que le toca a esa persona ese día. Los asesores tienen
    horarios distintos entre sí, así que la puntualidad se juzga contra el
    suyo, nunca contra una hora fija."""
    m = momento or _ahora()
    if m.weekday() > 4:
        return "", ""
    d = DIAS[m.weekday()]
    return (str(fila.get(f"ENTRADA_{d}", "") or "").strip(),
            str(fila.get(f"SALIDA_{d}", "") or "").strip())


def _a_minutos(hhmm: str):
    p = str(hhmm or "").strip().split(":")
    if len(p) < 2 or not p[0].strip().isdigit():
        return None
    try:
        return int(p[0]) * 60 + int(p[1])
    except ValueError:
        return None


def evaluar_puntualidad(hora_real: str, hora_esperada: str) -> str:
    esp, real = _a_minutos(hora_esperada), _a_minutos(hora_real)
    if esp is None or real is None:
        return "SIN HORARIO"
    diff = real - esp
    if diff <= TOLERANCIA_MIN:
        return "EN TIEMPO"
    return f"RETARDO {diff} MIN"


# ─────────────────────────────────────────────
# REGISTRO
# ─────────────────────────────────────────────
def _ws():
    import gspread
    sh = _client().open_by_key(_deps["sheet_id"])
    try:
        return sh.worksheet(TAB_ASISTENCIA)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(TAB_ASISTENCIA, rows=2, cols=len(COLUMNAS_ASISTENCIA))
        ws.append_row(COLUMNAS_ASISTENCIA)
        return ws


@st.cache_data(ttl=60)
def _cargar_asistencia(centro: str, dias: int = 30) -> pd.DataFrame:
    try:
        datos = _ws().get_all_records(numericise_ignore=["all"])
    except Exception:
        return pd.DataFrame(columns=COLUMNAS_ASISTENCIA)
    if not datos:
        return pd.DataFrame(columns=COLUMNAS_ASISTENCIA)
    df = pd.DataFrame(datos)
    if centro and "CENTRO" in df.columns:
        df = df[df["CENTRO"].astype(str).str.strip() == str(centro).strip()]
    if dias and "FECHA" in df.columns:
        corte = (_ahora() - timedelta(days=dias)).strftime("%Y-%m-%d")
        df = df[df["FECHA"].astype(str) >= corte]
    return df


def registrar(centro: str, rfc: str, nombre: str, metodo: str,
              registrado_por: str = "", motivo: str = "") -> tuple[bool, str]:
    """Marca entrada o salida. El sistema decide cuál según lo ya registrado
    hoy: un solo código sirve para ambas y no hay forma de equivocarse.

    La hora la pone el servidor. TS_* guarda además el instante real de
    escritura: si alguien registra el lunes a las 11:40, eso queda aunque la
    hora declarada sea otra.
    """
    ahora = _ahora()
    hoy = ahora.strftime("%Y-%m-%d")
    hora = ahora.strftime("%H:%M")
    sello = ahora.strftime("%Y-%m-%d %H:%M:%S")
    rfc = str(rfc).upper().strip()

    try:
        ws = _ws()
        filas = ws.get_all_records(numericise_ignore=["all"])
    except Exception as e:
        return False, f"No se pudo leer la asistencia: {e}"

    encabezados = COLUMNAS_ASISTENCIA
    fila_num = None
    for i, r in enumerate(filas, start=2):
        if (str(r.get("FECHA", "")).strip() == hoy
                and str(r.get("RFC", "")).strip().upper() == rfc):
            fila_num = i
            existente = r
            break

    try:
        if fila_num is None:
            ws.append_row([hoy, centro, rfc, nombre, hora, "", metodo, "",
                           registrado_por, motivo, sello, ""],
                          value_input_option="USER_ENTERED")
            _cargar_asistencia.clear()
            return True, f"Entrada registrada a las {hora}"

        if str(existente.get("HORA_SALIDA", "")).strip():
            return False, (f"Ya tienes entrada ({existente.get('HORA_ENTRADA')}) "
                           f"y salida ({existente.get('HORA_SALIDA')}) de hoy.")

        from gspread.cell import Cell
        celdas = [
            Cell(fila_num, encabezados.index("HORA_SALIDA") + 1, hora),
            Cell(fila_num, encabezados.index("METODO_SALIDA") + 1, metodo),
            Cell(fila_num, encabezados.index("TS_SALIDA") + 1, sello),
        ]
        ws.update_cells(celdas, value_input_option="USER_ENTERED")
        _cargar_asistencia.clear()
        return True, f"Salida registrada a las {hora}"
    except Exception as e:
        return False, f"No se pudo guardar: {e}"


# ─────────────────────────────────────────────
# VISTA 1 · PANTALLA DEL CENTRO (coordinador)
# ─────────────────────────────────────────────
def vista_pantalla(url_app: str):
    """Se proyecta en una pantalla del centro. Muestra el QR y lo renueva."""
    rfc = str(st.session_state.get("rfc", "")).upper().strip()
    centro = centro_del_responsable(rfc)
    if st.session_state.get("rol") == "admin" and not centro:
        opciones = centros_disponibles()
        if not opciones:
            st.info("El padrón no tiene personal de Centros de Maestros.")
            return
        centro = st.selectbox("Centro (vista admin)", opciones, key="aqr_centro_p")
    if not centro:
        st.error("Solo el Responsable del centro puede mostrar esta pantalla.")
        return

    try:
        import qrcode
        from io import BytesIO
        codigo = generar_codigo(centro)
    except Exception as e:
        _error(e, "al generar el código")
        return

    st.markdown(f"### {centro}")
    st.caption("Deja esta pantalla encendida. El código se renueva solo.")

    qr = qrcode.QRCode(box_size=12, border=2)
    qr.add_data(f"{url_app}/?asistencia={codigo}")
    qr.make(fit=True)
    buf = BytesIO()
    qr.make_image(fill_color="black", back_color="white").save(buf, format="PNG")

    c1, c2 = st.columns([2, 1])
    c1.image(buf.getvalue(), use_container_width=True)
    c2.metric("Renovación", f"{SEGUNDOS_VENTANA} s")
    c2.caption(f"Actualizado {_ahora().strftime('%H:%M:%S')}")
    if c2.button("🔄 Renovar ahora", use_container_width=True, key="aqr_renovar"):
        st.rerun()
    c2.caption("Coloca la pantalla donde no se vea desde fuera del centro: "
               "un código visible desde la puerta se puede escanear sin entrar.")

    hoy = _ahora().strftime("%Y-%m-%d")
    df = _cargar_asistencia(centro, dias=1)
    if not df.empty:
        df = df[df["FECHA"].astype(str) == hoy]
    st.divider()
    total = len(asesores_de(centro))
    st.metric("Registrados hoy", f"{len(df)} de {total}")
    if not df.empty:
        st.dataframe(df[["NOMBRE", "HORA_ENTRADA", "HORA_SALIDA", "METODO_ENTRADA"]],
                     use_container_width=True, hide_index=True)


# ─────────────────────────────────────────────
# VISTA 2 · EL ASESOR ESCANEA
# ─────────────────────────────────────────────
def procesar_escaneo(codigo: str):
    """Se invoca cuando llega ?asistencia=<codigo>. El código dice el centro y
    el momento; la sesión dice quién es. Esa separación es la que impide
    registrar a un compañero ausente."""
    st.markdown("### 📍 Registro de asistencia")

    rfc = str(st.session_state.get("rfc", "")).upper().strip()
    if not rfc:
        st.warning("Inicia sesión para registrar tu asistencia y vuelve a "
                   "escanear el código.")
        return

    cm = _padron_cm()
    yo = cm[cm["RFC"].astype(str).str.upper() == rfc] if not cm.empty else pd.DataFrame()
    if yo.empty:
        st.error("No estás registrado como personal de un Centro de Maestros.")
        return
    yo = yo.iloc[0]
    centro = str(yo["ADSCRIPCION_REAL"])

    ok, motivo = validar_codigo(codigo, centro)
    if not ok:
        st.error(motivo)
        return

    exito, mensaje = registrar(centro, rfc, str(yo["NOMBRE"]), "QR")
    if not exito:
        st.warning(mensaje)
        return

    ahora = _ahora()
    st.success(f"✅ {mensaje}")
    st.markdown(f"**{yo['NOMBRE']}** · {centro}")
    st.markdown(f"### {ahora.strftime('%H:%M')}")
    st.caption(ahora.strftime("%A %d de %B de %Y"))

    entrada, _ = horario_del_dia(yo, ahora)
    if entrada and mensaje.startswith("Entrada"):
        estado = evaluar_puntualidad(ahora.strftime("%H:%M"), entrada)
        (st.success if estado == "EN TIEMPO" else st.warning)(
            f"Tu horario de hoy inicia a las {entrada} · **{estado}**")

    st.info("Anota esta misma hora en la lista de asistencia física.")


# ─────────────────────────────────────────────
# VISTA 3 · REGISTRO ASISTIDO Y CONSULTA
# ─────────────────────────────────────────────
def vista_coordinador():
    rfc = str(st.session_state.get("rfc", "")).upper().strip()
    es_admin = st.session_state.get("rol") == "admin"
    centro = centro_del_responsable(rfc)
    if es_admin and not centro:
        opciones = centros_disponibles()
        if not opciones:
            st.info("El padrón no tiene personal de Centros de Maestros.")
            return
        centro = st.selectbox("Centro (vista admin)", opciones, key="aqr_centro_c")
    if not centro:
        st.error("Solo el Responsable del centro puede entrar aquí.")
        return

    st.markdown(f"## {centro}")
    t1, t2 = st.tabs(["🤝 Registro asistido", "📋 Consultar"])

    with t1:
        st.caption("Para quien no puede escanear: sin celular, sin batería, o "
                   "porque no maneja la aplicación. Queda constancia de quién "
                   "lo registró y por qué, así que la excepción es visible.")
        aseg = asesores_de(centro)
        if aseg.empty:
            st.warning("No hay personal registrado para este centro en el padrón.")
        else:
            nombres = aseg["NOMBRE"].astype(str).tolist()
            sel = st.selectbox("Persona", nombres, key="aqr_persona")
            razon = st.text_input("Motivo", key="aqr_motivo",
                                  placeholder="No trae celular, batería agotada...")
            if st.button("Registrar", type="primary", key="aqr_btn_asistido"):
                if not razon.strip():
                    st.warning("El motivo es obligatorio: es lo que distingue "
                               "una excepción de un hueco.")
                else:
                    fila = aseg[aseg["NOMBRE"].astype(str) == sel].iloc[0]
                    ok, msg = registrar(centro, fila["RFC"], sel, "ASISTIDO",
                                        st.session_state.get("nombre", rfc),
                                        razon.strip())
                    (st.success if ok else st.warning)(msg)
                    if ok:
                        st.rerun()

    with t2:
        df = _cargar_asistencia(centro, dias=45)
        if df.empty:
            st.info("Sin registros todavía.")
            return
        fechas = sorted(df["FECHA"].astype(str).unique(), reverse=True)
        f = st.selectbox("Fecha", fechas, key="aqr_fecha")
        dia = df[df["FECHA"].astype(str) == f]

        total = len(asesores_de(centro))
        asistidos = int((dia["METODO_ENTRADA"].astype(str) == "ASISTIDO").sum())
        m1, m2, m3 = st.columns(3)
        m1.metric("Registrados", f"{len(dia)} de {total}")
        m2.metric("Por QR", len(dia) - asistidos)
        m3.metric("Asistidos", asistidos)

        st.dataframe(
            dia[["NOMBRE", "HORA_ENTRADA", "HORA_SALIDA", "METODO_ENTRADA",
                 "MOTIVO", "REGISTRADO_POR"]],
            use_container_width=True, hide_index=True)

        # Quién no aparece ese día: es el dato que la lista de papel no da.
        presentes = set(dia["RFC"].astype(str).str.upper())
        faltantes = asesores_de(centro)
        faltantes = faltantes[~faltantes["RFC"].astype(str).str.upper().isin(presentes)]
        if not faltantes.empty:
            with st.expander(f"⚠️ {len(faltantes)} sin registro ese día"):
                for _, r in faltantes.iterrows():
                    st.write(f"• {r['NOMBRE']}")

        if es_admin and asistidos:
            st.caption("Un centro donde el registro asistido es frecuente "
                       "merece revisión: puede haber un problema de pantalla, "
                       "de señal, o de uso.")
