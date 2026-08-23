"""
oficios_module.py — Minutario interno de oficios de RH · DFC

Control paralelo (espejo) de los números de oficio que RH solicita al
minutario de la Dirección. El consecutivo NO lo asigna este sistema:
el folio se captura manualmente porque la fuente de verdad es el Sheet
que controla la Dirección. Este módulo registra lo que pasa por RH,
genera el QR de trazabilidad y resguarda el escaneo en Drive.

Patrón de integración idéntico a checador_module.render_checador(deps).

Secrets requeridos:
    sheet_oficios_id      = "..."            # Sheet nuevo, tabs: oficios, log_oficios
    drive_oficios_folder  = "..."            # Carpeta en Unidad Compartida
    rfcs_oficios          = ["RFC1", "..."]  # Quiénes pueden emitir
"""

import streamlit as st
import pandas as pd
import hashlib
from io import BytesIO
from datetime import datetime

import pytz
import qrcode

# ─────────────────────────────────────────────
# CONSTANTES
# ─────────────────────────────────────────────
TAB_OFICIOS = "oficios"
TAB_LOG = "log_oficios"

COLUMNAS_OFICIOS = [
    "ID_OFICIO", "AÑO", "FOLIO", "FECHA_SOLICITUD", "FECHA_OFICIO",
    "EMISOR_RFC", "EMISOR_NOMBRE", "ASUNTO", "DIRIGIDO_A", "CARGO_DESTINO",
    "ESTADO", "URL", "SHA256", "FECHA_ESCANEO", "OBSERVACIONES",
]

COLUMNAS_LOG = [
    "TIMESTAMP", "ID_OFICIO", "ACCION", "RFC", "NOMBRE", "DETALLE",
]

ESTADOS = ["RESERVADO", "EMITIDO", "ESCANEADO", "ACUSE", "CANCELADO"]

URL_APP = "https://gestion-personal-dfc.streamlit.app"

TZ = pytz.timezone("America/Mexico_City")


def _ahora() -> str:
    return datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S")


def _hoy() -> str:
    return datetime.now(TZ).strftime("%Y-%m-%d")


def construir_id(anio, folio: str) -> str:
    """ID canónico del oficio. El folio viene del minutario de la Dirección,
    por eso se normaliza a 4 dígitos sin inventar consecutivos."""
    folio_norm = str(folio).strip().zfill(4)
    return f"DFC-{anio}-{folio_norm}"


# ─────────────────────────────────────────────
# ACCESO A DATOS
# ─────────────────────────────────────────────
def _abrir_sheet(get_client):
    client = get_client()
    return client.open_by_key(st.secrets["sheet_oficios_id"])


@st.cache_data(ttl=300)
def _cargar_oficios_cached(_get_client) -> pd.DataFrame:
    """numericise_ignore=['all'] es obligatorio: los folios con ceros a la
    izquierda se corromperían si gspread los convierte a número."""
    sh = _abrir_sheet(_get_client)
    ws = sh.worksheet(TAB_OFICIOS)
    registros = ws.get_all_records(numericise_ignore=["all"])
    if not registros:
        return pd.DataFrame(columns=COLUMNAS_OFICIOS)
    return pd.DataFrame(registros)


def cargar_oficios(get_client) -> pd.DataFrame:
    return _cargar_oficios_cached(get_client)


def _registrar_log(get_client, id_oficio: str, accion: str, detalle: str = ""):
    """Bitácora append-only. Es lo que respalda a Ángel si alguien afirma
    que no solicitó un folio."""
    try:
        sh = _abrir_sheet(get_client)
        ws = sh.worksheet(TAB_LOG)
        ws.append_row([
            _ahora(),
            id_oficio,
            accion,
            str(st.session_state.get("rfc", "")).upper(),
            st.session_state.get("nombre", ""),
            detalle,
        ], value_input_option="USER_ENTERED")
    except Exception:
        # La bitácora nunca debe tumbar la operación principal.
        pass


def _error_amable(e: Exception, contexto: str = ""):
    if "429" in str(e) or "uota" in str(e):
        st.error("⏳ El sistema está ocupado. Espera 15 segundos y reintenta. "
                 "Tu información no se perdió.")
    else:
        st.error(f"Error {contexto}: {e}")


# ─────────────────────────────────────────────
# QR
# ─────────────────────────────────────────────
def generar_qr_png(id_oficio: str) -> bytes:
    """QR con corrección de error alta (H): sobrevive sello, engrapado y dobleces."""
    url = f"{URL_APP}/?validar_oficio={id_oficio}"
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_H,
        box_size=10,
        border=2,
    )
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    buf = BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────
# MOTOR PDF: estampado y lectura
# ─────────────────────────────────────────────
CM = 28.3465  # 1 cm en puntos PDF

# Import protegido: si falta la librería, el resto del módulo sigue vivo.
try:
    import pymupdf
    _ERROR_PDF = ""
except Exception as _e_pdf:  # pragma: no cover
    pymupdf = None
    _ERROR_PDF = str(_e_pdf)

try:
    import zxingcpp
    _ERROR_ZX = ""
except Exception as _e_zx:  # pragma: no cover
    zxingcpp = None
    _ERROR_ZX = str(_e_zx)


def estampar_qr_pdf(pdf_bytes: bytes, id_oficio: str,
                    lado_cm: float = 2.3, margen_cm: float = 1.0) -> bytes:
    """Inserta el QR en la esquina superior derecha de la PRIMERA página y
    escribe el ID en claro debajo, por si el QR se destruye al engrapar.
    Devuelve el PDF listo para imprimir y pasar a firma."""
    if pymupdf is None:
        raise RuntimeError(f"PyMuPDF no disponible: {_ERROR_PDF}")

    png = generar_qr_png(id_oficio)
    doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    try:
        pagina = doc[0]
        lado, margen = lado_cm * CM, margen_cm * CM
        x_der = pagina.rect.width - margen
        rect = pymupdf.Rect(x_der - lado, margen, x_der, margen + lado)
        pagina.insert_image(rect, stream=png)
        pagina.insert_text((rect.x0, rect.y1 + 7), id_oficio, fontsize=6)
        return doc.tobytes()
    finally:
        doc.close()


def extraer_texto_pdf(pdf_bytes: bytes, max_paginas: int = 2) -> str:
    """Texto de las primeras páginas. Los oficios salen de Word, así que
    traen capa de texto: no hace falta OCR."""
    if pymupdf is None:
        return ""
    doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    try:
        partes = [doc[i].get_text() for i in range(min(max_paginas, doc.page_count))]
        return "\n".join(partes).strip()
    finally:
        doc.close()


def leer_qr_pdf(archivo_bytes: bytes, nombre: str, dpi: int = 200) -> list[tuple[int, str]]:
    """Devuelve [(num_pagina, ID_OFICIO), ...] de cada página donde haya un QR
    del minutario. Es lo que permite procesar un escaneo en lote sin separadores."""
    if zxingcpp is None:
        raise RuntimeError(f"zxing-cpp no disponible: {_ERROR_ZX}")

    from PIL import Image

    def _ids_de_imagen(img, num_pag):
        hallados = []
        for res in zxingcpp.read_barcodes(img):
            texto = str(res.text)
            if "validar_oficio=" in texto:
                hallados.append((num_pag, texto.split("validar_oficio=")[-1].strip()))
        return hallados

    encontrados = []
    if nombre.lower().endswith(".pdf"):
        if pymupdf is None:
            raise RuntimeError(f"PyMuPDF no disponible: {_ERROR_PDF}")
        doc = pymupdf.open(stream=archivo_bytes, filetype="pdf")
        try:
            for i in range(doc.page_count):
                pix = doc[i].get_pixmap(dpi=dpi)
                img = Image.open(BytesIO(pix.tobytes("png")))
                encontrados.extend(_ids_de_imagen(img, i))
        finally:
            doc.close()
    else:
        encontrados.extend(_ids_de_imagen(Image.open(BytesIO(archivo_bytes)), 0))
    return encontrados


def partir_pdf_por_qr(pdf_bytes: bytes, marcas: list[tuple[int, str]]) -> dict[str, bytes]:
    """Corta un PDF de lote en documentos individuales. Cada página con QR
    inicia un oficio nuevo; las siguientes sin QR se le anexan."""
    if pymupdf is None:
        raise RuntimeError(f"PyMuPDF no disponible: {_ERROR_PDF}")

    doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    try:
        inicios = sorted(marcas, key=lambda m: m[0])
        salida = {}
        for idx, (pag_ini, id_of) in enumerate(inicios):
            pag_fin = inicios[idx + 1][0] - 1 if idx + 1 < len(inicios) else doc.page_count - 1
            nuevo = pymupdf.open()
            nuevo.insert_pdf(doc, from_page=pag_ini, to_page=pag_fin)
            salida[id_of] = nuevo.tobytes()
            nuevo.close()
        return salida
    finally:
        doc.close()


# ─────────────────────────────────────────────
# EXTRACCIÓN DE DATOS DEL OFICIO
#
# Método por defecto: expresiones regulares sobre la estructura fija del
# oficio institucional. Es determinista, auditable y NO envía el contenido
# del oficio a ningún servicio externo. La IA queda como opción apagada.
# ─────────────────────────────────────────────
import re

CAMPOS_EXTRAIBLES = ["folio", "fecha_oficio", "asunto", "dirigido_a", "cargo_destino"]

MESES = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
    "julio": 7, "agosto": 8, "septiembre": 9, "setiembre": 9, "octubre": 10,
    "noviembre": 11, "diciembre": 12,
}

# Tratamientos que anteceden al nombre del destinatario en un oficio oficial.
TRATAMIENTOS = (
    r"(?:LIC|MTRO|MTRA|ING|DR|DRA|PROF|PROFR|PROFRA|C|CP|LIC\.?A|ARQ|"
    r"MAESTRO|MAESTRA|LICENCIADO|LICENCIADA|DOCTOR|DOCTORA)"
)

_RE_FOLIO = re.compile(
    r"\b[A-ZÁÉÍÓÚÑ]{2,8}(?:\s*/\s*[A-ZÁÉÍÓÚÑ0-9]{1,10})*\s*/\s*(\d{1,6})\s*/\s*(\d{4})\b"
)
_RE_FOLIO_ALT = re.compile(
    r"oficio\s*(?:n[uú]m\.?(?:ero)?|n[oº°]\.?|#)?\s*:?\s*(\d{1,6})\s*/\s*(\d{4})",
    re.IGNORECASE,
)
_RE_FECHA_TEXTO = re.compile(
    r"\b(\d{1,2})\s+de\s+([a-záéíóú]+)\s+(?:de|del)\s+(\d{4})", re.IGNORECASE
)
_RE_FECHA_ISO = re.compile(r"\b(\d{4})-(\d{2})-(\d{2})\b")
_RE_ASUNTO = re.compile(
    r"asunto\s*:?\s*(.+?)(?:\n\s*\n|\n(?=\s*(?:" + TRATAMIENTOS + r")\b)|$)",
    re.IGNORECASE | re.DOTALL,
)
_RE_PRESENTE = re.compile(r"^\s*P\s*R\s*E\s*S\s*E\s*N\s*T\s*E\s*\.?\s*$", re.IGNORECASE)
_RE_DESTINATARIO = re.compile(
    r"^\s*(" + TRATAMIENTOS + r")\.?\s+([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ\s\.]{4,60})\s*$"
)


def _limpiar(texto: str) -> str:
    return re.sub(r"\s+", " ", str(texto or "")).strip(" .:,;-")


def extraer_datos_regex(texto: str) -> dict:
    """Lee los campos del oficio con reglas explícitas. Todo campo que no
    encuentre queda vacío: nunca adivina."""
    datos = {c: "" for c in CAMPOS_EXTRAIBLES}
    if not texto or not texto.strip():
        return datos

    lineas = [l.rstrip() for l in texto.splitlines()]

    # ── Folio ──
    m = _RE_FOLIO.search(texto) or _RE_FOLIO_ALT.search(texto)
    if m:
        datos["folio"] = m.group(1).zfill(4)

    # ── Fecha ──
    m = _RE_FECHA_TEXTO.search(texto)
    if m:
        mes = MESES.get(m.group(2).lower())
        if mes:
            datos["fecha_oficio"] = f"{m.group(3)}-{mes:02d}-{int(m.group(1)):02d}"
    if not datos["fecha_oficio"]:
        m = _RE_FECHA_ISO.search(texto)
        if m:
            datos["fecha_oficio"] = m.group(0)

    # ── Asunto ──
    m = _RE_ASUNTO.search(texto)
    if m:
        datos["asunto"] = _limpiar(m.group(1))[:250]

    # ── Destinatario y cargo ──
    # El cargo va entre el nombre y la línea "P R E S E N T E".
    idx_presente = next(
        (i for i, l in enumerate(lineas) if _RE_PRESENTE.match(l)), None
    )
    idx_nombre = None
    tope = idx_presente if idx_presente is not None else len(lineas)
    for i in range(tope - 1, -1, -1):
        mm = _RE_DESTINATARIO.match(lineas[i])
        if mm:
            datos["dirigido_a"] = _limpiar(f"{mm.group(1)}. {mm.group(2)}")
            idx_nombre = i
            break

    if idx_nombre is not None and idx_presente is not None:
        cargo = [_limpiar(l) for l in lineas[idx_nombre + 1:idx_presente]]
        datos["cargo_destino"] = " ".join(c for c in cargo if c)[:200]

    return datos


def extraer_datos_ia(texto: str) -> dict:
    """OPCIONAL y APAGADO por defecto. Solo corre si en secrets existe
    usar_ia_oficios = true. Enviaría el texto del oficio a un servicio
    externo, por eso requiere activación explícita."""
    vacio = {c: "" for c in CAMPOS_EXTRAIBLES}
    if not st.secrets.get("usar_ia_oficios", False):
        return vacio
    if not texto.strip() or "GEMINI_API_KEY" not in st.secrets:
        return vacio
    try:
        import json
        import google.generativeai as genai

        prompt = (
            "Extrae los datos de este oficio institucional mexicano.\n"
            "Responde SOLO con JSON, sin markdown:\n"
            '{"folio":"","fecha_oficio":"","asunto":"","dirigido_a":"","cargo_destino":""}\n'
            '- "folio": solo dígitos (de "DFC/RH/0100/2026" extrae "0100").\n'
            '- "fecha_oficio": AAAA-MM-DD.\n'
            "- Si un dato no aparece, cadena vacía. No inventes.\n\nTEXTO:\n"
        )
        genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
        modelo = genai.GenerativeModel("gemini-2.5-flash-lite")
        crudo = str(modelo.generate_content(prompt + texto[:6000]).text).strip()
        if "```" in crudo:
            crudo = crudo.split("```")[1].replace("json", "", 1).strip()
        d = json.loads(crudo)
        return {k: str(d.get(k, "") or "") for k in vacio}
    except Exception:
        return vacio


def extraer_datos_oficio(texto: str) -> tuple[dict, str]:
    """Punto de entrada único. Devuelve (datos, método usado).
    La IA solo complementa campos que las reglas dejaron vacíos, y solo
    si está habilitada explícitamente en secrets."""
    datos = extraer_datos_regex(texto)
    metodo = "reglas"

    faltantes = [c for c in CAMPOS_EXTRAIBLES if not datos[c]]
    if faltantes and st.secrets.get("usar_ia_oficios", False):
        sugeridos = extraer_datos_ia(texto)
        completados = [c for c in faltantes if sugeridos.get(c)]
        if completados:
            for c in completados:
                datos[c] = sugeridos[c]
            metodo = "reglas + IA"
    return datos, metodo


def subir_bytes_drive(contenido: bytes, nombre_archivo: str, mimetype: str) -> str:
    """Sube contenido en memoria a la carpeta de oficios. Sin permiso 'anyone':
    el acceso se hereda de la membresía de la Unidad Compartida."""
    try:
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaIoBaseUpload
        from google.oauth2.service_account import Credentials
        import io

        creds = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=["https://www.googleapis.com/auth/drive"],
        )
        service = build("drive", "v3", credentials=creds)
        media = MediaIoBaseUpload(io.BytesIO(contenido), mimetype=mimetype)
        meta = {"name": nombre_archivo,
                "parents": [st.secrets["drive_oficios_folder"]]}
        creado = service.files().create(
            body=meta, media_body=media, fields="id", supportsAllDrives=True
        ).execute()
        return f"https://drive.google.com/file/d/{creado.get('id')}/view"
    except Exception as e:
        return f"ERROR: {e}"


def subir_oficio_drive(archivo, id_oficio: str) -> str:
    """Sube un archivo cargado por la persona usuaria."""
    ext = archivo.name.split(".")[-1].lower()
    return subir_bytes_drive(archivo.getvalue(), f"{id_oficio}.{ext}",
                             archivo.type or "application/octet-stream")


def sha256_bytes(contenido: bytes) -> str:
    """Huella del archivo. Si alguien lo sustituye en Drive, deja de coincidir."""
    return hashlib.sha256(contenido).hexdigest()


def reservar_oficio(get_client, datos: dict) -> tuple[bool, str]:
    """Alta de un folio ya asignado por el minutario de la Dirección.
    Rechaza duplicados: es el único detector de que alguien copió mal
    el número del minutario."""
    id_oficio = construir_id(datos["anio"], datos["folio"])

    df = cargar_oficios(get_client)
    if not df.empty and "ID_OFICIO" in df.columns:
        if id_oficio in df["ID_OFICIO"].astype(str).values:
            fila = df[df["ID_OFICIO"].astype(str) == id_oficio].iloc[0]
            return False, (
                f"El folio {id_oficio} ya está registrado por "
                f"{fila.get('EMISOR_NOMBRE', '?')} el {fila.get('FECHA_SOLICITUD', '?')} "
                f"— asunto: {fila.get('ASUNTO', '?')}"
            )

    try:
        sh = _abrir_sheet(get_client)
        ws = sh.worksheet(TAB_OFICIOS)
        ws.append_row([
            id_oficio,
            str(datos["anio"]),
            str(datos["folio"]).strip().zfill(4),
            _hoy(),
            datos.get("fecha_oficio", ""),
            str(st.session_state.get("rfc", "")).upper(),
            st.session_state.get("nombre", ""),
            datos.get("asunto", ""),
            datos.get("dirigido_a", ""),
            datos.get("cargo_destino", ""),
            "RESERVADO",
            "", "", "",
            datos.get("observaciones", ""),
        ], value_input_option="USER_ENTERED")
        st.cache_data.clear()
        _registrar_log(get_client, id_oficio, "RESERVADO", datos.get("asunto", ""))
        return True, id_oficio
    except Exception as e:
        return False, f"Error al guardar: {e}"


def _actualizar_fila(get_client, id_oficio: str, cambios: dict) -> bool:
    """Actualiza celdas por nombre de columna, sin asumir el orden del Sheet."""
    from gspread.cell import Cell

    sh = _abrir_sheet(get_client)
    ws = sh.worksheet(TAB_OFICIOS)
    headers = ws.row_values(1)
    registros = ws.get_all_records(numericise_ignore=["all"])

    for i, row in enumerate(registros, start=2):
        if str(row.get("ID_OFICIO", "")) == id_oficio:
            celdas = []
            for col, valor in cambios.items():
                if col in headers:
                    celdas.append(Cell(i, headers.index(col) + 1, valor))
            if celdas:
                ws.update_cells(celdas, value_input_option="USER_ENTERED")
            st.cache_data.clear()
            return True
    return False


def cancelar_oficio(get_client, id_oficio: str, motivo: str) -> bool:
    """Un folio pedido y no usado deja un hueco en el consecutivo de la
    Dirección. Cancelarlo aquí le pone dueño y motivo a ese hueco."""
    ok = _actualizar_fila(get_client, id_oficio, {
        "ESTADO": "CANCELADO",
        "OBSERVACIONES": motivo,
    })
    if ok:
        _registrar_log(get_client, id_oficio, "CANCELADO", motivo)
    return ok


# ─────────────────────────────────────────────
# ANÁLISIS
# ─────────────────────────────────────────────
def detectar_huecos(df: pd.DataFrame, anio: str) -> list[int]:
    """Folios faltantes dentro del rango que RH ha manejado este año.
    No implica error: pueden ser oficios de otras áreas de la DFC."""
    if df.empty or "AÑO" not in df.columns:
        return []
    del_anio = df[df["AÑO"].astype(str) == str(anio)]
    if del_anio.empty:
        return []
    nums = sorted(
        int(f) for f in del_anio["FOLIO"].astype(str)
        if str(f).strip().isdigit()
    )
    if not nums:
        return []
    return [n for n in range(min(nums), max(nums) + 1) if n not in set(nums)]


# ─────────────────────────────────────────────
# VALIDACIÓN PÚBLICA (QR)
# ─────────────────────────────────────────────
def render_validacion_oficio(get_client, id_oficio: str):
    """Vista que abre el QR. Muestra solo lo mínimo para confirmar
    procedencia: sin nombres de destinatarios ni contenido del oficio."""
    st.markdown("### 🔎 Verificación de oficio · RH · DFC")
    try:
        df = cargar_oficios(get_client)
    except Exception as e:
        _error_amable(e, "al consultar el minutario")
        return

    fila = pd.DataFrame()
    if not df.empty and "ID_OFICIO" in df.columns:
        fila = df[df["ID_OFICIO"].astype(str).str.strip() == str(id_oficio).strip()]

    if fila.empty:
        st.error(f"El identificador **{id_oficio}** no corresponde a ningún "
                 "oficio emitido por Recursos Humanos de la DFC.")
        return

    r = fila.iloc[0]
    if str(r.get("ESTADO", "")) == "CANCELADO":
        st.warning(f"El oficio **{id_oficio}** fue **cancelado** y no debe surtir efectos.")
        return

    st.success(f"Oficio **{id_oficio}** emitido por Recursos Humanos de la DFC.")
    st.write(f"**Fecha del oficio:** {r.get('FECHA_OFICIO') or r.get('FECHA_SOLICITUD', '—')}")
    st.write(f"**Estado:** {r.get('ESTADO', '—')}")
    st.caption("Esta verificación confirma únicamente la procedencia del folio. "
               "No acredita el contenido del documento.")


# ─────────────────────────────────────────────
# VISTA PRINCIPAL
# ─────────────────────────────────────────────
def render_oficios(deps: dict):
    get_client = deps["get_client"]
    rfcs_autorizados = deps.get("rfcs_autorizados", [])

    rfc_actual = str(st.session_state.get("rfc", "")).upper().strip()
    es_admin = st.session_state.get("rol") == "admin"
    if not es_admin and rfc_actual not in rfcs_autorizados:
        st.error("No tienes acceso al minutario de oficios.")
        return

    st.title("📄 Minutario de oficios · RH")
    st.caption("Registro interno de RH. El consecutivo oficial lo asigna la "
               "Dirección: aquí se captura el folio que ya te asignaron.")

    try:
        df = cargar_oficios(get_client)
    except Exception as e:
        _error_amable(e, "al cargar el minutario")
        return

    tab_reg, tab_esc, tab_cons = st.tabs(
        ["📤 Emitir oficio", "📎 Escaneo firmado", "📋 Consultar"]
    )

    # ── Emitir oficio ────────────────────────
    with tab_reg:
        if pymupdf is None:
            st.error(f"Falta PyMuPDF: {_ERROR_PDF}")
        else:
            st.markdown("**1.** Sube el oficio en PDF · **2.** Confirma los datos · "
                        "**3.** Descarga el PDF con QR y pásalo a firma")

            pdf_in = st.file_uploader("Oficio en PDF (sin firmar)", type=["pdf"],
                                      key="of_pdf_in")

            if pdf_in is not None:
                pdf_bytes = pdf_in.getvalue()

                huella = hashlib.md5(pdf_bytes).hexdigest()
                if st.session_state.get("_of_huella") != huella:
                    with st.spinner("Leyendo el oficio..."):
                        texto = extraer_texto_pdf(pdf_bytes)
                        sug, metodo = extraer_datos_oficio(texto)
                        st.session_state["_of_sug"] = sug
                        st.session_state["_of_metodo"] = metodo
                        st.session_state["_of_huella"] = huella
                        st.session_state["_of_hay_texto"] = bool(texto.strip())

                sug = st.session_state.get("_of_sug", {})
                metodo = st.session_state.get("_of_metodo", "reglas")
                if not st.session_state.get("_of_hay_texto", True):
                    st.warning("El PDF no trae capa de texto (parece escaneado). "
                               "Captura los datos a mano.")
                elif any(sug.values()):
                    st.info(f"Datos detectados por **{metodo}**. "
                            "Revísalos y corrígelos antes de continuar.")
                    if metodo == "reglas":
                        st.caption("El contenido del oficio se procesó localmente. "
                                   "No se envió a ningún servicio externo.")

                anio_actual = datetime.now(TZ).year
                c1, c2, c3 = st.columns([1, 1, 2])
                anio = c1.number_input("Año", min_value=2020, max_value=2100,
                                       value=anio_actual, step=1, key="of_anio")
                folio = c2.text_input("Folio del minutario",
                                      value=sug.get("folio", ""), key="of_folio",
                                      help="El número que te asignó la Dirección.")
                try:
                    f_def = datetime.strptime(sug.get("fecha_oficio", ""), "%Y-%m-%d").date()
                except Exception:
                    f_def = datetime.now(TZ).date()
                fecha_of = c3.date_input("Fecha del oficio", value=f_def, key="of_fecha")

                asunto = st.text_input("Asunto", value=sug.get("asunto", ""),
                                       key="of_asunto")
                c4, c5 = st.columns(2)
                dirigido = c4.text_input("Dirigido a", value=sug.get("dirigido_a", ""),
                                         key="of_dirigido")
                cargo = c5.text_input("Cargo del destinatario",
                                      value=sug.get("cargo_destino", ""), key="of_cargo")
                obs = st.text_area("Observaciones", key="of_obs", height=70)

                if st.button("Confirmar y estampar QR", type="primary", key="of_btn_reg"):
                    if not str(folio).strip().isdigit():
                        st.warning("El folio debe ser numérico.")
                    elif not asunto.strip() or not dirigido.strip():
                        st.warning("Asunto y destinatario son obligatorios.")
                    else:
                        ok, resultado = reservar_oficio(get_client, {
                            "anio": int(anio),
                            "folio": folio,
                            "fecha_oficio": fecha_of.strftime("%Y-%m-%d"),
                            "asunto": asunto.strip(),
                            "dirigido_a": dirigido.strip(),
                            "cargo_destino": cargo.strip(),
                            "observaciones": obs.strip(),
                        })
                        if not ok:
                            st.error(resultado)
                        else:
                            try:
                                with st.spinner("Estampando QR y resguardando..."):
                                    sellado = estampar_qr_pdf(pdf_bytes, resultado)
                                    url = subir_bytes_drive(
                                        sellado, f"{resultado}_SINFIRMA.pdf",
                                        "application/pdf")
                                    _actualizar_fila(get_client, resultado, {
                                        "ESTADO": "EMITIDO",
                                        "URL": url if not url.startswith("ERROR:") else "",
                                    })
                                    _registrar_log(get_client, resultado, "EMITIDO", url)
                                st.success(f"**{resultado}** listo. Imprime este PDF "
                                           "y pásalo a firma.")
                                st.download_button(
                                    "⬇️ Descargar PDF con QR",
                                    data=sellado,
                                    file_name=f"{resultado}_conQR.pdf",
                                    mime="application/pdf", key="of_dl_pdf",
                                )
                                if url.startswith("ERROR:"):
                                    st.warning(f"Se registró, pero falló Drive: {url}")
                            except Exception as e:
                                _error_amable(e, "al estampar el QR")

    # ── Escaneo firmado ──────────────────────
    with tab_esc:
        if zxingcpp is None:
            st.error(f"Falta zxing-cpp: {_ERROR_ZX}")
        elif df.empty:
            st.info("Aún no hay oficios emitidos.")
        else:
            st.caption("Sube el oficio ya firmado y sellado. El sistema lee el QR "
                       "y lo asocia solo. Si el PDF trae varios oficios, los separa.")
            es_acuse = st.checkbox("Son acuses sellados de recibido", key="of_acuse")
            scan = st.file_uploader("Escaneo (PDF de uno o varios oficios, o imagen)",
                                    type=["pdf", "jpg", "jpeg", "png"], key="of_file")

            if scan is not None:
                scan_bytes = scan.getvalue()
                try:
                    with st.spinner("Buscando códigos QR..."):
                        marcas = leer_qr_pdf(scan_bytes, scan.name)
                except Exception as e:
                    marcas = []
                    _error_amable(e, "al leer el QR")

                conocidos = set(df["ID_OFICIO"].astype(str))
                validas = [(p, i) for p, i in marcas if i in conocidos]
                huerfanas = [i for _, i in marcas if i not in conocidos]

                if huerfanas:
                    st.error("QR que **no** corresponden a este minutario: "
                             + ", ".join(sorted(set(huerfanas))))

                if not validas:
                    st.warning("No se detectó ningún QR del minutario. Verifica que el "
                               "escaneo sea de al menos 200 dpi y que el QR esté completo.")
                else:
                    st.success(f"Detectados: {', '.join(i for _, i in validas)}")
                    if st.button("Guardar escaneos", type="primary", key="of_btn_esc"):
                        es_pdf = scan.name.lower().endswith(".pdf")
                        piezas = (partir_pdf_por_qr(scan_bytes, validas) if es_pdf
                                  else {validas[0][1]: scan_bytes})
                        estado = "ACUSE" if es_acuse else "ESCANEADO"
                        sufijo = "ACUSE" if es_acuse else "FIRMADO"
                        mime = "application/pdf" if es_pdf else (scan.type or "image/png")
                        ext = "pdf" if es_pdf else scan.name.split(".")[-1].lower()

                        barra = st.progress(0.0)
                        for n, (id_of, contenido) in enumerate(piezas.items(), start=1):
                            url = subir_bytes_drive(
                                contenido, f"{id_of}_{sufijo}.{ext}", mime)
                            if url.startswith("ERROR:"):
                                st.error(f"{id_of}: {url}")
                            else:
                                _actualizar_fila(get_client, id_of, {
                                    "URL": url,
                                    "SHA256": sha256_bytes(contenido),
                                    "FECHA_ESCANEO": _hoy(),
                                    "ESTADO": estado,
                                })
                                _registrar_log(get_client, id_of, estado, url)
                                st.write(f"✅ {id_of} → [Drive]({url})")
                            barra.progress(n / len(piezas))
                        st.cache_data.clear()

    # ── Consultar ────────────────────────────
    with tab_cons:
        if df.empty:
            st.info("Aún no hay folios registrados.")
        else:
            anios = sorted(df["AÑO"].astype(str).unique(), reverse=True)
            anio_f = st.selectbox("Año", options=anios, key="of_anio_f")
            vista = df[df["AÑO"].astype(str) == anio_f].copy()

            if not es_admin:
                vista = vista[vista["EMISOR_RFC"].astype(str).str.upper() == rfc_actual]

            m1, m2, m3 = st.columns(3)
            m1.metric("Oficios", len(vista))
            m2.metric("Con escaneo",
                      int((vista["URL"].astype(str).str.startswith("http")).sum()))
            m3.metric("Cancelados",
                      int((vista["ESTADO"].astype(str) == "CANCELADO").sum()))

            st.dataframe(
                vista[["ID_OFICIO", "FECHA_OFICIO", "EMISOR_NOMBRE", "ASUNTO",
                       "DIRIGIDO_A", "ESTADO", "URL"]],
                use_container_width=True, hide_index=True,
            )

            if es_admin:
                huecos = detectar_huecos(df, anio_f)
                if huecos:
                    st.warning(
                        f"**Folios ausentes en el rango de RH ({anio_f}):** "
                        + ", ".join(str(h).zfill(4) for h in huecos)
                    )
                    st.caption("Pueden pertenecer a otras áreas de la DFC. "
                               "Contrástalos contra el minutario de la Dirección.")

                st.divider()
                with st.expander("Cancelar un folio"):
                    activos = vista[vista["ESTADO"].astype(str) != "CANCELADO"]
                    if activos.empty:
                        st.caption("No hay folios activos.")
                    else:
                        id_can = st.selectbox("Oficio a cancelar",
                                              options=activos["ID_OFICIO"].astype(str).tolist(),
                                              key="of_sel_can")
                        motivo = st.text_input("Motivo", key="of_motivo_can")
                        if st.button("Cancelar folio", key="of_btn_can"):
                            if not motivo.strip():
                                st.warning("El motivo es obligatorio.")
                            elif cancelar_oficio(get_client, id_can, motivo.strip()):
                                st.success(f"{id_can} cancelado.")
                                st.rerun()
                            else:
                                st.error("No se pudo cancelar.")
