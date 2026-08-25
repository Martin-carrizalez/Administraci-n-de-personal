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
import re
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
    "ESTADO", "URL_EMITIDO", "SHA256_EMITIDO", "FECHA_ESCANEO",
    "OBSERVACIONES", "TOKEN", "URL_ESCANEO", "SHA256_ESCANEO",
]

# Longitud del token opaco del QR. 12 caracteres base32 ≈ 60 bits: no es
# enumerable, que es justo el punto — con el ID visible (DFC-2026-1009)
# cualquiera podría probar 1010, 1011, 1012 y cosechar información.
LONGITUD_TOKEN = 12

COLUMNAS_LOG = [
    "TIMESTAMP", "ID_OFICIO", "ACCION", "RFC", "NOMBRE", "DETALLE",
]

ESTADOS = ["RESERVADO", "EMITIDO", "ESCANEADO", "ACUSE", "CANCELADO", "HISTORICO"]

# Oficios emitidos antes de que existiera este sistema. No llevan QR: el
# papel ya circuló y estamparlo en una copia no protegería al original.
ESTADO_HISTORICO = "HISTORICO"

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


def _faltan_columnas(df: pd.DataFrame, requeridas: list, pestana: str) -> bool:
    """Avisa qué encabezados faltan en lugar de dejar que pandas lance un
    KeyError críptico. El Sheet se edita a mano, así que un encabezado mal
    escrito o sin renombrar es la causa más probable de un fallo."""
    faltantes = [c for c in requeridas if c not in df.columns]
    if not faltantes:
        return False
    st.error(f"A la pestaña **{pestana}** del Sheet le faltan estos "
             f"encabezados: **{', '.join(faltantes)}**")
    st.caption("Agrégalos en la fila 1, respetando mayúsculas y acentos. "
               "Si renombraste columnas, revisa que coincidan exactamente.")
    if df.columns.tolist():
        st.caption("Encabezados actuales: " + ", ".join(df.columns.tolist()))
    return True


def _error_amable(e: Exception, contexto: str = ""):
    if "429" in str(e) or "uota" in str(e):
        st.error("⏳ El sistema está ocupado. Espera 15 segundos y reintenta. "
                 "Tu información no se perdió.")
    else:
        st.error(f"Error {contexto}: {e}")


# ─────────────────────────────────────────────
# QR
# ─────────────────────────────────────────────
def generar_token() -> str:
    """Token aleatorio e irrepetible. Va en el QR en lugar del ID para que
    la página de validación pueda mostrar datos sin quedar expuesta a que
    alguien recorra folios consecutivos."""
    import secrets
    alfabeto = "abcdefghijkmnpqrstuvwxyz23456789"  # sin l, o, 0, 1: se confunden
    return "".join(secrets.choice(alfabeto) for _ in range(LONGITUD_TOKEN))


def generar_qr_png(token: str, color: str = "#888888") -> bytes:
    """QR con corrección de error alta (H): sobrevive sello, engrapado y
    dobleces. Lleva la URL completa con el token para que al escanearlo con
    el celular abra directo la página de validación. Eso obliga a 2.0 cm
    mínimo (41 módulos); no se sacrifica por hacerlo más pequeño."""
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_H,
        box_size=10,
        border=2,
    )
    qr.add_data(f"{URL_APP}/?validar_oficio={token}")
    qr.make(fit=True)
    img = qr.make_image(fill_color=color, back_color="white")
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


POSICIONES_QR = {
    "Superior izquierda": "SI",
    "Superior derecha": "SD",
    "Inferior izquierda": "II",
    "Inferior derecha": "ID",
}


def estampar_qr_pdf(pdf_bytes: bytes, token: str, posicion: str = "ID",
                    lado_cm: float = 2.3, margen_cm: float = 1.0,
                    color: str = "#888888", etiqueta: str = "") -> bytes:
    """Inserta el QR discreto en la esquina indicada de la PRIMERA página.

    Medido sobre escaneo degradado a 150 dpi (contraste pobre, inclinación,
    JPEG): con la URL completa (49 módulos), 2.3 cm es el mínimo fiable
    (12/12 lecturas). A 2.0 cm falla siempre (0/12). El color casi no afecta la lectura, así
    que el gris da discreción sin costo; el blanco no se lee nunca.
    """
    if pymupdf is None:
        raise RuntimeError(f"PyMuPDF no disponible: {_ERROR_PDF}")

    png = generar_qr_png(token, color)
    doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    try:
        pagina = doc[0]
        lado, margen = lado_cm * CM, margen_cm * CM
        ancho, alto = pagina.rect.width, pagina.rect.height

        x0 = margen if posicion in ("SI", "II") else ancho - margen - lado
        y0 = margen if posicion in ("SI", "SD") else alto - margen - lado

        rect = pymupdf.Rect(x0, y0, x0 + lado, y0 + lado)
        pagina.insert_image(rect, stream=png)
        if etiqueta:
            # Respaldo por si el QR se destruye: permite buscar a mano.
            pagina.insert_text((rect.x0, rect.y1 + 5), etiqueta, fontsize=4.5,
                               color=(0.6, 0.6, 0.6))
        return doc.tobytes()
    finally:
        doc.close()


def vista_previa_pdf(pdf_bytes: bytes, dpi: int = 90) -> bytes:
    """PNG de la primera página, para revisar dónde cayó el QR antes de confirmar."""
    if pymupdf is None:
        return b""
    doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    try:
        return doc[0].get_pixmap(dpi=dpi).tobytes("png")
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


_RE_TOKEN = re.compile(r"^[a-hjkmnp-z2-9]{" + str(LONGITUD_TOKEN) + r"}$")


def leer_qr_pdf(archivo_bytes: bytes, nombre: str, dpi: int = 300) -> list[tuple[int, str]]:
    """Devuelve [(num_pagina, token), ...] de cada página con un QR del
    minutario. Es lo que permite procesar un escaneo en lote sin separadores.
    300 dpi porque el QR es pequeño: a menos resolución se pierde."""
    if zxingcpp is None:
        raise RuntimeError(f"zxing-cpp no disponible: {_ERROR_ZX}")

    from PIL import Image

    def _tokens_de_imagen(img, num_pag):
        hallados = []
        for res in zxingcpp.read_barcodes(img):
            texto = str(res.text).strip()
            if "validar_oficio=" in texto:
                texto = texto.split("validar_oficio=")[-1].strip()
            if _RE_TOKEN.match(texto):
                hallados.append((num_pag, texto))
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
                encontrados.extend(_tokens_de_imagen(img, i))
        finally:
            doc.close()
    else:
        encontrados.extend(_tokens_de_imagen(Image.open(BytesIO(archivo_bytes)), 0))
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

_RE_FOLIO_ENC = re.compile(
    r"oficio\s*(?::|n[uú]m\.?(?:ero)?|n[oº°]\.?|#)?\s*:?\s*"
    r"(\d{1,6})\s*/\s*(?:[\dA-Za-z]{1,8}\s*/\s*)?(\d{4})",
    re.IGNORECASE,
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
# [ \t]* y no \s*: \s incluye el salto de línea y se comería el fragmento
# siguiente cuando el PDF parte "Asunto:" y su contenido en dos líneas.
_RE_ASUNTO = re.compile(r"^[ \t]*asunto[ \t]*:?[ \t]*(\S.*)$",
                        re.IGNORECASE | re.MULTILINE)
_RE_PRESENTE = re.compile(r"^\s*P\s*R\s*E\s*S\s*E\s*N\s*T\s*E\s*\.?\s*$", re.IGNORECASE)
# Líneas que pertenecen al encabezado y por tanto NO son del destinatario.
_RE_ENCABEZADO = re.compile(
    r"^\s*(?:oficio|asunto|expediente|exp)\s*[:.]|"
    r"\b\d{1,2}\s+de\s+[a-záéíóú]+\s+(?:de|del)\s+\d{4}",
    re.IGNORECASE,
)


def _limpiar(texto: str) -> str:
    return re.sub(r"\s+", " ", str(texto or "")).strip(" .:,;-")


# Palabras con que arranca el cargo del destinatario. Sirven para separar
# nombre y cargo cuando el PDF los entrega pegados ("BELTRÁNJEFA DEL...").
INICIOS_CARGO = (
    "JEFA", "JEFE", "DIRECTOR", "DIRECTORA", "ENCARGADO", "ENCARGADA",
    "COORDINADOR", "COORDINADORA", "SUBDIRECTOR", "SUBDIRECTORA",
    "TITULAR", "SECRETARIO", "SECRETARIA", "DELEGADO", "DELEGADA",
    "SUPERVISOR", "SUPERVISORA", "RESPONSABLE", "ASESOR", "ASESORA",
    "ADMINISTRADOR", "ADMINISTRADORA", "PRESIDENTE", "PRESIDENTA",
)

# Sobre texto aplanado los espacios son poco fiables, así que todo separador
# se trata como opcional (\s*) en vez de obligatorio.
_RE_P_FOLIO = re.compile(
    r"oficio\s*(?::|n[uú]m\.?(?:ero)?|n[oº°]\.?|#)?\s*"
    r"([\d\s]{1,10}?)\s*/\s*(?:[\dA-Za-z]{1,8}\s*/\s*)?(\d{4})",
    re.IGNORECASE,
)
_RE_P_FECHA = re.compile(
    r"(\d{1,2})\s*de\s*([a-záéíóúñ]+?)\s*de\s*l?\s*(\d{4})", re.IGNORECASE
)
_RE_P_ASUNTO = re.compile(r"asunto\s*:?\s*(.{3,200}?)\s*\.", re.IGNORECASE)
_RE_P_PRESENTE = re.compile(r"P\s*R\s*E\s*S\s*E\s*N\s*T\s*E", re.IGNORECASE)


def _aplanar(texto: str) -> str:
    """Une las líneas SIN separador. Suena contraintuitivo, pero los PDFs
    rotos parten los números a media cifra ('de 202' + '6'); unir con espacio
    los dejaría irreparables ('202 6'), mientras que unir sin él los repara."""
    return "".join(texto.splitlines())


def _separar_nombre_cargo(bloque: str) -> tuple[str, str]:
    """Parte 'ROSA ESTELA MATA BELTRÁNJEFA DEL DEPARTAMENTO...' en nombre y
    cargo, buscando dónde empieza una palabra de cargo conocida."""
    bloque = _limpiar(bloque)
    mejor = None
    for palabra in INICIOS_CARGO:
        pos = bloque.upper().find(palabra)
        if pos > 0 and (mejor is None or pos < mejor):
            mejor = pos
    if mejor:
        return _limpiar(bloque[:mejor]), _limpiar(bloque[mejor:])
    return bloque, ""


def extraer_datos_plano(texto: str) -> dict:
    """Respaldo para PDFs cuya capa de texto viene fragmentada. No sustituye
    al extractor por líneas: solo rellena lo que aquel dejó vacío."""
    datos = {c: "" for c in CAMPOS_EXTRAIBLES}
    plano = _aplanar(texto)
    if not plano.strip():
        return datos

    m = _RE_P_FOLIO.search(plano)
    if m:
        digitos = re.sub(r"\s+", "", m.group(1))
        if digitos.isdigit():
            datos["folio"] = digitos.zfill(4)

    m = _RE_P_FECHA.search(plano[:600]) or _RE_P_FECHA.search(plano)
    if m:
        mes = MESES.get(m.group(2).lower())
        if mes:
            datos["fecha_oficio"] = f"{m.group(3)}-{mes:02d}-{int(m.group(1)):02d}"

    m = _RE_P_ASUNTO.search(plano)
    if m:
        datos["asunto"] = _limpiar(m.group(1))[:250]

    # El destinatario vive entre el asunto y la línea "P R E S E N T E".
    m_pres = _RE_P_PRESENTE.search(plano)
    if m_pres:
        inicio = 0
        m_as = re.search(r"asunto\s*:?\s*.{3,200}?\.", plano, re.IGNORECASE)
        if m_as:
            inicio = m_as.end()
        bloque = plano[inicio:m_pres.start()]
        if 4 < len(bloque) < 300:
            nombre, cargo = _separar_nombre_cargo(bloque)
            datos["dirigido_a"] = nombre[:120]
            datos["cargo_destino"] = cargo[:200]

    return datos


def extraer_datos_regex(texto: str) -> dict:
    """Lee los campos del oficio con reglas explícitas. Todo campo que no
    encuentre queda vacío: nunca adivina."""
    datos = {c: "" for c in CAMPOS_EXTRAIBLES}
    if not texto or not texto.strip():
        return datos

    lineas = [l.rstrip() for l in texto.splitlines()]

    # ── Folio ──
    m = _RE_FOLIO_ENC.search(texto) or _RE_FOLIO.search(texto) or _RE_FOLIO_ALT.search(texto)
    if m:
        datos["folio"] = m.group(1).zfill(4)

    # ── Fecha ──
    # Se lee del encabezado APLANADO y en un tramo corto. Dos razones:
    #   1. El cuerpo cita otras fechas (los días que se justifican) y ganarían.
    #   2. En PDFs fragmentados la fecha del encabezado viene partida a media
    #      cifra ("de 202" + "6") y solo se recompone al aplanar.
    encabezado = _aplanar(texto)[:400]
    m = _RE_P_FECHA.search(encabezado)
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
    # Se ancla en la línea "P R E S E N T E" y se camina hacia atrás: la
    # primera línea del bloque es el nombre, las siguientes son el cargo.
    # No se exige tratamiento (LIC., MTRA.): muchos oficios no lo usan.
    idx_presente = next(
        (i for i, l in enumerate(lineas) if _RE_PRESENTE.match(l)), None
    )
    if idx_presente is not None:
        bloque = []
        for i in range(idx_presente - 1, max(-1, idx_presente - 7), -1):
            linea = lineas[i].strip()
            if not linea:
                break
            if _RE_ENCABEZADO.search(linea):
                break
            bloque.insert(0, _limpiar(linea))
        if bloque:
            datos["dirigido_a"] = bloque[0][:120]
            if len(bloque) > 1:
                datos["cargo_destino"] = " ".join(bloque[1:])[:200]

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

    # PDFs con capa de texto rota: el extractor por líneas falla, el plano no.
    faltantes = [c for c in CAMPOS_EXTRAIBLES if not datos[c]]
    if faltantes:
        plano = extraer_datos_plano(texto)
        rellenados = [c for c in faltantes if plano.get(c)]
        if rellenados:
            for c in rellenados:
                datos[c] = plano[c]
            metodo = "reglas (texto fragmentado)"

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


def sha256_bytes(contenido: bytes) -> str:
    """Huella del archivo. Si alguien lo sustituye en Drive, deja de coincidir."""
    return hashlib.sha256(contenido).hexdigest()


def buscar_por_token(df: pd.DataFrame, token: str):
    """Localiza el oficio por su token del QR. Devuelve la fila o None."""
    if df.empty or "TOKEN" not in df.columns:
        return None
    hit = df[df["TOKEN"].astype(str).str.strip() == str(token).strip()]
    return None if hit.empty else hit.iloc[0]


def reservar_oficio(get_client, datos: dict) -> tuple[bool, str, str]:
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
            ), ""

    token = generar_token()
    if not df.empty and "TOKEN" in df.columns:
        usados = set(df["TOKEN"].astype(str))
        while token in usados:
            token = generar_token()

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
            token,
            "", "",   # URL_ESCANEO, SHA256_ESCANEO
        ], value_input_option="USER_ENTERED")
        st.cache_data.clear()
        _registrar_log(get_client, id_oficio, "RESERVADO", datos.get("asunto", ""))
        return True, id_oficio, token
    except Exception as e:
        return False, f"Error al guardar: {e}", ""


def registrar_oficio_historico(get_client, datos: dict, archivo=None) -> tuple[bool, str]:
    """Alta de un oficio emitido antes de este sistema. No genera token ni
    estampa QR: el papel ya circuló. El escaneo del acuse es opcional para
    que se pueda capturar el registro completo primero y subir después."""
    id_oficio = construir_id(datos["anio"], datos["folio"])

    df = cargar_oficios(get_client)
    if not df.empty and "ID_OFICIO" in df.columns:
        if id_oficio in df["ID_OFICIO"].astype(str).values:
            return False, f"El folio {id_oficio} ya está registrado."

    url, huella = "", ""
    if archivo is not None:
        contenido = archivo.getvalue()
        ext = archivo.name.split(".")[-1].lower()
        url = subir_bytes_drive(contenido, f"{id_oficio}_ACUSE.{ext}",
                                archivo.type or "application/octet-stream")
        if url.startswith("ERROR:"):
            return False, url
        huella = sha256_bytes(contenido)

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
            ESTADO_HISTORICO,
            "", "",                                   # URL_EMITIDO, SHA256_EMITIDO
            _hoy() if url else "",                    # FECHA_ESCANEO
            datos.get("observaciones", ""),
            "",                                       # TOKEN: los históricos no llevan
            url, huella,                              # URL_ESCANEO, SHA256_ESCANEO
        ], value_input_option="USER_ENTERED")
        st.cache_data.clear()
        _registrar_log(get_client, id_oficio, "HISTORICO", datos.get("asunto", ""))
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
# ACUSES SIN FOLIO
#
# Documentos que NO llevan número de oficio de la DFC (propuestas de
# interinos, asignaciones temporales, constancias laborales). Viven en su
# propia pestaña del Sheet: mezclarlos con el minutario confundiría el
# consecutivo oficial, que no es de RH sino de la Dirección.
# ─────────────────────────────────────────────
TAB_ACUSES = "acuses"

COLUMNAS_ACUSES = [
    "ID_ACUSE", "AÑO", "CONSECUTIVO", "TIPO", "REFERENCIA",
    "FECHA_DOCUMENTO", "FECHA_REGISTRO", "REGISTRO_RFC", "REGISTRO_NOMBRE",
    "DESCRIPCION", "RELACIONADO_CON", "URL", "SHA256", "OBSERVACIONES",
]

TIPOS_ACUSE = {
    "Propuesta de interinos": "PROP",
    "Asignación temporal": "ASIG",
    "Constancia laboral": "CONS",
    "Otro documento": "OTRO",
}


@st.cache_data(ttl=300)
def _cargar_acuses_cached(_get_client) -> pd.DataFrame:
    sh = _abrir_sheet(_get_client)
    ws = sh.worksheet(TAB_ACUSES)
    registros = ws.get_all_records(numericise_ignore=["all"])
    if not registros:
        return pd.DataFrame(columns=COLUMNAS_ACUSES)
    return pd.DataFrame(registros)


def cargar_acuses(get_client) -> pd.DataFrame:
    return _cargar_acuses_cached(get_client)


def siguiente_consecutivo(df: pd.DataFrame, anio, clave_tipo: str) -> int:
    """Consecutivo interno por tipo y año. A diferencia del minutario, este
    sí lo controla RH, así que se calcula aquí."""
    if df.empty or "TIPO" not in df.columns:
        return 1
    prev = df[(df["AÑO"].astype(str) == str(anio)) &
              (df["TIPO"].astype(str) == clave_tipo)]
    if prev.empty:
        return 1
    nums = [int(c) for c in prev["CONSECUTIVO"].astype(str) if str(c).strip().isdigit()]
    return (max(nums) + 1) if nums else 1


def registrar_acuse(get_client, datos: dict, archivo) -> tuple[bool, str]:
    """Registra un acuse y resguarda su escaneo. El documento ya viene
    firmado, así que no se estampa QR: aquí el valor es el resguardo."""
    df = cargar_acuses(get_client)
    clave = datos["tipo"]
    consecutivo = siguiente_consecutivo(df, datos["anio"], clave)
    id_acuse = f"DFC-{datos['anio']}-{clave}-{consecutivo:04d}"

    contenido = archivo.getvalue()
    ext = archivo.name.split(".")[-1].lower()
    url = subir_bytes_drive(contenido, f"{id_acuse}.{ext}",
                            archivo.type or "application/octet-stream")
    if url.startswith("ERROR:"):
        return False, url

    try:
        sh = _abrir_sheet(get_client)
        ws = sh.worksheet(TAB_ACUSES)
        ws.append_row([
            id_acuse,
            str(datos["anio"]),
            str(consecutivo).zfill(4),
            clave,
            datos.get("referencia", ""),
            datos.get("fecha_documento", ""),
            _hoy(),
            str(st.session_state.get("rfc", "")).upper(),
            st.session_state.get("nombre", ""),
            datos.get("descripcion", ""),
            datos.get("relacionado_con", ""),
            url,
            sha256_bytes(contenido),
            datos.get("observaciones", ""),
        ], value_input_option="USER_ENTERED")
        st.cache_data.clear()
        _registrar_log(get_client, id_acuse, "ACUSE_REGISTRADO",
                       datos.get("descripcion", ""))
        return True, id_acuse
    except Exception as e:
        return False, f"Error al guardar: {e}"



def ficha_trazabilidad(get_client, id_oficio: str) -> bytes:
    """Hoja de una página con todo el rastro de un oficio, para presentar
    cuando alguien cuestiona su procedencia. Prioriza lo que una persona
    puede verificar (fechas, registro, resguardo) sobre lo criptográfico."""
    if pymupdf is None:
        raise RuntimeError(f"PyMuPDF no disponible: {_ERROR_PDF}")

    df = cargar_oficios(get_client)
    fila = df[df["ID_OFICIO"].astype(str) == str(id_oficio)]
    if fila.empty:
        raise ValueError(f"No existe {id_oficio} en el minutario.")
    r = fila.iloc[0]

    doc = pymupdf.open()
    pg = doc.new_page(width=612, height=792)
    x, y = 60, 70

    def linea(txt, tam=10, negrita=False, salto=16, gris=0.0):
        nonlocal y
        pg.insert_text((x, y), txt, fontsize=tam,
                       fontname="Helvetica-Bold" if negrita else "Helvetica",
                       color=(gris, gris, gris))
        y += salto

    linea("FICHA DE TRAZABILIDAD DE OFICIO", 15, True, 10)
    linea("Recursos Humanos · Dirección de Formación Continua · SEJ", 9, False, 24, 0.4)
    pg.draw_line(pymupdf.Point(x, y - 10), pymupdf.Point(552, y - 10),
                 color=(0.7, 0.7, 0.7), width=0.7)

    linea(str(r.get("ID_OFICIO", "")), 17, True, 26)

    for etiqueta, valor in [
        ("Folio del minutario", r.get("FOLIO", "")),
        ("Fecha del oficio", r.get("FECHA_OFICIO", "")),
        ("Registrado en el sistema", r.get("FECHA_SOLICITUD", "")),
        ("Emitido por", r.get("EMISOR_NOMBRE", "")),
        ("Asunto", r.get("ASUNTO", "")),
        ("Dirigido a", r.get("DIRIGIDO_A", "")),
        ("Cargo", r.get("CARGO_DESTINO", "")),
        ("Estado", r.get("ESTADO", "")),
        ("Fecha de escaneo", r.get("FECHA_ESCANEO", "") or "—"),
    ]:
        linea(f"{etiqueta}:", 9, True, 13, 0.35)
        linea(str(valor or "—")[:95], 11, False, 20)

    y += 8
    pg.draw_line(pymupdf.Point(x, y), pymupdf.Point(552, y),
                 color=(0.7, 0.7, 0.7), width=0.7)
    y += 20
    linea("ELEMENTOS DE VERIFICACIÓN", 11, True, 20)
    linea("1. El ejemplar original resguardado el día de la emisión permite", 9, False, 12, 0.2)
    linea("   comparar el texto contra cualquier documento cuestionado.", 9, False, 16, 0.2)
    linea("2. El archivo en la unidad compartida conserva su fecha de creación.", 9, False, 16, 0.2)
    linea("3. El folio puede contrastarse contra el minutario de la Dirección,", 9, False, 12, 0.2)
    linea("   que es un registro independiente de este sistema.", 9, False, 16, 0.2)
    linea("4. La bitácora del sistema conserva quién registró el oficio y cuándo.", 9, False, 20, 0.2)

    linea("Huellas digitales de los archivos (SHA-256)", 9, True, 13, 0.35)
    for et, h in [("Original emitido", r.get("SHA256_EMITIDO", "")),
                  ("Ejemplar firmado", r.get("SHA256_ESCANEO", ""))]:
        linea(f"{et}: {str(h or '—')[:64]}", 6.5, False, 11, 0.45)

    y += 10
    linea(f"Ficha generada el {_ahora()} · documento informativo interno",
          7.5, False, 12, 0.5)

    salida = doc.tobytes()
    doc.close()
    return salida


# ─────────────────────────────────────────────
# VALIDACIÓN PÚBLICA (QR)
# ─────────────────────────────────────────────
def render_validacion_oficio(get_client, token: str):
    """Vista que abre el QR. Muestra los datos que la persona puede CONTRASTAR
    contra el papel que tiene enfrente: si alguien copió este QR a otro
    documento, el asunto y el destinatario no van a coincidir.

    Se accede por token opaco, no por ID: si la URL llevara DFC-2026-1009,
    cualquiera podría recorrer 1010, 1011, 1012 y cosechar información.
    """
    st.markdown("### 🔎 Verificación de oficio · RH · DFC")
    try:
        df = cargar_oficios(get_client)
    except Exception as e:
        _error_amable(e, "al consultar el minutario")
        return

    r = buscar_por_token(df, token)
    if r is None:
        st.error("Este código **no corresponde** a ningún oficio emitido por "
                 "Recursos Humanos de la Dirección de Formación Continua.")
        return

    id_oficio = str(r.get("ID_OFICIO", ""))
    if str(r.get("ESTADO", "")) == "CANCELADO":
        st.warning(f"El oficio **{id_oficio}** fue **cancelado** "
                   "y no debe surtir efectos.")
        return

    st.success(f"Oficio **{id_oficio}** emitido por Recursos Humanos de la DFC.")
    st.info("**Compara estos datos con el documento impreso.** Si no coinciden, "
            "el código fue copiado de otro oficio.")

    c1, c2 = st.columns(2)
    c1.markdown(f"**Fecha del oficio**  \n{r.get('FECHA_OFICIO') or '—'}")
    c2.markdown(f"**Estado**  \n{r.get('ESTADO', '—')}")
    st.markdown(f"**Asunto**  \n{r.get('ASUNTO') or '—'}")
    st.markdown(f"**Dirigido a**  \n{r.get('DIRIGIDO_A') or '—'}"
                + (f"  \n{r.get('CARGO_DESTINO')}" if r.get("CARGO_DESTINO") else ""))

    url_emitido = str(r.get("URL_EMITIDO", ""))
    url_escaneo = str(r.get("URL_ESCANEO", ""))
    if url_emitido.startswith("http") or url_escaneo.startswith("http"):
        st.markdown("**Documentos resguardados**")
        if url_emitido.startswith("http"):
            st.markdown(f"📄 [Original emitido por RH]({url_emitido}) "
                        "— la versión que salió de esta oficina")
        if url_escaneo.startswith("http"):
            st.markdown(f"🖊️ [Ejemplar firmado y sellado]({url_escaneo})")
        st.caption("El acceso a los archivos depende de los permisos de la "
                   "unidad compartida de la Dirección.")

    st.divider()
    st.caption("Esta verificación acredita que el folio fue emitido por RH y "
               "cuáles son sus datos registrados. Contrástalos con el papel.")


# ─────────────────────────────────────────────
# VISTA PRINCIPAL
# ─────────────────────────────────────────────
def _tab_historico(get_client):
    """Captura de oficios anteriores a este sistema. Diseñada para cargar
    varios seguidos: el escaneo es opcional y el formulario se limpia solo,
    de modo que se pueda capturar todo el año y subir acuses después."""
    st.caption("Oficios emitidos antes de este sistema. Quedan registrados con "
               "estado HISTORICO, sin QR: el papel ya circuló.")

    anio_actual = datetime.now(TZ).year
    h1, h2, h3 = st.columns([1, 1, 2])
    anio_h = h1.number_input("Año", min_value=2020, max_value=2100,
                             value=anio_actual, step=1, key="oh_anio")
    folio_h = h2.text_input("Folio", key="oh_folio")
    fecha_h = h3.date_input("Fecha del oficio", key="oh_fecha")

    asunto_h = st.text_input("Asunto", key="oh_asunto")
    h4, h5 = st.columns(2)
    dirig_h = h4.text_input("Dirigido a", key="oh_dirigido")
    cargo_h = h5.text_input("Cargo del destinatario", key="oh_cargo")
    obs_h = st.text_area("Observaciones", key="oh_obs", height=68)

    arch_h = st.file_uploader("Acuse escaneado (opcional)",
                              type=["pdf", "jpg", "jpeg", "png"], key="oh_file")
    st.caption("Puedes capturar solo los datos ahora y subir los acuses después.")

    if st.button("Registrar oficio histórico", type="primary", key="oh_btn"):
        if not str(folio_h).strip().isdigit():
            st.warning("El folio debe ser numérico.")
        elif not asunto_h.strip() or not dirig_h.strip():
            st.warning("Asunto y destinatario son obligatorios.")
        else:
            with st.spinner("Registrando..."):
                ok, res = registrar_oficio_historico(get_client, {
                    "anio": int(anio_h),
                    "folio": folio_h,
                    "fecha_oficio": fecha_h.strftime("%Y-%m-%d"),
                    "asunto": asunto_h.strip(),
                    "dirigido_a": dirig_h.strip(),
                    "cargo_destino": cargo_h.strip(),
                    "observaciones": obs_h.strip(),
                }, arch_h)
            if ok:
                st.success(f"Registrado: **{res}**")
                # Se limpian los campos del documento, no el año: al capturar
                # en serie casi siempre es el mismo ejercicio.
                for k in ("oh_folio", "oh_asunto", "oh_dirigido",
                          "oh_cargo", "oh_obs"):
                    st.session_state.pop(k, None)
                st.rerun()
            else:
                st.error(res)


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

    if not df.empty and _faltan_columnas(df, COLUMNAS_OFICIOS, TAB_OFICIOS):
        return

    tab_reg, tab_esc, tab_cons, tab_acu = st.tabs(
        ["📤 Emitir oficio", "📎 Escaneo firmado", "📋 Consultar",
         "🗂️ Acuses sin folio"]
    )

    # ── Emitir oficio ────────────────────────
    with tab_reg:
        modo = st.radio(
            "¿Qué vas a registrar?",
            ["Oficio nuevo (se le estampa QR)",
             "Oficio histórico (ya firmado y entregado)"],
            key="of_modo", horizontal=True,
            help="Los históricos son los emitidos antes de este sistema: se "
                 "registran para tener el año completo, pero no llevan QR.")
        st.divider()

        if modo.startswith("Oficio histórico"):
            _tab_historico(get_client)
        elif pymupdf is None:
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

                st.divider()
                cp1, cp2 = st.columns([1, 2])
                pos_nom = cp1.radio("Posición del QR",
                                    options=list(POSICIONES_QR.keys()),
                                    index=3, key="of_pos",
                                    help="Elige una zona libre de sello y firma.")
                pos = POSICIONES_QR[pos_nom]
                lado = cp1.slider("Tamaño (cm)", 2.3, 3.0, 2.3, 0.1, key="of_lado",
                                  help="2.3 cm es el mínimo medido para que el "
                                       "QR abra la página de validación.")
                discreto = cp1.checkbox("Gris discreto", value=True, key="of_gris")
                color = "#888888" if discreto else "#000000"
                try:
                    prev = vista_previa_pdf(estampar_qr_pdf(
                        pdf_bytes, "abcdefghjkmn", pos, lado, color=color))
                    cp2.image(prev, caption="Vista previa · así se imprimirá",
                              use_container_width=True)
                except Exception as e:
                    cp2.warning(f"No se pudo generar la vista previa: {e}")

                if st.button("Confirmar y estampar QR", type="primary", key="of_btn_reg"):
                    if not str(folio).strip().isdigit():
                        st.warning("El folio debe ser numérico.")
                    elif not asunto.strip() or not dirigido.strip():
                        st.warning("Asunto y destinatario son obligatorios.")
                    else:
                        ok, resultado, token = reservar_oficio(get_client, {
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
                            # Registrado pero sin PDF en mano (p. ej. se perdió
                            # la descarga): se re-estampa con el token que ya
                            # tiene, sin duplicar el registro ni cambiar el QR.
                            id_dup = construir_id(int(anio), folio)
                            fila_dup = df[df["ID_OFICIO"].astype(str) == id_dup] \
                                if "ID_OFICIO" in df.columns else pd.DataFrame()
                            tok_dup = (str(fila_dup.iloc[0].get("TOKEN", ""))
                                       if not fila_dup.empty else "")
                            if tok_dup:
                                if st.button("Volver a estampar y descargar este oficio",
                                             key="of_btn_reestampar"):
                                    try:
                                        st.session_state["_of_pdf_listo"] = estampar_qr_pdf(
                                            pdf_bytes, tok_dup, pos, lado, color=color)
                                        st.session_state["_of_pdf_nombre"] = id_dup
                                        _registrar_log(get_client, id_dup,
                                                       "REESTAMPADO", "")
                                        st.rerun()
                                    except Exception as e:
                                        _error_amable(e, "al re-estampar")
                            else:
                                st.caption("Ese folio no tiene código QR "
                                           "(es histórico o se registró sin token).")
                        else:
                            try:
                                with st.spinner("Estampando QR y resguardando..."):
                                    sellado = estampar_qr_pdf(
                                        pdf_bytes, token, pos, lado, color=color)
                                    url = subir_bytes_drive(
                                        sellado, f"{resultado}_SINFIRMA.pdf",
                                        "application/pdf")
                                    _actualizar_fila(get_client, resultado, {
                                        "ESTADO": "EMITIDO",
                                        "URL_EMITIDO": url if not url.startswith("ERROR:") else "",
                                        "SHA256_EMITIDO": sha256_bytes(sellado),
                                    })
                                    _registrar_log(get_client, resultado, "EMITIDO", url)
                                st.session_state["_of_pdf_listo"] = sellado
                                st.session_state["_of_pdf_nombre"] = resultado
                                if url.startswith("ERROR:"):
                                    st.warning(f"Se registró, pero falló Drive: {url}")
                            except Exception as e:
                                _error_amable(e, "al estampar el QR")

                # El PDF vive en sesión, no dentro del if del botón: al
                # presionar "Descargar" Streamlit relanza el script y el
                # archivo se perdería si dependiera de la ejecución anterior.
                if st.session_state.get("_of_pdf_listo"):
                    nombre_listo = st.session_state.get("_of_pdf_nombre", "oficio")
                    st.success(f"**{nombre_listo}** listo. Imprime este PDF "
                               "y pásalo a firma.")
                    st.download_button(
                        "⬇️ Descargar PDF con QR",
                        data=st.session_state["_of_pdf_listo"],
                        file_name=f"{nombre_listo}_conQR.pdf",
                        mime="application/pdf", key="of_dl_pdf",
                    )

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

                # El QR trae el token; se traduce al ID por el Sheet.
                mapa = {}
                if "TOKEN" in df.columns:
                    mapa = dict(zip(df["TOKEN"].astype(str).str.strip(),
                                    df["ID_OFICIO"].astype(str)))
                validas = [(p, mapa[t]) for p, t in marcas if t in mapa]
                huerfanas = [t for _, t in marcas if t not in mapa]

                if huerfanas:
                    st.error(f"Se detectaron {len(huerfanas)} código(s) QR que "
                             "**no** corresponden a este minutario.")

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
                                    "URL_ESCANEO": url,
                                    "SHA256_ESCANEO": sha256_bytes(contenido),
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
                      int((vista["URL_ESCANEO"].astype(str).str.startswith("http")).sum()))
            m3.metric("Cancelados",
                      int((vista["ESTADO"].astype(str) == "CANCELADO").sum()))

            st.dataframe(
                vista[["ID_OFICIO", "FECHA_OFICIO", "EMISOR_NOMBRE", "ASUNTO",
                       "DIRIGIDO_A", "ESTADO", "URL_EMITIDO", "URL_ESCANEO"]],
                use_container_width=True, hide_index=True,
            )

            if es_admin:
                with st.expander("📑 Ficha de trazabilidad (para un reclamo)"):
                    st.caption("Hoja con el rastro completo de un oficio, "
                               "lista para imprimir y presentar.")
                    id_ficha = st.selectbox(
                        "Oficio", options=vista["ID_OFICIO"].astype(str).tolist(),
                        key="of_sel_ficha")
                    if st.button("Generar ficha", key="of_btn_ficha"):
                        try:
                            pdf_ficha = ficha_trazabilidad(get_client, id_ficha)
                            st.download_button(
                                "⬇️ Descargar ficha", data=pdf_ficha,
                                file_name=f"FICHA_{id_ficha}.pdf",
                                mime="application/pdf", key="of_dl_ficha")
                        except Exception as e:
                            _error_amable(e, "al generar la ficha")

                # Los HISTORICO no llevan token por diseño: no deben aparecer aquí.
                if "TOKEN" in vista.columns:
                    sin_token = vista[
                        (vista["TOKEN"].astype(str).str.strip() == "") &
                        (vista["ESTADO"].astype(str) != ESTADO_HISTORICO)
                    ]
                else:
                    sin_token = vista[vista["ESTADO"].astype(str) != ESTADO_HISTORICO]
                if not sin_token.empty:
                    with st.expander(f"⚠️ {len(sin_token)} oficio(s) sin código QR"):
                        st.caption("Registrados antes de que existiera el token. "
                                   "Genera el código y vuelve a estampar el PDF.")
                        id_sin = st.selectbox(
                            "Oficio", options=sin_token["ID_OFICIO"].astype(str).tolist(),
                            key="of_sel_token")
                        if st.button("Generar código", key="of_btn_token"):
                            nuevo = generar_token()
                            if _actualizar_fila(get_client, id_sin, {"TOKEN": nuevo}):
                                _registrar_log(get_client, id_sin, "TOKEN_GENERADO", "")
                                st.success(f"Código generado para {id_sin}. "
                                           "Vuelve a estampar y reimprimir el oficio.")
                                st.rerun()
                            else:
                                st.error("No se pudo guardar el código.")

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

    # ── Acuses sin folio ─────────────────────
    with tab_acu:
        st.caption("Documentos sin número de oficio de la DFC: propuestas de "
                   "interinos, asignaciones temporales, constancias laborales. "
                   "Llevan consecutivo propio de RH.")

        try:
            df_acu = cargar_acuses(get_client)
        except Exception as e:
            df_acu = pd.DataFrame(columns=COLUMNAS_ACUSES)
            _error_amable(e, "al cargar los acuses")

        if not df_acu.empty and _faltan_columnas(df_acu, COLUMNAS_ACUSES, TAB_ACUSES):
            df_acu = pd.DataFrame(columns=COLUMNAS_ACUSES)

        sub_alta, sub_lista = st.tabs(["➕ Registrar acuse", "📋 Ver acuses"])

        with sub_alta:
            ca1, ca2 = st.columns([2, 1])
            tipo_nom = ca1.selectbox("Tipo de documento",
                                     options=list(TIPOS_ACUSE.keys()),
                                     key="ac_tipo")
            clave = TIPOS_ACUSE[tipo_nom]
            anio_a = ca2.number_input("Año", min_value=2020, max_value=2100,
                                      value=datetime.now(TZ).year, step=1,
                                      key="ac_anio")

            prox = siguiente_consecutivo(df_acu, int(anio_a), clave)
            st.caption(f"Se registrará como **DFC-{int(anio_a)}-{clave}-{prox:04d}**")

            desc = st.text_input("Descripción", key="ac_desc",
                                 help="De qué se trata el documento.")
            cb1, cb2 = st.columns(2)
            ref = cb1.text_input("Referencia externa", key="ac_ref",
                                 help="Folio o identificador propio del "
                                      "documento, si lo tiene. Opcional.")
            fecha_doc = cb2.date_input("Fecha del documento", key="ac_fecha")
            relac = st.text_input("Relacionado con", key="ac_rel",
                                  help="Persona, Centro de Maestros o proceso "
                                       "al que corresponde. Opcional.")
            obs_a = st.text_area("Observaciones", key="ac_obs", height=68)

            arch = st.file_uploader("Escaneo del acuse (PDF o imagen)",
                                    type=["pdf", "jpg", "jpeg", "png"],
                                    key="ac_file")

            if st.button("Registrar acuse", type="primary", key="ac_btn"):
                if arch is None:
                    st.warning("Falta el escaneo del documento.")
                elif not desc.strip():
                    st.warning("La descripción es obligatoria.")
                else:
                    with st.spinner("Resguardando en Drive..."):
                        ok, res = registrar_acuse(get_client, {
                            "anio": int(anio_a),
                            "tipo": clave,
                            "referencia": ref.strip(),
                            "fecha_documento": fecha_doc.strftime("%Y-%m-%d"),
                            "descripcion": desc.strip(),
                            "relacionado_con": relac.strip(),
                            "observaciones": obs_a.strip(),
                        }, arch)
                    if ok:
                        st.success(f"Registrado como **{res}**")
                        st.rerun()
                    else:
                        st.error(res)

        with sub_lista:
            if df_acu.empty:
                st.info("Aún no hay acuses registrados.")
            else:
                cf1, cf2 = st.columns(2)
                anios_a = sorted(df_acu["AÑO"].astype(str).unique(), reverse=True)
                anio_fa = cf1.selectbox("Año", options=anios_a, key="ac_anio_f")
                tipos_f = cf2.multiselect(
                    "Tipo", options=list(TIPOS_ACUSE.keys()),
                    default=list(TIPOS_ACUSE.keys()), key="ac_tipo_f")
                claves_f = [TIPOS_ACUSE[t] for t in tipos_f]

                va = df_acu[(df_acu["AÑO"].astype(str) == anio_fa) &
                            (df_acu["TIPO"].astype(str).isin(claves_f))].copy()
                if not es_admin:
                    va = va[va["REGISTRO_RFC"].astype(str).str.upper() == rfc_actual]

                st.metric("Acuses", len(va))
                st.dataframe(
                    va[["ID_ACUSE", "FECHA_DOCUMENTO", "TIPO", "DESCRIPCION",
                        "RELACIONADO_CON", "REGISTRO_NOMBRE", "URL"]],
                    use_container_width=True, hide_index=True,
                )
