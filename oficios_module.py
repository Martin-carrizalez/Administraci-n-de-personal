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
# DRIVE
# ─────────────────────────────────────────────
def subir_oficio_drive(archivo, id_oficio: str) -> str:
    """Sube el escaneo a la carpeta de oficios. Sin permiso 'anyone':
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

        ext = archivo.name.split(".")[-1].lower()
        media = MediaIoBaseUpload(io.BytesIO(archivo.getvalue()), mimetype=archivo.type)
        meta = {
            "name": f"{id_oficio}.{ext}",
            "parents": [st.secrets["drive_oficios_folder"]],
        }
        creado = service.files().create(
            body=meta, media_body=media, fields="id", supportsAllDrives=True
        ).execute()
        return f"https://drive.google.com/file/d/{creado.get('id')}/view"
    except Exception as e:
        return f"ERROR: {e}"


def sha256_archivo(archivo) -> str:
    """Huella del escaneo. Si alguien sustituye el archivo en Drive,
    el hash deja de coincidir."""
    return hashlib.sha256(archivo.getvalue()).hexdigest()


# ─────────────────────────────────────────────
# ESCRITURA
# ─────────────────────────────────────────────
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


def registrar_escaneo(get_client, id_oficio: str, archivo, es_acuse: bool) -> tuple[bool, str]:
    url = subir_oficio_drive(archivo, id_oficio)
    if url.startswith("ERROR:"):
        return False, url

    estado = "ACUSE" if es_acuse else "ESCANEADO"
    ok = _actualizar_fila(get_client, id_oficio, {
        "URL": url,
        "SHA256": sha256_archivo(archivo),
        "FECHA_ESCANEO": _hoy(),
        "ESTADO": estado,
    })
    if ok:
        _registrar_log(get_client, id_oficio, estado, url)
        return True, url
    return False, "No se encontró el oficio en el Sheet."


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
        ["➕ Registrar folio", "📎 Subir escaneo", "📋 Consultar"]
    )

    # ── Registrar ────────────────────────────
    with tab_reg:
        anio_actual = datetime.now(TZ).year
        c1, c2, c3 = st.columns([1, 1, 2])
        anio = c1.number_input("Año", min_value=2020, max_value=2100,
                               value=anio_actual, step=1, key="of_anio")
        folio = c2.text_input("Folio del minutario", key="of_folio",
                              help="El número que te asignó la Dirección.")
        fecha_of = c3.date_input("Fecha del oficio", key="of_fecha")

        asunto = st.text_input("Asunto", key="of_asunto")
        c4, c5 = st.columns(2)
        dirigido = c4.text_input("Dirigido a", key="of_dirigido")
        cargo = c5.text_input("Cargo del destinatario", key="of_cargo")
        obs = st.text_area("Observaciones", key="of_obs", height=70)

        if st.button("Registrar folio", type="primary", key="of_btn_reg"):
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
                if ok:
                    st.success(f"Registrado: **{resultado}**")
                    png = generar_qr_png(resultado)
                    cq1, cq2 = st.columns([1, 3])
                    cq1.image(png, width=150)
                    cq2.download_button(
                        "⬇️ Descargar QR (PNG)", data=png,
                        file_name=f"QR_{resultado}.png", mime="image/png",
                        key="of_dl_qr",
                    )
                    cq2.caption("Insértalo en el oficio a 2–2.5 cm, esquina superior "
                                "derecha, **antes** de imprimir y firmar.")
                else:
                    st.error(resultado)

    # ── Subir escaneo ────────────────────────
    with tab_esc:
        if df.empty:
            st.info("Aún no hay folios registrados.")
        else:
            pendientes = df[df["ESTADO"].astype(str) != "CANCELADO"]
            opciones = pendientes["ID_OFICIO"].astype(str).tolist()
            sel = st.selectbox("Oficio", options=opciones, key="of_sel_esc")
            es_acuse = st.checkbox("Es acuse sellado de recibido", key="of_acuse")
            archivo = st.file_uploader("Escaneo (PDF o imagen)",
                                       type=["pdf", "jpg", "jpeg", "png"],
                                       key="of_file")
            if archivo and st.button("Guardar escaneo", type="primary", key="of_btn_esc"):
                with st.spinner("Subiendo a Drive..."):
                    ok, res = registrar_escaneo(get_client, sel, archivo, es_acuse)
                if ok:
                    st.success("Escaneo resguardado.")
                    st.markdown(f"[Abrir en Drive]({res})")
                else:
                    st.error(res)

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
