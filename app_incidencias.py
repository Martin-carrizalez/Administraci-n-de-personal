import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
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
    "PSE": "Pase de salida / entrada",
    "COM": "Comisión",
    "CHO": "Cambio de horario",
}
if HABILITAR_CUMPLEANOS:
    TIPO_LABELS["CUM"] = "Día de cumpleaños"

DIAS_SEMANA = ["LUN", "MAR", "MIE", "JUE", "VIE"]

DICT_JEFES = {
    "Martín Ángel Carrizalez Piña":                         "Martín Ángel Carrizalez Piña",
    "Maricela Esquivel Domínguez — Dir. Desarrollo Académico": "Maricela Esquivel Domínguez<br/>Directora de Desarrollo Académico",
    "Ignacio Aguilar García — Dir. Gestión y Evaluación":    "Ignacio Aguilar García<br/>Director de Gestión y Evaluación de Profesionales de la Educación",
    "Héctor Manuel Ramos Rico":                              "Héctor Manuel Ramos Rico",
    "Hugo Verduzco Zuñiga":                                  "Hugo Verduzco Zuñiga",
    "Erika Soledad Castillo Flores":                         "Erika Soledad Castillo Flores",
}

NOMBRE_DIRECTORA = "Claudia Gisela Ramírez Monroy<br/>Encargada del Despacho de la Dirección de Formación Continua"
DRIVE_ANEXOS_FOLDER = "1LnQjrhjEKgKxFTJD8USLoiCpCKQHfOQC"

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

@st.cache_data(ttl=300)
def cargar_usuarios():
    client = get_client()
    sh = client.open_by_key(st.secrets["sheet_checador_id"])
    ws = sh.worksheet("Usuarios")
    data = ws.get_all_records(numericise_ignore=["all"])
    return pd.DataFrame(data) if data else pd.DataFrame(columns=["RFC", "Correo electrónico institucional"])

@st.cache_data(ttl=300)
def cargar_solicitudes_eco():
    client = get_client()
    sh = client.open_by_key(st.secrets["sheet_economicos_id"])
    ws = sh.worksheet("Solicitudes")
    data = ws.get_all_records(numericise_ignore=["all"])
    return pd.DataFrame(data) if data else pd.DataFrame()

@st.cache_data(ttl=300)
def cargar_horarios():
    client = get_client()
    sh = client.open_by_key(st.secrets["sheet_checador_id"])
    ws = sh.worksheet("empleados")
    data = ws.get_all_records(numericise_ignore=["all"])
    return pd.DataFrame(data) if data else pd.DataFrame()

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

@st.cache_data(ttl=300)
def cargar_festivos():
    client = get_client()
    sh = client.open_by_key(st.secrets["sheet_checador_id"])
    ws = sh.worksheet("festivos")
    data = ws.get_all_records(numericise_ignore=["all"])
    return pd.DataFrame(data) if data else pd.DataFrame()

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

def dias_habiles_entre(fecha_inicio: date, fecha_fin: date, festivos_df: pd.DataFrame) -> int:
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
    total = 0
    d = fecha_inicio
    while d <= fecha_fin:
        if d.weekday() < 5 and d not in festivos_set:
            total += 1
        d += timedelta(days=1)
    return total

def dias_economicos_usados(rfc: str, solicitudes_df: pd.DataFrame) -> int:
    df = solicitudes_df[solicitudes_df["RFC"].astype(str).str.upper() == rfc.upper()]
    df_aprobados = df[df["Aprobado Por"].astype(str).str.strip() != ""]
    total = 0
    for _, row in df_aprobados.iterrows():
        try:
            total += int(row["Dias Solicitados"])
        except Exception:
            pass
    return total

def horas_pases_mes(rfc: str, incidencias_df: pd.DataFrame) -> float:
    """Suma las horas de pases autorizados del mes en curso usando FECHA_INICIO."""
    import pytz
    tz_mx = pytz.timezone("America/Mexico_City")
    ahora = datetime.now(pytz.utc).astimezone(tz_mx)
    df = incidencias_df[
        (incidencias_df["RFC"].astype(str).str.upper() == rfc.upper()) &
        (incidencias_df["TIPO"] == "PSE") &
        (incidencias_df["ESTADO"].astype(str).str.contains("AUTORIZADO"))
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
        anio_completo = 2000 + anio if anio <= 25 else 1900 + anio
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
    firma_interesado  = Paragraph(f"<br/><br/>___________________________<br/><b>Firma del Interesado</b><br/>{nombre_interesado}", estilo_firma)
    firma_jefe        = Paragraph(f"<br/><br/>___________________________<br/><b>Autoriza Jefe(a) Inmediato</b><br/>{jefe_pdf_texto}", estilo_firma)
    firma_vob         = Paragraph(f"<br/><br/>___________________________<br/><b>Vo.Bo. Titular del Área</b><br/>{NOMBRE_DIRECTORA}", estilo_firma)

    if datos["tipo"] in ["ECO", "CHO"]:
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
def subir_anexo_drive(archivo, folio: str, rfc: str) -> str:
    """Sube un archivo a Google Drive y retorna el link de visualización."""
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

        ext = archivo.name.split(".")[-1].lower()
        nombre_archivo = f"{folio}_{rfc}.{ext}"
        media = MediaIoBaseUpload(io.BytesIO(archivo.read()), mimetype=archivo.type)
        file_metadata = {
            "name": nombre_archivo,
            "parents": [DRIVE_ANEXOS_FOLDER],
        }
        archivo_drive = service.files().create(
            body=file_metadata,
            media_body=media,
            fields="id",
            supportsAllDrives=True
        ).execute()

        file_id = archivo_drive.get("id")
        try:
            service.permissions().create(
                fileId=file_id,
                body={"type": "anyone", "role": "reader"},
                supportsAllDrives=True
            ).execute()
        except Exception:
            pass  # En Shared Drive los permisos los maneja el admin

        return f"https://drive.google.com/file/d/{file_id}/view"
    except Exception as e:
        return f"ERROR: {e}"

def guardar_dia_economico(datos: dict):
    """Guarda día económico en tab Solicitudes del Sheet económicos — igual que app3."""
    client = get_client()
    sh = client.open_by_key(st.secrets["sheet_economicos_id"])
    ws = sh.worksheet("Solicitudes")
    todas = ws.get_all_records(numericise_ignore=["all"])
    nuevo_id = len(todas) + 1
    emp_id = ""
    try:
        empleados = cargar_empleados()
        match = empleados[empleados["RFC"].astype(str).str.upper() == datos["rfc"].upper()]
        if not match.empty:
            emp_id = match.iloc[0].get("ID", "")
    except Exception:
        pass
    fila = [
        nuevo_id,                    # ID
        emp_id,                      # EmpleadoID
        datos["rfc"],                 # RFC
        datos["nombre"],              # Nombre Completo
        "economico",                 # Tipo Permiso
        datos["fecha_inicio"],        # Fecha Inicio
        datos["fecha_fin"],           # Fecha Fin
        datos["dias"],                # Dias Solicitados
        datos["motivo"],              # Motivo
        datos["fecha_solicitud"],     # Fecha Registro
        "",                          # Aprobado Por
        datos["nombre"],              # Registrado Por
        datos["folio"],               # FOLIO
    ]
    ws.append_row(fila, value_input_option="USER_ENTERED")

def guardar_incidencia(datos: dict):
    client = get_client()
    sh = client.open_by_key(st.secrets["sheet_checador_id"])
    ws = sh.worksheet("Incidencias")
    todas = ws.get_all_records(numericise_ignore=["all"])
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
        datos["motivo"],
        "SÍ" if datos["tiene_anexo"] else "NO",
        datos.get("link_anexo", ""),
        "PENDIENTE",
        "", "", "",
    ]
    ws.append_row(fila, value_input_option="USER_ENTERED")
    cargar_incidencias.clear()

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
        for i, row in enumerate(data, start=2):
            if str(row.get("FOLIO", "")) == folio:
                ws.update_cell(i, headers.index("ESTADO") + 1,             "AUTORIZADO")
                ws.update_cell(i, headers.index("AUTORIZADO_POR") + 1,     "admin")
                ws.update_cell(i, headers.index("FECHA_AUTORIZACION") + 1, ahora)
                if obs and "OBSERVACIONES" in headers:
                    ws.update_cell(i, headers.index("OBSERVACIONES") + 1, obs)
                break
        cargar_incidencias.clear()
    except gspread.exceptions.APIError:
        st.error("⏳ Google Sheets está saturado temporalmente. Espera 5 segundos e intenta de nuevo.")
    except Exception as e:
        st.error(f"Error al autorizar: {e}")

def rechazar_incidencia(folio: str, obs: str):
    client = get_client()
    sh = client.open_by_key(st.secrets["sheet_checador_id"])
    ws = sh.worksheet("Incidencias")
    headers = ws.row_values(1)
    data = ws.get_all_records(numericise_ignore=["all"])
    for i, row in enumerate(data, start=2):
        if str(row.get("FOLIO", "")) == folio:
            ws.update_cell(i, headers.index("ESTADO") + 1,             "RECHAZADO")
            ws.update_cell(i, headers.index("AUTORIZADO_POR") + 1,     "admin")
            ws.update_cell(i, headers.index("FECHA_AUTORIZACION") + 1, datetime.now(__import__("pytz").utc).astimezone(__import__("pytz").timezone("America/Mexico_City")).strftime("%Y-%m-%d %H:%M"))
            ws.update_cell(i, headers.index("OBSERVACIONES") + 1,      obs)
            break
    cargar_incidencias.clear()

def rechazar_dia_economico(row_idx: int, obs: str):
    try:
        import pytz
        tz_mx = pytz.timezone("America/Mexico_City")
        ahora = datetime.now(pytz.utc).astimezone(tz_mx).strftime("%Y-%m-%d %H:%M")
        client = get_client()
        sh = client.open_by_key(st.secrets["sheet_economicos_id"])
        ws = sh.worksheet("Solicitudes")
        headers = [h.upper().strip() for h in ws.row_values(1)]
        col_aprobado = headers.index("APROBADO POR") + 1 if "APROBADO POR" in headers else None
        col_motivo   = headers.index("MOTIVO") + 1 if "MOTIVO" in headers else None
        if col_aprobado:
            ws.update_cell(row_idx, col_aprobado, f"RECHAZADO — {obs}")
        cargar_solicitudes_eco.clear()
        st.warning("Solicitud rechazada.")
    except Exception as e:
        st.error(f"Error al rechazar: {e}")

def aprobar_dia_economico(row_idx: int, nombre_admin: str):
    """Escribe el nombre del admin en Aprobado Por de la tab Solicitudes."""
    try:
        import pytz
        tz_mx = pytz.timezone("America/Mexico_City")
        ahora = datetime.now(pytz.utc).astimezone(tz_mx).strftime("%Y-%m-%d %H:%M")
        client = get_client()
        sh = client.open_by_key(st.secrets["sheet_economicos_id"])
        ws = sh.worksheet("Solicitudes")
        headers = [h.upper().strip() for h in ws.row_values(1)]
        col_aprobado = headers.index("APROBADO POR") + 1 if "APROBADO POR" in headers else None
        col_fecha    = headers.index("FECHA REGISTRO") + 1 if "FECHA REGISTRO" in headers else None
        if col_aprobado:
            ws.update_cell(row_idx, col_aprobado, nombre_admin)
        if col_fecha:
            ws.update_cell(row_idx, col_fecha, ahora)
        cargar_solicitudes_eco.clear()
        st.success("Día económico autorizado correctamente.")
    except Exception as e:
        st.error(f"Error al aprobar: {e}")

def autorizar_cambio_horario(emp_id: str, horario_nuevo: dict, folio: str, fecha_inicio_cho: str = None):
    client = get_client()
    sh = client.open_by_key(st.secrets["sheet_checador_id"])
    ws_emp = sh.worksheet("empleados")
    registros = ws_emp.get_all_records(numericise_ignore=["all"])
    headers   = ws_emp.row_values(1)
    fecha_hoy = datetime.now().strftime("%Y-%m-%d")
    fecha_aplicacion = fecha_inicio_cho if fecha_inicio_cho else fecha_hoy

    for i, row in enumerate(registros, start=2):
        if str(row.get("RFC", "")).upper().strip() == str(emp_id).upper().strip() or str(row.get("NOMBRE", "")).upper().strip() == str(emp_id).upper().strip():
            # Guardar horario ANTERIOR en HISTORIAL_HORARIOS antes de sobrescribir
            try:
                ws_hist = sh.worksheet("HISTORIAL_HORARIOS")
                hist_data = ws_hist.get_all_records(numericise_ignore=["all"])
                fila_hist = [
                    str(row.get("RFC", emp_id)),
                    str(row.get("NOMBRE", "")),
                    "",  # FECHA_INICIO — se llenará con la del registro anterior
                    (datetime.strptime(fecha_aplicacion, "%Y-%m-%d") - timedelta(days=1)).strftime("%Y-%m-%d"),  # FECHA_FIN — día anterior al cambio
                    str(row.get("ENTRADA_LUN", "")), str(row.get("SALIDA_LUN", "")),
                    str(row.get("ENTRADA_MAR", "")), str(row.get("SALIDA_MAR", "")),
                    str(row.get("ENTRADA_MIE", "")), str(row.get("SALIDA_MIE", "")),
                    str(row.get("ENTRADA_JUE", "")), str(row.get("SALIDA_JUE", "")),
                    str(row.get("ENTRADA_VIE", "")), str(row.get("SALIDA_VIE", "")),
                ]
                ws_hist.append_row(fila_hist, value_input_option="USER_ENTERED")
                # Agregar también el nuevo horario con FECHA_INICIO
                fila_nuevo = [
                    str(row.get("RFC", emp_id)),
                    str(row.get("NOMBRE", "")),
                    str(fecha_aplicacion),  # FECHA_INICIO
                    "",  # FECHA_FIN — vacío porque es el vigente
                    horario_nuevo.get("LUN",{}).get("entrada",""), horario_nuevo.get("LUN",{}).get("salida",""),
                    horario_nuevo.get("MAR",{}).get("entrada",""), horario_nuevo.get("MAR",{}).get("salida",""),
                    horario_nuevo.get("MIE",{}).get("entrada",""), horario_nuevo.get("MIE",{}).get("salida",""),
                    horario_nuevo.get("JUE",{}).get("entrada",""), horario_nuevo.get("JUE",{}).get("salida",""),
                    horario_nuevo.get("VIE",{}).get("entrada",""), horario_nuevo.get("VIE",{}).get("salida",""),
                ]
                ws_hist.append_row(fila_nuevo, value_input_option="USER_ENTERED")
            except Exception as e:
                st.warning(f"No se pudo guardar historial: {e}")

            # Sobrescribir horario actual en tab empleados
            for dia, (col_e, col_s) in COLUMNAS_HORARIO.items():
                if dia in horario_nuevo:
                    if col_e in headers:
                        ws_emp.update_cell(i, headers.index(col_e) + 1, horario_nuevo[dia]["entrada"])
                    if col_s in headers:
                        ws_emp.update_cell(i, headers.index(col_s) + 1, horario_nuevo[dia]["salida"])
            break
    ws_inc = sh.worksheet("Incidencias")
    inc_h  = ws_inc.row_values(1)
    inc_d  = ws_inc.get_all_records(numericise_ignore=["all"])
    for i, row in enumerate(inc_d, start=2):
        if str(row.get("FOLIO", "")) == folio:
            ws_inc.update_cell(i, inc_h.index("ESTADO") + 1,             "AUTORIZADO")
            ws_inc.update_cell(i, inc_h.index("AUTORIZADO_POR") + 1,     "admin")
            ws_inc.update_cell(i, inc_h.index("FECHA_AUTORIZACION") + 1, datetime.now(__import__("pytz").utc).astimezone(__import__("pytz").timezone("America/Mexico_City")).strftime("%Y-%m-%d %H:%M"))
            break
    st.cache_data.clear()

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

        admin_correo = st.secrets.get("admin_correo", "")
        admin_rfc    = st.secrets.get("admin_rfc", "")
        if correo.lower() == admin_correo.lower() and rfc_input.upper() == admin_rfc.upper():
            st.session_state["rol"]    = "admin"
            st.session_state["correo"] = correo
            st.session_state["rfc"]    = rfc_input.upper()
            st.session_state["nombre"] = "Administrador RH"
            st.rerun()
            return

        usuarios  = cargar_usuarios()
        match_usr = usuarios[usuarios["RFC"].astype(str).str.upper() == rfc_input.upper()]
        if match_usr.empty:
            st.error("RFC no encontrado. Verifica tus datos.")
            return

        usr_row           = match_usr.iloc[0]
        correo_registrado = str(usr_row.get("Correo electrónico institucional", "")).strip().lower()
        if correo_registrado and correo.strip().lower() != correo_registrado:
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
        jefe_pdf_auto = DICT_JEFES.get(jefe_auto, jefe_auto) if jefe_auto else ""
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
def render_checador():
    import xlrd
    import io as _io
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    DIAS_NUM_CH  = {0:"LUN",1:"MAR",2:"MIE",3:"JUE",4:"VIE",5:"SAB",6:"DOM"}
    DIAS_NOMBRE  = {"LUN":"Lunes","MAR":"Martes","MIE":"Miércoles","JUE":"Jueves","VIE":"Viernes","SAB":"Sábado","DOM":"Domingo"}

    def get_client_ch():
        creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=SCOPES)
        return gspread.authorize(creds)

    @st.cache_data(ttl=300)
    def load_emp_ch():
        client = get_client_ch()
        ws = client.open_by_key(st.secrets["sheet_checador_id"]).worksheet("empleados")
        data = ws.get_all_records(numericise_ignore=["all"])
        df = pd.DataFrame(data).fillna("")
        df["ID"] = df["ID"].astype(str).str.strip()
        if "ACTIVO" not in df.columns:
            df["ACTIVO"] = "SI"
        return df

    @st.cache_data(ttl=300)
    def load_fest_ch():
        client = get_client_ch()
        ws = client.open_by_key(st.secrets["sheet_checador_id"]).worksheet("festivos")
        data = ws.get_all_records()
        df = pd.DataFrame(data).fillna("")
        fechas = set()
        if "FECHA_INICIO" in df.columns:
            for _, r in df.iterrows():
                fi = pd.to_datetime(r.get("FECHA_INICIO",""), errors="coerce")
                ff = pd.to_datetime(r.get("FECHA_FIN","") or r.get("FECHA_INICIO",""), errors="coerce")
                if pd.isna(fi): continue
                if pd.isna(ff): ff = fi
                d = fi
                while d <= ff:
                    fechas.add(d.date())
                    d += timedelta(days=1)
        return fechas

    @st.cache_data(ttl=300)
    def load_hist_ch():
        try:
            client = get_client_ch()
            ws = client.open_by_key(st.secrets["sheet_checador_id"]).worksheet("HISTORIAL_HORARIOS")
            data = ws.get_all_records(numericise_ignore=["all"])
            return pd.DataFrame(data).fillna("") if data else pd.DataFrame()
        except:
            return pd.DataFrame()

    @st.cache_data(ttl=300)
    def load_economicos_ch(fi_str, ff_str):
        try:
            client = get_client_ch()
            ws = client.open_by_key(st.secrets["sheet_economicos_id"]).worksheet("Solicitudes")
            data = ws.get_all_records(numericise_ignore=["all"])
            df = pd.DataFrame(data).fillna("")
            if df.empty: return pd.DataFrame()
            df["Fecha Inicio"] = pd.to_datetime(df["Fecha Inicio"], errors="coerce", dayfirst=True)
            df["Fecha Fin"]    = pd.to_datetime(df["Fecha Fin"],    errors="coerce", dayfirst=True)
            fi = datetime.strptime(fi_str, "%Y-%m-%d")
            ff = datetime.strptime(ff_str, "%Y-%m-%d")
            mask = (df["Fecha Inicio"] <= ff) & (df["Fecha Fin"] >= fi)
            # Solo aprobados
            aprobados = df[mask & (df["Aprobado Por"].astype(str).str.strip() != "")]
            return aprobados
        except Exception as e:
            return pd.DataFrame()

    @st.cache_data(ttl=300)
    def load_incapacidades_ch(fi_str, ff_str):
        try:
            client = get_client_ch()
            ws = client.open_by_key(st.secrets["sheet_economicos_id"]).worksheet("Incapacidades")
            data = ws.get_all_records(numericise_ignore=["all"])
            df = pd.DataFrame(data).fillna("")
            if df.empty: return pd.DataFrame()
            df["Fecha Inicio"]  = pd.to_datetime(df["Fecha Inicio"],  errors="coerce", dayfirst=True)
            df["Fecha Termino"] = pd.to_datetime(df["Fecha Termino"], errors="coerce", dayfirst=True)
            fi = datetime.strptime(fi_str, "%Y-%m-%d")
            ff = datetime.strptime(ff_str, "%Y-%m-%d")
            mask = (df["Fecha Inicio"] <= ff) & (df["Fecha Termino"] >= fi)
            return df[mask].copy()
        except:
            return pd.DataFrame()

    @st.cache_data(ttl=300)
    def load_incidencias_ch(fi_str, ff_str):
        """Carga pases y comisiones autorizados del Sheet del checador."""
        try:
            client = get_client_ch()
            ws = client.open_by_key(st.secrets["sheet_checador_id"]).worksheet("Incidencias")
            data = ws.get_all_records(numericise_ignore=["all"])
            df = pd.DataFrame(data).fillna("")
            if df.empty: return pd.DataFrame()
            df = df[df["ESTADO"] == "AUTORIZADO"]
            df["FECHA_INICIO"] = pd.to_datetime(df["FECHA_INICIO"], errors="coerce")
            df["FECHA_FIN"]    = pd.to_datetime(df["FECHA_FIN"],    errors="coerce")
            fi = datetime.strptime(fi_str, "%Y-%m-%d")
            ff = datetime.strptime(ff_str, "%Y-%m-%d")
            mask = (df["FECHA_INICIO"] <= ff) & (df["FECHA_FIN"] >= fi)
            return df[mask].copy()
        except:
            return pd.DataFrame()

    def build_justificantes_rfc(economicos, incapacidades, incidencias_df, fi, ff):
        """
        Retorna dict: {RFC_upper -> {date -> tipo_justificante}}
        Incluye: días económicos, incapacidades, comisiones, pases de entrada
        """
        resultado = {}

        # Días económicos aprobados
        if not economicos.empty:
            col_rfc = next((c for c in economicos.columns if "RFC" in c.upper()), None)
            col_fi  = next((c for c in economicos.columns if "INICIO" in c.upper()), None)
            col_ff  = next((c for c in economicos.columns if "FIN" in c.upper() and "INICIO" not in c.upper()), None)
            if col_rfc and col_fi and col_ff:
                for _, r in economicos.iterrows():
                    rfc = str(r.get(col_rfc,"")).strip().upper()
                    if not rfc: continue
                    try:
                        d = r[col_fi]
                        f = r[col_ff]
                        while d <= f:
                            if fi <= d.date() <= ff:
                                resultado.setdefault(rfc, {})[d.date()] = "Día económico"
                            d += timedelta(days=1)
                    except: pass

        # Incapacidades
        if not incapacidades.empty:
            col_rfc  = next((c for c in incapacidades.columns if "RFC" in c.upper()), None)
            col_fi   = next((c for c in incapacidades.columns if "INICIO" in c.upper()), None)
            col_ff   = next((c for c in incapacidades.columns if "TERMINO" in c.upper() or ("FIN" in c.upper() and "INICIO" not in c.upper())), None)
            col_tipo = next((c for c in incapacidades.columns if "TIPO" in c.upper()), None)
            if col_rfc and col_fi and col_ff:
                for _, r in incapacidades.iterrows():
                    rfc = str(r.get(col_rfc,"")).strip().upper()
                    if not rfc: continue
                    tipo = str(r.get(col_tipo,"Incapacidad")).strip() if col_tipo else "Incapacidad"
                    try:
                        d = r[col_fi]
                        f = r[col_ff]
                        while d <= f:
                            if fi <= d.date() <= ff:
                                resultado.setdefault(rfc, {})[d.date()] = tipo or "Incapacidad"
                            d += timedelta(days=1)
                    except: pass

        # Comisiones y pases autorizados desde Incidencias
        if not incidencias_df.empty:
            for _, r in incidencias_df.iterrows():
                rfc  = str(r.get("RFC","")).strip().upper()
                tipo = str(r.get("TIPO","")).strip()
                if not rfc: continue
                try:
                    d = r["FECHA_INICIO"]
                    f = r["FECHA_FIN"]
                    if pd.isna(d) or pd.isna(f): continue
                    while d <= f:
                        if fi <= d.date() <= ff:
                            if tipo == "COM":
                                resultado.setdefault(rfc, {})[d.date()] = "Comisión"
                            elif tipo == "PSE":
                                motivo_pse = str(r.get("MOTIVO","")).lower()
                                hora_ret   = str(r.get("HORA_RETORNO","")).strip()
                                if "sin retorno" in motivo_pse:
                                    resultado.setdefault(rfc, {})[d.date()] = "Pase de salida sin retorno"
                                elif hora_ret:
                                    resultado.setdefault(rfc, {})[d.date()] = f"Pase de salida|retorno:{hora_ret}"
                                else:
                                    resultado.setdefault(rfc, {}).setdefault(d.date(), "Pase de entrada")
                        d += timedelta(days=1)
                except: pass

        return resultado

    def get_horario_fecha(emp_row, fecha, historial_df):
        rfc = str(emp_row.get("RFC","")).upper().strip()
        if historial_df.empty or not rfc:
            return emp_row
        hist_emp = historial_df[historial_df["RFC"].astype(str).str.upper().str.strip() == rfc]
        if hist_emp.empty:
            return emp_row
        for _, h in hist_emp.sort_values("FECHA_INICIO", ascending=False).iterrows():
            try:
                fi_h = datetime.strptime(str(h["FECHA_INICIO"]), "%Y-%m-%d").date() if h["FECHA_INICIO"] else None
                ff_h = datetime.strptime(str(h["FECHA_FIN"]),    "%Y-%m-%d").date() if h["FECHA_FIN"]    else None
                if fi_h and ff_h and fi_h <= fecha <= ff_h:
                    return h
                if fi_h and not ff_h and fecha >= fi_h:
                    return h
            except: pass
        return emp_row

    def parse_checada_ch(val):
        s = str(val).strip()
        if not s or s == "nan" or len(s) < 5:
            return None, None
        return s[:5], (s[5:10] if len(s) >= 10 else "")

    def parse_report_ch(file_bytes, empleados, festivos, historial_df, just_rfc, umbral=10):
        wb = xlrd.open_workbook(file_contents=file_bytes)
        if "Reporte de Asistencia" not in wb.sheet_names():
            raise ValueError("No encontré hoja 'Reporte de Asistencia'.")
        sh = wb.sheet_by_name("Reporte de Asistencia")
        periodo = ""
        for i in range(5):
            for j in range(sh.ncols):
                v = str(sh.cell_value(i,j)).strip()
                if "~" in v and "-" in v:
                    periodo = v; break
        fecha_ini = datetime.strptime(periodo.split("~")[0].strip(), "%Y-%m-%d")
        fecha_fin = datetime.strptime(periodo.split("~")[1].strip(), "%Y-%m-%d")
        dias = []
        d = fecha_ini
        while d <= fecha_fin:
            dias.append(d); d += timedelta(days=1)
        col_to_day = {}
        for i in range(5):
            row = sh.row_values(i)
            if sum(1 for v in row if isinstance(v, float) and 1 <= v <= 31) >= 10:
                for j, v in enumerate(row):
                    if isinstance(v, float) and 1 <= v <= 31:
                        col_to_day[j] = int(v)
                break
        checadas = {}
        i = 0
        while i < sh.nrows:
            row = sh.row_values(i)
            if str(row[0]).strip() == "ID:" and str(row[4]).strip().replace(".0",""):
                uid = str(row[4]).strip().replace(".0","")
                checadas[uid] = {}
                if i+1 < sh.nrows:
                    dr = sh.row_values(i+1)
                    for col, dn in col_to_day.items():
                        val = str(dr[col]).strip() if col < len(dr) else ""
                        if val and val != "nan":
                            e, s = parse_checada_ch(val)
                            if e: checadas[uid][dn] = (e, s)
                i += 2; continue
            i += 1

        activos = empleados[empleados["ACTIVO"]=="SI"]
        resumen, detalle_faltas, detalle_retardos = [], [], []

        for _, emp_row in activos.iterrows():
            uid    = emp_row["ID"]
            nombre = emp_row["NOMBRE"]
            rfc    = str(emp_row.get("RFC","")).upper().strip()
            dias_prog = asistidos = faltas = retardos = retardos_min = justif = 0
            uid_ch     = checadas.get(uid, {})
            just_emp   = just_rfc.get(rfc, {})
            dias_falta = []
            dias_just  = []
            det_retardos_emp = []

            for dia_dt in dias:
                if dia_dt.date() in festivos:
                    continue
                nd = DIAS_NUM_CH[dia_dt.weekday()]
                hor = get_horario_fecha(emp_row, dia_dt.date(), historial_df)
                ep  = str(hor.get(f"ENTRADA_{nd}","")).strip()
                if not ep:
                    continue
                dias_prog += 1
                ch  = uid_ch.get(dia_dt.day)
                just_tipo = just_emp.get(dia_dt.date())

                if not ch:
                    if just_tipo and not just_tipo.startswith("Pase de entrada"):
                        justif += 1
                        dias_just.append(f"{dia_dt.strftime('%d/%m')} ({just_tipo.split('|')[0]})")
                    else:
                        faltas += 1
                        dias_falta.append(dia_dt.strftime("%d/%m"))
                    continue

                asistidos += 1
                entrada_real, salida_real = ch
                mins   = 0
                estado = "Asistió"
                try:
                    ep_dt = datetime.strptime(ep, "%H:%M")
                    er_dt = datetime.strptime(entrada_real, "%H:%M")
                    mins  = (er_dt.hour - ep_dt.hour)*60 + (er_dt.minute - ep_dt.minute)
                    if mins > umbral:
                        if just_tipo and just_tipo.startswith("Pase de entrada"):
                            estado = "Asistió"  # pase de entrada: no retardo
                        else:
                            retardos += 1
                            retardos_min += mins
                            estado = "Retardo"
                            det_retardos_emp.append({
                                "Empleado": nombre,
                                "Fecha": dia_dt.strftime("%d/%m/%Y"),
                                "Día": DIAS_NOMBRE.get(nd, nd),
                                "Hora prog.": ep,
                                "Hora real": entrada_real,
                                "Minutos tarde": mins,
                                "Observación": ""
                            })
                except: pass

                # Detectar regreso tardío en pase de salida con retorno
                if just_tipo and "retorno:" in str(just_tipo) and salida_real:
                    try:
                        hora_ret_aut = just_tipo.split("retorno:")[1].strip()
                        ret_aut_dt   = datetime.strptime(hora_ret_aut, "%H:%M")
                        sal_real_dt  = datetime.strptime(salida_real,  "%H:%M")
                        mins_tarde   = (sal_real_dt.hour - ret_aut_dt.hour)*60 + (sal_real_dt.minute - ret_aut_dt.minute)
                        if mins_tarde > umbral:
                            retardos += 1
                            retardos_min += mins_tarde
                            det_retardos_emp.append({
                                "Empleado": nombre,
                                "Fecha": dia_dt.strftime("%d/%m/%Y"),
                                "Día": DIAS_NOMBRE.get(nd, nd),
                                "Hora prog.": hora_ret_aut,
                                "Hora real": salida_real,
                                "Minutos tarde": mins_tarde,
                                "Observación": "Regreso tardío de pase de salida"
                            })
                    except: pass

                detalle_faltas.append({
                    "nombre": nombre, "fecha": dia_dt.date(), "nd": nd,
                    "prog_entrada": ep, "checada_entrada": entrada_real,
                    "checada_salida": salida_real, "retardo_min": max(0,mins),
                    "estado": estado, "justificante": just_tipo or ""
                })
                detalle_retardos.extend(det_retardos_emp)
                det_retardos_emp = []

            pct = round((asistidos+justif)/dias_prog*100) if dias_prog > 0 else 0
            resumen.append({
                "Empleado": nombre,
                "Días prog.": dias_prog,
                "Asistidos": asistidos,
                "Faltas": faltas,
                "Justificadas": justif,
                "Retardos": retardos,
                "% Asistencia": f"{pct}%",
                "Días que faltó": ", ".join(dias_falta) if dias_falta else "—",
                "Días justificados": ", ".join(dias_just) if dias_just else "—",
            })

        return pd.DataFrame(resumen), pd.DataFrame(detalle_retardos), periodo, fecha_ini, fecha_fin

    def generar_excel_reporte(df_res, df_ret, periodo, fecha_ini, fecha_fin, umbral):
        wb = Workbook()
        AZUL    = "002F6C"
        BLANCO  = "FFFFFF"
        GRIS    = "F4F6F9"
        VERDE   = "D4EDDA"
        AMARILLO= "FFF3CD"
        ROJO    = "F8D7DA"
        AZUL_C  = "D0E4F7"

        def hdr_font(bold=True, color=BLANCO, sz=11):
            return Font(bold=bold, color=color, size=sz, name="Arial")
        def fill(color):
            return PatternFill("solid", fgColor=color)
        def center():
            return Alignment(horizontal="center", vertical="center", wrap_text=True)
        def left():
            return Alignment(horizontal="left", vertical="center", wrap_text=True)
        def border():
            s = Side(style="thin", color="CCCCCC")
            return Border(left=s, right=s, top=s, bottom=s)

        # ── Hoja 1: Reporte Ejecutivo ──────────────────────────────────────
        ws1 = wb.active
        ws1.title = "Reporte Ejecutivo"
        mes_nombre = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"][fecha_ini.month-1]

        ws1.merge_cells("A1:I1")
        ws1["A1"] = "SECRETARÍA DE EDUCACIÓN JALISCO"
        ws1["A1"].font = Font(bold=True, size=13, name="Arial", color=AZUL)
        ws1["A1"].alignment = center()

        ws1.merge_cells("A2:I2")
        ws1["A2"] = "DIRECCIÓN DE FORMACIÓN CONTINUA"
        ws1["A2"].font = Font(bold=True, size=12, name="Arial", color=AZUL)
        ws1["A2"].alignment = center()

        ws1.merge_cells("A3:I3")
        ws1["A3"] = f"REPORTE DE ASISTENCIA — {mes_nombre.upper()} {fecha_ini.year}"
        ws1["A3"].font = Font(bold=True, size=11, name="Arial")
        ws1["A3"].alignment = center()

        ws1.merge_cells("A4:I4")
        ws1["A4"] = (f"Período: {fecha_ini.strftime('%d/%m/%Y')} — {fecha_fin.strftime('%d/%m/%Y')}  |  "
                     f"Justificantes: días económicos, incapacidades, comisiones y pases aplicados  |  "
                     f"Tolerancia retardo: {umbral} min")
        ws1["A4"].font = Font(italic=True, size=9, name="Arial", color="555555")
        ws1["A4"].alignment = center()
        ws1.row_dimensions[4].height = 30

        ws1.append([])

        hdrs = ["Empleado","Días prog.","Asistidos","Faltas","Justificadas","Retardos","% Asistencia","Días que faltó","Días justificados"]
        ws1.append(hdrs)
        hrow = ws1.max_row
        for col, _ in enumerate(hdrs, 1):
            c = ws1.cell(hrow, col)
            c.font      = hdr_font()
            c.fill      = fill(AZUL)
            c.alignment = center()
            c.border    = border()

        for _, r in df_res.sort_values("Faltas", ascending=False).iterrows():
            faltas = r["Faltas"]
            pct    = int(str(r["% Asistencia"]).replace("%","") or 0)
            bg = ROJO if faltas >= 10 or pct < 75 else (AMARILLO if faltas >= 3 or pct < 90 else (VERDE if r["Justificadas"] > 0 else BLANCO))
            row_data = [r["Empleado"], r["Días prog."], r["Asistidos"], r["Faltas"],
                        r["Justificadas"], r["Retardos"], r["% Asistencia"],
                        r["Días que faltó"], r["Días justificados"]]
            ws1.append(row_data)
            drow = ws1.max_row
            for col in range(1, 10):
                c = ws1.cell(drow, col)
                c.fill      = fill(bg)
                c.border    = border()
                c.font      = Font(size=9, name="Arial")
                c.alignment = center() if col > 1 else left()

        ws1.append([])
        ws1.append(["Leyenda:", "0-2 faltas/≥90%", "3-9 faltas/75-89%", "≥10 faltas/<75%", "Justificado"])
        lrow = ws1.max_row
        for col, bg in [(2,BLANCO),(3,AMARILLO),(4,ROJO),(5,VERDE)]:
            c = ws1.cell(lrow, col)
            c.fill = fill(bg)
            c.font = Font(size=8, name="Arial", bold=True)
            c.alignment = center()

        ws1.column_dimensions["A"].width = 35
        for col in ["B","C","D","E","F","G"]:
            ws1.column_dimensions[col].width = 12
        ws1.column_dimensions["H"].width = 45
        ws1.column_dimensions["I"].width = 45

        # ── Hoja 2: Detalle faltas ──────────────────────────────────────────
        ws2 = wb.create_sheet("Detalle faltas y justificantes")
        ws2.merge_cells("A1:E1")
        ws2["A1"] = f"DETALLE DE FALTAS Y JUSTIFICANTES — {mes_nombre.upper()} {fecha_ini.year}"
        ws2["A1"].font = Font(bold=True, size=11, name="Arial", color=AZUL)
        ws2["A1"].alignment = center()
        ws2.append([])

        hdrs2 = ["Empleado","Días prog.","Faltas","Justificadas","Días que faltó / Justificante"]
        ws2.append(hdrs2)
        hrow2 = ws2.max_row
        for col, _ in enumerate(hdrs2, 1):
            c = ws2.cell(hrow2, col)
            c.font = hdr_font(); c.fill = fill(AZUL); c.alignment = center(); c.border = border()

        df_faltas = df_res[df_res["Faltas"] > 0].sort_values("Faltas", ascending=False)
        for _, r in df_faltas.iterrows():
            detalle_txt = ""
            if r["Días que faltó"] != "—":
                detalle_txt += "FALTA: " + r["Días que faltó"]
            if r["Días justificados"] != "—":
                detalle_txt += (" | " if detalle_txt else "") + "JUST: " + r["Días justificados"]
            ws2.append([r["Empleado"], r["Días prog."], r["Faltas"], r["Justificadas"], detalle_txt])
            drow = ws2.max_row
            for col in range(1,6):
                c = ws2.cell(drow, col)
                c.border = border(); c.font = Font(size=9, name="Arial")
                c.alignment = center() if col > 1 else left()

        ws2.column_dimensions["A"].width = 35
        for col in ["B","C","D"]:
            ws2.column_dimensions[col].width = 12
        ws2.column_dimensions["E"].width = 70

        # ── Hoja 3: Detalle retardos ───────────────────────────────────────
        ws3 = wb.create_sheet("Detalle de retardos")
        ws3.merge_cells("A1:G1")
        ws3["A1"] = f"DETALLE DE RETARDOS — {mes_nombre.upper()} {fecha_ini.year}  (tolerancia {umbral} min)"
        ws3["A1"].font = Font(bold=True, size=11, name="Arial", color=AZUL)
        ws3["A1"].alignment = center()
        ws3.append([])

        hdrs3 = ["Empleado","Fecha","Día","Hora prog.","Hora real","Minutos tarde","Observación"]
        ws3.append(hdrs3)
        hrow3 = ws3.max_row
        for col, _ in enumerate(hdrs3, 1):
            c = ws3.cell(hrow3, col)
            c.font = hdr_font(); c.fill = fill(AZUL); c.alignment = center(); c.border = border()

        if not df_ret.empty:
            for _, r in df_ret.iterrows():
                ws3.append([r["Empleado"],r["Fecha"],r["Día"],r["Hora prog."],r["Hora real"],r["Minutos tarde"],r.get("Observación","")])
                drow = ws3.max_row
                for col in range(1,8):
                    c = ws3.cell(drow, col)
                    c.border = border(); c.font = Font(size=9, name="Arial")
                    c.alignment = center() if col > 1 else left()
                    if r["Minutos tarde"] > 30:
                        c.fill = fill(ROJO)
                    elif r["Minutos tarde"] > 15:
                        c.fill = fill(AMARILLO)

        ws3.column_dimensions["A"].width = 35
        for col in ["B","C","D","E","F","G"]:
            ws3.column_dimensions[col].width = 14

        buf = _io.BytesIO()
        wb.save(buf)
        return buf.getvalue()

    # ── UI ────────────────────────────────────────────────────────────────
    st.markdown("### 🕐 Reloj Checador")
    col_up, col_tol = st.columns([3,1])
    with col_up:
        uploaded_xls = st.file_uploader("StandardReport.xls del checador", type=["xls"])
    with col_tol:
        umbral_r = st.number_input("Tolerancia retardo (min)", min_value=1, max_value=30, value=10)

    if uploaded_xls:
        # Pre-leer período
        try:
            import xlrd as _xlrd
            _wb = _xlrd.open_workbook(file_contents=uploaded_xls.getvalue())
            _sh = _wb.sheet_by_name("Reporte de Asistencia")
            _periodo = ""
            for _i in range(5):
                for _j in range(_sh.ncols):
                    _v = str(_sh.cell_value(_i,_j)).strip()
                    if "~" in _v and "-" in _v:
                        _periodo = _v; break
            _fi_str, _ff_str = [x.strip() for x in _periodo.split("~")]
        except Exception as e:
            st.error(f"No se pudo leer el archivo: {e}")
            return

        with st.spinner("Cargando datos..."):
            empleados_ch  = load_emp_ch()
            festivos_ch   = load_fest_ch()
            historial_df  = load_hist_ch()
            economicos    = load_economicos_ch(_fi_str, _ff_str)
            incapacidades = load_incapacidades_ch(_fi_str, _ff_str)
            incidencias_p = load_incidencias_ch(_fi_str, _ff_str)

            fi = datetime.strptime(_fi_str, "%Y-%m-%d").date()
            ff = datetime.strptime(_ff_str, "%Y-%m-%d").date()
            just_rfc = build_justificantes_rfc(economicos, incapacidades, incidencias_p, fi, ff)

        try:
            df_res, df_ret, periodo, fecha_ini, fecha_fin = parse_report_ch(
                uploaded_xls.getvalue(), empleados_ch, festivos_ch,
                historial_df, just_rfc, umbral_r)

            mes_nombre = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio",
                          "Agosto","Septiembre","Octubre","Noviembre","Diciembre"][fecha_ini.month-1]

            st.success(f"📅 {periodo} — {mes_nombre} {fecha_ini.year}")

            c1,c2,c3,c4,c5 = st.columns(5)
            c1.metric("Empleados activos", len(df_res))
            c2.metric("Asistencia prom.", df_res["% Asistencia"].apply(lambda x: int(str(x).replace("%","") or 0)).mean().__round__(1).__str__() + "%")
            c3.metric("Total faltas",   int(df_res["Faltas"].sum()))
            c4.metric("Total retardos", int(df_res["Retardos"].sum()))
            c5.metric("Justificadas",   int(df_res["Justificadas"].sum()))

            st.dataframe(
                df_res.sort_values("Faltas", ascending=False),
                use_container_width=True, hide_index=True
            )

            if not df_ret.empty:
                with st.expander(f"📋 Detalle de retardos ({len(df_ret)} registros)"):
                    st.dataframe(df_ret, use_container_width=True, hide_index=True)

            excel_bytes = generar_excel_reporte(df_res, df_ret, periodo, fecha_ini, fecha_fin, umbral_r)
            st.download_button(
                "⬇️ Descargar reporte Excel (3 hojas)",
                data=excel_bytes,
                file_name=f"Asistencia_{mes_nombre}_{fecha_ini.year}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Error procesando el archivo: {e}")
    else:
        st.info("Sube el StandardReport.xls para procesar la asistencia del período.")



def vista_empleado():
    rfc      = st.session_state["rfc"]
    nombre   = st.session_state["nombre"]
    emp_data = st.session_state.get("empleado_row", {})

    st.markdown(f'### Sesión iniciada como: {nombre}')
    st.caption('Filiación: ' + rfc_oculto(rfc) + ' · ' + emp_data.get('PUESTO', ''))

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
    col1, col2, col3 = st.columns(3)
    col1.metric("📅 Días económicos disponibles", dias_disponibles, delta=f"-{dias_usados} ejercidos")
    col2.metric("⏱️ Horas de pases este mes",     f"{horas_pases}h")
    col3.metric("🚨 Retardos este mes",            0)

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
    with st.expander("📋 Mis solicitudes registradas (último mes)", expanded=False):
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
                if sol_hist.empty:
                    sol_hist = sol_hist_all  # si falla el filtro mostrar todo
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
                lambda x: mapear_emojis_estado("AUTORIZADO" if str(x).strip() != "" else "PENDIENTE")
            )
            sol_hist["FECHA_AUTORIZACION"] = ""

        if not mis_inc.empty:
            mis_inc["ESTADO"] = mis_inc["ESTADO"].apply(mapear_emojis_estado)

        cols_mostrar = ["FOLIO", "TIPO", "FECHA_INICIO", "FECHA_FIN", "DIAS", "HORAS_PASE", "ESTADO", "FECHA_AUTORIZACION"]
        frames = []
        if not sol_hist.empty:
            frames.append(sol_hist[[c for c in cols_mostrar if c in sol_hist.columns]])
        if not mis_inc.empty:
            frames.append(mis_inc[[c for c in cols_mostrar if c in mis_inc.columns]])

        if not frames:
            st.info("Aún no tienes solicitudes registradas.")
        else:
            st.dataframe(pd.concat(frames, ignore_index=True), use_container_width=True, hide_index=True)

    st.divider()

    # ── Faltas informativas ──────────────────────
    with st.expander("📊 Mis faltas este mes", expanded=False):
        st.info("ℹ️ Este contador es informativo. Las faltas no justificadas las gestiona Recursos Humanos.")
        c1, c2, c3 = st.columns(3)
        c1.metric("Faltas totales",        0)
        c2.metric("Faltas justificadas",   0)
        c3.metric("Faltas no justificadas",0)
        st.caption("La integración con el reloj checador estará disponible próximamente.")

    st.divider()
    st.markdown("### Nueva solicitud")

    jefe_guardado = st.session_state.get("jefe_inmediato", "")
    if jefe_guardado:
        jefe_pdf = jefe_guardado
        st.caption(f"👤 Jefe inmediato: {jefe_guardado.split(chr(60))[0]}")
    else:
        jefe_sel = st.selectbox("👤 Jefe inmediato que autoriza", options=list(DICT_JEFES.keys()))
        jefe_pdf = DICT_JEFES[jefe_sel]

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
        col1, col2 = st.columns(2)
        with col1:
            fi = st.date_input("Fecha inicio", value=date.today())
        with col2:
            ff = st.date_input("Fecha fin",    value=date.today())
        if ff < fi:
            st.error("La fecha fin no puede ser anterior a la fecha inicio.")
            return
        festivos = cargar_festivos()
        dias_hab = dias_habiles_entre(fi, ff, festivos)
        st.caption(f"Días hábiles solicitados: **{dias_hab}**")
        if dias_hab > dias_disponibles:
            st.error(f"Excedes tus días disponibles ({dias_disponibles}).")
            return
        motivo = st.text_area("Motivo (opcional)", max_chars=300)
        st.caption("El día económico no requiere anexo digital. Trae el formato físico original.")
        enviar_solicitud(rfc, nombre, tipo, fi, ff, dias_hab, 0.0, motivo, False, incidencias, jefe_inmediato=jefe_pdf)

    # ── PASE DE SALIDA / ENTRADA ────────────────
    elif tipo == "PSE":
        subtipo = st.radio("Subtipo", ["Pase de salida sin retorno", "Pase de entrada", "Pase de salida"])
        fecha   = st.date_input("Fecha del pase", value=date.today())

        col1, col2 = st.columns(2)
        hora_salida = hora_entrada = hora_retorno = None
        with col1:
            if subtipo in ["Pase de salida sin retorno", "Pase de salida"]:
                hora_salida = st.text_input("Hora de salida (HH:MM)", placeholder="08:37", max_chars=5)
                hora_salida = hora_salida.strip() if hora_salida else None
        with col2:
            if subtipo == "Pase de entrada":
                hora_entrada = st.text_input("Hora de entrada (HH:MM)", placeholder="08:37", max_chars=5)
                hora_entrada = hora_entrada.strip() if hora_entrada else None
            elif subtipo == "Pase de salida":
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
                st.caption(f"Horas de ausencia: **{horas_pase}h** · Acumulado mes: **{horas_pases + horas_pase}h**")
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
                    st.caption(f"Horas de ausencia (hasta fin de jornada {salida_prog}): **{horas_pase}h** · Acumulado mes: **{horas_pases + horas_pase}h**")
                else:
                    horas_pase = 0.0
            except:
                horas_pase = 0.0
                st.warning("No se pudo calcular automáticamente.")

        motivo = st.text_area("Motivo", max_chars=300)
        archivo_anexo = st.file_uploader("Adjuntar justificante (opcional)", type=["pdf","png","jpg","jpeg"])
        tiene_anexo   = archivo_anexo is not None
        detalle = subtipo
        if hora_salida:  detalle += f" | Salida: {hora_salida}"
        if hora_entrada: detalle += f" | Entrada: {hora_entrada}"
        if hora_retorno: detalle += f" | Retorno: {hora_retorno}"
        motivo_completo = (detalle + "\n" + motivo).strip()
        enviar_solicitud(rfc, nombre, tipo, fecha, fecha, 0, horas_pase, motivo_completo, tiene_anexo, incidencias, archivo_anexo, subtipo_label=subtipo, hora_retorno=hora_retorno or "", jefe_inmediato=jefe_pdf)

    # ── COMISIÓN ────────────────────────────────
    elif tipo == "COM":
        col1, col2 = st.columns(2)
        with col1:
            fi = st.date_input("Fecha inicio comisión", value=date.today())
        with col2:
            ff = st.date_input("Fecha fin comisión",    value=date.today())
        if ff < fi:
            st.error("La fecha fin no puede ser anterior a la fecha inicio.")
        else:
            festivos = cargar_festivos()
            dias_hab = dias_habiles_entre(fi, ff, festivos)
            st.caption(f"Días de comisión: **{dias_hab}**")
        motivo      = st.text_area("Motivo de la comisión", max_chars=300)
        archivo_anexo = st.file_uploader("Adjuntar constancia/oficio (opcional)", type=["pdf","png","jpg","jpeg"])
        tiene_anexo   = archivo_anexo is not None
        if not tiene_anexo:
            st.caption("Recuerda que sin anexo tu solicitud puede quedar sin soporte documental.")
        enviar_solicitud(rfc, nombre, tipo, fi, ff, dias_hab, 0.0,
                         motivo, tiene_anexo, incidencias, archivo_anexo, jefe_inmediato=jefe_pdf)
    # ── CAMBIO DE HORARIO ───────────────────────
    elif tipo == "CHO":
        fecha_inicio_cho = st.date_input("¿A partir de qué fecha aplica el cambio?", value=date.today())
        motivo      = st.text_area("Motivo del cambio de horario", max_chars=300)
        tiene_anexo = st.checkbox("¿Traerás documento de soporte (oficio, etc.)?")
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
                         motivo_completo, tiene_anexo, incidencias, jefe_inmediato=jefe_pdf)

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
    key_btn = f"btn_registrar_{tipo}_{str(fi)}_{str(ff)}"
    if st.button("Registrar solicitud", type="primary", key=key_btn):
        if st.session_state.get(f"registrado_{key_btn}"):
            st.warning("Esta solicitud ya fue registrada.")
            return
        st.session_state[f"registrado_{key_btn}"] = True
        folio     = generar_folio(tipo, incidencias_df)
        import pytz
        tz_mx = pytz.timezone("America/Mexico_City")
        fecha_sol = datetime.now(pytz.utc).astimezone(tz_mx).strftime("%Y-%m-%d %H:%M")
        link_anexo = ""
        if archivo_anexo is not None:
            with st.spinner("Subiendo anexo a Drive..."):
                link_anexo = subir_anexo_drive(archivo_anexo, folio, rfc)
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
            "hora_retorno":    hora_retorno,
        }
        if tipo == "ECO":
            guardar_dia_economico(datos)
        else:
            guardar_incidencia(datos)
        pdf_bytes = generar_comprobante_pdf(datos)
        st.success(f"✅ Solicitud registrada con folio **{folio}**")
        st.download_button(
            label="⬇️ Descargar comprobante PDF",
            data=pdf_bytes,
            file_name=f"Comprobante_{folio}.pdf",
            mime="application/pdf",
        )
        st.info("Imprime tu comprobante y preséntalo en RH junto con el documento físico original.")

# ─────────────────────────────────────────────
# VISTA ADMIN
# ─────────────────────────────────────────────
def vista_admin():
    st.markdown("### Panel de administración · Incidencias")

    incidencias = cargar_incidencias()
    horarios_df = cargar_horarios()

    tab1, tab2, tab3, tab4 = st.tabs(["Pendientes", "Historial completo", "Reporte mensual", "🕐 Reloj Checador"])

    with tab1:
        solicitudes_eco = cargar_solicitudes_eco()
        col_aprobado = next((c for c in solicitudes_eco.columns if "Aprobado" in c), None)
        col_folio    = next((c for c in solicitudes_eco.columns if "FOLIO" in c.upper() or "Folio" in c), None)
        eco_pend = pd.DataFrame()
        if col_aprobado and not solicitudes_eco.empty:
            eco_pend = solicitudes_eco[solicitudes_eco[col_aprobado].astype(str).str.strip() == ""]

        pendientes = incidencias[incidencias["ESTADO"] == "PENDIENTE"]
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
                    with col_a:
                        if st.button("✅ Aprobar", key=f"eco_{idx}", type="primary"):
                            aprobar_dia_economico(idx + 2, "admin")
                            st.rerun()
                    with col_r:
                        if st.button("❌ Rechazar", key=f"rec_eco_{idx}"):
                            if not obs_eco:
                                st.error("Escribe una observación para rechazar.")
                            else:
                                rechazar_dia_economico(idx + 2, obs_eco)
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
                            autorizar_cambio_horario(row["RFC"], nuevo_horario, row["FOLIO"])
                            st.success("Horario actualizado y solicitud autorizada.")
                            st.rerun()

    with tab2:
        filtro_tipo   = st.selectbox("Filtrar por tipo",   ["TODOS"] + list(TIPO_LABELS.keys()))
        filtro_estado = st.selectbox("Filtrar por estado", ["TODOS", "PENDIENTE", "AUTORIZADO", "RECHAZADO"])
        df_hist = incidencias.copy()
        if filtro_tipo   != "TODOS": df_hist = df_hist[df_hist["TIPO"]   == filtro_tipo]
        if filtro_estado != "TODOS": df_hist = df_hist[df_hist["ESTADO"] == filtro_estado]
        st.dataframe(df_hist, use_container_width=True, hide_index=True)
        st.download_button("⬇️ Exportar CSV", data=df_hist.to_csv(index=False).encode("utf-8"),
                           file_name="incidencias.csv", mime="text/csv")

    with tab4:
        render_checador()

    with tab3:
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
                    Autorizadas=("ESTADO", lambda x: (x=="AUTORIZADO").sum()),
                    Rechazadas=("ESTADO",  lambda x: (x=="RECHAZADO").sum()),
                ).reset_index()
                st.dataframe(detalle, use_container_width=True, hide_index=True)
                st.download_button("⬇️ Descargar CSV",
                                   data=df_rep.to_csv(index=False).encode("utf-8"),
                                   file_name=f"Reporte_{mes_nombre}_{anio_rep}.csv",
                                   mime="text/csv")

# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
def vista_directorio():
    import pytz

    @st.cache_data(ttl=300)
    def cargar_directorio():
        client = get_client()
        sh = client.open_by_key(st.secrets["sheet_checador_id"])
        ws = sh.worksheet("Directorio")
        data = ws.get_all_records(numericise_ignore=["all"])
        return pd.DataFrame(data).fillna("")

    st.markdown("## 📞 Directorio interno DFC 2026")
    st.caption("Toca un área para ver su equipo · Busca por nombre, extensión o correo")

    df = cargar_directorio()
    if df.empty:
        st.warning("No se pudo cargar el directorio.")
        return

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

    def tarjeta(row):
        bg, tc = color(row["AREA"])
        ini = (str(row["NOMBRE"])[0] + (str(row["NOMBRE"]).split()[1][0] if len(str(row["NOMBRE"]).split()) > 1 else "")).upper()
        ext_html = f"📞 **{row['EXTENSION']}**" if row["EXTENSION"] else ""
        email_html = f"" if row["CORREO"] else ""
        dept_html = f" · *{row['DEPARTAMENTO']}*" if row["DEPARTAMENTO"] else ""

        col1, col2 = st.columns([3,1])
        with col1:
            st.markdown(
                f"<div style='display:flex;align-items:center;gap:10px;padding:4px 0'>"
                f"<div style='width:34px;height:34px;border-radius:50%;background:{bg};color:{tc};"
                f"display:flex;align-items:center;justify-content:center;font-weight:600;font-size:12px;flex-shrink:0'>{ini}</div>"
                f"<div><div style='font-size:14px;font-weight:500'>{row['NOMBRE']}</div>"
                f"<div style='font-size:12px;color:gray'>{row['AREA']}{dept_html}</div></div></div>",
                unsafe_allow_html=True
            )
        with col2:
            if ext_html: st.markdown(ext_html)
            if email_html: st.caption(row["CORREO"])
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
        ]
        ICONOS = {
            "Despacho":                    "🏢",
            "Recursos Humanos":            "👥",
            "Viáticos":                    "🧾",
            "Recursos Materiales":         "📦",
            "Licitaciones y Presupuestos": "📋",
            "Comunicación Social":         "📢",
            "Dir. Desarrollo Académico":   "🎓",
        }
        # Normalizar: si DEPARTAMENTO está vacío usar AREA como dept
        df["DEPT_VISTA"] = df.apply(lambda r: r["DEPARTAMENTO"] if r["DEPARTAMENTO"] else r["AREA"], axis=1)

        depts_en_datos = df["DEPT_VISTA"].unique().tolist()
        # Mostrar en orden acordado + cualquier otro que no esté en la lista
        orden_final = [d for d in ORDEN_AREAS if d in depts_en_datos] + [d for d in depts_en_datos if d not in ORDEN_AREAS]

        for dept in orden_final:
            personas = df[df["DEPT_VISTA"] == dept]
            icono = ICONOS.get(dept, "📁")
            bg, tc = color(df[df["DEPT_VISTA"]==dept]["AREA"].iloc[0])
            with st.expander(f"{icono} {dept}  ·  {len(personas)} personas"):
                for _, row in personas.iterrows():
                    tarjeta(row)


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
                st.write(f"**Motivo:** {motivo_reg}")
        else:
            st.error("🚨 ALERTA: DOCUMENTO NO ENCONTRADO O ALTERADO")
            st.warning("Este folio no existe en los registros oficiales de la DFC. El formato impreso podría ser falso o modificado de manera ilícita.")
        if st.button("Ir al inicio de sesión"):
            st.query_params.clear()
            st.rerun()
        return

    if "rol" not in st.session_state:
        st.markdown("<style>[data-testid='stSidebar']{display:none}</style>", unsafe_allow_html=True)
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
        if st.button("📞 Directorio DFC"):
            st.session_state["vista"] = "directorio"
            st.rerun()
        if st.button("🏠 Inicio"):
            st.session_state["vista"] = "inicio"
            st.rerun()
        if st.button("Cerrar sesión"):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()

    vista = st.session_state.get("vista", "inicio")
    if vista == "directorio":
        vista_directorio()
    elif st.session_state["rol"] == "admin":
        vista_admin()
    else:
        vista_empleado()


if __name__ == "__main__":
    main()