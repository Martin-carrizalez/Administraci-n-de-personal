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
    "FECHA_INICIO", "FECHA_FIN", "DIAS", "HORAS_PASE",
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
    """Suma las horas de pases de salida/entrada del mes en curso."""
    ahora = datetime.now()
    df = incidencias_df[
        (incidencias_df["RFC"].astype(str).str.upper() == rfc.upper()) &
        (incidencias_df["TIPO"] == "PSE")
    ].copy()
    if df.empty:
        return 0.0
    df["FECHA_DT"] = pd.to_datetime(df["FECHA_SOLICITUD"], errors="coerce")
    df = df[(df["FECHA_DT"].dt.year == ahora.year) & (df["FECHA_DT"].dt.month == ahora.month)]
    total = 0.0
    for _, row in df.iterrows():
        try:
            total += float(row.get("HORAS_PASE", 0) or 0)
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
    estilo_folio    = ParagraphStyle("f", parent=styles["Normal"], fontSize=16, fontName="Helvetica-Bold", alignment=TA_CENTER, textColor=colors.HexColor("#4B3FA0"), spaceAfter=6)
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
    elementos.append(Spacer(1, 0.2*cm))
    elementos.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#4B3FA0")))
    elementos.append(Spacer(1, 0.3*cm))
    elementos.append(Paragraph("COMPROBANTE DE CAPTURA DE INCIDENCIA", estilo_titulo))
    elementos.append(Spacer(1, 0.2*cm))
    elementos.append(Paragraph(f"FOLIO: {datos['folio']}", estilo_folio))
    elementos.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#CCCCCC")))
    elementos.append(Spacer(1, 0.4*cm))

    tabla_datos = [
        ["Nombre completo:", datos["nombre"]],
        ["Filiación:", rfc_oculto(datos["rfc"])],
        ["Tipo de incidencia:", datos.get("subtipo_label", datos["tipo_label"])],
        ["Fecha de solicitud:", datos["fecha_solicitud"]],
        ["Fecha inicio:", datos["fecha_inicio"]],
        ["Fecha fin:", datos["fecha_fin"]],
        ["Días solicitados:", "N/A" if datos["tipo"] == "PSE" else str(datos["dias"])],
        ["Horas de pase:", (f"{datos.get('horas_pase', 0)}h" if datos.get("horas_pase", 0) else "No registradas") if datos["tipo"] == "PSE" else "N/A"],
        ["Motivo / Descripción:", datos["motivo"]],
        ["Documento anexo:", "Sí, se presentará en RH" if datos["tiene_anexo"] else "No aplica"],
    ]

    t = Table(tabla_datos, colWidths=[5*cm, 11*cm])
    t.setStyle(TableStyle([
        ("FONTNAME",    (0, 0), (0, -1), "Helvetica-Bold"),
        ("FONTNAME",    (1, 0), (1, -1), "Helvetica"),
        ("FONTSIZE",    (0, 0), (-1, -1), 9),
        ("VALIGN",      (0, 0), (-1, -1), "TOP"),
        ("ROWBACKGROUNDS", (0, 0), (-1, -1), [colors.HexColor("#F7F5FF"), colors.white]),
        ("GRID",        (0, 0), (-1, -1), 0.3, colors.HexColor("#DDDDDD")),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING",(0, 0), (-1, -1), 8),
        ("TOPPADDING",  (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 5),
    ]))
    elementos.append(t)
    elementos.append(Spacer(1, 0.6*cm))
    elementos.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#CCCCCC")))
    elementos.append(Spacer(1, 0.3*cm))

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
    elementos.append(Spacer(1, 0.6*cm))
    # ── Bloque de firmas ─────────────────────────
    elementos.append(Paragraph("<b>REQUISITO OBLIGATORIO: SECCIÓN DE FIRMAS</b>",
        ParagraphStyle("tf", parent=styles["Normal"], fontSize=9, fontName="Helvetica-Bold", spaceAfter=6)))

    estilo_firma = ParagraphStyle("firmas", parent=styles["Normal"], fontSize=8, fontName="Helvetica", alignment=TA_CENTER)
    firma_interesado = Paragraph("<br/><br/>___________________________<br/><b>Firma del Interesado</b><br/>Servidor(a) Público(a)", estilo_firma)
    firma_jefe       = Paragraph("<br/><br/>___________________________<br/><b>Autoriza Jefe(a) Inmediato</b><br/>Nombre y Firma", estilo_firma)
    firma_vob        = Paragraph("<br/><br/>___________________________<br/><b>Vo.Bo. Titular del Área</b><br/>Nombre y Firma", estilo_firma)

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
    elementos.append(Spacer(1, 0.3*cm))
    elementos.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#CCCCCC")))
    elementos.append(Spacer(1, 0.3*cm))

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
    elementos.append(Spacer(1, 0.4*cm))
    elementos.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#CCCCCC")))
    elementos.append(Spacer(1, 0.2*cm))
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
        datos["motivo"],
        "SÍ" if datos["tiene_anexo"] else "NO",
        datos.get("link_anexo", ""),
        "PENDIENTE",
        "", "", "",
    ]
    ws.append_row(fila, value_input_option="USER_ENTERED")
    st.cache_data.clear()

def autorizar_incidencia(folio: str, obs: str = ""):
    client = get_client()
    sh = client.open_by_key(st.secrets["sheet_checador_id"])
    ws = sh.worksheet("Incidencias")
    headers = ws.row_values(1)
    data = ws.get_all_records(numericise_ignore=["all"])
    for i, row in enumerate(data, start=2):
        if str(row.get("FOLIO", "")) == folio:
            ws.update_cell(i, headers.index("ESTADO") + 1,             "AUTORIZADO")
            ws.update_cell(i, headers.index("AUTORIZADO_POR") + 1,     "admin")
            ws.update_cell(i, headers.index("FECHA_AUTORIZACION") + 1, datetime.now().strftime("%Y-%m-%d %H:%M"))
            if obs:
                ws.update_cell(i, headers.index("OBSERVACIONES") + 1, obs)
            break
    st.cache_data.clear()

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
            ws.update_cell(i, headers.index("FECHA_AUTORIZACION") + 1, datetime.now().strftime("%Y-%m-%d %H:%M"))
            ws.update_cell(i, headers.index("OBSERVACIONES") + 1,      obs)
            break
    st.cache_data.clear()

def autorizar_cambio_horario(emp_id: str, horario_nuevo: dict, folio: str):
    client = get_client()
    sh = client.open_by_key(st.secrets["sheet_checador_id"])
    ws_emp = sh.worksheet("empleados")
    registros = ws_emp.get_all_records(numericise_ignore=["all"])
    headers   = ws_emp.row_values(1)
    for i, row in enumerate(registros, start=2):
        if str(row.get("RFC", "")).upper().strip() == str(emp_id).upper().strip():
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
            ws_inc.update_cell(i, inc_h.index("FECHA_AUTORIZACION") + 1, datetime.now().strftime("%Y-%m-%d %H:%M"))
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

        st.session_state["rol"]          = "empleado"
        st.session_state["correo"]       = correo
        st.session_state["rfc"]          = rfc_input.upper()
        st.session_state["nombre"]       = nombre_completo
        st.session_state["empleado_row"] = emp_dict
        st.rerun()

# ─────────────────────────────────────────────
# VISTA EMPLEADO
# ─────────────────────────────────────────────
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
    with st.expander("📋 Mis solicitudes registradas", expanded=False):
        mis_inc  = incidencias[incidencias["RFC"].astype(str).str.upper() == rfc].copy()
        sol_hist = solicitudes[solicitudes["RFC"].astype(str).str.upper() == rfc].copy()

        if not sol_hist.empty:
            sol_hist = sol_hist.rename(columns={
                "Tipo Permiso":     "TIPO",
                "Fecha Inicio":     "FECHA_INICIO",
                "Fecha Fin":        "FECHA_FIN",
                "Dias Solicitados": "DIAS",
                "Aprobado Por":     "AUTORIZADO_POR",
                "Fecha Registro":   "FECHA_SOLICITUD",
            })
            sol_hist["FOLIO"]              = "HISTÓRICO"
            sol_hist["HORAS_PASE"]         = ""
            sol_hist["ESTADO"]             = sol_hist["AUTORIZADO_POR"].apply(
                lambda x: "✅ AUTORIZADO" if str(x).strip() != "" else "🟡 PENDIENTE"
            )
            sol_hist["FECHA_AUTORIZACION"] = ""

        if not mis_inc.empty:
            mis_inc["ESTADO"] = mis_inc["ESTADO"].map({
                "AUTORIZADO": "✅ AUTORIZADO",
                "PENDIENTE":  "🟡 PENDIENTE",
                "RECHAZADO":  "🔴 RECHAZADO",
            }).fillna(mis_inc["ESTADO"])

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
        enviar_solicitud(rfc, nombre, tipo, fi, ff, dias_hab, 0.0, motivo, False, incidencias)

    # ── PASE DE SALIDA / ENTRADA ────────────────
    elif tipo == "PSE":
        subtipo = st.radio("Subtipo", ["Pase de salida sin retorno", "Pase de entrada", "Pase de salida"])
        fecha   = st.date_input("Fecha del pase", value=date.today())

        col1, col2 = st.columns(2)
        hora_salida = hora_entrada = hora_retorno = None
        with col1:
            if subtipo in ["Pase de salida sin retorno", "Pase de salida"]:
                hora_salida = st.time_input("Hora de salida")
        with col2:
            if subtipo == "Pase de entrada":
                hora_entrada = st.time_input("Hora de entrada")
            elif subtipo == "Pase de salida":
                hora_retorno = st.time_input("Hora estimada de retorno")

        # Calcular horas del pase
        horas_pase = 0.0
        if hora_salida and hora_retorno:
            sal  = datetime.combine(fecha, hora_salida)
            ret  = datetime.combine(fecha, hora_retorno)
            diff = (ret - sal).total_seconds() / 3600
            horas_pase = round(max(diff, 0), 2)
            st.caption(f"Horas de ausencia estimadas: **{horas_pase}h** · Acumulado del mes: **{horas_pases + horas_pase}h**")
        elif hora_salida:
            horas_pase_inp = st.number_input("Horas aproximadas de ausencia", min_value=0.5, max_value=8.0, step=0.5, value=1.0)
            horas_pase     = horas_pase_inp
            st.caption(f"Acumulado del mes: **{horas_pases + horas_pase}h**")

        motivo      = st.text_area("Motivo", max_chars=300)
        archivo_anexo = st.file_uploader("Adjuntar justificante (opcional)", type=["pdf","png","jpg","jpeg"])
        tiene_anexo   = archivo_anexo is not None

        detalle = subtipo
        if hora_salida:  detalle += f" | Salida: {hora_salida}"
        if hora_entrada: detalle += f" | Entrada: {hora_entrada}"
        if hora_retorno: detalle += f" | Retorno: {hora_retorno}"
        motivo_completo = (detalle + '\n' + motivo).strip()
        enviar_solicitud(rfc, nombre, tipo, fecha, fecha, 0, horas_pase, motivo_completo, tiene_anexo, incidencias, archivo_anexo, subtipo_label=subtipo)


    # ── COMISIÓN ────────────────────────────────
    elif tipo == "COM":
        col1, col2 = st.columns(2)
        with col1:
            fi = st.date_input("Fecha inicio comisión", value=date.today())
        with col2:
            ff = st.date_input("Fecha fin comisión",    value=date.today())
        if ff < fi:
            st.error("La fecha fin no puede ser anterior a la fecha inicio.")
            return
        festivos = cargar_festivos()
        dias_hab = dias_habiles_entre(fi, ff, festivos)
        st.caption(f"Días de comisión: **{dias_hab}**")
        motivo      = st.text_area("Motivo de la comisión", max_chars=300)
        archivo_anexo = st.file_uploader("Adjuntar constancia/oficio (opcional)", type=["pdf","png","jpg","jpeg"])
        tiene_anexo   = archivo_anexo is not None
        if not tiene_anexo:
            st.caption("Recuerda que sin anexo tu solicitud puede quedar sin soporte documental.")
        enviar_solicitud(rfc, nombre, tipo, fi, ff, dias_hab, 0.0,
                         motivo, tiene_anexo, incidencias, archivo_anexo)
    # ── CAMBIO DE HORARIO ───────────────────────
    elif tipo == "CHO":
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
        enviar_solicitud(rfc, nombre, tipo, date.today(), date.today(), 0, 0.0,
                         motivo_completo, tiene_anexo, incidencias)

    # ── CUMPLEAÑOS (oculto hasta autorización) ──
    if HABILITAR_CUMPLEANOS and tipo == "CUM":
        cumple = cumpleanos_laboral(rfc)
        if cumple:
            st.info(f"Tu día de cumpleaños hábil es: **{cumple.strftime('%d/%m/%Y')}**")
            fi = cumple
            ff = cumple
            motivo = "Día de cumpleaños"
            st.caption("Solo el día hábil correspondiente a tu fecha de nacimiento.")
            enviar_solicitud(rfc, nombre, tipo, fi, ff, 1, 0.0, motivo, False, incidencias)
        else:
            st.error("No se pudo calcular tu fecha de cumpleaños desde el RFC.")


def enviar_solicitud(rfc, nombre, tipo, fi, ff, dias, horas_pase, motivo, tiene_anexo, incidencias_df, archivo_anexo=None, subtipo_label=None):
    if st.button("Registrar solicitud", type="primary"):
        folio     = generar_folio(tipo, incidencias_df)
        fecha_sol = datetime.now().strftime("%Y-%m-%d %H:%M")
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
            "subtipo_label":   subtipo_label if subtipo_label else None,
            "subtipo_label":   subtipo_label if subtipo_label else TIPO_LABELS[tipo],
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

    tab1, tab2, tab3 = st.tabs(["Pendientes", "Historial completo", "Reporte mensual"])

    with tab1:
        pendientes = incidencias[incidencias["ESTADO"] == "PENDIENTE"]
        if pendientes.empty:
            st.success("No hay solicitudes pendientes.")
        else:
            st.caption(f"{len(pendientes)} solicitud(es) pendiente(s)")
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
                        nuevo_horario = {}
                        cols_dias = st.columns(7)
                        for idx, dia in enumerate(DIAS_SEMANA):
                            col_e, col_s = COLUMNAS_HORARIO[dia]
                            actual_e = emp_hor.iloc[0].get(col_e, "") if not emp_hor.empty else ""
                            actual_s = emp_hor.iloc[0].get(col_s, "") if not emp_hor.empty else ""
                            with cols_dias[idx]:
                                st.caption(dia)
                                ne = st.text_input("Entrada", value=str(actual_e), key=f"ne_{row['FOLIO']}_{dia}", max_chars=5)
                                ns = st.text_input("Salida",  value=str(actual_s), key=f"ns_{row['FOLIO']}_{dia}", max_chars=5)
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
        if st.button("Cerrar sesión"):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()

    if st.session_state["rol"] == "admin":
        vista_admin()
    else:
        vista_empleado()


if __name__ == "__main__":
    main()