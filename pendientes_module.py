"""
MÓDULO: Pendientes RH DFC + Seguimiento de propuestas — DFC/SEJ
================================================================
Dos apartados en una sola vista, con el Sheet detrás para que lo que tú
cambies lo vean tus colaboradores al instante:

  1. PENDIENTES — la lista de cierre de año. Se puede pegar completa desde
     cualquier lado (un MD, un WhatsApp, lo que sea): el módulo la compara
     con lo que ya hay, agrega lo nuevo, marca lo tachado y NO duplica. Así
     no importa dónde trabajes la lista momento a momento.

  2. PROPUESTAS — seguimiento por etapas: Elaboración → Firma de la
     directora → Firma de la subsecretaria → Concluido. Cada trámite avanza
     con un botón y deja su nota.

Cero datos de servidores públicos escritos en el código: todo vive en el
Sheet, en las tabs Pendientes_RH y Propuestas_RH, que se crean solas la
primera vez.

Conexión desde app_incidencias.py:
    from pendientes_module import render_pendientes_rh
    render_pendientes_rh(get_client)
"""

import streamlit as st
import re
import unicodedata
import urllib.parse
from datetime import datetime, date, timedelta

# ── Tabs y columnas ──────────────────────────────────────────────────────
TAB_PEND = "Pendientes_RH"
COLS_PEND = ["ID", "FECHA_ALTA", "PENDIENTE", "ESTADO",
             "RESPONSABLE", "CREADO_POR", "FECHA_CIERRE", "FECHA_LIMITE"]

MESES_ES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
            "agosto", "septiembre", "octubre", "noviembre", "diciembre"]

TAB_PROP = "Propuestas_RH"
COLS_PROP = ["ID", "NOMBRE", "ETAPA", "SUBSISTEMA", "NOTA",
             "ACTUALIZADO", "CREADO_POR"]

# Los documentos firmados NO van aquí: van al Sheet del minutario, en su
# propia hoja. El minutario es donde ya vive el resguardo de oficios y
# acuses, y partir los documentos en dos lugares es como se pierden.
TAB_DOCS = "propuestas"
COLS_DOCS = ["ID_PROPUESTA", "PROPUESTA", "TIPO", "PERIODO",
             "PERIODO_INICIO", "PERIODO_FIN", "ARCHIVO", "URL",
             "SHA256", "SUBIDO_POR", "FECHA"]
TIPOS_DOC = {"FIRMADA": "Propuesta firmada", "ACUSE": "Acuse"}

MESES_CORTO = ["ene", "feb", "mar", "abr", "may", "jun",
               "jul", "ago", "sep", "oct", "nov", "dic"]

# El orden de las etapas ES el flujo real del trámite. Cambiar esta lista
# cambia el seguimiento completo; la ETAPA se guarda como número (0..3) para
# que renombrar una etapa no invalide lo ya capturado.
ETAPAS = ["Elaboración",
          "Firma de la directora",
          "Firma de la subsecretaria",
          "Entregado a RH",
          "Concluido"]

# Quién recibe en RH según el subsistema. Si un día cambia la persona, se
# cambia aquí y el sello, el botón y la nota se actualizan solos.
RECIBE_EN_RH = {"ESTATAL": "Aurora", "FEDERAL": "Alejandra"}


def quien_recibe(subsistema):
    return RECIBE_EN_RH.get(str(subsistema or "").upper().strip(), "")


def sello_de(etapa, subsistema=""):
    """Lo que dice el sello según dónde va el trámite."""
    base = ["EN ELABORACIÓN", "EN FIRMA · DIRECTORA", "EN FIRMA · SUBSECRETARÍA",
            "ENTREGADO A RH", "CONCLUIDO"]
    if etapa == 3:
        quien = quien_recibe(subsistema)
        return f"EN RH · {quien.upper()}" if quien else "ENTREGADO A RH"
    return base[etapa]


def boton_de(etapa, subsistema=""):
    """El botón dice lo que ACABA de pasar, no "siguiente": así se aprieta
    cuando el trámite cambia de manos de verdad y no por error."""
    quien = quien_recibe(subsistema)
    destino = f"RH · {quien}" if quien else "RH"
    return ["Ya se elaboró · pasa a firma de la directora",
            "La directora firmó · pasa a la subsecretaria",
            f"La subsecretaria firmó · se entrega a {destino}",
            f"{quien or 'RH'} ya revisó · concluido"][etapa]

VERDE = "#1B5E3F"
GRIS = "#8A94A6"


# ── Utilidades ───────────────────────────────────────────────────────────
def _norm(txt) -> str:
    """Texto comparable: sin acentos, sin mayúsculas, sin espacios de más.

    Es lo que evita que al volver a pegar la lista se dupliquen los renglones
    por una tilde o un espacio. Lo que se GUARDA es el texto original; esto
    solo sirve para comparar.
    """
    t = unicodedata.normalize("NFKD", str(txt or "").lower())
    t = "".join(c for c in t if not unicodedata.combining(c))
    return " ".join(t.split())


def detectar_fecha(texto):
    """Saca la fecha límite del propio texto del pendiente. "" si no hay.

    La gente la escribe dentro de la tarea, no en un campo aparte: "a más
    tardar el 25 de noviembre", "(fecha límite: martes 17 de noviembre)".
    Leerla de ahí evita capturarla dos veces.

    Sin año, se asume el próximo que todavía no pasa: en noviembre, "17 de
    noviembre" es de este año y "3 de enero" del que entra. Dar por hecho el
    año en curso pondría la fecha en el pasado y el recordatorio no serviría.
    """
    t = _norm(texto)
    m = re.search(r"(\d{1,2})\s*[/-]\s*(\d{1,2})\s*[/-]\s*(\d{2,4})", t)
    if m:
        d, mes, a = int(m.group(1)), int(m.group(2)), int(m.group(3))
        a = a + 2000 if a < 100 else a
        try:
            return date(a, mes, d).isoformat()
        except ValueError:
            return ""
    m = re.search(r"(\d{1,2})\s+de\s+([a-z]+)(?:\s+(?:de\s+)?(\d{4}))?", t)
    if not m or m.group(2) not in MESES_ES:
        return ""
    d, mes = int(m.group(1)), MESES_ES.index(m.group(2)) + 1
    hoy = date.today()
    anio = int(m.group(3)) if m.group(3) else hoy.year
    try:
        f = date(anio, mes, d)
    except ValueError:
        return ""
    if not m.group(3) and f < hoy:
        try:
            f = date(anio + 1, mes, d)
        except ValueError:
            return ""
    return f.isoformat()


def liga_calendar(titulo, fecha_iso, detalle=""):
    """Liga que abre Google Calendar con el evento ya lleno, para que él solo
    confirme. "" si la fecha no sirve.

    Es a propósito una liga y no una creación automática: la cuenta de
    servicio que usa la app entra a los Sheets, pero NO al calendario
    personal de nadie — eso necesitaría otro permiso y que cada quien
    comparta su calendario. Con la liga, cada persona crea el evento en SU
    calendario con un clic y nadie entrega accesos. Es el mismo camino del
    botón de Gmail que ya usa el módulo de nómina.
    """
    try:
        f = datetime.strptime(str(fecha_iso), "%Y-%m-%d").date()
    except (ValueError, TypeError):
        return ""
    # Evento de día completo: Calendar pide el día siguiente como fin.
    ini = f.strftime("%Y%m%d")
    fin = (f + timedelta(days=1)).strftime("%Y%m%d")
    return ("https://calendar.google.com/calendar/render?action=TEMPLATE"
            + "&text=" + urllib.parse.quote(f"RH DFC: {titulo}")
            + f"&dates={ini}/{fin}"
            + "&details=" + urllib.parse.quote(
                detalle or "Pendiente registrado en el Portal de Gestión de Incidencias (DFC)."))


def _hoy():
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def _quien():
    return st.session_state.get("nombre", "RH")


@st.cache_resource(show_spinner=False)
def _hoja(_get_client, sheet_id, tab, cols_tuple):
    """La worksheet, guardada entre recargas.

    Cada open_by_key() y cada worksheet() es UNA llamada a la API de Google,
    y Streamlit vuelve a correr todo el script con cada clic. Sin esta caché
    se gastaban dos llamadas por tab por recarga nada más para llegar a los
    datos, y el límite son 60 lecturas por minuto: tres clics y 429.
    El guion bajo de _get_client le dice a Streamlit que no intente usarlo
    como parte de la llave de caché.
    """
    import gspread
    sh = _get_client().open_by_key(sheet_id)
    try:
        return sh.worksheet(tab)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(tab, rows=2, cols=len(cols_tuple))
        ws.append_row(list(cols_tuple))
        return ws


@st.cache_data(ttl=300, show_spinner=False)
def _valores(_ws, sheet_id, tab):
    """El contenido de la tab. ttl=300 como el resto de la app.

    Lo que vuelve es una copia: si quien la recibe la modificara, estaría
    modificando lo que ve el resto de la pantalla. La caché se limpia a mano
    después de cada escritura, así que un cambio propio se ve de inmediato;
    los 30 segundos solo afectan a lo que cambie otra persona.
    """
    return [list(f) for f in _ws.get_all_values()]


def _refrescar():
    """Tira la caché de lecturas. Se llama después de CADA escritura."""
    try:
        st.cache_data.clear()
    except Exception:
        pass


def _abrir(get_client, tab, cols, sheet_id=None):
    """Devuelve (ws, valores). Crea la tab con sus encabezados si no existe,
    para que el módulo funcione desde el primer día sin preparar nada."""
    sid = sheet_id or st.secrets["sheet_checador_id"]
    ws = _hoja(get_client, sid, tab, tuple(cols))
    valores = _valores(ws, sid, tab)
    if not valores:
        ws.append_row(cols)
        _refrescar()
        return ws, [cols]

    # Migración de encabezados. Una tab creada por una versión anterior no
    # tiene las columnas que se agregaron después (SUBSISTEMA, FECHA_LIMITE).
    # Sin esto, _indices devuelve -1, el código arma una celda en la columna
    # 0 y gspread revienta con un error ilegible tipo "(2, 0)" justo al
    # guardar — que es como se perdió el avance de una propuesta.
    #
    # Las columnas nuevas se AGREGAN al final: así no se mueve ningún dato ya
    # capturado y las que ya existen se quedan donde están.
    actuales = [h.strip().upper() for h in valores[0]]
    faltan = [c for c in cols if c not in actuales]
    if faltan:
        nuevos = valores[0] + faltan
        ws.update([nuevos], f"A1", value_input_option="USER_ENTERED")
        _refrescar()
        valores[0] = nuevos
    return ws, valores


def _indices(valores, cols):
    headers = [h.strip().upper() for h in valores[0]]
    return {c: (headers.index(c) if c in headers else -1) for c in cols}


def _celda(fila, i):
    return fila[i] if 0 <= i < len(fila) else ""


# ── 1. PENDIENTES ────────────────────────────────────────────────────────
def leer_pendientes(get_client):
    """[(n_fila, texto, estado, responsable)] de la tab Pendientes_RH."""
    ws, valores = _abrir(get_client, TAB_PEND, COLS_PEND)
    idx = _indices(valores, COLS_PEND)
    filas = []
    for nf, f in enumerate(valores[1:], start=2):
        texto = str(_celda(f, idx["PENDIENTE"])).strip()
        if not texto:
            continue
        filas.append((nf, texto,
                      str(_celda(f, idx["ESTADO"])).strip().upper() or "PENDIENTE",
                      str(_celda(f, idx["RESPONSABLE"])).strip(),
                      str(_celda(f, idx["FECHA_LIMITE"])).strip()))
    return ws, idx, filas


def parsear_lista(texto):
    """Saca (hecho, pendiente) de una lista pegada.

    Acepta las formas en las que la gente de verdad manda una lista: ✅/⬜,
    casillas de markdown (- [x] / - [ ]), viñetas y renglones pelones. Lo que
    no trae marca se toma como pendiente, que es la lectura prudente: más
    vale que aparezca por hacer a darlo por cerrado sin que nadie lo dijera.
    """
    hechos, pendientes = [], []
    MARCAS_HECHO = ("✅", "✔", "☑", "[x]", "[X]")
    MARCAS_PEND = ("⬜", "☐", "⭕", "[ ]", "[]")
    for linea in str(texto or "").splitlines():
        t = linea.strip()
        if not t:
            continue
        # Un encabezado tipo "Pendientes de cierre de año – RH:" no es tarea.
        if t.endswith(":") and len(t.split()) <= 8 and not any(
                m in t for m in MARCAS_HECHO + MARCAS_PEND):
            continue
        hecho = any(m in t for m in MARCAS_HECHO)
        # Se quitan marcas, viñetas y numeración del principio
        for m in MARCAS_HECHO + MARCAS_PEND:
            t = t.replace(m, " ")
        t = t.lstrip("-*•·0123456789.) \t").strip()
        if not t:
            continue
        (hechos if hecho else pendientes).append(t)
    return hechos, pendientes


def sincronizar(get_client, texto):
    """Mete la lista pegada al Sheet. Devuelve (nuevos, cerrados, iguales).

    Nunca borra: lo que ya no viene en el pegado se queda como estaba. Si
    alguien pega media lista no se pierde la otra mitad — que es justo el
    accidente que haría inservible esto.
    """
    from gspread.cell import Cell
    hechos, pendientes = parsear_lista(texto)
    if not hechos and not pendientes:
        return 0, 0, 0

    ws, idx, filas = leer_pendientes(get_client)
    por_texto = {_norm(t): (nf, est) for nf, t, est, _, _ in filas}

    nuevas, cierres, iguales = [], [], 0
    for lista, marcado in ((pendientes, False), (hechos, True)):
        for t in lista:
            k = _norm(t)
            if k in por_texto:
                nf, est = por_texto[k]
                if marcado and est != "HECHO":
                    cierres.append(nf)
                else:
                    iguales += 1
                continue
            por_texto[k] = (None, "HECHO" if marcado else "PENDIENTE")
            # La fecha límite se saca del propio texto al dar de alta: así
            # el botón de agendar aparece solo, sin capturar nada aparte.
            nuevas.append([f"P{len(nuevas) + len(filas) + 1}", _hoy(), t,
                           "HECHO" if marcado else "PENDIENTE", "",
                           _quien(), _hoy() if marcado else "",
                           detectar_fecha(t)])

    if nuevas:
        ws.append_rows(nuevas, value_input_option="USER_ENTERED")
        _refrescar()
    if cierres:
        celdas = [Cell(nf, idx["ESTADO"] + 1, "HECHO") for nf in cierres]
        if idx["FECHA_CIERRE"] >= 0:
            celdas += [Cell(nf, idx["FECHA_CIERRE"] + 1, _hoy()) for nf in cierres]
        ws.update_cells(celdas, value_input_option="USER_ENTERED")
        _refrescar()
    return len(nuevas), len(cierres), iguales


def _barra(hechos, total, etiqueta):
    """Barra de avance con su porcentaje, igual arriba de los dos apartados."""
    pct = int(round(hechos * 100 / total)) if total else 0
    st.markdown(f"""
<div style="border:1px solid #E3E6EB;border-radius:12px;padding:14px 18px;margin-bottom:14px;">
  <div style="display:flex;justify-content:space-between;align-items:baseline;">
    <span style="font-size:0.95rem;font-weight:600;color:#1F2A37;">Avance general</span>
    <span style="font-size:1.05rem;font-weight:700;color:#1F2A37;">{pct}%</span>
  </div>
  <div style="background:#E9ECF1;border-radius:6px;height:9px;margin:10px 0 6px;">
    <div style="background:{VERDE};width:{pct}%;height:9px;border-radius:6px;"></div>
  </div>
  <div style="font-size:0.8rem;color:{GRIS};">{etiqueta}</div>
</div>""", unsafe_allow_html=True)


def _apartado_pendientes(get_client, puede_pegar):
    st.markdown("### 📌 Pendientes RH DFC")
    try:
        ws, idx, filas = leer_pendientes(get_client)
    except Exception as e:
        st.error(f"No se pudo leer la lista: {e}")
        return

    abiertos = [f for f in filas if f[2] != "HECHO"]
    hechos = [f for f in filas if f[2] == "HECHO"]
    if filas:
        _barra(len(hechos), len(filas),
               f"{len(hechos)} de {len(filas)} pendientes cerrados.")

    # Pegar la lista completa: es el puente con lo que trabajes fuera.
    if puede_pegar:
        with st.expander("📥 Pegar la lista completa (la sincroniza sin duplicar)"):
            st.caption("Pega la lista como la tengas: con ✅ y ⬜, con casillas "
                       "de markdown o a secas. Lo nuevo se agrega, lo marcado "
                       "se cierra, y nada se borra.")
            # El vaciado del recuadro se pide con una bandera y se ejecuta
            # AQUÍ, antes de crear el widget. Streamlit prohíbe asignarle
            # valor a una key después de instanciarla, y hacerlo reventaba la
            # sincronización: los pendientes se guardaban, pero el error
            # tapaba el rerun y la pantalla decía que no había ninguno.
            if st.session_state.pop("_limpiar_pegar", False):
                st.session_state["pend_pegar"] = ""
            txt = st.text_area("Lista", height=200, key="pend_pegar",
                               placeholder="✅ Entregar reportes a la directora\n⬜ Enviar los reclamos de pago")
            if st.button("Sincronizar con el Sheet", type="primary", key="pend_sync"):
                try:
                    n, c, ig = sincronizar(get_client, txt)
                    st.session_state["_msg_pend"] = (
                        f"Listo: {n} nuevo(s), {c} cerrado(s), {ig} sin cambio.")
                    st.session_state["_limpiar_pegar"] = True
                    st.rerun()
                except Exception as e:
                    st.error(f"No se pudo sincronizar: {e}")

    # Alta suelta: para que un colaborador agregue el suyo sin pegar nada.
    with st.form("pend_alta", clear_on_submit=True):
        c1, c2, c3 = st.columns([5, 2, 1])
        nuevo = c1.text_input("Nuevo pendiente", label_visibility="collapsed",
                              placeholder="Escribe un pendiente y agrégalo")
        resp = c2.text_input("Responsable", label_visibility="collapsed",
                             placeholder="Responsable (opcional)")
        if c3.form_submit_button("Agregar") and nuevo.strip():
            try:
                if _norm(nuevo) in {_norm(f[1]) for f in filas}:
                    st.warning("Ese pendiente ya está en la lista.")
                else:
                    ws.append_row([f"P{len(filas) + 1}", _hoy(), nuevo.strip(),
                                   "PENDIENTE", resp.strip(), _quien(), "",
                                   detectar_fecha(nuevo)],
                                  value_input_option="USER_ENTERED")
                    _refrescar()
                    st.session_state["_msg_pend"] = "Pendiente agregado."
                    st.rerun()
            except Exception as e:
                st.error(f"No se pudo agregar: {e}")

    if st.session_state.get("_msg_pend"):
        st.success(st.session_state.pop("_msg_pend"))

    if not filas:
        st.info("Todavía no hay pendientes. Pega la lista o agrega el primero.")
        return

    # Se marcan varios y se cierran de un golpe: una sola escritura al Sheet
    # en vez de una por cada palomita.
    st.caption(f"{len(abiertos)} por hacer")
    marcar = []
    for nf, texto, estado, resp, limite in abiertos:
        etiqueta = f"{texto}" + (f"  ·  {resp}" if resp else "")
        liga = liga_calendar(texto, limite)
        # Con fecha, la casilla se encoge para dejarle lugar al botón de
        # agendar; sin fecha ocupa todo, para no dejar un hueco vacío.
        c_chk, c_cal = st.columns([6, 1]) if liga else (st.container(), None)
        with c_chk:
            if st.checkbox(etiqueta, key=f"pend_chk_{nf}"):
                marcar.append(nf)
        if liga:
            _dd = "/".join(reversed(limite.split("-")))
            c_cal.markdown(
                f'<a href="{liga}" target="_blank" title="Agendar el {_dd}" '
                f'style="display:inline-block;padding:2px 10px;border:1px solid {VERDE};'
                f'color:{VERDE};border-radius:6px;text-decoration:none;font-size:0.78rem;'
                f'white-space:nowrap;">📅 {_dd[:5]}</a>', unsafe_allow_html=True)
    if marcar and st.button(f"✅ Marcar {len(marcar)} como hecho(s)",
                            type="primary", key="pend_cerrar"):
        try:
            from gspread.cell import Cell
            celdas = [Cell(nf, idx["ESTADO"] + 1, "HECHO") for nf in marcar]
            if idx["FECHA_CIERRE"] >= 0:
                celdas += [Cell(nf, idx["FECHA_CIERRE"] + 1, _hoy()) for nf in marcar]
            ws.update_cells(celdas, value_input_option="USER_ENTERED")
            _refrescar()
            st.session_state["_msg_pend"] = f"{len(marcar)} cerrado(s)."
            st.rerun()
        except Exception as e:
            st.error(f"No se pudo cerrar: {e}")

    if hechos:
        with st.expander(f"✅ Ya hechos ({len(hechos)})"):
            for _, texto, _, _, _ in hechos:
                st.markdown(f"<span style='color:{GRIS};text-decoration:line-through;'>"
                            f"{texto}</span>", unsafe_allow_html=True)


# ── 2. PROPUESTAS ────────────────────────────────────────────────────────
def leer_propuestas(get_client):
    """[(n_fila, id, nombre, etapa, subsistema, nota, actualizado)].

    El ID viaja porque los documentos del minutario se atan a él: con el
    número de renglón se despegarían al borrar una propuesta de arriba, y con
    el nombre, al corregirle una falta de ortografía.
    """
    ws, valores = _abrir(get_client, TAB_PROP, COLS_PROP)
    idx = _indices(valores, COLS_PROP)
    filas = []
    for nf, f in enumerate(valores[1:], start=2):
        nombre = str(_celda(f, idx["NOMBRE"])).strip()
        if not nombre:
            continue
        try:
            etapa = int(str(_celda(f, idx["ETAPA"])).strip() or 0)
        except ValueError:
            etapa = 0
        etapa = max(0, min(etapa, len(ETAPAS) - 1))
        id_prop = str(_celda(f, idx["ID"])).strip() or f"T{nf - 1}"
        filas.append((nf, id_prop, nombre, etapa,
                      str(_celda(f, idx["SUBSISTEMA"])).strip().upper(),
                      str(_celda(f, idx["NOTA"])).strip(),
                      str(_celda(f, idx["ACTUALIZADO"])).strip()))
    return ws, idx, filas


def _escalera(etapa):
    """Dibuja las cuatro etapas con su punto y su línea.

    Es HTML a mano porque Streamlit no trae nada parecido, y un dibujo así
    se lee de un vistazo: se ve dónde va el trámite sin leer una palabra.
    """
    html = ['<div style="margin:6px 0 14px;">']
    for i, nombre in enumerate(ETAPAS):
        hecha = i < etapa
        actual = i == etapa
        if hecha:
            punto = (f'<div style="width:26px;height:26px;border-radius:50%;'
                     f'background:{VERDE};color:#fff;display:flex;align-items:center;'
                     f'justify-content:center;font-size:14px;font-weight:700;">✓</div>')
            texto = f'<span style="color:#1F2A37;">{nombre}</span>'
        elif actual:
            punto = (f'<div style="width:26px;height:26px;border-radius:50%;'
                     f'border:2px solid {VERDE};color:{VERDE};display:flex;'
                     f'align-items:center;justify-content:center;font-size:13px;'
                     f'font-weight:700;">{i + 1}</div>')
            texto = (f'<span style="color:#1F2A37;font-weight:700;">{nombre}</span>'
                     f'<span style="color:{VERDE};font-size:0.82rem;"> · aquí va</span>')
        else:
            punto = (f'<div style="width:26px;height:26px;border-radius:50%;'
                     f'border:2px solid #D6DAE1;color:{GRIS};display:flex;'
                     f'align-items:center;justify-content:center;font-size:13px;">'
                     f'{i + 1}</div>')
            texto = f'<span style="color:{GRIS};">{nombre}</span>'
        html.append(f'<div style="display:flex;align-items:center;gap:12px;">'
                    f'{punto}<div style="font-size:0.94rem;">{texto}</div></div>')
        if i < len(ETAPAS) - 1:
            color = VERDE if hecha else "#D6DAE1"
            html.append(f'<div style="width:2px;height:16px;background:{color};'
                        f'margin-left:12px;"></div>')
    html.append("</div>")
    return "".join(html)


def etiqueta_periodo(ini, fin):
    """Texto legible del periodo que cubre el documento.

    Una prórroga se extiende por meses completos ("enero a marzo") o por días
    sueltos ("del 15 al 31 de enero"), y escribirlo de las dos formas a mano
    termina en cada quien poniéndolo distinto. Aquí se capturan DOS fechas y
    el texto se deriva: si el periodo empieza el día 1 y termina el último
    día del mes, se nombra por meses; si no, con los días. Las fechas exactas
    quedan guardadas aparte, que es lo que sirve para filtrar y ordenar.
    """
    import calendar
    if not ini or not fin:
        return ""
    mes_completo = (ini.day == 1
                    and fin.day == calendar.monthrange(fin.year, fin.month)[1])
    if mes_completo:
        if ini.year == fin.year and ini.month == fin.month:
            return f"{MESES_CORTO[ini.month - 1]} {ini.year}"
        if ini.year == fin.year:
            return (f"{MESES_CORTO[ini.month - 1]} a "
                    f"{MESES_CORTO[fin.month - 1]} {ini.year}")
        return (f"{MESES_CORTO[ini.month - 1]} {ini.year} a "
                f"{MESES_CORTO[fin.month - 1]} {fin.year}")
    if ini.year == fin.year and ini.month == fin.month:
        return f"{ini.day} al {fin.day} {MESES_CORTO[ini.month - 1]} {ini.year}"
    if ini.year == fin.year:
        return (f"{ini.day} {MESES_CORTO[ini.month - 1]} al "
                f"{fin.day} {MESES_CORTO[fin.month - 1]} {ini.year}")
    return (f"{ini.day} {MESES_CORTO[ini.month - 1]} {ini.year} al "
            f"{fin.day} {MESES_CORTO[fin.month - 1]} {fin.year}")


def _abrir_docs(get_client):
    """(ws, valores) de la hoja 'propuestas' DENTRO del Sheet del minutario.

    Ojo: este es otro Sheet (sheet_oficios_id), no el del checador donde vive
    el seguimiento. Es a propósito: el seguimiento es trabajo en curso y los
    documentos firmados son resguardo, que es justo lo que el minutario hace.
    """
    # Pasa por _abrir para heredar la caché y la migración de encabezados.
    # Tenía su propia copia de esa lógica y por eso seguía gastando una
    # apertura y una lectura en cada recarga, aunque las otras dos tabs ya
    # estuvieran cacheadas.
    return _abrir(get_client, TAB_DOCS, COLS_DOCS,
                  sheet_id=st.secrets["sheet_oficios_id"])


def leer_docs(get_client, id_propuesta=None):
    """[(id, tipo, archivo, url, subido_por, fecha, periodo)]."""
    ws, valores = _abrir_docs(get_client)
    idx = _indices(valores, COLS_DOCS)
    docs = []
    for f in valores[1:]:
        if id_propuesta and str(_celda(f, idx["ID_PROPUESTA"])).strip() != str(id_propuesta):
            continue
        if not str(_celda(f, idx["URL"])).strip():
            continue
        docs.append((str(_celda(f, idx["ID_PROPUESTA"])).strip(),
                     str(_celda(f, idx["TIPO"])).strip().upper(),
                     str(_celda(f, idx["ARCHIVO"])).strip(),
                     str(_celda(f, idx["URL"])).strip(),
                     str(_celda(f, idx["SUBIDO_POR"])).strip(),
                     str(_celda(f, idx["FECHA"])).strip(),
                     str(_celda(f, idx["PERIODO"])).strip()))
    return ws, docs


def guardar_doc(get_client, id_propuesta, nombre_propuesta, tipo, archivo,
                periodo_ini=None, periodo_fin=None):
    """Sube el archivo a Drive y lo registra en el minutario. (ok, mensaje).

    Reutiliza subir_bytes_drive de oficios_module en vez de copiarlo: es la
    misma carpeta de Unidad Compartida y la misma cuenta de servicio, y tener
    dos copias del mismo subidor es como se desincronizan los permisos.
    """
    try:
        from oficios_module import subir_bytes_drive, sha256_bytes
    except Exception as e:
        return False, f"No se pudo usar el subidor de oficios: {e}"

    contenido = archivo.getvalue()
    if not contenido:
        return False, "El archivo llegó vacío."
    ext = (archivo.name.rsplit(".", 1)[-1] or "pdf").lower()
    # El nombre lleva la propuesta y el tipo para que en Drive se entienda
    # qué es sin abrirlo, y la hora para no pisar un archivo anterior.
    limpio = "".join(c if c.isalnum() or c in " -_" else "_"
                     for c in nombre_propuesta)[:50].strip().replace(" ", "_")
    periodo = etiqueta_periodo(periodo_ini, periodo_fin)
    # El periodo va en el nombre del archivo: en Drive se distingue de un
    # vistazo la prorroga de enero-marzo de la de abril-junio, sin abrirlas.
    marca = "_" + periodo.replace(" ", "") if periodo else ""
    sello = datetime.now().strftime("%Y%m%d-%H%M%S")
    url = subir_bytes_drive(contenido, f"{limpio}_{tipo}{marca}_{sello}.{ext}",
                            archivo.type or "application/octet-stream")
    if str(url).startswith("ERROR:"):
        return False, url

    try:
        ws, _ = _abrir_docs(get_client)
        ws.append_row([str(id_propuesta), nombre_propuesta, tipo, periodo,
                       periodo_ini.isoformat() if periodo_ini else "",
                       periodo_fin.isoformat() if periodo_fin else "",
                       archivo.name, url, sha256_bytes(contenido),
                       _quien(), _hoy()],
                      value_input_option="USER_ENTERED")
        _refrescar()
    except Exception as e:
        # El archivo YA está en Drive: se devuelve la liga para que no se
        # pierda aunque el registro en el Sheet haya fallado.
        return False, f"Se subió a Drive pero no se registró en el minutario: {e} — {url}"
    return True, url


def _apartado_documentos(get_client, id_propuesta, nombre, docs=None):
    """Subida de firmadas y acuses.

    `docs` llega YA leído desde arriba. Antes esta función leía el minutario
    por su cuenta, o sea una lectura por cada propuesta en pantalla: con seis
    trámites eran seis lecturas por cada clic, y el límite de Google son 60
    por minuto. De ahí el 429. El minutario se lee UNA vez y se reparte.
    """
    if docs is None:
        try:
            _ws, docs = leer_docs(get_client, id_propuesta)
        except Exception as e:
            st.caption(f"No se pudo leer el minutario: {e}")
            docs = []

    firmadas = [d for d in docs if d[1] == "FIRMADA"]
    acuses = [d for d in docs if d[1] == "ACUSE"]
    st.markdown(f"**Documentos** · {len(firmadas)} firmada(s) · {len(acuses)} acuse(s)")

    for tipo, etiqueta in TIPOS_DOC.items():
        ya = [d for d in docs if d[1] == tipo]
        with st.expander(f"{'📄' if tipo == 'FIRMADA' else '🧾'} {etiqueta}"
                         + (f" ({len(ya)})" if ya else "")):
            for _i, _t, archivo, url, quien, cuando, periodo in ya:
                st.markdown(f"- **{periodo or 'sin periodo'}** — [{archivo}]({url})"
                            f"  ·  {quien} · {cuando}")

            # El periodo que CUBRE el documento, que no es la fecha en que se
            # sube: una prórroga de enero a marzo puede subirse en abril.
            cp1, cp2 = st.columns(2)
            # value=None para que nazcan VACÍOS. Sin esto, Streamlit los
            # prellena con la fecha de hoy: no se leen como una pregunta y,
            # peor, la validación de "es obligatorio" nunca se dispara, así
            # que se guardaría "hoy a hoy" como periodo sin que nadie lo note.
            ini = cp1.date_input("Desde", value=None, format="DD/MM/YYYY",
                                 key=f"pi_{tipo}_{id_propuesta}")
            fin = cp2.date_input("Hasta", value=None, format="DD/MM/YYYY",
                                 key=f"pf_{tipo}_{id_propuesta}")
            if not (ini and fin):
                st.caption("⚠️ Captura el periodo que cubre el documento.")
            if ini and fin:
                if fin < ini:
                    st.warning("La fecha de término es anterior a la de inicio.")
                else:
                    st.caption(f"Periodo: **{etiqueta_periodo(ini, fin)}**")

            subido = st.file_uploader(
                f"Subir {etiqueta.lower()}", key=f"up_{tipo}_{id_propuesta}",
                type=["pdf", "jpg", "jpeg", "png"],
                accept_multiple_files=True)
            if subido and st.button("Guardar en el minutario",
                                    key=f"gd_{tipo}_{id_propuesta}"):
                # Sin periodo no se guarda: un PDF de prórroga sin saber qué
                # meses cubre obliga a abrirlo para enterarse, y entonces el
                # registro del minutario no sirve de nada.
                if not (ini and fin) or fin < ini:
                    st.error("Captura el periodo que cubre el documento "
                             "antes de guardarlo.")
                else:
                    ok_n, fallos = 0, []
                    for arch in subido:
                        ok, msg = guardar_doc(get_client, id_propuesta, nombre,
                                              tipo, arch, ini, fin)
                        if ok:
                            ok_n += 1
                        else:
                            fallos.append(f"{arch.name}: {msg}")
                    st.session_state["_msg_prop"] = (
                        f"{ok_n} archivo(s) de {etiqueta_periodo(ini, fin)} "
                        f"en el minutario, hoja «{TAB_DOCS}»."
                        + (" Fallaron: " + " | ".join(fallos) if fallos else ""))
                    st.rerun()


def _guardar_sub(ws, idx, n_fila, subsistema):
    """Deja fijo si el trámite es estatal o federal: de ahí sale a quién se
    entrega en RH, y con eso el sello y el botón cambian solos."""
    from gspread.cell import Cell
    celdas = [Cell(n_fila, idx["SUBSISTEMA"] + 1, str(subsistema).upper())]
    if idx["ACTUALIZADO"] >= 0:
        celdas.append(Cell(n_fila, idx["ACTUALIZADO"] + 1, _hoy()))
    ws.update_cells(celdas, value_input_option="USER_ENTERED")
    _refrescar()


def _guardar_etapa(ws, idx, n_fila, etapa, nota=None):
    """Una sola escritura para etapa + fecha (+ nota si cambió)."""
    from gspread.cell import Cell
    celdas = [Cell(n_fila, idx["ETAPA"] + 1, str(etapa))]
    if idx["ACTUALIZADO"] >= 0:
        celdas.append(Cell(n_fila, idx["ACTUALIZADO"] + 1, _hoy()))
    if nota is not None and idx["NOTA"] >= 0:
        celdas.append(Cell(n_fila, idx["NOTA"] + 1, nota))
    ws.update_cells(celdas, value_input_option="USER_ENTERED")
    _refrescar()


def _apartado_propuestas(get_client):
    st.markdown("### 📑 Seguimiento de propuestas")
    st.caption("Así va la ruta de cada trámite: se elabora, lo firma la "
               "directora y luego la subsecretaria. Avanza con el botón "
               "cuando cambie de manos.")
    try:
        ws, idx, filas = leer_propuestas(get_client)
    except Exception as e:
        st.error(f"No se pudieron leer las propuestas: {e}")
        return

    if filas:
        # El avance cuenta ETAPAS cumplidas, no trámites: un trámite a medias
        # se ve a medias, que es lo honesto cuando falta una firma.
        cumplidas = sum(f[3] for f in filas)
        total = (len(ETAPAS) - 1) * len(filas)
        _barra(cumplidas, total,
               f"{cumplidas} de {total} etapas completadas en "
               f"{len(filas)} trámite{'s' if len(filas) != 1 else ''}.")

    if st.session_state.get("_msg_prop"):
        st.success(st.session_state.pop("_msg_prop"))

    # Una sola lectura del minutario para TODAS las propuestas de la
    # pantalla, agrupada por el ID de cada una.
    por_prop = {}
    try:
        _wsd, todos = leer_docs(get_client)
        for d in todos:
            por_prop.setdefault(d[0], []).append(d)
    except Exception as e:
        st.caption(f"No se pudo leer el minutario: {e}")

    for nf, id_prop, nombre, etapa, subsis, nota, actualizado in filas:
        with st.container(border=True):
            c1, c2 = st.columns([3, 2])
            c1.markdown(f"**{nombre}**")
            c2.markdown(
                f'<div style="text-align:right;"><span style="border:1.5px solid #C2410C;'
                f'color:#C2410C;border-radius:4px;padding:3px 10px;font-size:0.72rem;'
                f'font-weight:700;letter-spacing:0.04em;">{sello_de(etapa, subsis)}</span></div>',
                unsafe_allow_html=True)
            st.markdown(_escalera(etapa), unsafe_allow_html=True)

            # Sin subsistema no se sabe si va con Aurora o con Alejandra, y
            # el trámite no puede entregarse "a quien sea": se pide antes de
            # dejar la firma de la subsecretaria.
            if not subsis:
                nuevo_sub = st.radio(
                    "¿Es estatal o federal?", ["", "ESTATAL", "FEDERAL"],
                    horizontal=True, key=f"prop_sub_{nf}",
                    format_func=lambda v: "— elige —" if not v else
                    f"{v.capitalize()} · {RECIBE_EN_RH[v]}")
                if nuevo_sub and st.button("Guardar", key=f"prop_gs_{nf}"):
                    try:
                        _guardar_sub(ws, idx, nf, nuevo_sub)
                        st.session_state["_msg_prop"] = (
                            f"{nombre}: {nuevo_sub.capitalize()}, "
                            f"la recibe {RECIBE_EN_RH[nuevo_sub]}.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"No se pudo guardar: {e}")

            if etapa < len(ETAPAS) - 1:
                _falta_sub = (etapa == 2 and not subsis)
                if _falta_sub:
                    st.caption("Elige estatal o federal para poder entregarla.")
                if (not _falta_sub) and st.button(
                        boton_de(etapa, subsis), key=f"prop_av_{nf}",
                        type="primary", use_container_width=True):
                    try:
                        _guardar_etapa(ws, idx, nf, etapa + 1)
                        st.session_state["_msg_prop"] = f"{nombre}: {ETAPAS[etapa + 1]}."
                        st.rerun()
                    except Exception as e:
                        st.error(f"No se pudo avanzar: {e}")
            if etapa > 0:
                # Regresar existe porque un trámite se devuelve: sin esto, un
                # clic de más obliga a entrar al Sheet a corregir a mano.
                if st.button("Regresar una etapa", key=f"prop_re_{nf}",
                             use_container_width=True):
                    try:
                        _guardar_etapa(ws, idx, nf, etapa - 1)
                        st.session_state["_msg_prop"] = f"{nombre}: regresó a {ETAPAS[etapa - 1]}."
                        st.rerun()
                    except Exception as e:
                        st.error(f"No se pudo regresar: {e}")

            # Los dos botones de subida van SIEMPRE visibles, en cualquier
            # etapa. Antes solo aparecían en Concluido y eso los volvía
            # inalcanzables: un trámite tarda semanas en llegar ahí y los
            # documentos se tienen antes. Si alguien sube algo a destiempo,
            # se ve en la lista y se vuelve a subir; esconder el botón cuesta
            # mucho más que eso.
            _apartado_documentos(get_client, id_prop, nombre,
                                 por_prop.get(id_prop, []))

            nueva_nota = st.text_area("Nota", value=nota, height=80,
                                      key=f"prop_nota_{nf}")
            cols = st.columns([3, 1, 1])
            if actualizado:
                cols[0].caption(f"Actualizado: {actualizado}")
            if nueva_nota.strip() != nota and cols[1].button(
                    "Guardar nota", key=f"prop_gn_{nf}"):
                try:
                    _guardar_etapa(ws, idx, nf, etapa, nueva_nota.strip())
                    st.session_state["_msg_prop"] = "Nota guardada."
                    st.rerun()
                except Exception as e:
                    st.error(f"No se pudo guardar: {e}")
            with cols[2].popover("Eliminar"):
                st.caption(f"¿Quitar «{nombre}» del seguimiento?")
                if st.button("Sí, eliminar", key=f"prop_del_{nf}"):
                    try:
                        ws.delete_rows(nf)
                        st.session_state["_msg_prop"] = f"{nombre} eliminada."
                        st.rerun()
                    except Exception as e:
                        st.error(f"No se pudo eliminar: {e}")

    with st.form("prop_alta", clear_on_submit=True):
        c1, c2, c3 = st.columns([4, 2, 1])
        nombre = c1.text_input("Nueva propuesta", label_visibility="collapsed",
                               placeholder="Agregar propuesta por nombre...")
        # Se pregunta desde el alta porque decide quién la recibe en RH; se
        # puede dejar en blanco y elegirlo después, antes de entregarla.
        sub = c2.selectbox("Subsistema", ["", "ESTATAL", "FEDERAL"],
                           label_visibility="collapsed",
                           format_func=lambda v: "Estatal o federal..." if not v
                           else f"{v.capitalize()} · {RECIBE_EN_RH[v]}")
        if c3.form_submit_button("Agregar") and nombre.strip():
            try:
                if _norm(nombre) in {_norm(f[2]) for f in filas}:
                    st.warning("Ya hay una propuesta con ese nombre.")
                else:
                    ws.append_row([f"T{len(filas) + 1}", nombre.strip(), "0",
                                   sub, "", _hoy(), _quien()],
                                  value_input_option="USER_ENTERED")
                    _refrescar()
                    st.session_state["_msg_prop"] = f"{nombre.strip()} agregada."
                    st.rerun()
            except Exception as e:
                st.error(f"No se pudo agregar: {e}")


# ── Vista ────────────────────────────────────────────────────────────────
def render_pendientes_rh(get_client):
    st.markdown("#### Recursos Humanos · Cierre de año")
    # Pegar la lista completa la reescribe entera: queda para quien
    # administra. Agregar y marcar lo puede hacer cualquiera, que es el punto
    # de tenerlo aquí y no en un archivo que solo una persona toca.
    puede_pegar = st.session_state.get("rol") == "admin"
    t1, t2 = st.tabs(["📌 Pendientes", "📑 Propuestas"])
    with t1:
        _apartado_pendientes(get_client, puede_pegar)
    with t2:
        _apartado_propuestas(get_client)
