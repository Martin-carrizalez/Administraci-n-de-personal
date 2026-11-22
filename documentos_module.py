"""documentos_module.py — Constancias y comisiones desde ARI, solo para el admin.

DE DÓNDE VIENE ESTE CÓDIGO
La generación de los .docx es la de app3.py (la app de días económicos),
copiada VERBATIM: las 15 funciones de redacción dinámica, las tuplas de
vialidades y asentamientos, y los dos generadores. No se reescribió nada de
eso a propósito — ahí está resuelto el problema de las comas huérfanas en la
ubicación y el de concatenar los Word sin docxcompose, y reescribirlo de
memoria era la forma más fácil de reintroducir esos bugs.

Lo único que cambió respecto de app3:
  · el .docx se guarda en el temporal del sistema, no en la carpeta del repo;
  · los datos se leen con el get_client() de app_incidencias, no con un
    cliente propio: el Sheet está compartido con esa cuenta de servicio.

QUÉ HACE FALTA EN EL REPO
  1. La carpeta templates/ junto a este archivo, con las tres plantillas:
     plantilla.docx, PLANTILLA_ENCARGADOS_CM.docx y
     PLANTILLA_COMISIONES_GENERALES.docx.
  2. python-docx en requirements.txt.
  3. Para el PDF, 'libreoffice' en packages.txt. SIN ESO el Word se genera
     igual y solo falta el botón de PDF: no es bloqueante.

QUIÉN LO VE
Nadie más que el admin. El gate está en puede_documentos() y se comprueba en
los dos lados: en el panel y en el ejecutor de las herramientas, para que una
herramienta no se pueda disparar desde el chat de un empleado.
"""

VERSION = "documentos v3 · filtro por centro · 21-nov"

import os
import re
import tempfile
import unicodedata
from datetime import datetime, date, timedelta

import pandas as pd
import streamlit as st

# El Sheet es el de app3, compartido de lectura con la cuenta de servicio de
# app_incidencias. Se abre por NOMBRE, igual que en app3, para no tener que
# meter otro id en el código.
SHEET_DOCS = "Dias_Economicos_Formacion_Continua"
TAB_CONSTANCIAS = "Constancias"
TAB_COMISIONES = "Comisiones"

# Los dos tipos tal como están escritos EN LA COLUMNA tipo_comision del Sheet.
# Si alguien los renombra allá, aquí deja de filtrar: por eso están en una
# sola constante y no repartidos por el código.
TIPO_CM = "Encargado CM"
TIPO_GENERAL = "General"

_deps = {}


def configurar(**kw):
    """Inyección de dependencias, igual que los demás módulos.

    Espera get_client. Se aceptan claves de más y se ignoran: así una versión
    nueva de app_incidencias que mande algo extra no tumba la app, que es
    exactamente lo que pasó una vez con cm_module.
    """
    _deps.update(kw)


def puede_documentos(rol=None) -> bool:
    """Solo el admin. Se consulta en el panel Y en el ejecutor."""
    r = rol if rol is not None else st.session_state.get("rol")
    return r == "admin"


# ──────────────────────────────────────────────────────────────────────────
# DATOS
# ──────────────────────────────────────────────────────────────────────────
@st.cache_data(ttl=300, show_spinner=False)
def _cargar_tab(_client, tab):
    """Una lectura por tab cada 5 minutos, el estándar de la casa.

    _client va con guion bajo para que Streamlit no intente hashearlo.
    get_all_records() sin numericise_ignore, igual que app3: las plantillas
    esperan 'Hoja' como número y el teléfono como float.
    """
    hoja = _client.open(SHEET_DOCS).worksheet(tab)
    return pd.DataFrame(hoja.get_all_records())


def invalidar():
    _cargar_tab.clear()


def _norm(s):
    """Para comparar nombres: sin acentos, sin dobles espacios, minúsculas."""
    t = unicodedata.normalize("NFD", str(s or ""))
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    return " ".join(t.lower().split())


def _casa_nombre(pedido, candidato) -> bool:
    """¿Son la misma persona?

    Compara CONJUNTOS de palabras, no cadenas: en el padrón los nombres van
    "APELLIDO APELLIDO NOMBRE" y la gente los dice "Nombre Apellido". Pedir
    coincidencia exacta hacía que no apareciera nadie. Se exige que todas las
    palabras de lo que pidió el admin estén en el candidato, con un mínimo de
    dos para que un "manuel" no empate con cuatro personas.
    """
    p = set(_norm(pedido).split())
    c = set(_norm(candidato).split())
    if not p or not c:
        return False
    if len(p) < 2:
        return False
    return p.issubset(c)


def _resolver_personas(pedido, disponibles):
    """Traduce lo que dijo el admin a nombres reales del Sheet.

    Devuelve (encontrados, no_encontrados, ambiguos). No adivina: si un nombre
    empata con dos personas, se reporta como ambiguo y se le pregunta. Un
    oficio con el nombre de otra persona es un documento que hay que rehacer.
    """
    t = _norm(pedido)
    if not t or t in ("todos", "todas", "todo", "all"):
        return list(disponibles), [], []
    enc, falt, amb = [], [], []
    for pieza in re.split(r"\s*(?:,|;|\sy\s)\s*", pedido):
        pieza = pieza.strip()
        if not pieza:
            continue
        caben = [d for d in disponibles if _casa_nombre(pieza, d)]
        if len(caben) == 1:
            if caben[0] not in enc:
                enc.append(caben[0])
        elif len(caben) > 1:
            amb.append((pieza, caben))
        else:
            falt.append(pieza)
    return enc, falt, amb


# ──────────────────────────────────────────────────────────────────────────
# HERRAMIENTAS PARA EL AGENTE
#
# Mismo reparto que en los pases: el modelo decide CUÁNDO llamar y con qué
# nombres; todo lo demás —leer el Sheet, resolver los nombres, numerar los
# oficios, decir qué falta— lo hace el código. Si lo decidiera el modelo,
# diría "ya está" con la mitad de los datos, que es lo que pasaba antes.
# Las fechas llegan ya en AAAA-MM-DD: las normaliza ari_module antes de
# llamar aquí, con el mismo normalizador que usa para los pases.
# ──────────────────────────────────────────────────────────────────────────
HERRAMIENTAS_DOC = [
    {"type": "function", "function": {
        "name": "preparar_constancias",
        "description": ("Prepara las constancias de servicio de una quincena. "
                        "Léelas de la hoja Constancias. Úsala cuando el admin "
                        "pida constancias."),
        "parameters": {"type": "object", "properties": {
            "quincena": {"type": "integer",
                         "description": "Número de quincena, 1 a 24"},
            "anio": {"type": "integer", "description": "Año, por ejemplo 2026"},
            "centro": {"type": "string",
                       "description": ("Centro de maestros o adscripción, por "
                                       "número o por nombre. Déjalo vacío "
                                       "para todas las adscripciones.")},
            "personas": {"type": "string",
                         "description": ("Nombres separados por coma, o "
                                         "'todos' para todo el personal. "
                                         "PREFIERE 'todos' con un centro "
                                         "antes que listar nombres.")},
        }, "required": ["quincena", "anio", "personas"]}}},
    {"type": "function", "function": {
        "name": "preparar_comisiones",
        "description": ("Prepara los oficios de comisión. Úsala cuando el "
                        "admin pida comisiones de asesores, coordinadores o "
                        "encargados de Centro de Maestros."),
        "parameters": {"type": "object", "properties": {
            "tipo": {"type": "string",
                     "description": ("'Encargados CM' para los encargados de "
                                     "Centro de Maestros, 'Generales' para el "
                                     "resto (asesores, coordinadores).")},
            "oficio_inicial": {"type": "integer",
                               "description": "Número del primer oficio"},
            "fecha_inicio": {"type": "string",
                             "description": "Inicio de la comisión, AAAA-MM-DD"},
            "fecha_fin": {"type": "string",
                          "description": "Fin de la comisión, AAAA-MM-DD"},
            "centro": {"type": "string",
                       "description": ("Centro de maestros, por número "
                                       "(14-05) o por nombre (La Palmita, "
                                       "Tlaquepaque). Vacío = todos los "
                                       "centros.")},
            "personas": {"type": "string",
                         "description": ("Nombres separados por coma, o "
                                         "'todos'. PREFIERE 'todos' con un "
                                         "centro antes que listar nombres: "
                                         "'la coordinadora del centro de La "
                                         "Palmita' es tipo=coordinador, "
                                         "centro='La Palmita', "
                                         "personas='todos'.")},
        }, "required": ["tipo", "oficio_inicial", "fecha_inicio",
                        "fecha_fin", "personas"]}}},
    {"type": "function", "function": {
        "name": "mostrar_boton_generar",
        "description": ("Llámala cuando el admin confirme que los datos del "
                        "resumen están bien. Saca el botón de generar."),
        "parameters": {"type": "object", "properties": {}}}},
]


def _tipo_comision(valor):
    """Traduce lo que diga el modelo al rótulo que espera el generador.

    El generador de app3 compara contra la cadena exacta "Encargados CM", y
    el Sheet guarda "Encargado CM" en singular. Dos cadenas distintas para
    lo mismo: se normaliza aquí y no se vuelve a pensar.
    """
    t = _norm(valor)
    if not t:
        return ""
    # Son DOS tipos y la gente los nombra de varias formas:
    #   · coordinador / encargado de centro de maestros  -> Encargados CM
    #   · generales (las de los asesores)                -> Generales
    # "coordinador" es la palabra que se usa al hablar; "Encargado CM" la que
    # está escrita en la columna tipo_comision del Sheet. Las dos entran.
    if any(w in t for w in ("coordinad", "encargad", "cm", "centro")):
        return "Encargados CM"
    return "Generales"


def _casa_centro(pedido, valor) -> bool:
    """¿El registro es de ese centro de maestros?

    Acepta el número ("14-05", "1405") o el nombre ("La Palmita",
    "Tlaquepaque"), porque al hablar se usa el nombre y en el Sheet puede
    venir cualquiera de los dos. Compara por contención normalizada: en el
    Sheet dice "CENTRO DE MAESTROS 14-05 LA PALMITA" y tú escribes
    "la palmita".
    """
    p, v = _norm(pedido), _norm(valor)
    if not p or not v:
        return False
    if p in v:
        return True
    # El número con y sin guion: "14-05" y "1405" son el mismo centro.
    pd = re.sub(r"\D", "", p)
    vd = re.sub(r"\D", "", v)
    return bool(pd) and len(pd) >= 3 and pd in vd


def _filtrar_centro(df, columnas, centro):
    """Deja solo las filas del centro pedido. Si no se pidió, no filtra.

    Devuelve (df_filtrado, error). Si el centro no existe se avisa en vez de
    devolver un DataFrame vacío: generar cero documentos sin decir por qué es
    la clase de silencio que cuesta una mañana.
    """
    if not str(centro or "").strip():
        return df, None
    cols = [c for c in columnas if c in df.columns]
    if not cols:
        return df, ("La hoja no tiene columna de centro de maestros, "
                    "así que no puedo filtrar por centro.")
    mask = False
    for c in cols:
        m = df[c].astype(str).map(lambda v: _casa_centro(centro, v))
        mask = m if mask is False else (mask | m)
    dfx = df[mask]
    if dfx.empty:
        return dfx, f"No hay nadie del centro «{centro}» en esa hoja."
    return dfx, None


def _iso(v):
    return bool(re.fullmatch(r"\d{4}-\d{2}-\d{2}", str(v or "").strip()))


def _armar_pendiente(clase, base, personas):
    """Deja el documento listo en session_state. Los NOMBRES viven aquí, en
    el servidor, y se pintan en el panel: nunca entran en la respuesta de
    una herramienta, que es lo que viaja al modelo."""
    pend = dict(base)
    pend["clase"] = clase
    pend["personas"] = list(personas)
    pend.setdefault("fecha_doc", date.today().isoformat())
    st.session_state["doc_pend"] = pend
    for k in ("_doc_mostrar_boton", "doc_generado", "doc_ambiguos"):
        st.session_state.pop(k, None)
    return pend


def ejecutar(nombre, args, rol=None):
    """Ejecuta una herramienta de documentos. Devuelve siempre un dict.

    El gate del admin se comprueba AQUÍ además de en el panel: si algún día
    el ruteo deja pasar un mensaje de un empleado, la herramienta no corre.
    """
    if not puede_documentos(rol):
        return {"error": "Solo el administrador puede generar documentos."}
    if nombre == "preparar_constancias":
        return _t_constancias(args)
    if nombre == "preparar_comisiones":
        return _t_comisiones(args)
    if nombre == "mostrar_boton_generar":
        if not st.session_state.get("doc_pend"):
            return {"error": "Todavía no hay nada preparado."}
        st.session_state["_doc_mostrar_boton"] = True
        return {"listo": True}
    return {"error": f"Herramienta desconocida: {nombre}"}


def _t_constancias(args):
    cli = _deps.get("get_client")
    if not cli:
        return {"error": "El módulo no recibió get_client."}
    try:
        q = int(args.get("quincena") or 0)
        anio = int(args.get("anio") or 0)
    except Exception:
        return {"error": "La quincena y el año tienen que ser números."}
    falta = []
    if not 1 <= q <= 24:
        falta.append("la quincena (1 a 24)")
    if not 2024 <= anio <= 2030:
        falta.append("el año")
    if falta:
        return {"listo": False, "falta": falta}

    df = _cargar_tab(cli(), TAB_CONSTANCIAS)
    if df.empty or "Nombre Completo" not in df.columns:
        return {"error": ("La hoja Constancias está vacía o no tiene la "
                          "columna 'Nombre Completo'.")}
    df, err = _filtrar_centro(
        df, ("Se desempeña en", "C.C.T. ADSCRIPCIÓN", "centro_maestros"),
        args.get("centro"))
    if err:
        return {"listo": False, "error": err}
    disp = list(dict.fromkeys(df["Nombre Completo"].astype(str).tolist()))
    enc, no_hay, amb = _resolver_personas(args.get("personas") or "todos", disp)
    if amb:
        # Los candidatos NO se le mandan al modelo: se resuelven con un
        # selector en el panel, que corre dentro de la app.
        st.session_state["doc_ambiguos"] = {
            "clase": "constancias",
            "base": {"quincena": q, "anio": anio},
            "enc": enc, "amb": amb}
        return {"listo": False, "elegir": len(amb),
                "nota": ("Dile que hay nombres que pueden ser varias personas "
                         "y que los elija en el panel de abajo.")}
    if not enc:
        return {"listo": False,
                "error": "No encontré a nadie con esos nombres en la hoja.",
                "no_encontrados": no_hay}

    _armar_pendiente("constancias",
                     {"quincena": q, "anio": anio,
                      "centro": str(args.get("centro") or "")}, enc)
    # Solo CUÁNTAS. Los nombres se ven en el panel, no salen de la app.
    # no_encontrados son las palabras que escribió el propio admin, así que
    # devolverlas no expone nada que no estuviera ya en la conversación.
    return {"listo": True, "cuantas": len(enc),
            "quincena": q, "anio": anio,
            "no_encontrados": no_hay,
            "nota": ("Dile al admin la quincena, el año y cuántas "
                     "constancias son, y pídele que confirme. NO enumeres "
                     "nombres: los ve en el panel.")}


def _t_comisiones(args):
    cli = _deps.get("get_client")
    if not cli:
        return {"error": "El módulo no recibió get_client."}
    tipo = _tipo_comision(args.get("tipo"))
    try:
        of = int(args.get("oficio_inicial") or 0)
    except Exception:
        of = 0
    fi, ff = str(args.get("fecha_inicio") or ""), str(args.get("fecha_fin") or "")

    falta = []
    if not tipo:
        falta.append("el tipo (encargados de CM o generales)")
    if not 1 <= of <= 9999:
        falta.append("el número de oficio inicial")
    if not _iso(fi):
        falta.append("la fecha de inicio de la comisión")
    if not _iso(ff):
        falta.append("la fecha de fin de la comisión")
    if falta:
        return {"listo": False, "falta": falta}
    if ff < fi:
        return {"listo": False,
                "falta": ["la fecha de fin, que quedó antes de la de inicio"]}

    df = _cargar_tab(cli(), TAB_COMISIONES)
    if df.empty or "nombre_completo" not in df.columns:
        return {"error": ("La hoja Comisiones está vacía o no tiene la "
                          "columna 'nombre_completo'.")}
    if "tipo_comision" not in df.columns:
        return {"error": "La hoja Comisiones no tiene la columna 'tipo_comision'."}

    marca = TIPO_CM if tipo == "Encargados CM" else TIPO_GENERAL
    dfx = df[df["tipo_comision"].astype(str).str.strip() == marca]
    if dfx.empty:
        return {"error": f"No hay registros de tipo '{marca}' en la hoja."}
    dfx, err = _filtrar_centro(dfx, ("centro_maestros", "institucion"),
                               args.get("centro"))
    if err:
        return {"listo": False, "error": err}
    disp = list(dict.fromkeys(dfx["nombre_completo"].astype(str).tolist()))
    enc, no_hay, amb = _resolver_personas(args.get("personas") or "todos", disp)
    _base = {"tipo": tipo, "oficio_inicial": of,
             "fecha_inicio": fi, "fecha_fin": ff,
             "centro": str(args.get("centro") or "")}
    if amb:
        st.session_state["doc_ambiguos"] = {
            "clase": "comisiones", "base": _base, "enc": enc, "amb": amb}
        return {"listo": False, "elegir": len(amb),
                "nota": ("Dile que hay nombres que pueden ser varias personas "
                         "y que los elija en el panel de abajo.")}
    if not enc:
        return {"listo": False,
                "error": "No encontré a nadie con esos nombres en ese tipo.",
                "no_encontrados": no_hay}

    _armar_pendiente("comisiones", _base, enc)
    # El año del oficio sale de la FECHA DEL DOCUMENTO, no de la comisión:
    # así lo arma el generador de app3 (fecha_doc.year) y así se firma. Si
    # aquí usara el año de la comisión, la vista previa diría 2027 y el Word
    # 2026. Un número de oficio que no coincide con el documento es un error
    # de los que se detectan hasta que ya están firmados.
    _ao = st.session_state["doc_pend"]["fecha_doc"][:4]
    return {"listo": True, "tipo": tipo, "cuantos": len(enc),
            "oficios": f"{of}/52/{_ao} al {of + len(enc) - 1}/52/{_ao}",
            "del": fi, "al": ff,
            "no_encontrados": no_hay,
            "nota": ("Dile al admin el tipo, el rango de oficios, las fechas "
                     "y cuántos son, y pídele que confirme. NO enumeres "
                     "nombres: los ve en el panel.")}


# ──────────────────────────────────────────────────────────────────────────
# PANEL
# ──────────────────────────────────────────────────────────────────────────
def _resumen(pend):
    """La línea que ve el admin antes de generar. Lleva título propio porque
    el panel se redibuja al final del hilo y, si entre medias pregunta otra
    cosa, quedaría debajo de esa respuesta y parecería contestarla."""
    if pend["clase"] == "constancias":
        _c = f" · centro **{pend['centro']}**" if pend.get("centro") else ""
        return (f"📄 **Constancias pendientes de generar**\n\n"
                f"Quincena **{pend['quincena']}** de **{pend['anio']}**{_c} · "
                f"**{len(pend['personas'])}** constancias")
    of, n = pend["oficio_inicial"], len(pend["personas"])
    anio = pend["fecha_doc"][:4]
    return (f"📄 **Comisiones pendientes de generar**\n\n"
            f"**{pend['tipo']}** · oficios **{of}/52/{anio}** a "
            f"**{of + n - 1}/52/{anio}** · del **{pend['fecha_inicio']}** "
            f"al **{pend['fecha_fin']}** · **{n}** personas"
            + (f" · centro **{pend['centro']}**" if pend.get("centro") else ""))


def _generar(pend):
    """Llama al generador que corresponde. Devuelve la ruta del .docx.

    Los errores se traducen a algo accionable: la causa realista aquí es que
    falte una plantilla en templates/, y "FileNotFoundError" no le dice al
    admin qué archivo subir ni dónde.
    """
    cli = _deps.get("get_client")
    f_doc = datetime.strptime(pend["fecha_doc"], "%Y-%m-%d").date()
    if pend["clase"] == "constancias":
        df = _cargar_tab(cli(), TAB_CONSTANCIAS)
        dfx = df[df["Nombre Completo"].astype(str).isin(pend["personas"])].copy()
        return generar_constancias_word(dfx, pend["personas"],
                                        pend["quincena"], pend["anio"], f_doc)
    df = _cargar_tab(cli(), TAB_COMISIONES)
    marca = TIPO_CM if pend["tipo"] == "Encargados CM" else TIPO_GENERAL
    dfx = df[(df["tipo_comision"].astype(str).str.strip() == marca)
             & (df["nombre_completo"].astype(str).isin(pend["personas"]))].copy()
    # El orden manda: los oficios se numeran por posición, así que se respeta
    # el orden en que quedaron los nombres resueltos y no el del Sheet.
    dfx["_ord"] = dfx["nombre_completo"].astype(str).map(
        {n: i for i, n in enumerate(pend["personas"])})
    dfx = dfx.sort_values("_ord").drop(columns=["_ord"])
    return generar_comisiones_word(
        dfx, pend["tipo"], pend["oficio_inicial"], f_doc,
        datetime.strptime(pend["fecha_inicio"], "%Y-%m-%d").date(),
        datetime.strptime(pend["fecha_fin"], "%Y-%m-%d").date())


def limpiar():
    for k in ("doc_pend", "doc_generado", "_doc_mostrar_boton",
              "doc_ambiguos"):
        st.session_state.pop(k, None)


def _resolver_ambiguos(estado):
    """Desambigua con un selector, no preguntándole al modelo.

    Si "Juan" puede ser tres personas, los tres nombres se pintan AQUÍ, en la
    app. Mandárselos al modelo para que pregunte sería exponer el padrón a un
    servicio externo para resolver algo que el admin resuelve en dos clics.
    """
    st.markdown("📄 **Hay nombres que pueden ser varias personas**")
    elegidos = []
    for i, (pieza, candidatos) in enumerate(estado["amb"]):
        elegidos.append(st.selectbox(
            f"Escribiste «{pieza}». ¿Quién es?",
            ["— elige —"] + list(candidatos), key=f"doc_amb_{i}"))
    c1, c2 = st.columns([3, 1])
    with c1:
        if st.button("✅ Usar estos nombres", key="doc_amb_ok",
                     type="primary", use_container_width=True):
            malos = [e for e in elegidos if e == "— elige —"]
            if malos:
                st.warning("Falta elegir en todos los casos.")
                return
            personas = list(estado["enc"])
            for e in elegidos:
                if e not in personas:
                    personas.append(e)
            _armar_pendiente(estado["clase"], estado["base"], personas)
            st.rerun()
    with c2:
        if st.button("🗑️ Cancelar", key="doc_amb_no",
                     use_container_width=True):
            limpiar()
            st.rerun()


def panel_documento():
    """Resumen, generación y descargas. Solo admin."""
    if not puede_documentos():
        return
    amb = st.session_state.get("doc_ambiguos")
    if amb:
        _resolver_ambiguos(amb)
        return
    pend = st.session_state.get("doc_pend")
    if not pend:
        return

    st.markdown(_resumen(pend))

    with st.expander(f"👁️ Ver las {len(pend['personas'])} personas"):
        if pend["clase"] == "comisiones":
            of = pend["oficio_inicial"]
            anio = pend["fecha_doc"][:4]
            st.dataframe(
                [{"Oficio": f"{of + i}/52/{anio}", "Nombre": n}
                 for i, n in enumerate(pend["personas"])],
                use_container_width=True, hide_index=True)
        else:
            st.dataframe([{"Nombre": n} for n in pend["personas"]],
                         use_container_width=True, hide_index=True)

    hecho = st.session_state.get("doc_generado") or {}
    if not hecho:
        # app3 tiene un date_input "Fecha del Documento" y aquí hacía falta:
        # el oficio se numera con ESE año y es la fecha que va impresa. Por
        # defecto hoy, que es el caso normal, pero se puede cambiar —por
        # ejemplo para firmar en diciembre comisiones que arrancan en enero.
        _fd = st.date_input(
            "Fecha del documento",
            value=datetime.strptime(pend["fecha_doc"], "%Y-%m-%d").date(),
            key="doc_fecha_doc", format="DD/MM/YYYY")
        if _fd and _fd.isoformat() != pend["fecha_doc"]:
            pend["fecha_doc"] = _fd.isoformat()
            st.session_state["doc_pend"] = pend
            st.rerun()
        if not st.session_state.get("_doc_mostrar_boton"):
            # El agente sigue conversando: todavía no hay nada que apretar.
            if st.button("🗑️ Cancelar", key="doc_cancel_parcial"):
                limpiar()
                st.rerun()
            return
        c1, c2 = st.columns([3, 1])
        with c1:
            if st.button("✅ Generar los documentos", key="doc_generar",
                         type="primary", use_container_width=True):
                with st.spinner("Generando..."):
                    try:
                        ruta = _generar(pend)
                    except FileNotFoundError as e:
                        st.error(
                            "Falta la plantilla en el repo. Sube la carpeta "
                            "`templates/` junto a los .py, con "
                            "`plantilla.docx`, `PLANTILLA_ENCARGADOS_CM.docx` "
                            "y `PLANTILLA_COMISIONES_GENERALES.docx`.\n\n"
                            f"Detalle: {e}")
                        return
                    except Exception as e:
                        st.error(f"No se pudieron generar: {e}")
                        return
                st.session_state["doc_generado"] = {"word": ruta}
                st.rerun()
        with c2:
            if st.button("🗑️ Cancelar", key="doc_cancel",
                         use_container_width=True):
                limpiar()
                st.rerun()
        return

    # Ya generado: descargas. La ruta vive en el temporal del contenedor, así
    # que sobrevive a los reruns pero no a un reinicio de la app; si el
    # archivo ya no está se vuelve a generar en vez de dar un error raro.
    ruta = hecho.get("word")
    if not ruta or not os.path.exists(ruta):
        st.warning("El archivo ya no está en el servidor. Genéralo otra vez.")
        st.session_state.pop("doc_generado", None)
        return

    st.success(f"✅ Listo: {len(pend['personas'])} documentos en un solo Word.")
    c1, c2, c3 = st.columns(3)
    with c1:
        with open(ruta, "rb") as f:
            st.download_button(
                "📄 Descargar Word", f.read(),
                file_name=os.path.basename(ruta),
                mime=("application/vnd.openxmlformats-officedocument"
                      ".wordprocessingml.document"),
                key="doc_dl_word", use_container_width=True)
    with c2:
        pdf = hecho.get("pdf")
        if pdf and os.path.exists(pdf):
            with open(pdf, "rb") as f:
                st.download_button("📕 Descargar PDF", f.read(),
                                   file_name=os.path.basename(pdf),
                                   mime="application/pdf",
                                   key="doc_dl_pdf", use_container_width=True)
        elif st.button("📕 Convertir a PDF", key="doc_pdf",
                       use_container_width=True):
            with st.spinner("Convirtiendo..."):
                p = convertir_word_a_pdf(ruta)
            if p:
                hecho["pdf"] = p
                st.session_state["doc_generado"] = hecho
                st.rerun()
            else:
                # Causa realista: falta libreoffice en packages.txt. No es
                # bloqueante, el Word ya está descargable.
                st.info("No hay conversor de PDF en el servidor. Para "
                        "habilitarlo, agrega `libreoffice` a packages.txt.")
    with c3:
        if st.button("🗑️ Terminar", key="doc_fin", use_container_width=True):
            limpiar()
            st.rerun()


def panel_documento_en_chat():
    """El panel como un mensaje de ARI, igual que el de los pases."""
    if not puede_documentos():
        return
    if not (st.session_state.get("doc_pend")
            or st.session_state.get("doc_ambiguos")):
        return
    with st.chat_message("assistant"):
        panel_documento()


# ══════════════════════════════════════════════════════════════════════════
# NÚCLEO COPIADO VERBATIM DE app3.py
# A partir de aquí NO hay código nuevo: son las funciones de app3 tal como
# están allá. Si algo se arregla aquí, hay que arreglarlo allá también —o
# mejor, mover app3 a importar este módulo.
# ══════════════════════════════════════════════════════════════════════════

_VIALIDADES = (
    'calle', 'callejon', 'callejón', 'av', 'avenida', 'blvd', 'boulevard',
    'bulevar', 'calz', 'calzada', 'carretera', 'camino', 'prolongacion',
    'prolongación', 'andador', 'privada', 'circuito', 'paseo', 'periferico',
    'periférico', 'glorieta', 'retorno', 'cerrada', 'eje', 'via', 'vía'
)


_ASENTAMIENTOS = (
    'colonia', 'col', 'fraccionamiento', 'fracc', 'barrio', 'unidad',
    'residencial', 'ejido', 'ampliacion', 'ampliación', 'delegacion',
    'delegación', 'zona', 'sector', 'conjunto', 'rancheria', 'ranchería'
)


_VACIOS = ('', 'nan', 'none', 'null', 'n/a', 'na', '-', '--', 'sin dato', 'sin datos')


_SEDE_MASCULINA = (
    'centro', 'instituto', 'colegio', 'departamento', 'plantel', 'jardin',
    'jardín', 'edificio', 'organismo', 'consejo', 'archivo', 'auditorio'
)


_SEDE_FEMENINA = (
    'universidad', 'escuela', 'unidad', 'delegacion', 'delegación', 'secretaria',
    'secretaría', 'direccion', 'dirección', 'coordinacion', 'coordinación',
    'region', 'región', 'sede', 'normal', 'benemerita', 'benemérita', 'casa',
    'biblioteca', 'preparatoria', 'jefatura', 'oficina', 'supervision',
    'supervision', 'supervisión', 'zona', 'subdireccion', 'subdirección'
)


def _normaliza_clave(texto):
    """'INSTITUCION', 'Institución', 'institucion ' -> 'institucion'.
    Sirve para localizar una columna del Sheet sin importar mayúsculas,
    acentos ni espacios. SIN esto, un encabezado escrito distinto hacía que
    el dato desapareciera del oficio en silencio.
    """
    t = unicodedata.normalize('NFD', str(texto))
    t = ''.join(c for c in t if unicodedata.category(c) != 'Mn')
    return t.strip().lower().replace(' ', '_')


def valor_limpio(persona, columna):
    """Valor de una columna como texto limpio, tolerando variantes del
    encabezado. Regresa '' si no existe, viene vacía o trae NaN / 'nan'."""
    objetivo = _normaliza_clave(columna)
    valor = ''
    try:
        for clave, dato in persona.items():
            if _normaliza_clave(clave) == objetivo:
                valor = dato
                break
    except Exception:
        valor = ''

    if valor is None:
        return ''
    try:
        if pd.isna(valor):
            return ''
    except (TypeError, ValueError):
        pass

    texto = str(valor).strip()
    if texto.lower() in _VACIOS:
        return ''
    return texto


def _primera_palabra(texto):
    """Primera palabra en minúsculas y sin puntuación, para comparar."""
    if not texto:
        return ''
    return texto.strip().split()[0].strip('.,;:').lower()


def _frase_domicilio(domicilio):
    """'Juárez 123' -> 'la calle Juárez 123'   |   'Av. Vallarta 1234' -> 'Av. Vallarta 1234'"""
    if not domicilio:
        return ''
    if _primera_palabra(domicilio) in _VIALIDADES:
        return domicilio
    return 'la calle ' + domicilio


def _frase_colonia(colonia):
    """'Centro' -> 'colonia Centro'   |   'Fracc. Los Robles' -> 'Fracc. Los Robles'"""
    if not colonia:
        return ''
    if _primera_palabra(colonia) in _ASENTAMIENTOS:
        return colonia
    return 'colonia ' + colonia


def _run_destino(paragraph):
    """Devuelve el run donde se debe escribir el texto ya reemplazado.

    Antes se usaba siempre runs[0]. En la plantilla de Comisiones Generales
    el primer run es un salto de línea en NEGRITAS, así que TODO el párrafo
    salía en negritas. Ahora se elige el primer run que realmente tenga
    texto, que es el que trae el formato correcto del cuerpo.
    """
    for run in paragraph.runs:
        if run.text.strip():
            return run
    return paragraph.runs[0] if paragraph.runs else None


def _con_articulo(nombre):
    """Antepone el artículo al nombre de una sede, si se puede deducir.
    'Escuela Normal Superior de Jalisco' -> 'la Escuela Normal Superior...'
    'Centro de Maestros 1407'            -> 'el Centro de Maestros 1407'
    'UPN 143'                            -> se deja igual (no se adivina)
    """
    if not nombre:
        return ''
    primera = _primera_palabra(nombre)
    if primera in ('la', 'las', 'el', 'los'):
        return nombre
    if primera in _SEDE_FEMENINA:
        return 'la ' + nombre
    if primera in _SEDE_MASCULINA:
        return 'el ' + nombre
    return nombre


def _nombre_sede(persona):
    """La sede es SIEMPRE el Centro de Maestros cuando existe; la institución
    solo funge como sede si no hay Centro de Maestros capturado."""
    return valor_limpio(persona, 'centro_maestros') or valor_limpio(persona, 'institucion')


def construir_sede(persona, tipo_comision=None):
    """Nombre de la sede CON su preposición ya contraída.

    'Centro de Maestros 1407'  -> 'al Centro de Maestros 1407'
    'Escuela Normal Superior'  -> 'a la Escuela Normal Superior'
    """
    nombre = _nombre_sede(persona)
    if not nombre:
        return ''
    con_art = _con_articulo(nombre)
    bajo = con_art.lower()
    if bajo.startswith('el '):
        return 'al ' + con_art[3:]
    if bajo.startswith(('la ', 'las ', 'los ')):
        return 'a ' + con_art
    return 'al ' + con_art


def construir_direccion(persona, tipo_comision=None):
    """Todo lo que va DESPUÉS del nombre de la sede.

    Incluye la institución que aloja a la sede SOLO cuando es distinta de la
    sede misma. Así:
      Centro de Maestros 1407 (dentro de la Normal Superior)
        -> 'ubicado en la Escuela Normal Superior de Jalisco, con domicilio
            en la calle Lisboa 488, colonia Santa Elena Estadio, en el
            municipio de Guadalajara, Jalisco, C.P. 44230'
      UPN 143 Autlán (la sede ES la institución, no se repite)
        -> 'con domicilio en la calle Francisco González Bocanegra 177, ...'
    """
    institucion = valor_limpio(persona, 'institucion')
    sede = _nombre_sede(persona)
    domicilio = valor_limpio(persona, 'domicilio')
    colonia = valor_limpio(persona, 'colonia')
    municipio = valor_limpio(persona, 'municipio')
    # El código postal aparece con varios encabezados según la hoja
    cp = (valor_limpio(persona, 'cp') or valor_limpio(persona, 'codigo_postal')
          or valor_limpio(persona, 'c.p.') or valor_limpio(persona, 'c_p'))

    partes = []

    if institucion and institucion.strip().lower() != sede.strip().lower():
        if _primera_palabra(institucion) in ('en', 'dentro', 'al', 'a'):
            partes.append('ubicado ' + institucion)
        else:
            partes.append('ubicado en ' + _con_articulo(institucion))

    if domicilio:
        partes.append('con domicilio en ' + _frase_domicilio(domicilio))
    if colonia:
        partes.append(_frase_colonia(colonia))
    if municipio:
        partes.append('en el municipio de ' + municipio + ', Jalisco')
    elif partes:
        partes.append('en el estado de Jalisco')
    if cp:
        partes.append('C.P. ' + cp.replace('C.P.', '').replace('CP', '').strip())

    return ', '.join(partes)


def construir_ubicacion(persona):
    """Arma la frase de ubicación según los datos que EXISTAN.

    Casos que resuelve:
      CON institución:
        'ubicado en la Universidad Pedagógica Nacional Unidad 141,
         con domicilio en la calle Juárez 123, colonia Centro,
         en el municipio de Guadalajara, Jalisco'
      SIN institución:
        'ubicado en la calle Juárez 123, colonia Centro,
         en el municipio de Guadalajara, Jalisco'
      Sin domicilio, sin colonia, etc.: simplemente omite esa parte,
      sin dejar comas sueltas.
    """
    institucion = valor_limpio(persona, 'institucion')
    domicilio = valor_limpio(persona, 'domicilio')
    colonia = valor_limpio(persona, 'colonia')
    municipio = valor_limpio(persona, 'municipio')

    partes = []

    if institucion:
        # Si en el Sheet ya viene con preposición ("en la UPN...", "dentro de
        # la Normal...") no la duplicamos.
        if _primera_palabra(institucion) in ('en', 'dentro', 'al', 'a'):
            partes.append('ubicado ' + institucion)
        else:
            partes.append('ubicado en ' + _con_articulo(institucion))
        if domicilio:
            partes.append('con domicilio en ' + _frase_domicilio(domicilio))
    elif domicilio:
        partes.append('ubicado en ' + _frase_domicilio(domicilio))

    if colonia:
        partes.append(_frase_colonia(colonia))

    if municipio:
        partes.append('en el municipio de ' + municipio + ', Jalisco')
    elif partes:
        partes.append('en el estado de Jalisco')

    return ', '.join(partes)


def limpiar_redaccion(texto):
    """Red de seguridad: quita comas huérfanas y espacios dobles que deje
    cualquier marcador vacío (sirve también para la plantilla de Comisiones
    Generales). NO toca los saltos de línea del documento."""
    t = texto
    t = re.sub(r'[ \t]*\(\s*\)', '', t)          # paréntesis vacíos
    t = re.sub(r'[ \t]+([,;:.])', r'\1', t)      # " ," -> ","
    t = re.sub(r'(,[ \t]*){2,}', ', ', t)        # ", ," -> ", "
    t = re.sub(r',[ \t]*\.', '.', t)             # ", ." -> "."
    t = re.sub(r'(?<!\.)\.[ \t]*\.(?!\.)', '.', t)      # ".." -> "."
    t = re.sub(r'([,;])(?=[^\s\d,;])', r'\1 ', t)  # falta espacio tras coma (no toca C.P.)
    t = re.sub(r'[ \t]{2,}', ' ', t)             # espacios dobles
    t = re.sub(r'[ \t]+\n', '\n', t)
    t = re.sub(r'\n[ \t]+', '\n', t)
    return t.strip(' \t')


def generar_constancias_word(df_constancias, empleados_seleccionados, num_quincena, año, fecha_elaboracion):
    """Genera documento Word con constancias conservando formato e imágenes"""
    from docx import Document
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    import os
    
    plantilla_path = os.path.join(os.path.dirname(__file__), 'templates', 'plantilla.docx')
    
    if not os.path.exists(plantilla_path):
        raise FileNotFoundError(f"No se encontró la plantilla en: {plantilla_path}")
    
    meses = {
        1: 'enero', 2: 'febrero', 3: 'marzo', 4: 'abril',
        5: 'mayo', 6: 'junio', 7: 'julio', 8: 'agosto',
        9: 'septiembre', 10: 'octubre', 11: 'noviembre', 12: 'diciembre'
    }
    fecha_texto = f"{fecha_elaboracion.day} de {meses[fecha_elaboracion.month]} de {fecha_elaboracion.year}"
    
    # Lista para guardar documentos individuales
    docs = []
    nombres_unicos = list(dict.fromkeys(empleados_seleccionados))

    for nombre_empleado in nombres_unicos:
        # Filtrar TODOS los registros de este empleado
        registros_empleado = df_constancias[df_constancias['Nombre Completo'] == nombre_empleado]
        
        # Generar una constancia por CADA registro
        for idx, emp in registros_empleado.iterrows():
            # Abrir plantilla NUEVA para cada registro
            doc = Document(plantilla_path)
        
            tel_personal = str(emp['TEL. PERSONAL'])
            if '.' in tel_personal:
                try:
                    tel_personal = f"{int(float(tel_personal)):010d}"
                except:
                    pass
            
            reemplazos = {
                '<<QUINCENA>>': str(num_quincena),
                '<<AÑO>>': str(año),
                '<<FECHA>>': fecha_texto,
                '<<APELLIDO_PATERNO>>': str(emp['Apellido paterno']),
                '<<APELLIDO_MATERNO>>': str(emp['Apellido Materno']),
                '<<NOMBRE>>': str(emp['Nombre(s)']),
                '<<RFC>>': str(emp['RFC']),
                '<<FECHA_INGRESO>>': str(emp['INGRESOA LA SEJ']),
                '<<SE_DESEMPENA_EN>>': str(emp['Se desempeña en']),
                '<<DESCRIPCION_PUESTO>>': str(emp['Descripción de puesto']),
                '<<CCT>>': str(emp['C.C.T. ADSCRIPCIÓN']),
                '<<CLAVE_PRESUPUESTAL>>': str(emp['Clave Presupuestal']),
                '<<TEL_PERSONAL>>': tel_personal,
                '<<TEL_EXT>>': str(emp['TEL. ext.']),
                '<<HOJA>>': str(int(emp['Hoja']))
            }
            
            # Reemplazar en párrafos
            for paragraph in doc.paragraphs:
                texto_completo = paragraph.text
                for marcador, valor in reemplazos.items():
                    if marcador in texto_completo:
                        texto_completo = texto_completo.replace(marcador, valor)
                
                # Actualizar el texto
                if texto_completo != paragraph.text:
                    for run in paragraph.runs:
                        run.text = ''
                    if paragraph.runs:
                        paragraph.runs[0].text = texto_completo
                    else:
                        paragraph.add_run(texto_completo)
            
            # Reemplazar en tablas
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            texto_completo = paragraph.text
                            for marcador, valor in reemplazos.items():
                                if marcador in texto_completo:
                                    texto_completo = texto_completo.replace(marcador, valor)
                            
                            if texto_completo != paragraph.text:
                                for run in paragraph.runs:
                                    run.text = ''
                                if paragraph.runs:
                                    paragraph.runs[0].text = texto_completo
                                else:
                                    paragraph.add_run(texto_completo)
            
            docs.append(doc)
    
    # Combinar documentos SIN docxcompose
    doc_final = docs[0]
    
    for doc in docs[1:]:
        # NO agregar page_break - solo copiar contenido
        for element in list(doc.element.body):
            # SALTAR sectPr (propiedades de sección)
            if element.tag.endswith('sectPr'):
                continue
            doc_final.element.body.append(element)
    
    # A temporal, NO a la carpeta del repo: en app3 el .docx se guardaba
    # junto al código. Aquí no hace falta y ensucia el despliegue.
    output_path = os.path.join(tempfile.gettempdir(),
                               f'Constancias_Q{num_quincena}_{año}.docx')
    doc_final.save(output_path)

    return output_path


def generar_comisiones_word(df_comisiones, tipo_comision, oficio_inicial, fecha_doc, fecha_inicio, fecha_fin):
    """Genera documento Word de comisiones"""
    from docx import Document
    from datetime import datetime
    import os

    try:
        # Seleccionar plantilla según tipo
        if tipo_comision == "Encargados CM":
            plantilla_path = os.path.join(os.path.dirname(__file__), 'templates', 'PLANTILLA_ENCARGADOS_CM.docx')
        else:
            plantilla_path = os.path.join(os.path.dirname(__file__), 'templates', 'PLANTILLA_COMISIONES_GENERALES.docx')

        if not os.path.exists(plantilla_path):
            raise FileNotFoundError(f"Plantilla no encontrada: {plantilla_path}")

        # Generar un documento por cada persona
        docs = []
        oficio_actual = oficio_inicial

        for idx, persona in df_comisiones.iterrows():
            doc = Document(plantilla_path)

            # Diccionario de meses en español
            meses = {
                1: 'enero', 2: 'febrero', 3: 'marzo', 4: 'abril',
                5: 'mayo', 6: 'junio', 7: 'julio', 8: 'agosto',
                9: 'septiembre', 10: 'octubre', 11: 'noviembre', 12: 'diciembre'
            }

            # Convertir fechas a español
            fecha_doc_esp = f"{fecha_doc.day} de {meses[fecha_doc.month]} de {fecha_doc.year}"
            fecha_inicio_esp = f"{fecha_inicio.day} de {meses[fecha_inicio.month]} de {fecha_inicio.year}"
            fecha_fin_esp = f"{fecha_fin.day} de {meses[fecha_fin.month]} de {fecha_fin.year}"

            # ── Marcadores DINÁMICOS: cada columna del Sheet Comisiones queda
            # disponible como <<NOMBRE_COLUMNA>> (mayúsculas, espacios → _).
            # Para agregar un dato nuevo: columna al Sheet + marcador a la
            # plantilla. Sin tocar código. ──
            reemplazos = {}
            for _col, _val in persona.items():
                _marc = "<<" + str(_col).strip().upper().replace(" ", "_") + ">>"
                reemplazos[_marc] = "" if pd.isna(_val) else str(_val)

            # Marcadores fijos (pisan a los dinámicos si coinciden)
            reemplazos.update({
                '<<OFICIO>>': f"{oficio_actual}/52/{fecha_doc.year}",
                '<<FECHA>>': fecha_doc_esp,
                '<<NOMBRE_COMPLETO>>': valor_limpio(persona, 'nombre_completo'),
                '<<FECHA_INICIO>>': fecha_inicio_esp,
                '<<FECHA_FIN>>': fecha_fin_esp
            })

            # ── AQUÍ SE RESUELVE EL PROBLEMA DE LA INSTITUCIÓN ──
            # <<UBICACION>> trae la frase ya redactada, con o sin institución.
            # Los marcadores sueltos se mantienen por compatibilidad, pero la
            # plantilla debe usar <<UBICACION>>.
            reemplazos.update({
                '<<UBICACION>>': construir_ubicacion(persona),
                '<<SEDE>>': construir_sede(persona, tipo_comision),
                '<<DIRECCION>>': construir_direccion(persona, tipo_comision),
                # Cargo por persona. Si la columna 'cargo' del Sheet viene
                # vacía se usa el cargo estándar, así solo hay que capturar
                # las excepciones (p. ej. personal administrativo).
                '<<CARGO>>': (valor_limpio(persona, 'cargo')
                              or 'Asesor Pedagógico de la Dirección de Formación Continua'),
                '<<INSTITUCION>>': valor_limpio(persona, 'institucion'),
                '<<CENTRO_MAESTROS>>': valor_limpio(persona, 'centro_maestros'),
                '<<DOMICILIO>>': valor_limpio(persona, 'domicilio'),
                '<<COLONIA>>': valor_limpio(persona, 'colonia'),
                '<<MUNICIPIO>>': valor_limpio(persona, 'municipio'),
                '<<CP>>': valor_limpio(persona, 'cp')
            })

            # Reemplazar en párrafos
            for paragraph in doc.paragraphs:
                texto_completo = paragraph.text
                hubo_cambio = False
                for marcador, valor in reemplazos.items():
                    if marcador in texto_completo:
                        texto_completo = texto_completo.replace(marcador, str(valor))
                        hubo_cambio = True

                if hubo_cambio:
                    # Limpieza de comas/espacios que dejen los datos vacíos
                    texto_completo = limpiar_redaccion(texto_completo)

                if texto_completo != paragraph.text:
                    tiene_nombre = '<<NOMBRE_COMPLETO>>' in paragraph.text
                    destino = _run_destino(paragraph)
                    if destino is not None:
                        for run in paragraph.runs:
                            run.text = ''
                        destino.text = texto_completo
                        # Negritas SOLO al nombre. NUNCA asignar False: eso
                        # quitaba las negritas que la plantilla ya tenía en
                        # encabezados con <<OFICIO>>, <<FECHA>>, etc.
                        if tiene_nombre:
                            destino.bold = True

            # Reemplazar en tablas si existen
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            texto_completo = paragraph.text
                            hubo_cambio = False
                            for marcador, valor in reemplazos.items():
                                if marcador in texto_completo:
                                    texto_completo = texto_completo.replace(marcador, str(valor))
                                    hubo_cambio = True

                            if hubo_cambio:
                                texto_completo = limpiar_redaccion(texto_completo)

                            if texto_completo != paragraph.text:
                                destino = _run_destino(paragraph)
                                if destino is not None:
                                    for run in paragraph.runs:
                                        run.text = ''
                                    destino.text = texto_completo

            docs.append(doc)
            oficio_actual += 1

        # Combinar documentos SIN docxcompose — CADA COMISIÓN EN SU PROPIA PÁGINA
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn

        doc_final = docs[0]
        _body = doc_final.element.body
        # Insertar ANTES del sectPr (orden válido del XML de Word)
        _sectpr = _body.find(qn('w:sectPr'))

        def _insertar(el):
            if _sectpr is not None:
                _sectpr.addprevious(el)
            else:
                _body.append(el)

        def _es_parrafo_vacio(el):
            return el.tag == qn('w:p') and not ''.join(el.itertext()).strip() \
                and el.find('.//' + qn('w:drawing')) is None \
                and el.find('.//' + qn('w:pict')) is None

        def _podar_final(body):
            """Quita los párrafos vacíos del final. Sin esto, la hoja se
            desbordaba y aparecía una PÁGINA EN BLANCO entre comisión y
            comisión."""
            hijos = [e for e in body if not e.tag.endswith('sectPr')]
            for el in reversed(hijos):
                if _es_parrafo_vacio(el):
                    body.remove(el)
                else:
                    break

        def _marcar_salto(el):
            """Salto de página EN el primer párrafo del siguiente oficio
            (w:pageBreakBefore), no como párrafo aparte: así cada comisión
            inicia en hoja nueva sin generar hojas vacías."""
            if el.tag != qn('w:p'):
                return False
            pPr = el.find(qn('w:pPr'))
            if pPr is None:
                pPr = OxmlElement('w:pPr')
                el.insert(0, pPr)
            if pPr.find(qn('w:pageBreakBefore')) is None:
                pPr.insert(0, OxmlElement('w:pageBreakBefore'))
            return True

        _podar_final(_body)

        for doc in docs[1:]:
            _podar_final(doc.element.body)
            salto_puesto = False
            for element in list(doc.element.body):
                if element.tag.endswith('sectPr'):
                    continue
                if not salto_puesto:
                    salto_puesto = _marcar_salto(element)
                _insertar(element)

        # Guardar
        tipo_archivo = "Encargados_CM" if tipo_comision == "Encargados CM" else "Comisiones_Generales"
        # Igual que en constancias: a temporal, no al repo.
        output_path = os.path.join(tempfile.gettempdir(),
                                   f'{tipo_archivo}_{oficio_inicial}.docx')
        doc_final.save(output_path)

        return output_path

    except Exception as e:
        import traceback
        error_completo = traceback.format_exc()
        raise Exception(f"Error al generar comisiones: {str(e)}\n\nStack trace:\n{error_completo}")


def convertir_word_a_pdf(word_path):
    """Convierte Word a PDF usando LibreOffice directamente"""
    import subprocess
    import os
    
    # Definir ruta de salida
    output_dir = os.path.dirname(os.path.abspath(word_path))
    pdf_path = word_path.replace('.docx', '.pdf')
    
    try:
        # Comando directo de LibreOffice (lowriter es más específico para docs)
        comando = [
            'lowriter',
            '--headless',
            '--convert-to', 'pdf',
            '--outdir', output_dir,
            '-env:UserInstallation=file:///tmp/lo_conversion',
            os.path.abspath(word_path)
        ]
        
        # Ejecutar con timeout de 30 segundos
        result = subprocess.run(comando, capture_output=True, text=True, timeout=30)
        
        if os.path.exists(pdf_path):
            return pdf_path
        else:
            st.error(f"Error de LibreOffice: {result.stderr}")
            return None
            
    except Exception as e:
        st.error(f"No se pudo ejecutar la conversión: {e}")
        return None
