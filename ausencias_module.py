"""
MÓDULO: Quién no está — DFC/SEJ
================================
Para que cada coordinador y cada responsable vea, antes de que pase, quién de
su gente no va a estar —o no va a estar COMPLETO— y por qué: día económico,
incapacidad, comisión, pase de salida o de entrada, cambio de horario o
guardia.

Se llama "quién no está" y no "quién no viene" porque no es lo mismo: con un
pase de entrada o de salida la persona sí viene, solo que llega tarde o se va
antes, y para saber con cuánta gente se cuenta eso pesa igual.

El objetivo no es vigilar al personal: es que quien coordina sepa con cuánta
gente va a contar el martes, y pueda mover una actividad a tiempo en vez de
enterarse ese día.

Quién ve a quién: por la columna JEFE_INMEDIATO del padrón. Cada quien ve a
los que lo traen como jefe. El admin ve a todos.

Cero datos de personal en el código: todo sale de los Sheets.

Conexión desde app_incidencias.py:
    from ausencias_module import render_ausencias
    render_ausencias({...})
"""

import re
import unicodedata
from datetime import date, datetime, timedelta

import pandas as pd
import streamlit as st

# Cómo se nombra cada causa en pantalla. Las claves son las que devuelve
# build_justificantes_rfc del checador, que es quien cruza las fuentes.
ETIQUETA = {
    "ECONOMICO": ("🟡", "Día económico"),
    "ECO": ("🟡", "Día económico"),
    "INCAPACIDAD": ("🔴", "Incapacidad"),
    "COMISION": ("🔵", "Comisión"),
    "COM": ("🔵", "Comisión"),
    "PASE_ENTRADA": ("🟠", "Pase de entrada"),
    "PEN": ("🟠", "Pase de entrada"),
    "PASE_SALIDA": ("🟠", "Pase de salida"),
    "PSE": ("🟠", "Pase de salida sin retorno"),
    "PSR": ("🟠", "Pase de salida con retorno"),
    "CHO": ("🟣", "Cambio de horario"),
    "RGU": ("🟢", "Reposición de guardia"),
    "CUM": ("🎂", "Día de cumpleaños"),
    "EVENTO": ("⚪", "Evento institucional"),
}

DIAS_ES = ["lunes", "martes", "miércoles", "jueves", "viernes",
           "sábado", "domingo"]
MESES_ES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
            "agosto", "septiembre", "octubre", "noviembre", "diciembre"]


def _norm(txt) -> str:
    """Texto comparable: sin acentos, sin mayúsculas, sin espacios de más."""
    t = unicodedata.normalize("NFKD", str(txt or "").lower())
    t = "".join(c for c in t if not unicodedata.combining(c))
    return " ".join(t.split())


def _solo_nombre(jefe_raw) -> str:
    """Del valor 'Maricela Esquivel Domínguez — Dir. Desarrollo Académico'
    se queda con el nombre.

    La columna JEFE_INMEDIATO trae nombre y cargo pegados con un guion, y el
    cargo cambia (una dirección se renombra, alguien asciende) mientras el
    nombre no. Comparar la cadena completa haría que un equipo entero dejara
    de verse por un cambio de letrero.
    """
    t = str(jefe_raw or "").strip()
    for sep in (" — ", " – ", " - ", "—", "–"):
        if sep in t:
            t = t.split(sep, 1)[0]
            break
    return t.strip()


# Partículas que no identifican a nadie y que unos capturan y otros no.
_PARTICULAS = {"de", "del", "la", "las", "los", "y", "san", "mc"}


def _palabras(nombre) -> set:
    """El nombre vuelto conjunto de palabras, sin acentos ni partículas."""
    return {p for p in re.findall(r"[a-z0-9ñ]+", _norm(_solo_nombre(nombre)))
            if p not in _PARTICULAS and len(p) > 1}


def mismo_nombre(a, b) -> bool:
    """¿Dos formas de escribir a la misma persona?

    Compara CONJUNTOS de palabras, no cadenas: la sesión trae el nombre con
    los apellidos primero ("RAMOS RICO HECTOR MANUEL") y en el padrón los
    jefes están escritos empezando por el nombre ("Héctor Manuel Ramos
    Rico"). Son la misma persona y una comparación por orden las daba por
    distintas, que es como un coordinador se quedaba sin ver a su equipo.

    También vale que uno sea parte del otro —el padrón suele omitir el
    primer nombre—, pero exigiendo al menos dos palabras en común: con una
    sola, cualquier tocayo entraría al equipo equivocado.
    """
    pa, pb = _palabras(a), _palabras(b)
    if not pa or not pb:
        return False
    if pa == pb:
        return True
    chico, grande = (pa, pb) if len(pa) <= len(pb) else (pb, pa)
    return len(chico) >= 2 and chico <= grande




def _etiqueta(clave):
    """(emoji, texto) a partir de lo que devuelve build_justificantes_rfc.

    Esa función NO devuelve siglas: devuelve frases en español con datos
    pegados detrás de barras verticales, p. ej.

        "Pase de salida|retorno:14:00|motivo:trámite"

    Mapear siglas, como estaba antes, hacía que los pases cayeran al caso
    por omisión: salían con el emoji equivocado y con el "|motivo:" crudo a
    la vista. Aquí se parte la cadena y cada dato se escribe donde toca.

    Lo que no se reconoce viene de la columna TIPO de Incapacidades, que
    trae el diagnóstico. Se muestra como "Incapacidad" a secas: quien
    coordina necesita saber que la persona no está, no de qué está enferma.
    """
    bruto = str(clave or "").strip()
    if not bruto:
        return ("", "")
    partes = bruto.split("|")
    base = partes[0].strip()
    extras = {}
    for p in partes[1:]:
        if ":" in p:
            k, v = p.split(":", 1)
            extras[k.strip().lower()] = v.strip()

    b = _norm(base)
    if b.startswith("pase de entrada"):
        emoji, texto = "🟠", "Pase de entrada"
    elif b.startswith("pase de salida sin retorno"):
        emoji, texto = "🟠", "Pase de salida sin retorno"
    elif b.startswith("pase de salida"):
        emoji, texto = "🟠", "Pase de salida"
    elif b.startswith("dia economico"):
        emoji, texto = "🟡", "Día económico"
    elif b.startswith("comision"):
        emoji, texto = "🔵", "Comisión"
    elif b.startswith("reposicion"):
        emoji, texto = "🟢", "Reposición de guardias"
    elif "cumple" in b:
        emoji, texto = "🎂", "Día de cumpleaños"
    elif "evento" in b or "institucional" in b:
        emoji, texto = "⚪", base
    else:
        emoji, texto = "🔴", "Incapacidad"

    if extras.get("retorno"):
        texto += f", regresa {extras['retorno']}"
    # El motivo solo se muestra de lo que el propio empleado escribió al
    # pedir su pase o su día; de una incapacidad no se muestra nada.
    if extras.get("motivo") and emoji != "🔴":
        texto += f" — {extras['motivo']}"
    return emoji, texto

def _fecha_corta(f):
    return f"{DIAS_ES[f.weekday()][:3]} {f.day}"


def _rango(opcion, hoy=None):
    """(desde, hasta, titulo) de la ventana elegida.

    La semana arranca en lunes y termina en viernes: el personal no labora
    sábado ni domingo, y meterlos solo ensancha la tabla con columnas vacías.
    """
    hoy = hoy or date.today()
    if opcion == "Hoy":
        return hoy, hoy, f"hoy, {hoy.day} de {MESES_ES[hoy.month - 1]}"
    if opcion == "Esta semana":
        lun = hoy - timedelta(days=hoy.weekday())
        return lun, lun + timedelta(days=4), "esta semana"
    if opcion == "Próxima semana":
        lun = hoy - timedelta(days=hoy.weekday()) + timedelta(days=7)
        return lun, lun + timedelta(days=4), "la próxima semana"
    # Este mes: de hoy al último día, no desde el día 1. Lo que ya pasó no
    # sirve para decidir si alcanza el personal.
    if hoy.month == 12:
        ultimo = date(hoy.year, 12, 31)
    else:
        ultimo = date(hoy.year, hoy.month + 1, 1) - timedelta(days=1)
    return hoy, ultimo, f"lo que queda de {MESES_ES[hoy.month - 1]}"


def mi_gente(padron, nombre_jefe, es_admin=False):
    """DataFrame con el personal que reporta a esa persona.

    La comparación es por palabras y sin importar el orden (ver
    mismo_nombre): la sesión trae los apellidos primero y el padrón no.
    """
    if padron is None or padron.empty:
        return pd.DataFrame()
    if "JEFE_INMEDIATO" not in padron.columns:
        return pd.DataFrame()
    if es_admin:
        return padron.copy()
    if not _palabras(nombre_jefe):
        return pd.DataFrame()
    toca = padron["JEFE_INMEDIATO"].map(lambda v: mismo_nombre(v, nombre_jefe))
    return padron[toca].copy()


def jefes_del_padron(padron):
    """Nombres de jefe que aparecen en el padrón, para el selector del admin."""
    if padron is None or padron.empty or "JEFE_INMEDIATO" not in padron.columns:
        return []
    vistos = {}
    for v in padron["JEFE_INMEDIATO"]:
        n = _solo_nombre(v)
        if n and _norm(n) not in vistos:
            vistos[_norm(n)] = n
    return sorted(vistos.values())


def ausencias_por_dia(just_rfc, gente, desde, hasta):
    """[(fecha, [(nombre, emoji, causa), ...])] día por día.

    `just_rfc` es {RFC -> {fecha -> causa}}, tal como lo devuelve
    build_justificantes_rfc del checador. Aquí solo se filtra por la gente
    que toca y se ordena por día, sin recalcular nada: el cruce de fuentes ya
    lo hace esa función y duplicarlo sería invitar a que se desincronicen.
    """
    if gente is None or gente.empty:
        return []
    col_rfc = next((c for c in gente.columns if c.upper().strip() == "RFC"), None)
    if not col_rfc:
        return []
    nombres = {}
    for _, r in gente.iterrows():
        rfc = str(r.get(col_rfc, "")).upper().strip()
        if rfc:
            nombres[rfc] = _nombre_de(r)

    # Los eventos institucionales viajan bajo __GLOBAL__ y aplican a TODOS.
    # Sin esto, un día en que nadie labora se vería como día completo y el
    # tablero diría que hay personal cuando no hay ninguno.
    globales = just_rfc.get("__GLOBAL__") or {}

    salida = []
    f = desde
    while f <= hasta:
        dia = []
        for rfc, nombre in nombres.items():
            causa = (just_rfc.get(rfc) or {}).get(f) or globales.get(f)
            if causa:
                emoji, texto = _etiqueta(causa)
                dia.append((nombre, emoji, texto))
        salida.append((f, sorted(dia)))
        f += timedelta(days=1)
    return salida


def _nombre_de(fila):
    """El nombre completo, armado con las columnas que traiga el padrón."""
    for c in ("NOMBRE_COMPLETO", "Nombre completo"):
        if c in fila.index and str(fila.get(c, "")).strip():
            return str(fila[c]).strip()
    partes = [str(fila.get(c, "")).strip()
              for c in ("NOMBRE", "PATERNO", "MATERNO") if c in fila.index]
    junto = " ".join(p for p in partes if p)
    return junto or str(fila.get("RFC", "")).strip()


def tabla_saldos(gente, saldo_fn):
    """[(nombre, tomados, quedan, total)] y los que no tienen saldo dado de alta.

    El saldo NO se calcula aquí: lo da saldo_economicos_rfc, la misma función
    que ve el empleado en su pantalla. Recalcularlo acabaría en que el jefe
    vea un número y su gente otro, y la discusión sería sobre cuál pantalla
    tiene razón en vez de sobre si se autoriza el día.
    """
    filas, sin_alta = [], []
    if gente is None or gente.empty or saldo_fn is None:
        return filas, sin_alta
    col_rfc = next((c for c in gente.columns if c.upper().strip() == "RFC"), None)
    if not col_rfc:
        return filas, sin_alta
    for _, r in gente.iterrows():
        rfc = str(r.get(col_rfc, "")).upper().strip()
        nombre = _nombre_de(r)
        res = saldo_fn(rfc) if rfc else None
        if res is None:
            # Sin fila en la tab Empleados no hay saldo. Se dice con nombre
            # en vez de pintar un 0, que se leería como "ya no le quedan" y
            # significa justo lo contrario.
            sin_alta.append(nombre or rfc)
            continue
        disp, tot = res
        filas.append((nombre, max(0, tot - disp), disp, tot))
    filas.sort(key=lambda f: f[2])      # primero a quien menos le queda
    return filas, sin_alta


def _apartado_saldos(gente, saldo_fn):
    st.markdown("#### 🟡 Días económicos de tu equipo")
    if saldo_fn is None:
        st.info("Esta sección necesita que la app pase la función de saldos.")
        return
    filas, sin_alta = tabla_saldos(gente, saldo_fn)
    if not filas and not sin_alta:
        st.caption("Sin datos de saldo para esta gente.")
        return
    st.caption("Revisa antes de autorizar: son 9 al año, en 3 ocasiones de "
               "hasta 3 días separadas por al menos un mes.")
    if filas:
        st.dataframe(pd.DataFrame(
            [{"Persona": n, "Tomados": t, "Le quedan": q, "De": d}
             for n, t, q, d in filas]),
            use_container_width=True, hide_index=True)
        agotados = [n for n, _, q, _ in filas if q <= 0]
        pocos = [(n, q) for n, _, q, _ in filas if 0 < q <= 2]
        if agotados:
            st.error("Sin días disponibles: **" + "**, **".join(agotados)
                     + "**. No autorices más sin consultarlo con RH.")
        if pocos:
            st.warning("Les queda poco: "
                       + " · ".join(f"**{n}** {q}" for n, q in pocos))
    if sin_alta:
        st.caption("Sin saldo registrado en la tab Empleados: "
                   + ", ".join(sin_alta)
                   + ". No significa que no tengan días: significa que no "
                     "están dados de alta ahí.")


@st.cache_data(ttl=300, show_spinner=False)
def _cruce(_cli, desde, hasta, _f_eco, _f_inc, _f_ind, _f_cruce):
    """(justificantes, hora de lectura). Cacheada 300 s como el resto.

    La hora se calcula DENTRO de la función cacheada, así que es la hora de
    la lectura real y no la del reloj al pintar: si se devolviera fuera,
    diría "leído ahora" mostrando datos de hace cinco minutos.
    """
    vacio = pd.DataFrame()
    fi_s, ff_s = desde.isoformat(), hasta.isoformat()
    eco = _f_eco(_cli, fi_s, ff_s)
    inc = _f_inc(_cli, fi_s, ff_s)
    ind = _f_ind(_cli, fi_s, ff_s)
    just = _f_cruce(eco if eco is not None else vacio,
                    inc if inc is not None else vacio,
                    ind if ind is not None else vacio, desde, hasta)
    return just, datetime.now().strftime("%H:%M")


def _tarjeta_dia(f, ausentes, total, hoy):
    """Un día: cuántos faltan, cuántos quedan y quiénes son."""
    faltan = len(ausentes)
    quedan = max(0, total - faltan)
    # El color va por la proporción que falta, no por el número: que falten
    # 3 de 4 es grave y que falten 3 de 30 no.
    prop = (faltan / total) if total else 0
    color = "#B91C1C" if prop >= 0.5 else ("#B45309" if prop >= 0.25 else "#15803D")
    etiqueta_dia = f"{DIAS_ES[f.weekday()].capitalize()} {f.day}"
    if f == hoy:
        etiqueta_dia += " · hoy"
    st.markdown(
        f"<div style='border-left:4px solid {color};padding:6px 12px;margin:10px 0 4px;'>"
        f"<b>{etiqueta_dia}</b> — <span style='color:{color};font-weight:700;'>"
        f"{quedan} de {total} disponibles</span>"
        + (f" · faltan {faltan}" if faltan else " · nadie falta")
        + "</div>", unsafe_allow_html=True)
    for nombre, emoji, causa in ausentes:
        st.markdown(f"&nbsp;&nbsp;&nbsp;{emoji} **{nombre}** — {causa}",
                    unsafe_allow_html=True)


def render_ausencias(deps):
    """deps: cargar_padron, cargar_solicitudes_eco, get_client, rol, nombre."""
    st.markdown("### 👥 Quién no está")
    st.caption("Tu gente y los días en que falta o no está completa —incluye "
               "pases de entrada y de salida—, para que sepas con cuántos "
               "cuentas antes de comprometer una actividad.")

    es_admin = str(deps.get("rol", "")).lower() == "admin"
    try:
        padron = deps["cargar_padron"]()
    except Exception as e:
        st.error(f"No se pudo leer el padrón: {e}")
        return
    if padron is None or padron.empty:
        st.warning("El padrón llegó vacío.")
        return
    if "JEFE_INMEDIATO" not in padron.columns:
        st.warning("El padrón no trae la columna **JEFE_INMEDIATO**, que es "
                   "de donde sale quién reporta a quién. Sin ella no se puede "
                   "armar el equipo de cada coordinador.")
        return

    # El admin puede pararse en los zapatos de cualquiera: es la forma de
    # revisar que a cada coordinador le aparezca su gente y no la de otro.
    quien = deps.get("nombre", "")
    if es_admin:
        opciones = ["— Todo el personal —"] + jefes_del_padron(padron)
        elegido = st.selectbox("Ver el equipo de", opciones)
        gente = (mi_gente(padron, "", True) if elegido == opciones[0]
                 else mi_gente(padron, elegido))
        titular = "todo el personal" if elegido == opciones[0] else elegido
    else:
        gente = mi_gente(padron, quien)
        titular = quien

    if gente.empty:
        st.info(f"No hay personal que tenga a **{titular}** como jefe "
                "inmediato en el padrón. Si debería haberlo, revisa cómo está "
                "escrito el nombre en la columna JEFE_INMEDIATO.")
        return

    t_dias, t_saldo = st.tabs(["📅 Quién no está", "🟡 Días económicos"])
    with t_saldo:
        _apartado_saldos(gente, deps.get("saldo_economicos"))

    with t_dias:
        _vista_dias(deps, gente, titular, es_admin)


def _vista_dias(deps, gente, titular, es_admin):
    col1, col2 = st.columns([2, 3])
    opcion = col1.radio("Periodo", ["Hoy", "Esta semana", "Próxima semana",
                                    "Este mes"], horizontal=False)
    hoy = date.today()
    desde, hasta, titulo = _rango(opcion, hoy)
    col2.caption(f"**{len(gente)}** persona(s) a cargo de {titular}.")
    col2.caption(f"Mostrando {titulo}: del {desde.strftime('%d/%m')} al "
                 f"{hasta.strftime('%d/%m')}.")

    # Se reutiliza el cruce del checador en vez de rehacerlo. Son tres
    # fuentes distintas —'Solicitudes' los días económicos, 'Incapacidades'
    # las incapacidades e 'Incidencias' las comisiones y pases— y esa función
    # ya sabe cosas que no se ven a simple vista: por ejemplo que un MOTIVO
    # con "Fechas: ..." justifica SOLO esos días y no el rango completo.
    # Copiarla aquí sería tener dos versiones que tarde o temprano discrepan.
    try:
        from checador_module import (load_incapacidades_ch, load_economicos_ch,
                                     load_incidencias_ch, build_justificantes_rfc)
    except Exception as e:
        st.error(f"No se pudo usar el cruce de incidencias del checador: {e}")
        return

    # Se lee SIEMPRE la misma ventana ancha (de esta semana al fin del mes
    # siguiente) y los cuatro periodos se recortan en memoria. Si se leyera
    # el rango elegido, cada cambio de periodo sería una llave de caché
    # nueva y otras 9 llamadas a Google: cuatro clics del coordinador se
    # comían media cuota del minuto, y esa cuota la comparte con los
    # empleados que están capturando sus incidencias.
    _lun = hoy - timedelta(days=hoy.weekday())
    _fin = (date(hoy.year + (hoy.month == 12), (hoy.month % 12) + 1, 1)
            - timedelta(days=1))
    try:
        just_rfc, leido = _cruce(deps["get_client"](), _lun, _fin,
                                 load_economicos_ch, load_incapacidades_ch,
                                 load_incidencias_ch, build_justificantes_rfc)
    except Exception as e:
        st.error(f"No se pudieron leer las incidencias: {e}")
        return

    # Streamlit NO recarga solo: con la pantalla abierta y sin tocar nada, lo
    # que se ve es de cuando se abrió. Y aunque se toque algo, la lectura se
    # reusa hasta 5 minutos. Por eso se dice A QUÉ HORA se leyó y se deja un
    # botón para releer: un tablero de disponibilidad que miente sobre su
    # antigüedad es peor que no tenerlo.
    cabecera, boton = st.columns([4, 1])
    cabecera.caption(f"Datos leídos a las **{leido}**. "
                     "Solo cuentan los pases ya AUTORIZADOS.")
    if boton.button("🔄 Actualizar", key="aus_refrescar",
                    use_container_width=True):
        # Solo esta lectura. st.cache_data.clear() borraría la caché de
        # TODA la app —padrón, empleados, solicitudes— y un coordinador
        # tocando este botón obligaría a releerlo todo, justo cuando los
        # empleados están capturando.
        try:
            _cruce.clear()
        except Exception:
            pass
        st.rerun()

    dias = ausencias_por_dia(just_rfc, gente, desde, hasta)
    total = len(gente)

    # Resumen arriba: el dato que de verdad se usa para decidir es el peor
    # día del periodo, no el promedio.
    con_faltas = [(f, a) for f, a in dias if a]
    if not con_faltas:
        st.success(f"✅ Nadie de tu equipo tiene ausencias ni pases registrados "
                   f"{titulo}. Cuentas con los {total}.")
    else:
        peor = max(con_faltas, key=lambda x: len(x[1]))
        st.warning(f"El día más apretado es el **{DIAS_ES[peor[0].weekday()]} "
                   f"{peor[0].day}**: faltan {len(peor[1])} de {total}.")

    for f, ausentes in dias:
        if opcion == "Este mes" and not ausentes:
            continue   # en un mes entero, los días completos solo estorban
        _tarjeta_dia(f, ausentes, total, hoy)

    st.caption("Fuente: días económicos, incapacidades, comisiones y pases "
               "registrados en el portal. Una ausencia que no se haya "
               "capturado aquí no aparece.")

