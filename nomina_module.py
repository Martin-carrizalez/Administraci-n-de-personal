"""
MÓDULO: Pendientes de Firma de Nómina — DFC/SEJ
================================================
Pestaña del admin para notificar firmas de nómina pendientes.
Lee todo del Sheet (Directorio_Nomina): empleados, correos, jefes y CC.
Cero datos de servidores públicos hardcodeados.

Salida: vista previa + botón que abre Gmail con el correo prellenado
(destinatario, CC, asunto, cuerpo) + descarga de TXT como respaldo.

Conexión desde app_incidencias.py:
    from nomina_module import render_pendientes_nomina
    render_pendientes_nomina(cargar_directorio_nomina, get_client)
"""

import streamlit as st
import pandas as pd
import urllib.parse
import hashlib


def construir_cuerpo_nomina(nombre, pendientes_por_nomina, segundo_aviso=False):
    """pendientes_por_nomina: {nomina: [conceptos]}"""
    total = sum(len(v) for v in pendientes_por_nomina.values())
    bloques = []
    for nomina, lista in pendientes_por_nomina.items():
        if not lista:
            continue
        items = "\n".join(f"    • {c}" for c in lista)
        bloques.append(f"  Nómina {nomina}:\n{items}")
    bloque = "\n\n".join(bloques)
    if segundo_aviso:
        encabezado = ("Por medio del presente se le hace un SEGUNDO AVISO, toda vez que no se ha "
                      "presentado a regularizar su situación a pesar de haber sido notificado(a) "
                      "previamente mediante correo electrónico.")
        plazo = "2 días hábiles"
        cierre = ("\nDe no presentarse en el plazo indicado, se turnará su caso a la Dirección "
                  "de Pagos para los efectos que procedan.\n")
    else:
        encabezado = (f"Por medio del presente se le notifica que a la fecha cuenta con {total} "
                      "registros de nómina pendientes de firma en la Dirección de Formación Continua, "
                      "correspondientes a los siguientes conceptos y quincenas:")
        plazo = "3 días hábiles"
        cierre = ""
    cuerpo = f"""Estimado(a) C. {nombre}:

{encabezado}

{bloque}

Se le informa que la Dirección de Pagos únicamente permite un rezago máximo
de 2 quincenas. Su situación actual excede dicho límite, por lo que su presencia
para regularizar la firma es INDISPENSABLE.

Se le solicita presentarse en la Dirección de Formación Continua en un plazo no
mayor a {plazo} a partir de la recepción del presente correo.
{cierre}
Sin otro particular, quedo a sus órdenes.

Martín Ángel Carrizalez Piña
Enlace de Recursos Humanos de Dirección de Formación Continua"""
    return cuerpo.strip()


NOMINAS = ["14ADG1075P", "14FMP0001B"]

TAB_PEND_NOMINA = "Pendientes_Nomina"
COLS_PEND_NOMINA = ["FECHA_REGISTRO", "NOMBRE", "NOMINA", "CONCEPTO",
                    "ESTADO", "REGISTRADO_POR", "FECHA_FIRMA"]


def _norm_nom(txt: str) -> str:
    import unicodedata
    t = unicodedata.normalize("NFKD", str(txt or "").upper())
    return " ".join("".join(c for c in t if not unicodedata.combining(c)).split())


def guardar_pendientes_en_sheet(get_client, lista) -> tuple:
    """Deja los pendientes en la tab Pendientes_Nomina para que los
    coordinadores los vean en su botón de Pendientes (la lista de la sesión
    se borra al recargar). Una lectura y una escritura, sin duplicar lo que
    ya estaba registrado como PENDIENTE. Devuelve (nuevos, ya_estaban)."""
    import gspread
    from datetime import datetime
    sh = get_client().open_by_key(st.secrets["sheet_checador_id"])
    try:
        ws = sh.worksheet(TAB_PEND_NOMINA)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(TAB_PEND_NOMINA, rows=2, cols=len(COLS_PEND_NOMINA))
        ws.append_row(COLS_PEND_NOMINA)
    valores = ws.get_all_values()
    headers = valores[0] if valores else COLS_PEND_NOMINA
    idx = {c: headers.index(c) for c in COLS_PEND_NOMINA if c in headers}
    ya = set()
    for f in valores[1:]:
        def _v(c):
            i = idx.get(c, -1)
            return f[i] if 0 <= i < len(f) else ""
        if str(_v("ESTADO")).upper().strip() != "FIRMADO":
            ya.add((_norm_nom(_v("NOMBRE")), str(_v("NOMINA")).strip(), str(_v("CONCEPTO")).strip()))
    hoy = datetime.now().strftime("%Y-%m-%d %H:%M")
    quien = st.session_state.get("nombre", "RH")
    nuevas, repetidas = [], 0
    for x in lista:
        for nom, conceptos in x["pendientes"].items():
            for c in conceptos:
                clave = (_norm_nom(x["nombre"]), str(nom).strip(), str(c).strip())
                if clave in ya:
                    repetidas += 1
                    continue
                ya.add(clave)
                nuevas.append([hoy, x["nombre"], nom, c, "PENDIENTE", quien, ""])
    if nuevas:
        ws.append_rows(nuevas, value_input_option="USER_ENTERED")
    return len(nuevas), repetidas


def _leer_pendientes_abiertos(get_client):
    """(ws, headers, filas_abiertas) de la tab Pendientes_Nomina.
    filas_abiertas: [(n_fila, nombre, nomina, concepto)] con ESTADO != FIRMADO."""
    import gspread
    sh = get_client().open_by_key(st.secrets["sheet_checador_id"])
    try:
        ws = sh.worksheet(TAB_PEND_NOMINA)
    except gspread.WorksheetNotFound:
        return None, [], []
    valores = ws.get_all_values()
    if len(valores) < 2:
        return ws, (valores[0] if valores else []), []
    headers = [h.strip().upper() for h in valores[0]]
    def _i(c):
        return headers.index(c) if c in headers else -1
    i_n, i_nm, i_c, i_e = _i("NOMBRE"), _i("NOMINA"), _i("CONCEPTO"), _i("ESTADO")
    if min(i_n, i_nm, i_c, i_e) < 0:
        return ws, headers, []
    abiertas = []
    for nf, f in enumerate(valores[1:], start=2):
        def _v(i):
            return f[i] if 0 <= i < len(f) else ""
        if str(_v(i_e)).upper().strip() == "FIRMADO":
            continue
        abiertas.append((nf, str(_v(i_n)).strip(), str(_v(i_nm)).strip(), str(_v(i_c)).strip()))
    return ws, headers, abiertas


def reconstruir_lista_desde_sheet(get_client, directorio) -> int:
    """Rearma la lista de trabajo con lo guardado en el Sheet, para volver a
    sacar el PDF y los correos sin recapturar a los 120 empleados."""
    _ws, _h, abiertas = _leer_pendientes_abiertos(get_client)
    if not abiertas:
        return 0
    # Correo y jefe se recuperan del directorio cruzando por nombre normalizado
    datos = {}
    for _, r in directorio.iterrows():
        datos[_norm_nom(r.get("NOMBRE_COMPLETO", ""))] = r
    por_persona = {}
    for _nf, nombre, nomina, concepto in abiertas:
        d = por_persona.setdefault(nombre, {})
        d.setdefault(nomina, [])
        if concepto not in d[nomina]:
            d[nomina].append(concepto)
    lista = []
    for nombre, pend in por_persona.items():
        r = datos.get(_norm_nom(nombre))
        lista.append({
            "id": nombre,
            "nombre": nombre,
            "correo": str(r.get("CORREO", "")) if r is not None else "",
            "jefe": str(r.get("JEFE_INMEDIATO", "")) if r is not None else "",
            "correo_jefe": str(r.get("CORREO_JEFE", "")) if r is not None else "",
            "pendientes": pend,
        })
    st.session_state["lista_nomina"] = lista
    return sum(len(v) for x in lista for v in x["pendientes"].values())


def cerrar_quincena(get_client) -> int:
    """Cierra TODO lo abierto marcándolo FIRMADO con su fecha, para empezar la
    siguiente ronda en limpio. No borra filas: el histórico se conserva."""
    from datetime import datetime
    from gspread.cell import Cell
    ws, headers, abiertas = _leer_pendientes_abiertos(get_client)
    if ws is None or not abiertas:
        return 0
    i_e = headers.index("ESTADO")
    i_f = headers.index("FECHA_FIRMA") if "FECHA_FIRMA" in headers else -1
    ahora = datetime.now().strftime("%Y-%m-%d %H:%M")
    celdas = [Cell(nf, i_e + 1, "FIRMADO") for nf, _, _, _ in abiertas]
    if i_f >= 0:
        celdas += [Cell(nf, i_f + 1, ahora) for nf, _, _, _ in abiertas]
    ws.update_cells(celdas, value_input_option="USER_ENTERED")  # una sola escritura
    st.session_state["lista_nomina"] = []
    return len(abiertas)


def precargar_quincena(get_client, directorio, num_q: int, nominas: list) -> tuple:
    """Da de alta como PENDIENTE a TODO el personal del directorio de nómina
    para la quincena ordinaria: el día de pago nadie ha firmado todavía.
    Así nadie se escapa por un olvido de captura. Solo la quincena ordinaria:
    conceptos como Q17-RETRO aplican a unos cuantos y se siguen capturando a
    mano, para no inventarle una deuda a quien nunca tuvo ese pago.
    Devuelve (creados, ya_estaban, sin_cct)."""
    from datetime import datetime
    concepto = f"Q{num_q}"
    _ws, _h, abiertas = _leer_pendientes_abiertos(get_client)
    import gspread
    sh = get_client().open_by_key(st.secrets["sheet_checador_id"])
    try:
        ws = sh.worksheet(TAB_PEND_NOMINA)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(TAB_PEND_NOMINA, rows=2, cols=len(COLS_PEND_NOMINA))
        ws.append_row(COLS_PEND_NOMINA)
    # No duplicar: se revisa TODO el historial, incluido lo ya FIRMADO, para
    # no revivir a quien ya vino a firmar esta misma quincena.
    valores = ws.get_all_values()
    headers = [h.strip().upper() for h in valores[0]] if valores else COLS_PEND_NOMINA
    def _i(c):
        return headers.index(c) if c in headers else -1
    i_n, i_nm, i_c = _i("NOMBRE"), _i("NOMINA"), _i("CONCEPTO")
    existentes = set()
    for f in valores[1:]:
        def _v(i):
            return f[i] if 0 <= i < len(f) else ""
        existentes.add((_norm_nom(_v(i_n)), str(_v(i_nm)).strip(), str(_v(i_c)).strip()))
    hoy = datetime.now().strftime("%Y-%m-%d %H:%M")
    quien = st.session_state.get("nombre", "RH")
    nuevas, repetidas = [], 0
    sin_cct = []
    for _, r in directorio.iterrows():
        nombre = str(r.get("NOMBRE_COMPLETO", "")).strip()
        if not nombre:
            continue
        # Cada quien SOLO en el o los CCT donde realmente cobra. La celda trae
        # uno ("14ADG1075P") o los dos separados por diagonal
        # ("14ADG1075P / 14FMP0001B"); nunca se asume que cobra en ambos.
        cct = str(r.get("CCT_NOMINA", "") or "").upper()
        suyas = [n for n in nominas if n.upper() in cct]
        if not suyas:
            sin_cct.append(nombre)   # sin dato: no se inventa en cuál cobra
            continue
        for nomina in suyas:
            if (_norm_nom(nombre), nomina, concepto) in existentes:
                repetidas += 1
                continue
            nuevas.append([hoy, nombre, nomina, concepto, "PENDIENTE", quien, ""])
    if nuevas:
        ws.append_rows(nuevas, value_input_option="USER_ENTERED")  # una sola escritura
    return len(nuevas), repetidas, sin_cct


def borrar_precarga(get_client, concepto: str) -> int:
    """Borra del Sheet las filas PENDIENTE de ese concepto (p. ej. 'Q18').
    Para deshacer una precarga equivocada. No toca lo ya FIRMADO, que es
    historial, ni los conceptos capturados a mano de otras quincenas."""
    import gspread
    sh = get_client().open_by_key(st.secrets["sheet_checador_id"])
    try:
        ws = sh.worksheet(TAB_PEND_NOMINA)
    except gspread.WorksheetNotFound:
        return 0
    valores = ws.get_all_values()
    if len(valores) < 2:
        return 0
    headers = [h.strip().upper() for h in valores[0]]
    i_c = headers.index("CONCEPTO") if "CONCEPTO" in headers else -1
    i_e = headers.index("ESTADO") if "ESTADO" in headers else -1
    if i_c < 0 or i_e < 0:
        return 0
    def _v(f, i):
        return f[i] if 0 <= i < len(f) else ""
    quedan = [valores[0]] + [
        f for f in valores[1:]
        if not (str(_v(f, i_c)).strip().upper() == concepto.strip().upper()
                and str(_v(f, i_e)).strip().upper() != "FIRMADO")]
    borradas = len(valores) - len(quedan)
    if borradas:
        ws.clear()
        ws.update(quedan, value_input_option="USER_ENTERED")
        st.cache_data.clear()
    return borradas


def _panel_marcar_firmados(get_client):
    """Marca como FIRMADO lo que ya vino a firmar, para que deje de aparecerle
    al coordinador. NO se borra la fila: queda el histórico con su fecha."""
    import gspread
    from datetime import datetime
    from gspread.cell import Cell
    try:
        sh = get_client().open_by_key(st.secrets["sheet_checador_id"])
        ws = sh.worksheet(TAB_PEND_NOMINA)
        valores = ws.get_all_values()
    except gspread.WorksheetNotFound:
        st.caption("Aún no hay pendientes guardados en el Sheet.")
        return
    except Exception as e:
        st.error(f"No se pudo leer el registro: {e}")
        return
    if len(valores) < 2:
        st.caption("Aún no hay pendientes guardados en el Sheet.")
        return
    headers = [h.strip().upper() for h in valores[0]]
    def _i(col):
        return headers.index(col) if col in headers else -1
    i_nom, i_nomina, i_con, i_est = _i("NOMBRE"), _i("NOMINA"), _i("CONCEPTO"), _i("ESTADO")
    if min(i_nom, i_nomina, i_con, i_est) < 0:
        st.error("La tab Pendientes_Nomina no tiene las columnas esperadas.")
        return
    def _v(f, i):
        return f[i] if 0 <= i < len(f) else ""
    # fila real en el Sheet (base 2) de cada pendiente
    abiertos = {}
    for n_fila, f in enumerate(valores[1:], start=2):
        if str(_v(f, i_est)).upper().strip() == "FIRMADO":
            continue
        abiertos.setdefault(str(_v(f, i_nom)).strip(), []).append(
            (n_fila, str(_v(f, i_nomina)).strip(), str(_v(f, i_con)).strip()))
    if not abiertos:
        st.success("✅ No hay firmas pendientes registradas: todos al corriente.")
        return
    st.caption(f"{sum(len(v) for v in abiertos.values())} pendiente(s) de "
               f"{len(abiertos)} persona(s). Marca lo que ya vino a firmar.")
    marcadas = []
    for nombre in sorted(abiertos):
        with st.expander(f"{nombre} — {len(abiertos[nombre])} pendiente(s)"):
            if st.checkbox("Ya firmó TODO lo de esta persona", key=f"fmt_all_{nombre}"):
                marcadas += [n for n, _, _ in abiertos[nombre]]
            else:
                for n_fila, nomina, concepto in abiertos[nombre]:
                    if st.checkbox(f"{concepto} · nómina {nomina}", key=f"fmd_{n_fila}"):
                        marcadas.append(n_fila)
    # Cierre invertido: con la nómina física enfrente se marcan los POCOS que
    # no firmaron y un clic cierra a todos los demás. Al revés (palomear a los
    # 100 que sí firmaron) sería más trabajo que capturarlos a mano.
    with st.popover("✅ Los NO marcados ya firmaron (cerrar el resto)", use_container_width=True):
        _resto = [n for n, _, _, _ in
                  [(nf, nb, nm, cc) for nb in abiertos for nf, nm, cc in abiertos[nb]]
                  if n not in marcadas]
        st.caption(f"Marca arriba solo a quienes NO firmaron. Al confirmar, los otros "
                   f"{len(_resto)} quedan como FIRMADO con la fecha de hoy.")
        if st.checkbox("Confirmo que el resto ya firmó", key="conf_resto"):
            if st.button("Cerrar el resto", type="primary") and _resto:
                i_fecha = _i("FECHA_FIRMA")
                ahora = datetime.now().strftime("%Y-%m-%d %H:%M")
                celdas = [Cell(n, i_est + 1, "FIRMADO") for n in _resto]
                if i_fecha >= 0:
                    celdas += [Cell(n, i_fecha + 1, ahora) for n in _resto]
                try:
                    ws.update_cells(celdas, value_input_option="USER_ENTERED")
                    st.cache_data.clear()
                    st.success(f"Cerrados {len(_resto)}. Quedan pendientes los {len(marcadas)} que marcaste.")
                    st.rerun()
                except Exception as e:
                    st.error(f"No se pudo cerrar: {e}")

    if marcadas and st.button(f"✅ Marcar {len(marcadas)} como FIRMADO", type="primary"):
        i_fecha = _i("FECHA_FIRMA")
        ahora = datetime.now().strftime("%Y-%m-%d %H:%M")
        celdas = [Cell(n, i_est + 1, "FIRMADO") for n in marcadas]
        if i_fecha >= 0:
            celdas += [Cell(n, i_fecha + 1, ahora) for n in marcadas]
        try:
            ws.update_cells(celdas, value_input_option="USER_ENTERED")  # una sola escritura
            st.cache_data.clear()
            st.success(f"Listo: {len(marcadas)} firma(s) marcadas. "
                       "Ya no le aparecen al coordinador.")
            st.rerun()
        except Exception as e:
            st.error(f"No se pudo actualizar: {e}")


def render_pendientes_nomina(cargar_directorio_nomina, get_client=None,
                             quincena_actual=None):
    st.markdown("### 📋 Pendientes de Firma de Nómina")
    directorio = cargar_directorio_nomina()
    if get_client is not None and not directorio.empty:
        cbt1, cbt2 = st.columns(2)
        # La lista de la sesión se pierde al recargar; esto la rearma desde el
        # Sheet para volver a generar PDF y correos sin recapturar todo.
        if cbt1.button("📥 Cargar pendientes guardados", use_container_width=True):
            try:
                n = reconstruir_lista_desde_sheet(get_client, directorio)
                st.session_state["_msg_nomina"] = (
                    f"Lista recuperada del Sheet: {n} pendiente(s)." if n
                    else "No hay pendientes abiertos en el Sheet.")
                st.rerun()
            except Exception as e:
                st.error(f"No se pudo cargar: {e}")
        with cbt2.popover("🔄 Regenerar lista (cerrar quincena)", use_container_width=True):
            st.caption("Cierra TODOS los pendientes abiertos (quedan como FIRMADO "
                       "con su fecha, no se borran) y deja la lista en blanco para "
                       "empezar la siguiente quincena.")
            if st.checkbox("Confirmo que ya se solventaron", key="conf_regen"):
                if st.button("Sí, cerrar y empezar de nuevo", type="primary"):
                    try:
                        n = cerrar_quincena(get_client)
                        st.session_state["_msg_nomina"] = f"Cerrados {n} pendiente(s). Lista en blanco."
                        st.rerun()
                    except Exception as e:
                        st.error(f"No se pudo cerrar: {e}")
        if quincena_actual:
            _q = quincena_actual()
            if _q:
                _num, _fpago = _q
                with st.popover(f"📅 Precargar Q{_num} a todo el personal", use_container_width=True):
                    st.caption(f"Quincena vigente según el calendario de pagos: **Q{_num}** "
                               f"(pago del {_fpago.strftime('%d/%m/%Y')}). Se dará de alta "
                               "como PENDIENTE a todo el personal en nómina; después marcas "
                               "a quienes van firmando. No se duplica si ya se precargó.")
                    if st.checkbox(f"Deshacer: borrar lo precargado de Q{_num}", key="chk_borrar_prec"):
                        if st.button(f"Borrar pendientes de Q{_num}", key="btn_borrar_prec"):
                            try:
                                n_b = borrar_precarga(get_client, f"Q{_num}")
                                st.session_state["_msg_nomina"] = f"Borrados {n_b} registro(s) de Q{_num}."
                                st.rerun()
                            except Exception as e:
                                st.error(f"No se pudo borrar: {e}")
                    if st.button("Precargar ahora", type="primary", key="btn_precarga"):
                        try:
                            n_new, n_rep, sin_cct = precargar_quincena(
                                get_client, directorio, _num, NOMINAS)
                            msg = f"Precargados {n_new} registro(s) de Q{_num}."
                            if n_rep:
                                msg += f" {n_rep} ya existían (no se duplicaron)."
                            if sin_cct:
                                msg += (f" ⚠️ {len(sin_cct)} sin CCT_NOMINA en el padrón, "
                                        f"no se precargaron: {', '.join(sin_cct[:5])}"
                                        + ("..." if len(sin_cct) > 5 else ""))
                            st.session_state["_msg_nomina"] = msg
                            st.rerun()
                        except Exception as e:
                            st.error(f"No se pudo precargar: {e}")
        with st.expander("✅ Marcar firmas ya recibidas (quitarlas de pendientes)"):
            _panel_marcar_firmados(get_client)
    if directorio.empty:
        st.warning("No se encontró la tab **Directorio_Nomina** o está vacía. "
                   "Crea esa hoja con columnas: ID, NOMBRE_COMPLETO, CORREO, "
                   "JEFE_INMEDIATO, CORREO_JEFE, CC_FIJO.")
        return


    # 1. Conceptos del período (los escribe el usuario), por nómina
    st.markdown("#### 1. Conceptos pendientes de este período")
    st.caption("Escribe los conceptos separados por coma. Ej: Q9 Aguinaldo, Q9 Ayuda de Libros, Q9 RB, Q10, Q11")
    conceptos_por_nomina = {}
    cols = st.columns(len(NOMINAS))
    for i, nom in enumerate(NOMINAS):
        with cols[i]:
            txt = st.text_input(f"Conceptos {nom}", key=f"conceptos_{nom}")
            conceptos_por_nomina[nom] = [c.strip() for c in txt.split(",") if c.strip()]

    # 2. Captura por empleado (selector, sin cruces)
    st.markdown("#### 2. Agregar empleado con pendientes")
    # Cada opción apunta a su FILA exacta, no al ID: si varios empleados traían
    # el ID vacío, todos se resolvían al primero de la lista (bug de nómina).
    directorio = directorio.reset_index(drop=True)
    def _clave_emp(r):
        _id = str(r.get("ID", "")).strip()
        return _id if _id else f"NOM:{str(r.get('NOMBRE_COMPLETO', '')).strip()}"
    opciones = {}
    for i, r in directorio.iterrows():
        _id = str(r.get("ID", "")).strip()
        etiqueta = f"{r['NOMBRE_COMPLETO']}  ·  {_id}" if _id else f"{r['NOMBRE_COMPLETO']}"
        if etiqueta in opciones:          # homónimos sin ID: no se pisan
            etiqueta = f"{etiqueta}  ·  #{i + 1}"
        opciones[etiqueta] = i
    sel = st.selectbox("Busca y elige al empleado", options=["—"] + list(opciones.keys()))

    if "lista_nomina" not in st.session_state:
        st.session_state["lista_nomina"] = []
    if st.session_state.get("_msg_nomina"):
        st.success(st.session_state.pop("_msg_nomina"))

    if sel != "—":
        emp = directorio.iloc[opciones[sel]]
        emp_id = _clave_emp(emp)
        st.caption(f"📧 {emp.get('CORREO','(sin correo)')}  ·  Jefe: {emp.get('JEFE_INMEDIATO','(sin jefe)')}")
        _ya = next((x for x in st.session_state["lista_nomina"] if str(x["id"]) == str(emp_id)), None)
        if _ya:
            st.info("Ya registrado: " + " · ".join(
                f"{n}: {', '.join(c)}" for n, c in _ya["pendientes"].items() if c))
        # Lo marcado se guarda EN EL MOMENTO en un borrador por empleado, no
        # solo al presionar el botón. Antes vivía únicamente en las casillas:
        # si Streamlit las volvía a crear (recarga, caché del directorio que
        # expira, cambio de un dato de arriba) se desmarcaban solas y se
        # perdían los conceptos ya elegidos.
        borrador = st.session_state.setdefault("pend_borrador", {})
        draft = borrador.setdefault(emp_id, {})
        pend_emp = {}
        for nom in NOMINAS:
            if conceptos_por_nomina[nom]:
                st.markdown(f"**{nom}** — marca lo que debe:")
                marcados = []
                previos_nom = draft.get(nom, [])
                ccols = st.columns(min(len(conceptos_por_nomina[nom]), 4) or 1)
                for j, concepto in enumerate(conceptos_por_nomina[nom]):
                    with ccols[j % len(ccols)]:
                        # value= restaura lo ya marcado si la casilla se recreó.
                        if st.checkbox(concepto, key=f"chk_{emp_id}_{nom}_{concepto}",
                                       value=concepto in previos_nom):
                            marcados.append(concepto)
                draft[nom] = marcados
                if marcados:
                    pend_emp[nom] = marcados
        _tot_draft = sum(len(v) for v in draft.values())
        if _tot_draft:
            st.caption(f"✔️ Marcado ahora ({_tot_draft}): " + " · ".join(
                f"{n}: {', '.join(c)}" for n, c in draft.items() if c))
        if st.button("➕ Agregar a la lista", key=f"add_{emp_id}"):
            if not pend_emp:
                st.warning("No marcaste ningún concepto para este empleado.")
            else:
                # COMBINAR, no reemplazar: antes se borraba el registro previo del
                # empleado y solo quedaba lo último marcado (Juan en Q15 perdía su
                # Q14). Ahora lo nuevo se SUMA a lo que ya tenía, sin duplicar.
                previo = next((x for x in st.session_state["lista_nomina"]
                               if str(x["id"]) == str(emp_id)), None)
                if previo is None:
                    st.session_state["lista_nomina"].append({
                        "id": emp_id,
                        "nombre": emp.get("NOMBRE_COMPLETO",""),
                        "correo": emp.get("CORREO",""),
                        "jefe": emp.get("JEFE_INMEDIATO",""),
                        "correo_jefe": emp.get("CORREO_JEFE",""),
                        "pendientes": pend_emp,
                    })
                    st.session_state["_msg_nomina"] = f"Agregado: {emp.get('NOMBRE_COMPLETO','')}"
                else:
                    for nom, conceptos in pend_emp.items():
                        acumulados = previo["pendientes"].setdefault(nom, [])
                        for c in conceptos:
                            if c not in acumulados:
                                acumulados.append(c)
                    total = sum(len(v) for v in previo["pendientes"].values())
                    st.session_state["_msg_nomina"] = (
                        f"Actualizado: {emp.get('NOMBRE_COMPLETO','')} — "
                        f"ahora tiene {total} pendiente(s) acumulados")
                # Ya quedó guardado en la lista: se limpian borrador y casillas
                # para que la siguiente captura empiece en blanco.
                borrador.pop(emp_id, None)
                for _k in [k for k in st.session_state
                           if str(k).startswith(f"chk_{emp_id}_")]:
                    st.session_state.pop(_k, None)
                st.rerun()

    # 3. Lista capturada
    lista = st.session_state["lista_nomina"]
    if not lista:
        st.info("Aún no has agregado empleados a la lista.")
        return
    st.markdown("#### 3. Empleados en la lista")
    # Detalle visible: antes solo se veía el TOTAL y había que abrir el correo
    # para saber qué conceptos quedaron guardados.
    resumen = [{"Nombre": x["nombre"],
                "Pendientes": sum(len(v) for v in x["pendientes"].values()),
                "Detalle": " | ".join(f"{n}: {', '.join(c)}"
                                      for n, c in x["pendientes"].items() if c),
                "Correo": x["correo"]} for x in lista]
    st.dataframe(pd.DataFrame(resumen).sort_values("Pendientes", ascending=False),
                 use_container_width=True, hide_index=True)
    if get_client is not None:
        if st.button("💾 Guardar para los coordinadores", type="primary"):
            try:
                nuevos, repetidos = guardar_pendientes_en_sheet(get_client, lista)
                msg = f"Guardados {nuevos} pendiente(s) en el Sheet."
                if repetidos:
                    msg += f" {repetidos} ya estaban registrados (no se duplicaron)."
                st.success(msg + " Los coordinadores ya los ven en su botón de Pendientes.")
            except Exception as e:
                st.error(f"No se pudieron guardar: {e}")
    if st.button("🗑️ Vaciar lista"):
        st.session_state["lista_nomina"] = []
        st.rerun()

    # CC fijo desde el Sheet (primeras filas de la columna CC_FIJO)
    cc_fijos = [str(c).strip() for c in directorio.get("CC_FIJO", []) if str(c).strip()]

    segundo = st.checkbox("Marcar como SEGUNDO AVISO (tono firme, plazo 2 días)")

    # 4. Vista previa
    st.markdown("#### 4. Vista previa")
    for x in lista:
        total = sum(len(v) for v in x["pendientes"].values())
        para = [c for c in [x["correo"], x["correo_jefe"]] if c]
        prefijo = "SEGUNDO AVISO: " if segundo else ("URGENTE: " if total >= 5 else "")
        asunto = f"{prefijo}Firma de nómina pendiente — {total} registros sin firmar | DFC RH"
        cuerpo = construir_cuerpo_nomina(x["nombre"], x["pendientes"], segundo)
        with st.expander(f"{x['nombre']} — {total} pendientes"):
            st.text(f"Para: {', '.join(para) or '(sin correo)'}")
            st.text(f"CC: {', '.join(cc_fijos) or '(sin CC)'}")
            st.text(f"Asunto: {asunto}")
            # La key incluye una huella del contenido: con una key fija,
            # Streamlit conservaba el texto viejo y la vista previa mostraba
            # menos pendientes de los que realmente llevaba el correo.
            _huella = hashlib.md5(cuerpo.encode("utf-8")).hexdigest()[:8]
            st.text_area("Cuerpo", value=cuerpo, height=280,
                         key=f"prev_{x['id']}_{_huella}")
            # Botón que abre Gmail con todo prellenado (sin permisos de admin)
            if para:
                gmail_url = (
                    "https://mail.google.com/mail/?view=cm&fs=1"
                    + "&to=" + urllib.parse.quote(",".join(para))
                    + ("&cc=" + urllib.parse.quote(",".join(cc_fijos)) if cc_fijos else "")
                    + "&su=" + urllib.parse.quote(asunto)
                    + "&body=" + urllib.parse.quote(cuerpo)
                )
                st.markdown(
                    f'<a href="{gmail_url}" target="_blank" '
                    f'style="display:inline-block;padding:8px 16px;background:#002F6C;'
                    f'color:white;border-radius:6px;text-decoration:none;font-weight:bold;">'
                    f'📧 Abrir en Gmail (revisa y envía)</a>',
                    unsafe_allow_html=True
                )
            else:
                st.warning("Sin correo registrado para este empleado.")

    # 5. Descargar TXT (respaldo siempre disponible)
    txt_lines = ["="*80, "CORREOS DE NOTIFICACIÓN — FIRMA DE NÓMINA PENDIENTE",
                 "Dirección de Formación Continua | SEJ", "="*80, ""]
    for n, x in enumerate(sorted(lista, key=lambda y: sum(len(v) for v in y["pendientes"].values()), reverse=True), 1):
        total = sum(len(v) for v in x["pendientes"].values())
        para = [c for c in [x["correo"], x["correo_jefe"]] if c]
        prefijo = "SEGUNDO AVISO: " if segundo else ("URGENTE: " if total >= 5 else "")
        asunto = f"{prefijo}Firma de nómina pendiente — {total} registros sin firmar | DFC RH"
        txt_lines += ["─"*80, f"#{n} — {x['nombre']} | {total} pendientes",
                      f"PARA: {', '.join(para)}", f"CC: {', '.join(cc_fijos)}",
                      f"ASUNTO: {asunto}", "─"*80,
                      construir_cuerpo_nomina(x["nombre"], x["pendientes"], segundo), "", ""]
    txt_final = "\n".join(txt_lines)
    st.download_button("⬇️ Descargar correos en TXT", data=txt_final,
                       file_name="Correos_Pendientes_Nomina.txt", mime="text/plain")

    # 6. PDF de cartas para imprimir (agrupado por nómina)
    st.markdown("#### 5. Relación de pendientes en PDF (para imprimir)")
    try:
        pdf_bytes = generar_pdf_cartas_nomina(lista, NOMINAS, segundo, conceptos_por_nomina)
        if pdf_bytes:
            st.download_button("📄 Descargar relación en PDF (una hoja por nómina)",
                               data=pdf_bytes, file_name="Relacion_Pendientes_Nomina.pdf",
                               mime="application/pdf")
        else:
            st.warning("No se generó el PDF: revisa que los empleados tengan conceptos marcados.")
    except Exception as e:
        st.error(f"Error generando el PDF: {e}")


def generar_pdf_cartas_nomina(lista, nominas, segundo_aviso=False, conceptos_por_nomina=None):
    """Genera la RELACIÓN de pendientes de firma en PDF, clonando el formato
    del Excel de control (Control_Firmas_Nomina):
      · Fila 1: clave de la nómina (negritas, centrado)
      · Fila 2: QUINCENA Q{min} A Q{max} DEL {año}
      · Headers de concepto con fondo gris, en el ORDEN en que se capturaron
      · Nombres en orden alfabético por columna
      · Pie por columna: PENDIENTES: N (negritas)
    Una página por nómina."""
    if not lista:
        return None
    import io
    import re
    from datetime import datetime
    from reportlab.lib.pagesizes import letter, landscape
    from reportlab.lib.units import cm
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER

    buf = io.BytesIO()
    PAGE = landscape(letter)
    doc = SimpleDocTemplate(buf, pagesize=PAGE,
                            leftMargin=1.5*cm, rightMargin=1.5*cm, topMargin=1.5*cm, bottomMargin=1.5*cm)
    styles = getSampleStyleSheet()
    GRIS_HDR = colors.HexColor("#BFBFBF")
    st_nom = ParagraphStyle("nom", parent=styles["Normal"], fontSize=13, fontName="Helvetica-Bold",
                            alignment=TA_CENTER, spaceAfter=2)
    st_qna = ParagraphStyle("qna", parent=styles["Normal"], fontSize=10, fontName="Helvetica-Bold",
                            alignment=TA_CENTER, spaceAfter=8)

    def _con_q(c):
        c = str(c).strip()
        return c if c.upper().startswith("Q") else f"Q{c}"

    elems = []
    primera = True
    for nom in nominas:
        # dict {concepto: [nombres]} para esta nómina
        por_concepto = {}
        for x in lista:
            for concepto in x["pendientes"].get(nom, []):
                por_concepto.setdefault(concepto, []).append(x["nombre"])
        if not por_concepto:
            continue
        if not primera:
            elems.append(PageBreak())
        primera = False

        # Orden de columnas = orden en que se CAPTURARON los conceptos
        # (no el orden en que se agregaron empleados)
        # Columnas SIEMPRE por número de quincena, no por el orden en que se
        # capturaron: antes salían salteadas (Q16, Q17, Q18, Q15) y así nadie
        # las atiende. Variantes de la misma quincena quedan juntas
        # (Q17, Q17 GU1, Q17-RETRO) y lo que no trae número, al final.
        def _clave_orden(c):
            txt = str(c).strip().upper()
            m = re.search(r"(\d+)", txt)
            if not m:
                return (2, 0, txt)          # sin número: al final
            # Quincenas normales (Q15, Q16...) primero y en orden; después las
            # variantes (Q17-RETRO, Q17 GU1), también en orden. Así las cuatro
            # quincenas corridas quedan juntas en la misma hoja.
            simple = re.fullmatch(r"Q?\s*\d+", txt) is not None
            return (0 if simple else 1, int(m.group(1)), txt)
        orden = sorted(por_concepto.keys(), key=_clave_orden)

        # Nombres en orden alfabético por columna (como el Excel de control)
        for c in orden:
            por_concepto[c] = sorted(por_concepto[c])

        # Línea "QUINCENA Qx A Qy DEL {año}" a partir de los conceptos
        nums = []
        for c in orden:
            m = re.search(r"(\d+)", str(c))
            if m:
                nums.append(int(m.group(1)))
        anio = datetime.now().year
        if nums:
            linea_qna = f"QUINCENA Q{min(nums)} A Q{max(nums)} DEL {anio}"
        else:
            linea_qna = f"CONCENTRADO POR QUINCENA/CONCEPTO — {anio}"

        elems.append(Paragraph(str(nom), st_nom))
        elems.append(Paragraph(linea_qna, st_qna))

        # Máximo 4 columnas por página: con 6 conceptos las columnas quedaban
        # tan angostas que los nombres se encimaban y el PDF era ilegible.
        MAX_COLS = 4
        st_celda = ParagraphStyle("celda", parent=styles["Normal"], fontSize=8,
                                  leading=9.5, fontName="Helvetica")
        st_hdr = ParagraphStyle("hdr", parent=styles["Normal"], fontSize=9, leading=11,
                                fontName="Helvetica-Bold", alignment=TA_CENTER)
        bloques_cols = [orden[i:i + MAX_COLS] for i in range(0, len(orden), MAX_COLS)]
        for nb, cols_bloque in enumerate(bloques_cols):
            if nb:  # los conceptos que no cupieron siguen en otra página
                elems.append(PageBreak())
                elems.append(Paragraph(str(nom), st_nom))
                elems.append(Paragraph(linea_qna + "  (continúa)", st_qna))
            # Paragraph en cada celda para que los nombres largos se PARTAN
            # en varias líneas en vez de desbordarse sobre la columna vecina.
            data = [[Paragraph(_con_q(c), st_hdr) for c in cols_bloque]]
            max_filas = max(len(por_concepto[c]) for c in cols_bloque)
            for i in range(max_filas):
                data.append([Paragraph(por_concepto[c][i], st_celda)
                             if i < len(por_concepto[c]) else "" for c in cols_bloque])
            data.append([Paragraph(f"PENDIENTES: {len(por_concepto[c])}", st_hdr)
                         for c in cols_bloque])
            ancho_col = (PAGE[0] - 3*cm) / len(cols_bloque)
            tb = Table(data, colWidths=[ancho_col]*len(cols_bloque), repeatRows=1)
            ult = len(data) - 1
            tb.setStyle(TableStyle([
                ("BACKGROUND",(0,0),(-1,0), GRIS_HDR),
                ("BACKGROUND",(0,ult),(-1,ult), GRIS_HDR),
                ("VALIGN",(0,0),(-1,-1),"TOP"),
                ("GRID",(0,0),(-1,-1),0.4, colors.grey),
                ("LEFTPADDING",(0,0),(-1,-1),4),
                ("RIGHTPADDING",(0,0),(-1,-1),4),
                ("TOPPADDING",(0,0),(-1,-1),2),
                ("BOTTOMPADDING",(0,0),(-1,-1),2),
                # Filas alternas: ayuda a seguir el renglón con la vista
                ("ROWBACKGROUNDS",(0,1),(-1,ult-1),[colors.white, colors.HexColor("#F2F2F2")]),
            ]))
            elems.append(tb)

    if not elems:
        return None
    doc.build(elems)
    return buf.getvalue()