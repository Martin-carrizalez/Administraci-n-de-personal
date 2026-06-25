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


def render_pendientes_nomina(cargar_directorio_nomina):
    st.markdown("### 📋 Pendientes de Firma de Nómina")
    directorio = cargar_directorio_nomina()
    if directorio.empty:
        st.warning("No se encontró la tab **Directorio_Nomina** o está vacía. "
                   "Crea esa hoja con columnas: ID, NOMBRE_COMPLETO, CORREO, "
                   "JEFE_INMEDIATO, CORREO_JEFE, CC_FIJO.")
        return

    NOMINAS = ["14ADG1075P", "14FMP0001B"]

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
    opciones = {f"{r['NOMBRE_COMPLETO']}  ·  ID {r['ID']}": r["ID"] for _, r in directorio.iterrows()}
    sel = st.selectbox("Busca y elige al empleado", options=["—"] + list(opciones.keys()))

    if "lista_nomina" not in st.session_state:
        st.session_state["lista_nomina"] = []

    if sel != "—":
        emp_id = opciones[sel]
        emp = directorio[directorio["ID"].astype(str) == str(emp_id)].iloc[0]
        st.caption(f"📧 {emp.get('CORREO','(sin correo)')}  ·  Jefe: {emp.get('JEFE_INMEDIATO','(sin jefe)')}")
        pend_emp = {}
        for nom in NOMINAS:
            if conceptos_por_nomina[nom]:
                st.markdown(f"**{nom}** — marca lo que debe:")
                marcados = []
                ccols = st.columns(min(len(conceptos_por_nomina[nom]), 4) or 1)
                for j, concepto in enumerate(conceptos_por_nomina[nom]):
                    with ccols[j % len(ccols)]:
                        if st.checkbox(concepto, key=f"chk_{emp_id}_{nom}_{concepto}"):
                            marcados.append(concepto)
                if marcados:
                    pend_emp[nom] = marcados
        if st.button("➕ Agregar a la lista", key=f"add_{emp_id}"):
            if not pend_emp:
                st.warning("No marcaste ningún concepto para este empleado.")
            else:
                st.session_state["lista_nomina"] = [
                    x for x in st.session_state["lista_nomina"] if str(x["id"]) != str(emp_id)
                ]
                st.session_state["lista_nomina"].append({
                    "id": emp_id,
                    "nombre": emp.get("NOMBRE_COMPLETO",""),
                    "correo": emp.get("CORREO",""),
                    "jefe": emp.get("JEFE_INMEDIATO",""),
                    "correo_jefe": emp.get("CORREO_JEFE",""),
                    "pendientes": pend_emp,
                })
                st.success(f"Agregado: {emp.get('NOMBRE_COMPLETO','')}")

    # 3. Lista capturada
    lista = st.session_state["lista_nomina"]
    if not lista:
        st.info("Aún no has agregado empleados a la lista.")
        return
    st.markdown("#### 3. Empleados en la lista")
    resumen = [{"Nombre": x["nombre"],
                "Pendientes": sum(len(v) for v in x["pendientes"].values()),
                "Correo": x["correo"]} for x in lista]
    st.dataframe(pd.DataFrame(resumen).sort_values("Pendientes", ascending=False),
                 use_container_width=True, hide_index=True)
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
            st.text_area("Cuerpo", value=cuerpo, height=280, key=f"prev_{x['id']}")
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