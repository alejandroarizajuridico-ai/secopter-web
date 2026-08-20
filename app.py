"""
S.E.C.O.P.T.E.R. Web - Sistema Estratégico de Captura de Oportunidades Públicas
Versión: 10.4 (Persistencia Absoluta en Memoria y Depuración SMTP)
"""
import streamlit as st
from datetime import datetime, timedelta
import requests
import pandas as pd
import concurrent.futures
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText

st.set_page_config(page_title="S.E.C.O.P.T.E.R.", layout="wide")

TIPO_CONTRATO_API_LIKE = {
    "Obras general": ["%OBRA%"], "Obra Vías": ["%OBRA%"], "Obra APSB": ["%OBRA%"],
    "Prestación de servicios": ["%PRESTACI%N DE SERVICIO%", "%PRESTACI%N%"],
    "Consultoría": ["%CONSULTOR%A%", "%CONSULTORIA%"], "Interventoría": ["%INTERVENTOR%A%", "%INTERVENTORIA%"],
    "Suministro": ["%SUMINISTRO%"], "Compraventa": ["%COMPRAVENTA%"],
    "Decreto 092 de 2017": ["%092 DE 2017%", "%092%2017%"], "Seguros": ["%SEGURO%"], "Otro": []
}
KEYWORDS_SUBTIPO = {
    "Obra Vías": ['pavimento', 'pavimentación', 'vía', 'placa huella', 'andén', 'puerto', 'muelle', 'tunel', 'naútico', 'parqueadero', 'reparcheo'],
    "Obra APSB": ['alcantarillado', 'acueducto', 'ptar', 'pozos de inspección', 'suministro de agua', 'limpieza de aguas', 'recurso hídrico']
}
KEYWORDS_EXTRA = {
    "Manual de Contratación": ['manual de contr', 'manuales de contr', 'manual de supervis', 'estatuto de c'],
    "Jurídico": ['abogado', 'juridic', 'derecho', 'defensa', 'apoderado', 'litigio', 'asesoria', 'tutela', 'legal']
}
ESTADOS_ACTIVOS_SECOP_II = ['%BORRADOR%', '%PUBLICADO%', '%PUBLISHED%', '%PRESENTACI%N DE OBSERVACIONES%', '%MANIFESTACI%N DE INTER%S%', '%PRESENTACI%N DE OFERTA%', '%FASE DE OFERTAS%', '%EVALUACI%N%', '%CLOSEDFORREPLIES%']
ESTADOS_ACTIVOS_SECOP_I = ['%BORRADOR%', '%CONVOCADO%', '%EN EVALUACI%N%']
DEPARTAMENTOS = ["Nacional (Todos)", "Amazonas", "Antioquia", "Arauca", "Atlántico", "Bolívar", "Boyacá", "Caldas", "Caquetá", "Casanare", "Cauca", "Cesar", "Chocó", "Córdoba", "Cundinamarca", "Distrito Capital de Bogotá", "Huila", "La Guajira", "Magdalena", "Meta", "Nariño", "Norte de Santander", "Putumayo", "Quindío", "Risaralda", "San Andrés, Providencia y Santa Catalina", "Santander", "Sucre", "Tolima", "Valle del Cauca"]
MODALIDADES = ["Mínima cuantía", "Contratación régimen especial (con ofertas)", "Selección Abreviada de Menor Cuantía", "Seleccion Abreviada Menor Cuantia Sin Manifestacion Interes", "Selección abreviada subasta inversa", "Licitación pública", "Licitación pública (Obra pública)", "Concurso de méritos abierto", "Contratación Directa (Ley 1150 de 2007)", "Régimen Especial"]

def sanitizar_sql_like(texto):
    t = str(texto).upper().replace("'", "''").strip()
    for v in ['Á', 'É', 'Í', 'Ó', 'Ú']: t = t.replace(v, '%')
    return f"%{t}%"

def fetch_api(session, url, query, headers, plataforma):
    payload = {"query": query, "page": {"pageNumber": 1, "pageSize": 5000}, "includeSynthetic": False}
    try:
        resp = session.post(url, headers=headers, json=payload, timeout=60)
        if resp.status_code == 200:
            data = resp.json()
            for i in data: i['_plataforma_origen'] = plataforma
            return data
    except: pass
    return []

st.title("S.E.C.O.P.T.E.R. 🚁")
st.markdown("*Sistema Estratégico de Captura de Oportunidades Públicas y Tendencias Estatales Regionales*")

col1, col2 = st.columns(2)
with col1:
    plataforma = st.selectbox("Plataforma de Búsqueda", ["Ambos (SECOP I y II)", "SECOP II", "SECOP I"])
    tipo_contrato = st.selectbox("Tipo de Contrato", list(TIPO_CONTRATO_API_LIKE.keys()))
    tipo_otro = st.text_input("Escriba el tipo (Si eligió 'Otro')") if tipo_contrato == "Otro" else ""
    modalidades = st.multiselect("Modalidad(es)", MODALIDADES, default=[])
    modalidad_otra = st.text_input("Otra modalidad (Opcional)")

with col2:
    deptos = st.multiselect("Departamentos", DEPARTAMENTOS, default=["Nacional (Todos)"])
    opcion_palabras = st.selectbox("Palabras Inclusivas (En objeto)", ["Ninguna", "Manual de Contratación", "Jurídico", "Otra"])
    palabras_otra = st.text_input("Palabras separadas por coma") if opcion_palabras == "Otra" else ""
    palabras_excluir = st.text_input("Palabras a EXCLUIR (Anti-Ruido)", placeholder="Ej: colegio, salud, hospital")
    
dias = st.slider("Intervalo de búsqueda (Días atrás)", 1, 45, 15)

usar_presupuesto = st.checkbox("Habilitar Filtro de Valor (COP)")
p_min, p_max = 0.0, 999999999999.0
if usar_presupuesto:
    c_min, c_max = st.columns(2)
    p_min = c_min.number_input("Mínimo (COP)", value=0.0, step=1000000.0)
    p_max = c_max.number_input("Máximo (COP)", value=15000000.0, step=1000000.0)

# Inicializar memoria profunda de sesión
if "df_resultados" not in st.session_state:
    st.session_state.df_resultados = None
if "excel_buffer" not in st.session_state:
    st.session_state.excel_buffer = None

if st.button("▶ INICIAR EXTRACCIÓN", type="primary"):
    with st.spinner('Escaneando Datos Abiertos (SECOP I y II)...'):
        mods_seleccionadas = [sanitizar_sql_like(m) for m in modalidades]
        if modalidad_otra: mods_seleccionadas.append(sanitizar_sql_like(modalidad_otra))
        deps_seleccionados = [sanitizar_sql_like(d) for d in deptos]
        fecha_limite = (datetime.now() - timedelta(days=dias)).strftime('%Y-%m-%d')
        
        headers = {"Content-Type": "application/json", "X-App-Token": "bybeB97SohWfAjAaAXPjkPASr"}
        datos_totales = []
        
        with requests.Session() as req_session:
            with concurrent.futures.ThreadPoolExecutor(max_workers=2) as executor:
                tareas = {}
                if plataforma in ["SECOP II", "Ambos (SECOP I y II)"]:
                    where_s2 = [f"fecha_de_publicacion_del >= '{fecha_limite}T00:00:00.000'"]
                    if tipo_contrato == "Otro" and tipo_otro: where_s2.append(f"upper(tipo_de_contrato) LIKE '{sanitizar_sql_like(tipo_otro)}'")
                    elif tipo_contrato in TIPO_CONTRATO_API_LIKE and TIPO_CONTRATO_API_LIKE[tipo_contrato]:
                        where_s2.append("(" + " OR ".join([f"upper(tipo_de_contrato) LIKE '{t}'" for t in TIPO_CONTRATO_API_LIKE[tipo_contrato]]) + ")")
                    if deps_seleccionados and "%NACIONAL (TODOS)%" not in deps_seleccionados:
                        where_s2.append("(" + " OR ".join([f"upper(departamento_entidad) LIKE '{d}'" for d in deps_seleccionados]) + ")")
                    where_s2.append("(" + " OR ".join([f"upper(estado_resumen) LIKE '{e}'" for e in ESTADOS_ACTIVOS_SECOP_II]) + ")")
                    if mods_seleccionadas: where_s2.append("(" + " OR ".join([f"upper(modalidad_de_contratacion) LIKE '{m}'" for m in mods_seleccionadas]) + ")")
                    if usar_presupuesto: where_s2.append(f"precio_base >= {p_min:.0f} AND precio_base <= {p_max:.0f}")
                    
                    q_s2 = f"SELECT * WHERE {' AND '.join(where_s2)} ORDER BY fecha_de_publicacion_del DESC"
                    tareas[executor.submit(fetch_api, req_session, "https://www.datos.gov.co/api/v3/views/p6dx-8zbt/query.json", q_s2, headers, "SECOP II")] = "S2"

                if plataforma in ["SECOP I", "Ambos (SECOP I y II)"]:
                    where_s1 = [f"fecha_de_cargue_en_el_secop >= '{fecha_limite}T00:00:00.000'"]
                    if tipo_contrato == "Otro" and tipo_otro: where_s1.append(f"upper(tipo_de_contrato) LIKE '{sanitizar_sql_like(tipo_otro)}'")
                    elif tipo_contrato in TIPO_CONTRATO_API_LIKE and TIPO_CONTRATO_API_LIKE[tipo_contrato]:
                        where_s1.append("(" + " OR ".join([f"upper(tipo_de_contrato) LIKE '{t}'" for t in TIPO_CONTRATO_API_LIKE[tipo_contrato]]) + ")")
                    if deps_seleccionados and "%NACIONAL (TODOS)%" not in deps_seleccionados:
                        where_s1.append("(" + " OR ".join([f"upper(departamento_entidad) LIKE '{d}'" for d in deps_seleccionados]) + ")")
                    where_s1.append("(" + " OR ".join([f"upper(estado_del_proceso) LIKE '{e}'" for e in ESTADOS_ACTIVOS_SECOP_I]) + ")")
                    if mods_seleccionadas: where_s1.append("(" + " OR ".join([f"upper(modalidad_de_contratacion) LIKE '{m}'" for m in mods_seleccionadas]) + ")")
                    if usar_presupuesto: where_s1.append(f"cuantia_proceso >= {p_min:.0f} AND cuantia_proceso <= {p_max:.0f}")
                    
                    q_s1 = f"SELECT * WHERE {' AND '.join(where_s1)} ORDER BY fecha_de_cargue_en_el_secop DESC"
                    tareas[executor.submit(fetch_api, req_session, "https://www.datos.gov.co/api/v3/views/f789-7hwg/query.json", q_s1, headers, "SECOP I")] = "S1"

                for future in concurrent.futures.as_completed(tareas): datos_totales.extend(future.result())

        procesos_finales = []
        palabras_exclusion = [p.strip().lower() for p in palabras_excluir.split(',')] if palabras_excluir and "ej:" not in palabras_excluir.lower() else []

        for item in datos_totales:
            plat = item.get('_plataforma_origen', 'ND')
            if plat == 'SECOP II':
                nombre = str(item.get('nombre_del_procedimiento', '')).lower()
                desc = str(item.get('descripci_n_del_procedimiento', '')).lower()
                valor = item.get('precio_base', '0')
                estado = item.get('estado_resumen', 'ND')
                fecha_pub = str(item.get('fecha_de_publicacion_del', 'ND')).split('T')[0]
                url_obj = item.get('urlproceso', {})
                enlace = url_obj.get('url', 'Sin enlace') if isinstance(url_obj, dict) else str(url_obj)
            else: 
                nombre = str(item.get('objeto_a_contratar', '')).lower()
                desc = str(item.get('detalle_del_objeto_a_contratar', '')).lower()
                valor = item.get('cuantia_proceso', '0')
                estado = item.get('estado_del_proceso', 'ND')
                fecha_pub = str(item.get('fecha_de_cargue_en_el_secop', 'ND')).split('T')[0]
                enlace = item.get('ruta_proceso_en_secop_i', 'Sin enlace')

            texto_busqueda = f"{nombre} {desc}"
            es_valido = True

            if tipo_contrato in KEYWORDS_SUBTIPO and not any(k in texto_busqueda for k in KEYWORDS_SUBTIPO[tipo_contrato]): es_valido = False
            if es_valido and opcion_palabras != "Ninguna":
                if opcion_palabras in KEYWORDS_EXTRA and not any(k in texto_busqueda for k in KEYWORDS_EXTRA[opcion_palabras]): es_valido = False
                elif opcion_palabras == "Otra":
                    p_extra = [p.strip().lower() for p in palabras_otra.split(',') if p.strip()]
                    if p_extra and not any(p in texto_busqueda for p in p_extra): es_valido = False
            if es_valido and palabras_exclusion and any(ex in texto_busqueda for ex in palabras_exclusion): es_valido = False

            if es_valido:
                procesos_finales.append({
                    'Plataforma': plat, 'Entidad': item.get('entidad', item.get('nombre_entidad', 'ND')),
                    'Objeto': f"{nombre} - {desc}"[:500].capitalize(), 'Tipo Contrato': item.get('tipo_de_contrato', 'ND'),
                    'Valor (COP)': valor, 'Modalidad': item.get('modalidad_de_contratacion', 'ND'),
                    'Estado': estado, 'Enlace': enlace, 'Departamento': item.get('departamento_entidad', 'ND'), 'Fecha Publicación': fecha_pub
                })

        if procesos_finales:
            df = pd.DataFrame(procesos_finales)
            df['Valor (COP)'] = pd.to_numeric(df['Valor (COP)'], errors='coerce')
            st.session_state.df_resultados = df
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer: 
                df.to_excel(writer, index=False)
            st.session_state.excel_buffer = buffer.getvalue()
        else:
            st.session_state.df_resultados = None
            st.session_state.excel_buffer = None
            st.warning("No se encontraron procesos activos con estos filtros.")

# Si hay resultados guardados en memoria, desplegar la tabla interactiva y los botones
if st.session_state.df_resultados is not None:
    st.success(f"✅ Se capturaron {len(st.session_state.df_resultados)} procesos.")
    st.dataframe(st.session_state.df_resultados)
    
    c_descarga, c_correo = st.columns(2)
    c_descarga.download_button(
        label="📥 Descargar Excel", 
        data=st.session_state.excel_buffer, 
        file_name=f"SECOPTER_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx", 
        mime="application/vnd.ms-excel"
    )
    
    st.markdown("### ✉️ Envío Directo por Correo Electrónico")
    correo_destino = st.text_input("Ingrese su correo electrónico", value="", placeholder="ejemplo@gmail.com")
    
    if st.button("Enviar Reporte a mi Correo", type="secondary"):
        if not correo_destino or "@" not in correo_destino:
            st.warning("Por favor, ingrese un correo válido.")
        else:
            try:
                if "email_user" in st.secrets and "email_pass" in st.secrets:
                    remitente = st.secrets["email_user"]
                    password = st.secrets["email_pass"]
                    
                    msg = MIMEMultipart()
                    msg['From'] = remitente
                    msg['To'] = correo_destino
                    msg['Subject'] = f"Reporte S.E.C.O.P.T.E.R. - {datetime.now().strftime('%Y-%m-%d')}"
                    
                    cuerpo = "Hola,\n\nAdjunto encontrarás el reporte consolidado generado por S.E.C.O.P.T.E.R. con las oportunidades contractuales detectadas.No olvides ver mi blog juridico https://derechoyleycolombia.blogspot.com/ y mi portafolio en  contratación https://alejandro-ariza-contratacion.netlify.app y mi Twitter/X https://x.com/AlejoExitiuM .\n\nAtentamente,\nAlejandro Ariza - Asesor en Contratación Estatal /nS.E.C.O.P.T.E.R. - Inteligencia Contractual"
                    msg.attach(MIMEText(cuerpo, 'plain'))
                    
                    adjunto = MIMEApplication(st.session_state.excel_buffer, _subtype="xlsx")
                    adjunto.add_header('Content-Disposition', 'attachment', filename=f"SECOPTER_{datetime.now().strftime('%Y%m%d')}.xlsx")
                    msg.attach(adjunto)
                    
                    # Conexión SMTP y transmisión
                    server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
                    server.login(remitente, password)
                    server.send_message(msg)
                    server.quit()
                    
                    st.success(f"✅ ¡Excel enviado correctamente a {correo_destino}! Revisa tu bandeja de entrada o la carpeta de correo no deseado.")
                else:
                    st.error("Error crítico: Las credenciales (Secrets) no están configuradas en el servidor.")
            except smtplib.SMTPAuthenticationError:
                st.error("🚨 Error de Autenticación: Google rechazó la contraseña. Verifica que la clave de 16 letras esté sin espacios en los Secrets y corresponda a derechoyleycolombia@gmail.com.")
            except Exception as e:
                st.error(f"🚨 Falla en el servidor de correo: {e}")

st.markdown("---")
st.caption("Alejandro Ariza - Asesor en Contratación Estatal - alejandroarizajuridico@gmail.com . Los datos generados están sujetos a la disponibilidad de Datos Abiertos (Colombia Compra Eficiente).")
