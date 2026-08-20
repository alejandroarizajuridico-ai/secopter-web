import os
import requests
import pandas as pd
from datetime import datetime, timedelta
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
import concurrent.futures

def ejecutar_radar_automatico():
    # 1. Configuración de la Búsqueda (Fija para el Bot)
    dias = 1
    fecha_limite = (datetime.now() - timedelta(days=dias)).strftime('%Y-%m-%d')
    departamento = "%SANTANDER%"
    
    headers = {"Content-Type": "application/json"}
    datos_totales = []

    # 2. Extracción SECOP II
    where_s2 = f"fecha_de_publicacion_del >= '{fecha_limite}T00:00:00.000' AND upper(departamento_entidad) LIKE '{departamento}'"
    q_s2 = f"SELECT * WHERE {where_s2}"
    
    # 3. Extracción SECOP I
    where_s1 = f"fecha_de_cargue_en_el_secop >= '{fecha_limite}T00:00:00.000' AND upper(departamento_entidad) LIKE '{departamento}'"
    q_s1 = f"SELECT * WHERE {where_s1}"

    with requests.Session() as req_session:
        # SECOP II
        try:
            resp2 = req_session.post("https://www.datos.gov.co/api/v3/views/p6dx-8zbt/query.json", headers=headers, json={"query": q_s2, "page": {"pageNumber": 1, "pageSize": 5000}, "includeSynthetic": False})
            if resp2.status_code == 200:
                data2 = resp2.json()
                for i in data2: i['_plataforma'] = 'SECOP II'
                datos_totales.extend(data2)
        except Exception as e: print("Error en SECOP II:", e)

        # SECOP I
        try:
            resp1 = req_session.post("https://www.datos.gov.co/api/v3/views/f789-7hwg/query.json", headers=headers, json={"query": q_s1, "page": {"pageNumber": 1, "pageSize": 5000}, "includeSynthetic": False})
            if resp1.status_code == 200:
                data1 = resp1.json()
                for i in data1: i['_plataforma'] = 'SECOP I'
                datos_totales.extend(data1)
        except Exception as e: print("Error en SECOP I:", e)

    if not datos_totales:
        print("No se encontraron procesos nuevos hoy.")
        return

    # 4. Procesamiento de Datos
    procesos_finales = []
    for item in datos_totales:
        plat = item.get('_plataforma')
        if plat == 'SECOP II':
            nombre = item.get('nombre_del_procedimiento', '')
            desc = item.get('descripci_n_del_procedimiento', '')
            valor = item.get('precio_base', '0')
            estado = item.get('estado_resumen', 'ND')
            url_obj = item.get('urlproceso', {})
            enlace = url_obj.get('url', 'Sin enlace') if isinstance(url_obj, dict) else str(url_obj)
        else:
            nombre = item.get('objeto_a_contratar', '')
            desc = item.get('detalle_del_objeto_a_contratar', '')
            valor = item.get('cuantia_proceso', '0')
            estado = item.get('estado_del_proceso', 'ND')
            enlace = item.get('ruta_proceso_en_secop_i', 'Sin enlace')

        # Agregamos el campo Modalidad aquí
        modalidad = item.get('modalidad_de_contratacion', 'ND')

        procesos_finales.append({
            'Plataforma': plat,
            'Entidad': item.get('entidad', item.get('nombre_entidad', 'ND')),
            'Objeto': f"{nombre} - {desc}"[:500].capitalize(),
            'Modalidad': modalidad,
            'Valor (COP)': valor,
            'Estado': estado,
            'Enlace': enlace
        })

    df = pd.DataFrame(procesos_finales)
    df['Valor (COP)'] = pd.to_numeric(df['Valor (COP)'], errors='coerce')
    
    # 5. Envío Automático
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer: df.to_excel(writer, index=False)
    
    remitente = os.environ.get("EMAIL_USER")
    password = os.environ.get("EMAIL_PASS")
    # Este es tu correo destino
    destinatario = "alejandroarizajuridico@gmail.com" 

    if not remitente or not password:
        print("Error: Credenciales no encontradas.")
        return

    msg = MIMEMultipart()
    msg['From'] = remitente
    msg['To'] = destinatario
    msg['Subject'] = f"Reporte Automático S.E.C.O.P.T.E.R. - Santander ({datetime.now().strftime('%Y-%m-%d')})"
    
    cuerpo = f"Buen día.\n\nEl bot de S.E.C.O.P.T.E.R. ha capturado {len(df)} procesos en Santander durante las últimas 48 horas.\n\nAdjunto el reporte consolidado."
    msg.attach(MIMEText(cuerpo, 'plain'))
    
    adjunto = MIMEApplication(buffer.getvalue(), _subtype="xlsx")
    adjunto.add_header('Content-Disposition', 'attachment', filename="Reporte_Diario.xlsx")
    msg.attach(adjunto)
    
    try:
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.login(remitente, password)
        server.send_message(msg)
        server.quit()
        print("Reporte automático enviado con éxito.")
    except Exception as e:
        print(f"Error al enviar: {e}")

if __name__ == "__main__":
    ejecutar_radar_automatico()
