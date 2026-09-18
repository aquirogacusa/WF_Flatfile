import os
import base64
import zipfile
import logging
import smtplib
import sys
import traceback
import pandas as pd
from datetime import datetime
from logging.handlers import TimedRotatingFileHandler
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import argparse
import shutil
import time
import subprocess
import win32com.client
# ===== COLUMNA URL DEL PDF: DESHABILITADA TEMPORALMENTE =====
# Al habilitarla: descomentar estos imports, las funciones encrypt_invoice_no/build_url,
# la asignación de df_csv["URL"] (paso 3.2) y el campo "URL" en new_rows (paso 6).
# Requiere: pip install pycryptodome y configurar URL_BASE en .env
# from Crypto.Cipher import AES
# from Crypto.Util.Padding import pad
# ===========================================================
from dotenv import load_dotenv

# Cargar variables de entorno
load_dotenv()

# Configuración de encriptación AES-256-CBC
AES_SECRET_KEY = b"aURADJRBs1eVjCDTplxWUHDfYjmC61Hr"  # 32 bytes = AES-256
AES_IV = b"HQCDTon0AEqcHkGM"  # 16 bytes = block size AES
URL_BASE = os.getenv(
    "URL_BASE",
    "https://digitaldocs.gruponutresa.com/nrw/?t=facturadeventa&txId=",
)
DUMMY_INVOICE_NO = "99-99999999"

AGENT_NAME = "Wells Fargo Customer Agent"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_FILENAME = "wf_customer_agent.log"

logger = logging.getLogger("wf_customer_agent")

def setup_logging():
    # Registro de logs en archivo (rotación diaria, 30 días de histórico) y consola
    default_log_dir = os.path.join(BASE_DIR, "logs")
    log_dir = os.getenv("LOG_DIR") or default_log_dir
    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(log_dir, LOG_FILENAME)

    formatter = logging.Formatter(
        fmt="%(asctime)s | %(levelname)-8s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    file_handler = TimedRotatingFileHandler(
        log_file, when="midnight", interval=1, backupCount=30, encoding="utf-8"
    )
    file_handler.setFormatter(formatter)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)

    logger.setLevel(logging.INFO)
    logger.handlers.clear()
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    logger.info(f"Log de ejecución: {log_file}")
    return logger

def send_notification_email(subject, body, attachment_path=None):
    # Envía una notificación por correo usando la configuración SMTP del archivo .env
    smtp_host = os.getenv("SMTP_HOST")
    smtp_port = int(os.getenv("SMTP_PORT", "587"))
    smtp_user = os.getenv("SMTP_USER")
    smtp_password = os.getenv("SMTP_PASSWORD")
    use_ssl = os.getenv("SMTP_USE_SSL", "false").strip().lower() in ("1", "true", "yes")
    use_tls = os.getenv("SMTP_USE_TLS", "true").strip().lower() in ("1", "true", "yes")
    mail_from = os.getenv("ALERT_FROM") or smtp_user
    recipients = [
        r.strip()
        for r in os.getenv("ALERT_TO", "").replace(";", ",").split(",")
        if r.strip()
    ]

    if not smtp_host or not mail_from or not recipients:
        logger.error("Notificación NO enviada: configure SMTP_HOST, ALERT_FROM y ALERT_TO en el archivo .env")
        return False

    message = MIMEMultipart()
    message["From"] = mail_from
    message["To"] = ", ".join(recipients)
    message["Subject"] = subject
    message.attach(MIMEText(body, "plain", "utf-8"))

    if attachment_path and os.path.exists(attachment_path):
        with open(attachment_path, "rb") as f:
            part = MIMEApplication(f.read())
        part.add_header(
            "Content-Disposition",
            "attachment",
            filename=os.path.basename(attachment_path),
        )
        message.attach(part)

    try:
        if use_ssl:
            server = smtplib.SMTP_SSL(smtp_host, smtp_port, timeout=30)
        else:
            server = smtplib.SMTP(smtp_host, smtp_port, timeout=30)
        with server:
            if use_tls and not use_ssl:
                server.starttls()
            if smtp_user and smtp_password:
                server.login(smtp_user, smtp_password)
            server.sendmail(mail_from, recipients, message.as_string())
        logger.info(f"Notificación enviada a: {', '.join(recipients)}")
        return True
    except Exception as e:
        logger.error(f"No se pudo enviar la notificación por correo: {e}")
        return False

def notify_failure(step, detail):
    # Notifica por correo cualquier fallo que anule el proceso
    logger.error(f"Fallo en {step}: {detail}")
    subject = f"[ALERTA] {AGENT_NAME} - fallo: {step}"
    body = (
        f"Se presentó un fallo durante la ejecución del agente y el proceso fue ANULADO.\n\n"
        f"Fecha/hora: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        f"Etapa: {step}\n"
        f"Detalle: {detail}\n\n"
        f"El archivo de salida no se regeneró, por lo que Wells Fargo conserva la "
        f"información de la última carga exitosa.\n"
        f"Revise el log del agente para más información.\n"
    )
    return send_notification_email(subject, body)

def is_flatfile_empty(df):
    # El archivo se considera vacío si no tiene filas o si ninguna fila trae dato en la llave Auth
    if df.empty:
        return True
    if "Auth" in df.columns:
        return df["Auth"].notna().sum() == 0
    return False

# ===== COLUMNA URL DEL PDF: DESHABILITADA TEMPORALMENTE =====
# def encrypt_invoice_no(invoice_no):
#     # Quitar el guion intermedio, p.ej. 22-60331413 -> 2260331413
#     clean = str(invoice_no).strip().replace("-", "")
#     # AES-256-CBC con padding PKCS5 (equivale a PKCS7 en bloques de 16 bytes)
#     cipher = AES.new(AES_SECRET_KEY, AES.MODE_CBC, AES_IV)
#     padded = pad(clean.encode("utf-8"), AES.block_size)
#     encrypted = cipher.encrypt(padded)
#     # Codificar en Base64 estándar
#     return base64.b64encode(encrypted).decode("utf-8")
#
# def build_url(invoice_no):
#     # Solo aplica a filas originales; las dummy quedan vacías
#     if str(invoice_no).strip() == DUMMY_INVOICE_NO:
#         return ""
#     return URL_BASE + encrypt_invoice_no(invoice_no)
# ===========================================================

def ensure_sap_running():
    try:
        win32com.client.GetObject("SAPGUI")
        return True
    except:
        pass
        
    logger.info("SAP Logon no está abierto. Intentando iniciarlo automáticamente...")
    common_paths = [
        r"C:\Program Files (x86)\SAP\FrontEnd\SAPGUI\saplogon.exe",
        r"C:\Program Files\SAP\FrontEnd\SAPGUI\saplogon.exe"
    ]
    
    sap_path = None
    for path in common_paths:
        if os.path.exists(path):
            sap_path = path
            break
            
    if not sap_path:
        logger.error("No se pudo encontrar el ejecutable saplogon.exe en las rutas comunes.")
        return False
        
    try:
        subprocess.Popen([sap_path])
        logger.info("Abriendo SAP Logon. Esperando a que el sistema inicialice...")
        time.sleep(10) # Esperar a que SAP inicie y registre los objetos COM
        return True
    except Exception as e:
        logger.error(f"Error al intentar abrir SAP Logon: {e}")
        return False

def export_sap_report_to_file(session, output_file):
    # Exporta la lista ALV actual a un archivo local en formato Spreadsheet (UTF-16, tabs)
    output_file = os.path.abspath(output_file)
    out_dir = os.path.dirname(output_file)
    out_name = os.path.basename(output_file)

    logger.info(f"Exportando reporte a archivo de texto: {output_file}...")
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").select() # Local file...
    time.sleep(1)

    # Seleccionar formato Spreadsheet (texto tabulado)
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
    session.findById("wnd[1]/tbar[0]/btn[0]").press() # Continuar
    time.sleep(1)

    logger.info(f"Guardando en: {output_file}")
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = out_dir
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = out_name
    session.findById("wnd[1]/tbar[0]/btn[11]").press() # Replace (Sobrescribir si existe)
    time.sleep(3)

def extract_sap_report(base_path):
    logger.info("Iniciando extracción de reporte desde SAP...")
    
    # Credenciales y config (idealmente en .env)
    sap_user = os.getenv("SAP_USER")
    sap_password = os.getenv("SAP_PASSWORD")
    sap_connection = os.getenv("SAP_CONNECTION")
    sap_client = os.getenv("SAP_CLIENT", "300")
    sap_lang = os.getenv("SAP_LANGUAGE", "EN")
    
    try:
        # Verificar y abrir SAP si es necesario
        if not ensure_sap_running():
            logger.error("No fue posible iniciar SAP Logon.")
            return False
            
        logger.info("Conectando a SAP GUI...")
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        
        # Iniciar una nueva conexión
        logger.info(f"Abriendo conexión: {sap_connection}...")
        connection = application.OpenConnection(sap_connection, True)
        session = connection.Children(0)
        
        # Hacer Login
        logger.info("Realizando login...")
        session.findById("wnd[0]/usr/txtRSYST-MANDT").text = sap_client
        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = sap_user
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = sap_password
        session.findById("wnd[0]/usr/txtRSYST-LANGU").text = sap_lang
        session.findById("wnd[0]").sendVKey(0) # Enter
        
        # Verificar si hay popup de mensajes (ej. copyright)
        try:
            if session.findById("wnd[1]"):
                session.findById("wnd[1]").sendVKey(0)
        except:
            pass

        # Ir a la transacción
        logger.info("Ejecutando transacción ZSD_POS_1052...")
        session.StartTransaction("ZSD_POS_1052")
        time.sleep(1)
        
        # Cargar Variante CUSA-WF
        logger.info("Cargando variante CUSA-WF...")
        session.findById("wnd[0]").sendVKey(17) # Shift+F5 (Get Variant)
        time.sleep(1)
        session.findById("wnd[1]/usr/txtV-LOW").text = "CUSA-WF"
        session.findById("wnd[1]/usr/txtENAME-LOW").text = "" # Limpiar usuario para buscar globalmente
        session.findById("wnd[1]/tbar[0]/btn[8]").press() # Execute (variante)
        time.sleep(1)
        
        # Ejecutar reporte (Partner Functions = ZA según variante CUSA-WF)
        logger.info("Ejecutando reporte...")
        session.findById("wnd[0]/tbar[1]/btn[8]").press() # Ejecutar F8
        time.sleep(5) # Esperar a que cargue el reporte
        
        # Exportar a Archivo Local (TXT Spreadsheet) -> Customers_SAP.txt (proceso WF)
        export_sap_report_to_file(session, os.path.join(base_path, "Customers_SAP.txt"))

        # Segunda ejecución de la transacción con Partner Functions = Z5 -> Customers_SAP_Z5.txt
        # (lo consume otro proceso, no el agente WF)
        try:
            logger.info("Regresando a la pantalla de selección para la segunda ejecución...")
            session.findById("wnd[0]/tbar[0]/btn[3]").press() # Back (F3)
            time.sleep(2)
            logger.info("Cambiando Partner Functions a Z5...")
            session.findById("wnd[0]/usr/ctxtP_PARVW-LOW").text = "Z5"
            logger.info("Ejecutando reporte con Partner Functions = Z5...")
            session.findById("wnd[0]/tbar[1]/btn[8]").press() # Ejecutar F8
            time.sleep(5)
            export_sap_report_to_file(session, os.path.join(base_path, "Customers_SAP_Z5.txt"))
            logger.info("Reporte Z5 exportado correctamente.")
        except Exception as e:
            logger.error(f"No se pudo generar el archivo Customers_SAP_Z5.txt: {e}")
            send_notification_email(
                f"[ALERTA] {AGENT_NAME} - fallo al extraer reporte Z5 de SAP",
                f"No fue posible extraer el reporte con Partner Functions = Z5 "
                f"(transacción ZSD_POS_1052).\n\n"
                f"Fecha/hora: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
                f"Carpeta de datos: {os.path.abspath(base_path)}\n\n"
                f"El archivo Customers_SAP.txt (Partner Functions = ZA) sí fue generado y el "
                f"proceso de Wells Fargo continúa. El proceso que consume Customers_SAP_Z5.txt "
                f"se verá afectado. Revise el log del agente para más información.\n"
            )
        
        # Hacer logout
        logger.info("Cerrando sesión en SAP...")
        session.findById("wnd[0]/tbar[0]/btn[3]").press() # Back (F3)
        session.findById("wnd[0]/tbar[0]/btn[3]").press() # Back (F3) again to home
        session.findById("wnd[0]/tbar[0]/btn[15]").press() # Exit (Shift+F3)
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press() # Yes to log off
        
        # Cerrar completamente SAP Logon
        logger.info("Cerrando aplicación SAP Logon...")
        subprocess.run(["taskkill", "/F", "/IM", "saplogon.exe"], capture_output=True)
        
        logger.info("Extracción completada.")
        return True

    except Exception as e:
        logger.error(f"Error en automatización SAP: {str(e)}")
        return False

def process_wells_fargo_files(base_path):
    logger.info("=" * 70)
    logger.info(f"Inicio de ejecución - {AGENT_NAME}")
    logger.info(f"Carpeta de datos: {os.path.abspath(base_path)}")

    # 1. Ejecutar extracción de SAP
    if not extract_sap_report(base_path):
        logger.warning("La extracción desde SAP no finalizó correctamente; se continuará con el archivo Customers_SAP.txt existente si está disponible.")
        send_notification_email(
            f"[ALERTA] {AGENT_NAME} - fallo al extraer reporte de SAP",
            f"No fue posible extraer el reporte de SAP (transacción ZSD_POS_1052).\n\n"
            f"Fecha/hora: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            f"Carpeta de datos: {os.path.abspath(base_path)}\n\n"
            f"El agente continuó el proceso con el archivo Customers_SAP.txt existente "
            f"(puede estar desactualizado). Revise el log del agente para más información.\n"
        )
    
    input_zip_filename = os.getenv(
        "INPUT_ZIP_FILENAME",
        "Wells Fargo Bill File Original.zip",
    )
    output_zip_filename = os.getenv(
        "OUTPUT_ZIP_FILENAME",
        "Wells Fargo Bill File.zip",
    )
    txt_filename = "Customers_SAP.txt"
    csv_filename = "FlatFile.csv"
    
    input_zip_path = os.path.join(base_path, input_zip_filename)
    output_zip_path = os.path.join(base_path, output_zip_filename)
    txt_path = os.path.join(base_path, txt_filename)
    temp_csv_path = os.path.join(base_path, csv_filename)
    
    # 2. Descomprimir el archivo zip
    logger.info(f"Descomprimiendo {input_zip_path}...")
    if not os.path.exists(input_zip_path):
        notify_failure("lectura del archivo de entrada", f"No se encontró el archivo {input_zip_path}")
        return False

    try:
        with zipfile.ZipFile(input_zip_path, 'r') as zip_ref:
            if csv_filename not in zip_ref.namelist():
                notify_failure("lectura del archivo de entrada", f"El archivo {input_zip_filename} no contiene {csv_filename}")
                return False
            zip_ref.extract(csv_filename, path=base_path)
    except zipfile.BadZipFile:
        notify_failure("lectura del archivo de entrada", f"El archivo {input_zip_path} no es un ZIP válido o está corrupto.")
        return False
    
    # 3. Leer archivos
    logger.info(f"Leyendo archivos {csv_filename} y {txt_filename}...")
    if not os.path.exists(txt_path):
        notify_failure("lectura del archivo de SAP", f"No se encontró el archivo {txt_path}. Verifica que SAP exportó correctamente.")
        return False

    # Leer CSV con pandas, forzando tipos para la columna Auth
    try:
        df_csv = pd.read_csv(temp_csv_path, dtype={"Auth": str})
    except pd.errors.EmptyDataError:
        df_csv = pd.DataFrame()

    # 3.1 CONTINGENCIA: el FlatFile.csv viene vacío (solo encabezado) por falla en la recarga de SAP BO.
    # No se puede enviar un archivo vacío a Wells Fargo: se anula el proceso y se alerta por correo.
    if is_flatfile_empty(df_csv):
        logger.error(
            f"CONTINGENCIA ACTIVADA: {csv_filename} no contiene filas de información "
            f"(solo encabezado o archivo vacío). Se anula el proceso para no enviar "
            f"un archivo vacío a Wells Fargo."
        )
        subject = f"[ALERTA] {AGENT_NAME} - {csv_filename} vacío: proceso anulado"
        body = (
            f"Se detectó que el archivo {csv_filename} contenido en {input_zip_filename} "
            f"no tiene filas de información (solo encabezado o archivo vacío).\n\n"
            f"Fecha/hora de detección: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            f"Carpeta de datos: {os.path.abspath(base_path)}\n"
            f"Archivo origen: {input_zip_path}\n\n"
            f"Acción tomada: el proceso fue ANULADO. No se envió ni se regeneró el archivo "
            f"{output_zip_filename}, por lo que Wells Fargo conserva la información de la "
            f"última carga exitosa.\n\n"
            f"Causa probable: falla en la recarga de SAP BO que origina el archivo.\n"
            f"Por favor revisar la fuente y volver a ejecutar el agente.\n"
        )
        send_notification_email(subject, body, attachment_path=temp_csv_path)
        logger.warning(
            f"Proceso anulado. El ZIP original ({input_zip_filename}) y el ZIP de salida "
            f"({output_zip_filename}) quedaron intactos."
        )
        return False

    logger.info(f"{csv_filename} contiene {len(df_csv)} filas de información.")

    # ===== COLUMNA URL DEL PDF: DESHABILITADA TEMPORALMENTE =====
    # Para habilitarla: descomentar la siguiente línea y el campo "URL" en new_rows (paso 6).
    # df_csv["URL"] = df_csv["Biller Invoice No."].apply(build_url)
    # ===========================================================
    
    # Leer TXT exportado de SAP (Formato Spreadsheet es UTF-16, separado por tabs)
    # El layout del reporte ALV puede cambiar de orden: se localiza dinámicamente la fila
    # de encabezados y se toman únicamente las columnas necesarias por nombre.
    try:
        with open(txt_path, 'r', encoding='utf-16') as f:
            lines = f.readlines()
    except Exception as e:
        notify_failure("lectura del archivo de SAP", f"Error leyendo el archivo {txt_filename}: {e}")
        return False

    required_columns = ['Customer', 'Name 1', 'Terms Paym']
    header_idx = None
    for idx, line in enumerate(lines):
        fields = [field.strip() for field in line.rstrip('\r\n').split('\t')]
        if all(col in fields for col in required_columns):
            header_idx = idx
            break

    if header_idx is None:
        notify_failure(
            "lectura del archivo de SAP",
            f"No se encontraron las columnas {', '.join(required_columns)} en {txt_filename}. "
            f"Verifique que el layout del reporte en SAP conserve estos encabezados.",
        )
        return False

    try:
        df_sap_raw = pd.read_csv(
            txt_path, sep='\t', encoding='utf-16', skiprows=header_idx, header=0, on_bad_lines='skip'
        )
    except Exception as e:
        notify_failure("lectura del archivo de SAP", f"Error leyendo el archivo {txt_filename}: {e}")
        return False

    # 4. Procesar y filtrar datos de SAP
    # Se toman solo las columnas necesarias, sin importar su posición en el layout
    df_sap = df_sap_raw[required_columns].copy()
    df_sap.columns = ['Customer_SAP', 'Name 1', 'Terms of payment']
    
    # Convertir a string
    df_sap['Customer_SAP'] = df_sap['Customer_SAP'].astype(str)
    df_sap['Terms of payment'] = df_sap['Terms of payment'].astype(str)
    
    # Remover los 2 ceros a la izquierda del código (si tiene 10 caracteres y empieza con 00)
    def clean_customer_code(code):
        code = str(code).strip()
        # En caso de que se haya leído como float "0010380463.0"
        if code.endswith('.0'):
            code = code[:-2]
        
        # Eliminar 2 ceros a la izquierda
        if len(code) == 10 and code.startswith('00'):
            return code[2:]
        return code
        
    df_sap['Clean_Customer'] = df_sap['Customer_SAP'].apply(clean_customer_code)
    
    # 5. Identificar clientes faltantes
    # Llave CSV: Auth (segunda columna)
    # Llave SAP: Clean_Customer
    csv_auth_list = set(df_csv['Auth'].dropna().unique())
    
    # Obtener todos los clientes de SAP que no estén en el CSV y eliminar duplicados
    missing_customers = df_sap[~df_sap['Clean_Customer'].isin(csv_auth_list)].drop_duplicates(subset=['Clean_Customer'])
    
    if missing_customers.empty:
        logger.info("No se encontraron clientes faltantes en SAP.")
        df_updated = df_csv
    else:
        logger.info(f"Insertando {len(missing_customers)} clientes faltantes...")
        
        # 6. Preparar datos dummy
        today_str = datetime.now().strftime("%m/%d/%Y")
        
        new_rows = []
        for _, row in missing_customers.iterrows():
            new_rows.append({
                "Payer Code": row["Clean_Customer"],
                "Auth": row["Clean_Customer"],
                "Customer": row["Name 1"],
                "Due Date": today_str,
                "Amount Due": 0,
                "Biller Invoice No.": "99-99999999",
                "P.O.": "POAdvance",
                "Customer Name": row["Name 1"],
                "Bank Account": "4942472523",
                # "URL": ""   # COLUMNA URL DEL PDF: deshabilitada temporalmente
            })
        
        df_new = pd.DataFrame(new_rows)
        df_updated = pd.concat([df_csv, df_new], ignore_index=True)

    # Guardar CSV actualizado
    df_updated.to_csv(temp_csv_path, index=False, quoting=1) # quoting=1 (QUOTE_ALL)
    logger.info(f"CSV actualizado guardado con {len(df_updated)} filas en total.")

    # 7. Repackage y Cleanup
    logger.info("Repackaging ZIP y limpiando...")
    
    # Eliminar el zip original
    os.remove(input_zip_path)
    
    # Crear nuevo zip con el CSV actualizado
    logger.info(f"Creando {output_zip_path}...")
    with zipfile.ZipFile(output_zip_path, 'w', zipfile.ZIP_DEFLATED) as zip_new:
        zip_new.write(temp_csv_path, arcname=csv_filename)
    
    # Eliminar el CSV extraído
    os.remove(temp_csv_path)
    
    # 8. Notificación de éxito con resumen ejecutivo (se desactiva con SUCCESS_NOTIFICATION=false en .env)
    send_success = os.getenv("SUCCESS_NOTIFICATION", "true").strip().lower() in ("1", "true", "yes")
    if send_success:
        total_cartera = pd.to_numeric(df_csv["Amount Due"], errors="coerce").sum()
        success_body = (
            f"Resumen ejecutivo - {AGENT_NAME}\n\n"
            f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            f"Total valor de cartera: ${total_cartera:,.2f}\n"
            f"Clientes con cartera: {df_csv['Auth'].nunique()}\n"
            f"Clientes agregados sin cartera: {len(missing_customers)}\n\n"
            f"Se dejó en el path de red el archivo {output_zip_filename} para transmitir a Wells Fargo exitosamente.\n"
        )
        send_notification_email(f"[INFO] {AGENT_NAME} - proceso exitoso", success_body)
    else:
        logger.info("Notificación de éxito deshabilitada (SUCCESS_NOTIFICATION=false).")

    logger.info("Proceso completado exitosamente.")
    logger.info("=" * 70)
    return True

def main():
    parser = argparse.ArgumentParser(description="Wells Fargo Customer Agent")
    parser.add_argument(
        "--path",
        default=os.getenv("DATA_PATH", "."),
        help="Ruta donde se encuentran los archivos",
    )
    args = parser.parse_args()
    
    setup_logging()
    try:
        success = process_wells_fargo_files(args.path)
    except Exception:
        logger.exception("Error no controlado durante la ejecución del agente.")
        send_notification_email(
            f"[ALERTA] {AGENT_NAME} - error no controlado",
            f"El agente terminó de forma inesperada y el proceso fue ANULADO.\n\n"
            f"Fecha/hora: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            f"Carpeta de datos: {os.path.abspath(args.path)}\n\n"
            f"Detalle técnico:\n{traceback.format_exc()}\n"
        )
        sys.exit(1)
    
    sys.exit(0 if success else 1)

if __name__ == "__main__":
    main()
