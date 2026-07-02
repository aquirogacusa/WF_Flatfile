import os
import zipfile
import pandas as pd
from datetime import datetime
import argparse
import shutil
import time
import subprocess
import win32com.client
from dotenv import load_dotenv

# Cargar variables de entorno
load_dotenv()

def ensure_sap_running():
    try:
        win32com.client.GetObject("SAPGUI")
        return True
    except:
        pass
        
    print("SAP Logon no está abierto. Intentando iniciarlo automáticamente...")
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
        print("Error: No se pudo encontrar el ejecutable saplogon.exe en las rutas comunes.")
        return False
        
    try:
        subprocess.Popen([sap_path])
        print("Abriendo SAP Logon. Esperando a que el sistema inicialice...")
        time.sleep(10) # Esperar a que SAP inicie y registre los objetos COM
        return True
    except Exception as e:
        print(f"Error al intentar abrir SAP Logon: {e}")
        return False

def extract_sap_report(base_path):
    print("Iniciando extracción de reporte desde SAP...")
    
    # Credenciales y config (idealmente en .env)
    sap_user = os.getenv("SAP_USER")
    sap_password = os.getenv("SAP_PASSWORD")
    sap_connection = os.getenv("SAP_CONNECTION")
    sap_client = os.getenv("SAP_CLIENT", "300")
    sap_lang = os.getenv("SAP_LANGUAGE", "EN")
    
    try:
        # Verificar y abrir SAP si es necesario
        if not ensure_sap_running():
            return False
            
        print("Conectando a SAP GUI...")
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        
        # Iniciar una nueva conexión
        print(f"Abriendo conexión: {sap_connection}...")
        connection = application.OpenConnection(sap_connection, True)
        session = connection.Children(0)
        
        # Hacer Login
        print("Realizando login...")
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
        print("Ejecutando transacción ZSD_POS_1052...")
        session.StartTransaction("ZSD_POS_1052")
        time.sleep(1)
        
        # Cargar Variante CUSA-WF
        print("Cargando variante CUSA-WF...")
        session.findById("wnd[0]").sendVKey(17) # Shift+F5 (Get Variant)
        time.sleep(1)
        session.findById("wnd[1]/usr/txtV-LOW").text = "CUSA-WF"
        session.findById("wnd[1]/usr/txtENAME-LOW").text = "" # Limpiar usuario para buscar globalmente
        session.findById("wnd[1]/tbar[0]/btn[8]").press() # Execute (variante)
        time.sleep(1)
        
        # Ejecutar reporte
        print("Ejecutando reporte...")
        session.findById("wnd[0]/tbar[1]/btn[8]").press() # Ejecutar F8
        time.sleep(5) # Esperar a que cargue el reporte
        
        # Exportar a Archivo Local (TXT Spreadsheet)
        print("Exportando reporte a archivo de texto...")
        session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").select() # Local file...
        time.sleep(1)
        
        # Seleccionar formato Spreadsheet (texto tabulado)
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
        session.findById("wnd[1]/tbar[0]/btn[0]").press() # Continuar
        time.sleep(1)
        
        output_file = os.path.abspath(os.path.join(base_path, "Customers_SAP.txt"))
        out_dir = os.path.dirname(output_file)
        out_name = os.path.basename(output_file)
        
        print(f"Guardando en: {output_file}")
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = out_dir
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = out_name
        session.findById("wnd[1]/tbar[0]/btn[11]").press() # Replace (Sobrescribir si existe)
        
        time.sleep(3)
        
        # Hacer logout
        print("Cerrando sesión en SAP...")
        session.findById("wnd[0]/tbar[0]/btn[3]").press() # Back (F3)
        session.findById("wnd[0]/tbar[0]/btn[3]").press() # Back (F3) again to home
        session.findById("wnd[0]/tbar[0]/btn[15]").press() # Exit (Shift+F3)
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press() # Yes to log off
        
        # Cerrar completamente SAP Logon
        print("Cerrando aplicación SAP Logon...")
        subprocess.run(["taskkill", "/F", "/IM", "saplogon.exe"], capture_output=True)
        
        print("Extracción completada.")
        return True

    except Exception as e:
        print(f"Error en automatización SAP: {str(e)}")
        return False

def process_wells_fargo_files(base_path):
    # 1. Ejecutar extracción de SAP
    extract_sap_report(base_path)
    
    input_zip_filename = "Wells Fargo Bill File Original.zip"
    output_zip_filename = "Wells Fargo Bill File.zip"
    txt_filename = "Customers_SAP.txt"
    csv_filename = "FlatFile.csv"
    
    input_zip_path = os.path.join(base_path, input_zip_filename)
    output_zip_path = os.path.join(base_path, output_zip_filename)
    txt_path = os.path.join(base_path, txt_filename)
    temp_csv_path = os.path.join(base_path, csv_filename)
    
    # 2. Descomprimir el archivo zip
    print(f"Descomprimiendo {input_zip_path}...")
    if not os.path.exists(input_zip_path):
        print(f"Error: No se encontró el archivo {input_zip_path}")
        return

    with zipfile.ZipFile(input_zip_path, 'r') as zip_ref:
        zip_ref.extract(csv_filename, path=base_path)
    
    # 3. Leer archivos
    print(f"Leyendo archivos {csv_filename} y {txt_filename}...")
    if not os.path.exists(txt_path):
        print(f"Error: No se encontró el archivo {txt_path}. Verifica que SAP exportó correctamente.")
        return

    # Leer CSV con pandas, forzando tipos para la columna Auth
    df_csv = pd.read_csv(temp_csv_path, dtype={"Auth": str})
    
    # Leer TXT exportado de SAP (Formato Spreadsheet es UTF-16, separado por tabs)
    # Ignoramos las primeras 20 líneas de metadata del reporte ALV
    try:
        df_sap_raw = pd.read_csv(txt_path, sep='\t', encoding='utf-16', skiprows=20, header=None, on_bad_lines='skip')
    except Exception as e:
        print(f"Error leyendo el archivo TXT: {e}")
        return
        
    # 4. Procesar y filtrar datos de SAP
    # Columnas esperadas en el TXT: 1: Customer, 2: Name 1, 4: Terms of payment
    df_sap = df_sap_raw[[1, 2, 4]].copy()
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
        print("No se encontraron clientes faltantes en SAP.")
    else:
        print(f"Insertando {len(missing_customers)} clientes faltantes...")
        
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
                "Bank Account": "4942472523"
            })
        
        df_new = pd.DataFrame(new_rows)
        df_updated = pd.concat([df_csv, df_new], ignore_index=True)
        
        # Guardar CSV actualizado
        df_updated.to_csv(temp_csv_path, index=False, quoting=1) # quoting=1 (QUOTE_ALL)
        print("CSV actualizado guardado.")

    # 7. Repackage y Cleanup
    print("Repackaging ZIP y limpiando...")
    
    # Eliminar el zip original
    os.remove(input_zip_path)
    
    # Crear nuevo zip con el CSV actualizado
    print(f"Creando {output_zip_path}...")
    with zipfile.ZipFile(output_zip_path, 'w', zipfile.ZIP_DEFLATED) as zip_new:
        zip_new.write(temp_csv_path, arcname=csv_filename)
    
    # Eliminar el CSV extraído
    os.remove(temp_csv_path)
    
    print("Proceso completado exitosamente.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Wells Fargo Customer Agent")
    parser.add_argument("--path", default=".", help="Ruta donde se encuentran los archivos")
    args = parser.parse_args()
    
    process_wells_fargo_files(args.path)
