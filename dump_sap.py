import os
import time
import subprocess
import win32com.client
from dotenv import load_dotenv

load_dotenv()

def dump_ui():
    common_paths = [
        r"C:\Program Files (x86)\SAP\FrontEnd\SAPGUI\saplogon.exe",
        r"C:\Program Files\SAP\FrontEnd\SAPGUI\saplogon.exe"
    ]
    sap_path = next((p for p in common_paths if os.path.exists(p)), None)
    subprocess.Popen([sap_path])
    time.sleep(10)
    
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine
    connection = application.OpenConnection(os.getenv("SAP_CONNECTION"), True)
    session = connection.Children(0)
    
    session.findById("wnd[0]/usr/txtRSYST-MANDT").text = os.getenv("SAP_CLIENT", "300")
    session.findById("wnd[0]/usr/txtRSYST-BNAME").text = os.getenv("SAP_USER")
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = os.getenv("SAP_PASSWORD")
    session.findById("wnd[0]/usr/txtRSYST-LANGU").text = os.getenv("SAP_LANGUAGE", "EN")
    session.findById("wnd[0]").sendVKey(0)
    
    try:
        if session.findById("wnd[1]"):
            session.findById("wnd[1]").sendVKey(0)
    except:
        pass
        
    session.StartTransaction("ZSD_POS_1052")
    time.sleep(1)
    session.findById("wnd[0]").sendVKey(17)
    time.sleep(1)
    session.findById("wnd[1]/usr/txtV-LOW").text = "CUSA-WF"
    session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    time.sleep(1)
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    time.sleep(3)
    
    print("\nSeleccionando Local file...")
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").select()
    time.sleep(2)
    
    print("\n=== POPUP LOCAL FILE ===")
    try:
        for item in session.findById("wnd[1]/usr").Children:
            print(item.ID, getattr(item, "text", ""), getattr(item, "type", ""))
    except Exception as e:
        print("Error reading wnd[1]:", e)

    session.findById("wnd[0]/tbar[0]/btn[15]").press()
    try:
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
    except:
        pass

if __name__ == "__main__":
    dump_ui()
