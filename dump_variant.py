import os
import time
import win32com.client
from dotenv import load_dotenv

load_dotenv()

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "variant_screen_dump.txt")

INPUT_TYPES = ("GuiCTextField", "GuiTextField", "GuiCheckBox", "GuiComboBox", "GuiRadioButton")

def dump_component(component, lines, depth=0):
    try:
        children = component.Children
    except Exception:
        return
    if not children:
        return
    for child in children:
        try:
            cid = child.Id
            ctype = child.Type
            text = getattr(child, "Text", "")
        except Exception:
            continue
        lines.append(f"{'  ' * depth}{cid} | {ctype} | {text!r}")
        dump_component(child, lines, depth + 1)

def find_partner_candidates(component, lines, depth=0):
    try:
        children = component.Children
    except Exception:
        return
    if not children:
        return
    children = list(children)
    for i, child in enumerate(children):
        try:
            ctype = child.Type
            text = getattr(child, "Text", "") or ""
            cid = child.Id
        except Exception:
            continue
        if ctype == "GuiLabel" and "partner" in text.lower():
            lines.append(f"LABEL: {cid} | {text!r}")
            for nxt in children[i + 1:i + 5]:
                try:
                    lines.append(f"   SIGUIENTE: {nxt.Id} | {nxt.Type} | {getattr(nxt, 'Text', '')!r}")
                except Exception:
                    pass
        find_partner_candidates(child, lines, depth + 1)

def main():
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine
    connection = application.OpenConnection(os.getenv("SAP_CONNECTION"), True)
    session = connection.Children(0)

    lines = []
    try:
        session.findById("wnd[0]/usr/txtRSYST-MANDT").text = os.getenv("SAP_CLIENT", "300")
        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = os.getenv("SAP_USER")
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = os.getenv("SAP_PASSWORD")
        session.findById("wnd[0]/usr/txtRSYST-LANGU").text = os.getenv("SAP_LANGUAGE", "EN")
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(2)

        for intento in range(3):
            try:
                popup = session.findById("wnd[1]")
                lines.append(f"[LOGIN POPUP {intento}] {popup.Text!r}")
                popup.sendVKey(0)
                time.sleep(1)
            except Exception:
                break

        session.StartTransaction("ZSD_POS_1052")
        time.sleep(3)

        lines.append(f"=== PANTALLA TRAS StartTransaction: wnd[0] t={session.findById('wnd[0]').Text!r}")
        dump_component(session.findById("wnd[0]"), lines)
        try:
            popup = session.findById("wnd[1]")
            lines.append(f"=== POPUP wnd[1] t={popup.Text!r}")
            dump_component(popup, lines)
        except Exception:
            lines.append("=== No hay wnd[1] tras StartTransaction")

        try:
            session.findById("wnd[0]").sendVKey(17)
            time.sleep(2)
            session.findById("wnd[1]/usr/txtV-LOW").text = "CUSA-WF"
            session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
            session.findById("wnd[1]/tbar[0]/btn[8]").press()
            time.sleep(3)
            lines.append("")
            lines.append("=== PANTALLA wnd[0] TRAS CARGAR VARIANTE CUSA-WF ===")
            dump_component(session.findById("wnd[0]"), lines)
        except Exception as e:
            lines.append(f"=== ERROR cargando variante: {e}")

        lines.append("")
        lines.append("=== CANDIDATOS 'partner' ===")
        find_partner_candidates(session.findById("wnd[0]"), lines)
    except Exception as e:
        lines.append(f"=== ERROR GENERAL: {e}")
    finally:
        with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
        print(f"Dump escrito en: {OUTPUT_FILE} ({len(lines)} lineas)")

        try:
            session.findById("wnd[0]/tbar[0]/btn[15]").press()
            time.sleep(1)
        except Exception:
            pass
        try:
            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
        except Exception:
            pass
        print("Sesion cerrada.")

if __name__ == "__main__":
    main()
