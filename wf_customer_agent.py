import os
import re
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
import threading
import subprocess
import ctypes
from ctypes import wintypes
import pythoncom
import win32com.client
import win32gui
import win32con
import win32process
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

def _iter_sap_controls(component, max_depth=12):
    # Recorre recursivamente el arbol de controles de un componente de SAP GUI
    def walk(node, depth):
        if depth > max_depth:
            return
        try:
            children = list(node.Children)
        except Exception:
            return
        for child in children:
            yield child
            yield from walk(child, depth + 1)

    yield from walk(component, 0)

def _select_menu_item_by_text(session, menu_path, text_prefix):
    # Selecciona un item de menu por su texto para no depender de los indices (separadores)
    try:
        for item in session.findById(menu_path).Children:
            try:
                text = (item.Text or "").strip()
            except Exception:
                continue
            if text.lower().startswith(text_prefix.lower()):
                item.select()
                return True
    except Exception:
        pass
    return False

def _select_spreadsheet_format(session):
    # Si SAP muestra el dialogo "Select Spreadsheet", selecciona la opcion de Excel (preferir XXL).
    # Si el dialogo no aparece (formato recordado por el usuario), continua directo al guardado.
    try:
        popup = session.findById("wnd[1]")
    except Exception:
        return

    try:
        session.findById("wnd[1]/usr/ctxtDY_PATH")
        return  # Ya es el dialogo de guardado de archivo
    except Exception:
        pass

    excel_radio = None
    excel_xxl_radio = None
    for control in _iter_sap_controls(popup):
        try:
            if control.Type != "GuiRadioButton":
                continue
            text = (control.Text or "").strip()
        except Exception:
            continue
        lower = text.lower()
        if "excel" not in lower:
            continue
        if "xxl" in lower:
            excel_xxl_radio = control
            break
        if excel_radio is None:
            excel_radio = control

    chosen = excel_xxl_radio or excel_radio
    if chosen is not None:
        chosen.select()
        logger.info(f"Formato de exportación seleccionado: {chosen.Text}")
    else:
        logger.warning("No se identificó el formato Excel en el diálogo 'Select Spreadsheet'; se usará el formato por defecto.")

    session.findById("wnd[1]/tbar[0]/btn[0]").press() # Continuar
    time.sleep(1)

def _wait_for_save_dialog(session, timeout=20):
    # Espera el dialogo estandar de guardado de archivo, cerrando popups intermedios
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            session.findById("wnd[1]/usr/ctxtDY_PATH")
            return True
        except Exception:
            pass
        try:
            if session.findById("wnd[1]"):
                session.findById("wnd[1]").sendVKey(0)
        except Exception:
            pass
        time.sleep(0.5)
    return False

def _close_unexpected_popup(session):
    # Cierra popups (informativos o de confirmacion) que SAP pueda mostrar durante la exportacion.
    # Si hay un dialogo nativo de Windows abierto, las llamadas a SAP se bloquean: no tocar la sesion.
    if _find_windows_save_dialog(timeout=0):
        return
    try:
        popup = session.findById("wnd[1]")
    except Exception:
        return
    try:
        text = (popup.Text or "").strip()
    except Exception:
        text = ""
    logger.warning(f"Cerrando diálogo de SAP durante la exportación: {text!r}")
    for button_id in ("wnd[1]/usr/btnSPOP-OPTION1", "wnd[1]/tbar[0]/btn[0]"):
        try:
            session.findById(button_id).press()
            return
        except Exception:
            continue
    try:
        popup.sendVKey(0)
    except Exception:
        pass

def _wait_for_exported_file(session, output_file, since, timeout):
    # Espera a que SAP genere el archivo Excel (mtime posterior al inicio de la exportacion)
    base, ext = os.path.splitext(output_file)
    candidates = [output_file]
    for alt_ext in (".xlsx", ".xls"):
        for candidate in (base + alt_ext, output_file + alt_ext):
            # base + .xls cubre también .XLS (Windows no distingue mayúsculas);
            # output_file + alt_ext cubre el caso en que SAP duplique la extensión
            if candidate not in candidates:
                candidates.append(candidate)

    def new_file_size(candidate):
        try:
            if (
                os.path.exists(candidate)
                and os.path.getsize(candidate) > 0
                and os.path.getmtime(candidate) + 2 >= since
            ):
                return os.path.getsize(candidate)
        except OSError:
            pass
        return None

    deadline = time.time() + timeout
    last_sizes = {}
    while time.time() < deadline:
        for candidate in candidates:
            size = new_file_size(candidate)
            if size is None:
                continue
            if last_sizes.get(candidate) == size:
                return candidate
            last_sizes[candidate] = size
        _close_unexpected_popup(session)
        time.sleep(2)

    for candidate in candidates:
        if new_file_size(candidate) is not None:
            return candidate
    return None

def _sap_process_ids():
    # PIDs de los procesos del SAP GUI: el dialogo "Guardar como" de Windows pertenece a uno de ellos
    pids = set()

    def enum_callback(hwnd, _):
        try:
            class_name = win32gui.GetClassName(hwnd) or ""
            if class_name.startswith("SAP_FRONTEND") or class_name == "SapGuiShell":
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                pids.add(pid)
        except Exception:
            pass
        return True

    try:
        win32gui.EnumWindows(enum_callback, None)
    except Exception:
        pass
    return pids

def _find_child_windows(parent, class_name):
    matches = []

    def enum_callback(hwnd, _):
        try:
            if (win32gui.GetClassName(hwnd) or "") == class_name:
                matches.append(hwnd)
        except Exception:
            pass
        return True

    try:
        win32gui.EnumChildWindows(parent, enum_callback, None)
    except Exception:
        pass
    return matches

def _get_dlg_item(hwnd, control_id):
    # GetDlgItem lanza excepcion si el control no existe
    try:
        return win32gui.GetDlgItem(hwnd, control_id)
    except Exception:
        return 0

def _find_button_by_text(hwnd_dlg, texts):
    # Busca un boton por el texto (los dialogos modernos de Windows no usan IDs clasicos)
    for button in _find_child_windows(hwnd_dlg, "Button"):
        try:
            text = (win32gui.GetWindowText(button) or "").strip().lower().replace("&", "")
        except Exception:
            continue
        if text in texts:
            return button
    return 0

class _GuiThreadInfo(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.DWORD),
        ("flags", wintypes.DWORD),
        ("hwndActive", wintypes.HWND),
        ("hwndFocus", wintypes.HWND),
        ("hwndCapture", wintypes.HWND),
        ("hwndMenuOwner", wintypes.HWND),
        ("hwndMoveSize", wintypes.HWND),
        ("hwndCaret", wintypes.HWND),
        ("rcCaret", wintypes.RECT),
    ]

def _get_focused_hwnd(hwnd):
    # Control con el foco del hilo que posee la ventana
    try:
        tid, _ = win32process.GetWindowThreadProcessId(hwnd)
        info = _GuiThreadInfo()
        info.cbSize = ctypes.sizeof(_GuiThreadInfo)
        if ctypes.windll.user32.GetGUIThreadInfo(tid, ctypes.byref(info)):
            return info.hwndFocus or 0
    except Exception:
        pass
    return 0

def _find_filename_edit(hwnd_dlg):
    # Campo "Nombre de archivo" del dialogo de guardado
    combo = _get_dlg_item(hwnd_dlg, 1148)
    if combo:
        if (win32gui.GetClassName(combo) or "") == "Edit":
            return combo
        edits = _find_child_windows(combo, "Edit")
        if edits:
            return edits[0]
    for edit in _find_child_windows(hwnd_dlg, "Edit"):
        if win32gui.IsWindowVisible(edit) and (win32gui.GetWindowText(edit) or "").strip():
            return edit
    edits = [e for e in _find_child_windows(hwnd_dlg, "Edit") if win32gui.IsWindowVisible(e)]
    return edits[0] if edits else None

def _is_windows_save_dialog(hwnd):
    # Debe ser el dialogo de archivos: se valida por titulo o por la estructura
    # del dialogo moderno (DUIViewWndClassName), para no confundirlo con otras
    # ventanas #32770 del proceso SAP (p.ej. "SAP Logon 800")
    try:
        title = (win32gui.GetWindowText(hwnd) or "").strip().lower()
        if "save as" in title or "guardar como" in title:
            return True
        return bool(_find_child_windows(hwnd, "DUIViewWndClassName"))
    except Exception:
        return False

def _find_windows_save_dialog(timeout=60):
    # Busca el dialogo nativo de Windows "Guardar como" que SAP abre para el export Spreadsheet
    deadline = time.time() + timeout
    while time.time() < deadline:
        sap_pids = _sap_process_ids()
        candidates = []

        def enum_callback(hwnd, _):
            try:
                if not win32gui.IsWindowVisible(hwnd):
                    return True
                if (win32gui.GetClassName(hwnd) or "") != "#32770":
                    return True
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                if pid not in sap_pids:
                    return True
                if not _is_windows_save_dialog(hwnd):
                    return True
                candidates.append(hwnd)
            except Exception:
                pass
            return True

        try:
            win32gui.EnumWindows(enum_callback, None)
        except Exception:
            pass
        if candidates:
            return candidates[0]
        time.sleep(0.5)
    return None

def _has_confirm_dialog(sap_pids):
    # Indica si hay una confirmacion de sobrescritura pendiente
    found = []

    def enum_callback(hwnd, _):
        try:
            if not win32gui.IsWindowVisible(hwnd):
                return True
            if (win32gui.GetClassName(hwnd) or "") != "#32770":
                return True
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            if pid not in sap_pids:
                return True
            title = (win32gui.GetWindowText(hwnd) or "").strip().lower()
            if "confirm" in title or _get_dlg_item(hwnd, 6):
                found.append(hwnd)
        except Exception:
            pass
        return True

    try:
        win32gui.EnumWindows(enum_callback, None)
    except Exception:
        pass
    return found[0] if found else None

def _press_save_in_windows_dialog(hwnd_dlg, edit):
    # Intenta presionar "Guardar" por varias vias. Siempre con PostMessage (asincrono):
    # un click modal puede abrir la confirmacion de sobrescritura y un SendMessage
    # sincrono quedaria bloqueado esperando a que esa confirmacion se cierre.
    attempts = []
    save_button = _find_button_by_text(hwnd_dlg, ("save", "guardar")) or _get_dlg_item(hwnd_dlg, 1)
    if save_button:
        attempts.append(("click", save_button))
    if edit:
        attempts.append(("enter", edit))
    attempts.append(("idok", hwnd_dlg))

    sap_pids = _sap_process_ids()
    for kind, target in attempts:
        if not win32gui.IsWindow(hwnd_dlg):
            return True
        if _has_confirm_dialog(sap_pids):
            return True  # Ya apareció la confirmación: la maneja _confirm_windows_overwrite
        try:
            if kind == "click":
                win32gui.PostMessage(target, win32con.BM_CLICK, 0, 0)
            elif kind == "enter":
                win32gui.PostMessage(target, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
                win32gui.PostMessage(target, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
            else:
                win32gui.PostMessage(target, win32con.WM_COMMAND, 1, 0)  # IDOK (boton por defecto)
        except Exception:
            continue
        time.sleep(0.8)

    if win32gui.IsWindow(hwnd_dlg) and not _has_confirm_dialog(sap_pids):
        try:
            win32gui.SetForegroundWindow(hwnd_dlg)
        except Exception:
            pass
        win32com.client.Dispatch("WScript.Shell").SendKeys("{ENTER}")
    return True

def _uia_edit_text(edit):
    try:
        return edit.get_value() or ""
    except Exception:
        pass
    try:
        return edit.window_text() or ""
    except Exception:
        return ""

def _fill_windows_save_dialog_uia(hwnd_dlg, output_file):
    # Los dialogos modernos de Windows 10/11 (DirectUI) no aceptan WM_SETTEXT/GetWindowText
    # cross-proceso, por lo que se usa UI Automation (pywinauto) para escribir la ruta.
    try:
        from pywinauto import Application
    except ImportError:
        logger.warning("pywinauto no está instalado; se usará el método clásico para el diálogo de guardado.")
        return False

    try:
        app = Application(backend="uia").connect(handle=hwnd_dlg)
        dlg = app.window(handle=hwnd_dlg)

        edits = []
        for control in dlg.descendants(control_type="Edit"):
            try:
                edits.append(control)
            except Exception:
                continue
        if not edits:
            logger.warning("UIA: no se encontró el campo de nombre de archivo.")
            return False

        edit = None
        for control in edits:
            try:
                if control.element_info.automation_id == "1001":
                    edit = control
                    break
            except Exception:
                continue
        if edit is None:
            for control in edits:
                try:
                    name = (control.element_info.name or "").lower()
                    if name.startswith("file name") or name.startswith("nombre"):
                        edit = control
                        break
                except Exception:
                    continue
        if edit is None:
            edit = edits[0]

        text_ok = False
        try:
            edit.set_edit_text(output_file)
            time.sleep(0.4)
            text_ok = output_file.lower() in _uia_edit_text(edit).lower()
        except Exception:
            text_ok = False

        if not text_ok:
            edit.set_focus()
            time.sleep(0.2)
            dlg.type_keys("^a", pause=0.05)
            dlg.type_keys(output_file, with_spaces=True, pause=0.02)
            time.sleep(0.4)
            text_ok = output_file.lower() in _uia_edit_text(edit).lower()

        if not text_ok:
            logger.warning("UIA: no se pudo escribir ni verificar la ruta en el diálogo de guardado.")
            return False
        logger.info("Ruta escrita en el diálogo de guardado (UIA).")

        save_button = None
        for control in dlg.descendants(control_type="Button"):
            try:
                if control.window_text().strip().lower() in ("save", "guardar"):
                    save_button = control
                    break
            except Exception:
                continue

        if save_button is not None:
            save_button.invoke()
        else:
            dlg.type_keys("{ENTER}")

        # Verificación: el diálogo debe cerrarse (o quedar solo la confirmación de sobrescritura)
        time.sleep(2)
        if win32gui.IsWindow(hwnd_dlg) and not _has_confirm_dialog(_sap_process_ids()):
            logger.warning("El diálogo de guardado sigue abierto; reintentando con Enter.")
            try:
                dlg.type_keys("{ENTER}")
            except Exception:
                pass
            time.sleep(1)
        logger.info("Guardado confirmado en el diálogo de Windows (UIA).")
        return True
    except Exception as e:
        logger.warning(f"UI Automation falló en el diálogo de guardado: {e}")
        return False

def _fill_windows_save_dialog(hwnd_dlg, output_file):
    # Escribe la ruta completa en el campo de nombre y confirma el guardado.
    # Primero UI Automation (funciona con los diálogos modernos tipo DirectUI);
    # si no está disponible, se usa el método clásico con SendMessage/SendKeys.
    if _fill_windows_save_dialog_uia(hwnd_dlg, output_file):
        return

    edit = _find_filename_edit(hwnd_dlg)
    written = False
    if edit:
        try:
            win32gui.SendMessage(edit, win32con.WM_SETTEXT, 0, output_file)
            written = output_file.lower() in (win32gui.GetWindowText(edit) or "").lower()
        except Exception:
            written = False

    if written:
        logger.info("Ruta escrita en el diálogo de guardado de Windows.")
        _press_save_in_windows_dialog(hwnd_dlg, edit)
        return

    # Fallback: el campo de nombre tiene el foco al abrir el dialogo, se escribe con SendKeys
    logger.warning("No se pudo escribir la ruta directamente; se usará SendKeys.")
    try:
        win32gui.SetForegroundWindow(hwnd_dlg)
    except Exception:
        pass
    shell = win32com.client.Dispatch("WScript.Shell")
    time.sleep(0.3)
    shell.SendKeys("^a")
    shell.SendKeys(output_file)  # la ruta no contiene caracteres especiales de SendKeys
    time.sleep(0.5)
    shell.SendKeys("{ENTER}")

def _press_enter_in_dialog(hwnd_dlg):
    # Envia Enter al control con el foco del dialogo (el aviso interno de sobrescritura
    # de Windows 11 tiene el boton "Si" como accion por defecto)
    target = _get_focused_hwnd(hwnd_dlg) or hwnd_dlg
    try:
        win32gui.PostMessage(target, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
        win32gui.PostMessage(target, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
    except Exception:
        pass

def _confirm_windows_overwrite(hwnd_save_dialog, timeout=15):
    # Si el archivo ya existe, Windows pide confirmar la sobrescritura: puede ser una
    # ventana separada ("Confirmar guardar como") o un aviso dentro del mismo dialogo.
    # Devuelve True cuando ya no queda ninguna confirmacion pendiente.
    deadline = time.time() + timeout
    attempts = 0
    while time.time() < deadline:
        if not win32gui.IsWindow(hwnd_save_dialog):
            return True  # El dialogo se cerro: ya no hay confirmacion pendiente

        sap_pids = _sap_process_ids()
        confirm_hwnd = _has_confirm_dialog(sap_pids)
        if confirm_hwnd:
            yes_button = _find_button_by_text(
                confirm_hwnd, ("yes", "si", "sí", "aceptar", "continuar", "ok")
            ) or _get_dlg_item(confirm_hwnd, 6)  # Confirmacion clasica
            if yes_button:
                win32gui.PostMessage(yes_button, win32con.BM_CLICK, 0, 0)
            else:
                _press_enter_in_dialog(confirm_hwnd)
            logger.info("Confirmada la sobrescritura del archivo existente.")
            time.sleep(0.6)
            continue

        # Aviso interno dentro del dialogo: Enter equivale a "Si"
        attempts += 1
        if attempts in (3, 6, 9, 12):
            _press_enter_in_dialog(hwnd_save_dialog)
        time.sleep(0.5)
    return False

def _move_previous_output(output_file, previous_file):
    # Aparta el archivo anterior (respaldo temporal) para que Windows no pida
    # confirmar la sobrescritura al exportar
    try:
        if os.path.exists(previous_file):
            os.remove(previous_file)
        if os.path.exists(output_file):
            os.replace(output_file, previous_file)
            logger.info(f"Archivo anterior apartado temporalmente: {previous_file}")
            return True
    except Exception as e:
        logger.warning(f"No se pudo apartar el archivo anterior ({e}); se continuará sin respaldo.")
    return False

def _close_exported_excel_file(file_path):
    # SAP abre Excel automáticamente con el archivo exportado; se cierra para que no
    # quede bloqueado (otro proceso lo lee y la siguiente ejecución lo reemplaza).
    try:
        excel = win32com.client.GetObject(Class="Excel.Application")
    except Exception:
        return
    try:
        excel.DisplayAlerts = False
    except Exception:
        pass
    target = os.path.basename(file_path).lower()
    closed_any = False
    try:
        for index in range(excel.Workbooks.Count, 0, -1):
            try:
                book = excel.Workbooks(index)
                if os.path.basename(book.FullName).lower() == target:
                    book.Close(SaveChanges=False)
                    closed_any = True
            except Exception:
                continue
    except Exception as e:
        logger.warning(f"No se pudo cerrar el Excel del archivo exportado: {e}")
        return
    if closed_any:
        logger.info("Se cerró el Excel abierto automáticamente por SAP.")
    try:
        if excel.Workbooks.Count == 0:
            excel.Quit()
    except Exception:
        pass

def _close_app_windows_for_file(file_path, timeout=12):
    # SAP abre automáticamente el archivo exportado con la aplicación asociada. En los
    # equipos sin Excel se abre con otra app (p. ej. Notepad) o aparece el diálogo
    # "Elegir una aplicación": se cierran esas ventanas para no dejar el archivo bloqueado.
    name = os.path.basename(file_path).lower()
    dialogo_claves = (
        "select an app",
        "abrir con",
        "open with",
        "how do you want to open",
        "cómo quieres abrir",
        "selecciona una aplicaci",
    )
    deadline = time.time() + timeout
    closed_titles = []
    while time.time() < deadline:
        matches = []

        def enum_callback(hwnd, _):
            try:
                if not win32gui.IsWindowVisible(hwnd):
                    return True
                title = (win32gui.GetWindowText(hwnd) or "").strip()
                lower = title.lower()
                if name and name in lower:
                    matches.append((hwnd, title))
                elif ".xlsx" in lower and any(key in lower for key in dialogo_claves):
                    matches.append((hwnd, title))
            except Exception:
                pass
            return True

        try:
            win32gui.EnumWindows(enum_callback, None)
        except Exception:
            pass

        for hwnd, title in matches:
            try:
                win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                closed_titles.append(title)
            except Exception:
                pass

        if closed_titles and not matches:
            break
        time.sleep(0.5)

    if closed_titles:
        logger.info(f"Se cerró la ventana que abrió el archivo exportado: {closed_titles[0]!r}")

def _restore_previous_output(previous_file, output_file):
    # Si la exportación falló, se restaura el archivo anterior para no perder el respaldo
    try:
        if os.path.exists(previous_file) and not os.path.exists(output_file):
            os.replace(previous_file, output_file)
            logger.warning("La exportación no generó archivo nuevo; se restauró el archivo anterior.")
    except Exception as e:
        logger.error(f"No se pudo restaurar el archivo anterior {previous_file}: {e}")

def _refresh_sap_session(session=None):
    # El export a Excel invalida la referencia COM de la sesión de SAP Scripting
    # ("Object is not connected to server") aunque la sesión siga viva en SAP GUI:
    # se re-adquiere una referencia nueva para continuar trabajando.
    session_name = None
    try:
        session_name = session.Name
    except Exception:
        pass

    application = win32com.client.GetObject("SAPGUI").GetScriptingEngine
    target_connection = os.getenv("SAP_CONNECTION")
    fallback = None
    preferred = []
    for connection in application.Children:
        try:
            description = connection.Description
        except Exception:
            description = None
        try:
            sessions = list(connection.Children)
        except Exception:
            continue
        for candidate in sessions:
            if fallback is None:
                fallback = candidate
            if target_connection and description == target_connection:
                preferred.append(candidate)
                try:
                    if session_name and candidate.Name == session_name:
                        return candidate
                except Exception:
                    continue
    if preferred:
        return preferred[-1]  # la sesión más reciente de la conexión del agente
    if fallback is not None:
        return fallback
    raise RuntimeError("No hay ninguna sesión de SAP disponible.")

def _select_spreadsheet_menu(session):
    # Selecciona List > Export > Spreadsheet en un hilo aparte: en SAP GUI 8 el clic
    # puede quedarse bloqueado hasta que el diálogo modal de Windows se cierre.
    # Cada hilo que usa COM debe inicializar su propio apartamento (CoInitialize).
    try:
        pythoncom.CoInitialize()
    except Exception:
        pass
    try:
        try:
            _select_spreadsheet_menu_run(session)
        except Exception:
            # Reintento con una referencia COM creada en este mismo hilo
            sapgui = win32com.client.GetObject("SAPGUI")
            sess = sapgui.GetScriptingEngine.Children(0).Children(0)
            _select_spreadsheet_menu_run(sess)
    except Exception as e:
        logger.warning(f"No se pudo seleccionar el menú de exportación a Excel: {e}")
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

def _select_spreadsheet_menu_run(session):
    if not _select_menu_item_by_text(session, "wnd[0]/mbar/menu[0]/menu[3]", "Spreadsheet"):
        session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()

def _run_with_timeout(fn, timeout):
    # Ejecuta fn() en un hilo aparte y espera hasta timeout segundos.
    # Devuelve el resultado de fn(), o "timeout" si sigue bloqueada
    # (p.ej. una llamada a SAP GUI Scripting bloqueada por un diálogo modal).
    result = {}

    def runner():
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass
        try:
            result["value"] = fn()
        except Exception as e:
            result["error"] = e
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    thread = threading.Thread(target=runner, daemon=True)
    thread.start()
    thread.join(timeout)
    if thread.is_alive():
        return "timeout"
    if "error" in result:
        raise result["error"]
    return result.get("value")

def _export_via_sap_dialogs(session, out_dir, out_name):
    # Flujo para SAP GUI clásico: diálogos de SAP ("Select Spreadsheet" y archivo).
    # Devuelve True si el guardado se inició correctamente.
    _select_spreadsheet_format(session)
    if not _wait_for_save_dialog(session, timeout=10):
        return False
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = out_dir
    logger.info(f"Guardando en (diálogo SAP): {os.path.join(out_dir, out_name)}")
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = out_name
    try:
        session.findById("wnd[1]/usr/chkSCR-EXEC").selected = False  # No abrir Excel al finalizar
    except Exception:
        pass
    session.findById("wnd[1]/tbar[0]/btn[11]").press() # Replace (Sobrescribir si existe)
    return True

def export_sap_report_to_file(session, output_file, timeout=180):
    # Exporta la lista ALV actual a un archivo Excel real usando List > Export > Spreadsheet.
    # OJO: mientras hay un diálogo modal abierto (el "Guardar como" nativo de Windows),
    # las llamadas a SAP GUI Scripting quedan bloqueadas; por eso no se debe llamar a
    # session.findById() hasta que el diálogo se resuelva.
    output_file = os.path.abspath(output_file)
    out_dir = os.path.dirname(output_file)
    out_name = os.path.basename(output_file)

    logger.info(f"Exportando reporte a Excel (Spreadsheet): {output_file}...")

    # List > Export > Spreadsheet... Se lanza en segundo plano porque la llamada a
    # SAP puede quedar bloqueada hasta que el diálogo modal de Windows se cierre.
    threading.Thread(target=_select_spreadsheet_menu, args=(session,), daemon=True).start()
    time.sleep(2)

    previous_file = output_file + ".previo"
    moved_previous = False
    export_started = None

    # 1) Esperar el diálogo nativo de Windows usando solo Win32 (sin tocar SAP).
    #    La exportación puede tardar varios minutos en prepararse antes de mostrarlo.
    hwnd_dlg = _find_windows_save_dialog(timeout=30)

    # 2) Sin diálogo de Windows: probar el flujo clásico de SAP (GUI antiguos).
    #    Se ejecuta en un hilo con timeout por si la sesión está bloqueada.
    if hwnd_dlg is None:
        try:
            classic_done = _run_with_timeout(
                lambda: _export_via_sap_dialogs(session, out_dir, out_name), timeout=25
            )
        except Exception as e:
            logger.warning(f"El flujo clásico de SAP falló: {e}")
            classic_done = None
        if classic_done is True:
            export_started = time.time()
        else:
            # Sesión ocupada preparando la exportación: esperar con más paciencia
            hwnd_dlg = _find_windows_save_dialog(timeout=300)

    # 3) Resolver el diálogo nativo de Windows
    if hwnd_dlg and export_started is None:
        # Se aparta el archivo anterior para que Windows no muestre la confirmación
        # de sobrescritura (el proceso siempre debe sobrescribir)
        moved_previous = _move_previous_output(output_file, previous_file)

        logger.info(f"Guardando en (diálogo Windows): {output_file}")
        export_started = time.time()
        _fill_windows_save_dialog(hwnd_dlg, output_file)
        _confirm_windows_overwrite(hwnd_dlg)

    if export_started is None:
        logger.error("No apareció ningún diálogo de guardado (ni SAP ni Windows).")
        return None

    exported = _wait_for_exported_file(session, output_file, export_started, timeout)

    if exported:
        _close_exported_excel_file(exported)
        _close_app_windows_for_file(exported)

    if moved_previous:
        if exported:
            try:
                os.remove(previous_file)
            except Exception:
                pass
        else:
            _restore_previous_output(previous_file, output_file)
    if exported:
        logger.info(f"Reporte exportado: {exported}")
    else:
        logger.error(f"El archivo Excel no se generó en {timeout} segundos: {output_file}")
    return exported

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
        
        # Exportar a Excel (Spreadsheet) -> Customers_SAP.xlsx (proceso WF)
        exported = export_sap_report_to_file(session, os.path.join(base_path, "Customers_SAP.xlsx"))
        if not exported:
            raise RuntimeError("No fue posible exportar el reporte de clientes (Partner Function = ZA) desde SAP.")

        # La referencia COM de la sesión queda desconectada tras el export: re-adquirirla
        session = _refresh_sap_session(session)

        # Segunda ejecución de la transacción con Partner Functions = Z5 -> Customers_SAP_Z5.xlsx
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
            export_sap_report_to_file(session, os.path.join(base_path, "Customers_SAP_Z5.xlsx"))
            session = _refresh_sap_session(session)
            logger.info("Reporte Z5 exportado correctamente.")
        except Exception as e:
            logger.error(f"No se pudo generar el archivo Customers_SAP_Z5.xlsx: {e}")
            send_notification_email(
                f"[ALERTA] {AGENT_NAME} - fallo al extraer reporte Z5 de SAP",
                f"No fue posible extraer el reporte con Partner Functions = Z5 "
                f"(transacción ZSD_POS_1052).\n\n"
                f"Fecha/hora: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
                f"Carpeta de datos: {os.path.abspath(base_path)}\n\n"
                f"El archivo Customers_SAP.xlsx (Partner Functions = ZA) sí fue generado y el "
                f"proceso de Wells Fargo continúa. El proceso que consume Customers_SAP_Z5.xlsx "
                f"se verá afectado. Revise el log del agente para más información.\n"
            )
        
        # Hacer logout (mejor esfuerzo: un fallo aquí no debe invalidar una exportación exitosa)
        logger.info("Cerrando sesión en SAP...")
        try:
            session = _refresh_sap_session(session)
            session.findById("wnd[0]/tbar[0]/btn[3]").press() # Back (F3)
            session.findById("wnd[0]/tbar[0]/btn[3]").press() # Back (F3) again to home
            session.findById("wnd[0]/tbar[0]/btn[15]").press() # Exit (Shift+F3)
            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press() # Yes to log off
        except Exception as e:
            logger.warning(f"No se pudo cerrar la sesión de SAP correctamente: {e}")
        
        # Cerrar completamente SAP Logon
        logger.info("Cerrando aplicación SAP Logon...")
        subprocess.run(["taskkill", "/F", "/IM", "saplogon.exe"], capture_output=True)
        
        logger.info("Extracción completada.")
        return True

    except Exception as e:
        logger.error(f"Error en automatización SAP: {str(e)}")
        # Evitar dejar una sesión de SAP abierta que bloquee la próxima ejecución
        try:
            subprocess.run(["taskkill", "/F", "/IM", "saplogon.exe"], capture_output=True)
        except Exception:
            pass
        return False

def _parse_quoted_concatenated_flatfile(csv_path, encoding):
    # SAP BO puede exportar el FlatFile sin delimitador, con los campos entre comillas:
    #   "Payer Code""Auth""Customer"...
    # Se extraen los valores entre comillas de cada línea.
    rows = []
    with open(csv_path, "r", encoding=encoding, errors="replace", newline="") as f:
        for line in f:
            line = line.rstrip("\r\n")
            if not line:
                continue
            fields = re.findall(r'"([^"]*)"', line)
            if fields:
                rows.append(fields)
    if len(rows) < 2:
        return None
    header = rows[0]
    if "Auth" not in header:
        return None
    width = len(header)
    data = [row for row in rows[1:] if len(row) == width]
    return pd.DataFrame(data, columns=header)

def read_flatfile_csv(csv_path):
    # Lee el FlatFile.csv y devuelve (DataFrame, layout, encoding), o (None, None, None).
    # Soporta el CSV estándar (separado por comas) y el formato de SAP BO sin delimitador.
    encodings = ("utf-8-sig", "utf-8", "cp1252", "latin-1")
    for enc in encodings:
        try:
            df = pd.read_csv(csv_path, dtype={"Auth": str}, encoding=enc)
        except UnicodeDecodeError:
            logger.warning(f"No se pudo leer el FlatFile con codificación {enc}. Probando la siguiente...")
            continue
        except pd.errors.EmptyDataError:
            return pd.DataFrame(), "comma", enc
        except Exception as e:
            logger.warning(f"Error leyendo el FlatFile con codificación {enc}: {e}")
            continue

        if "Auth" in df.columns:
            if enc != "utf-8-sig":
                logger.info(f"FlatFile leído con codificación {enc} (CSV estándar).")
            return df, "comma", enc

        # El CSV estándar no encontró la columna Auth: probar el formato de SAP BO
        df_quoted = _parse_quoted_concatenated_flatfile(csv_path, enc)
        if df_quoted is not None:
            logger.info(f"FlatFile leído con codificación {enc} (formato SAP BO sin delimitador).")
            return df_quoted, "quoted", enc
    return None, None, None

def write_flatfile_csv(df, csv_path, layout, encoding="utf-8"):
    # Guarda el FlatFile preservando el formato del archivo original
    if layout == "quoted":
        lines = ['"' + '""'.join(str(col).replace('"', " ") for col in df.columns) + '"']
        for row in df.itertuples(index=False, name=None):
            fields = []
            for value in row:
                try:
                    empty = pd.isna(value)
                except Exception:
                    empty = False
                text = "" if empty else str(value)
                fields.append(text.replace('"', " "))  # las comillas romperían el formato
            lines.append('"' + '""'.join(fields) + '"')
        with open(csv_path, "w", encoding=encoding, errors="replace", newline="") as f:
            f.write("\r\n".join(lines) + "\r\n")
    else:
        df.to_csv(csv_path, index=False, quoting=1) # quoting=1 (QUOTE_ALL)

def _detect_sap_columns_without_header(raw, required_columns):
    # El export "Spreadsheet" de SAP GUI 8 no incluye la fila de encabezados: solo datos.
    # Se detectan las columnas por su contenido, usando como pista el layout conocido
    # del reporte (Customer=2, Name 1=4, Terms Paym=29).
    sample = raw.iloc[: min(len(raw), 500)]

    def column_values(idx):
        return [str(v).strip() for v in sample.iloc[:, idx].tolist() if pd.notna(v)]

    # Customer: códigos con relleno de SAP ("00" + 8 dígitos)
    customer_col = None
    best_ratio = 0.0
    for idx in range(raw.shape[1]):
        values = column_values(idx)
        if len(values) < 5:
            continue
        hits = sum(1 for v in values if re.fullmatch(r"00\d{8}", v))
        ratio = hits / len(values)
        if ratio > best_ratio:
            best_ratio = ratio
            customer_col = idx
    if customer_col is None or best_ratio < 0.5:
        # Pista de layout: columna 2
        values = [str(v).strip() for v in raw.iloc[:, 2].tolist() if pd.notna(v)] if raw.shape[1] > 2 else []
        hits = sum(1 for v in values if re.fullmatch(r"\d{8,10}", v))
        if values and hits / len(values) >= 0.5:
            customer_col = 2
            best_ratio = hits / len(values)
    if customer_col is None:
        return None

    # El resto de columnas se ubican respecto a Customer según el layout del reporte
    offsets = {"Name 1": 2, "Terms Paym": 27}
    detected = {"Customer": customer_col}
    for name in required_columns:
        if name == "Customer":
            continue
        offset = offsets.get(name)
        idx = customer_col + offset if offset is not None else None
        if idx is not None and 0 <= idx < raw.shape[1]:
            detected[name] = idx

    if "Terms Paym" in required_columns and "Terms Paym" not in detected:
        # Búsqueda de respaldo: columna con valores tipo C030 / C008
        best_ratio = 0.0
        for idx in range(raw.shape[1]):
            values = column_values(idx)
            if len(values) < 5:
                continue
            hits = sum(1 for v in values if re.fullmatch(r"[A-Z]\d{3}", v))
            ratio = hits / len(values)
            if ratio > best_ratio:
                best_ratio = ratio
                detected["Terms Paym"] = idx
        if best_ratio < 0.5:
            detected.pop("Terms Paym", None)

    logger.info(f"Columnas detectadas sin encabezado: {detected} (layout del reporte SAP)")
    return detected

def read_sap_report_excel(excel_path, required_columns):
    # Lee el Excel exportado de SAP (Spreadsheet). Devuelve un DataFrame con las columnas
    # requeridas, o None si no es posible identificarlas.
    last_error = None
    raw = None
    for attempt in range(3):
        # Se prueban los engines por si el contenido no coincide con la extensión del archivo
        for engine in (None, "openpyxl", "xlrd"):
            try:
                read_kwargs = {} if engine is None else {"engine": engine}
                raw = pd.read_excel(excel_path, header=None, dtype=object, **read_kwargs)
                break
            except Exception as e:
                last_error = e
        if raw is not None:
            break
        time.sleep(3)
    if raw is None:
        raise last_error

    # Caso 1: el archivo trae la fila de encabezados (formatos anteriores)
    header_idx = None
    for idx in range(min(len(raw), 200)):
        values = [str(v).strip() if pd.notna(v) else "" for v in raw.iloc[idx]]
        if all(col in values for col in required_columns):
            header_idx = idx
            break

    if header_idx is not None:
        df = raw.iloc[header_idx + 1:].copy()
        df.columns = [str(v).strip() if pd.notna(v) else "" for v in raw.iloc[header_idx]]
        return df.dropna(how="all")

    # Caso 2: el export moderno viene sin encabezados: se detectan las columnas por contenido
    detected = _detect_sap_columns_without_header(raw, required_columns)
    if detected is None:
        return None

    data = raw.dropna(how="all").reset_index(drop=True)
    result = pd.DataFrame()
    for name in required_columns:
        idx = detected.get(name)
        if idx is None:
            result[name] = ""
        else:
            result[name] = data.iloc[:, idx].values
    return result.dropna(how="all")

def resolve_sap_excel_path(base_path, filename):
    # SAP/Excel pueden guardar el reporte con extensión .xlsx o .xls: se resuelve cuál existe
    expected = os.path.join(base_path, filename)
    if os.path.exists(expected):
        return expected
    base = os.path.splitext(expected)[0]
    for candidate in (base + ".xlsx", base + ".xls"):
        if os.path.exists(candidate):
            logger.info(f"Archivo exportado de SAP encontrado: {candidate}")
            return candidate
    return expected

def process_wells_fargo_files(base_path):
    logger.info("=" * 70)
    logger.info(f"Inicio de ejecución - {AGENT_NAME}")
    logger.info(f"Carpeta de datos: {os.path.abspath(base_path)}")

    # 1. Ejecutar extracción de SAP
    if not extract_sap_report(base_path):
        logger.warning("La extracción desde SAP no finalizó correctamente; se continuará con el archivo Customers_SAP.xlsx existente si está disponible.")
        send_notification_email(
            f"[ALERTA] {AGENT_NAME} - fallo al extraer reporte de SAP",
            f"No fue posible extraer el reporte de SAP (transacción ZSD_POS_1052).\n\n"
            f"Fecha/hora: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            f"Carpeta de datos: {os.path.abspath(base_path)}\n\n"
            f"El agente continuó el proceso con el archivo Customers_SAP.xlsx existente "
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
    sap_filename = "Customers_SAP.xlsx"
    csv_filename = "FlatFile.csv"
    
    input_zip_path = os.path.join(base_path, input_zip_filename)
    output_zip_path = os.path.join(base_path, output_zip_filename)
    sap_path = resolve_sap_excel_path(base_path, sap_filename)
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
    logger.info(f"Leyendo archivos {csv_filename} y {sap_filename}...")
    if not os.path.exists(sap_path):
        notify_failure("lectura del archivo de SAP", f"No se encontró el archivo {sap_path}. Verifica que SAP exportó correctamente.")
        return False

    # Leer CSV. Puede venir en UTF-8 o Windows-1252, y en dos formatos:
    #  - CSV estándar separado por comas
    #  - Formato de SAP BO sin delimitador: "campo1""campo2""campo3"
    df_csv, csv_layout, csv_encoding = read_flatfile_csv(temp_csv_path)
    if df_csv is None:
        notify_failure(
            "lectura del archivo de entrada",
            f"No fue posible leer/decodificar {temp_csv_path} en ninguno de los formatos "
            f"ni codificaciones esperadas.",
        )
        return False

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
    
    # Leer Excel exportado de SAP (List > Export > Spreadsheet)
    # El layout del reporte ALV puede cambiar de orden: se localiza dinámicamente la fila
    # de encabezados y se toman únicamente las columnas necesarias por nombre.
    required_columns = ['Customer', 'Name 1', 'Terms Paym']
    try:
        df_sap_raw = read_sap_report_excel(sap_path, required_columns)
    except Exception as e:
        notify_failure("lectura del archivo de SAP", f"Error leyendo el archivo {sap_filename}: {e}")
        return False

    if df_sap_raw is None:
        notify_failure(
            "lectura del archivo de SAP",
            f"No se encontraron las columnas {', '.join(required_columns)} en {sap_filename}. "
            f"Verifique que el layout del reporte en SAP conserve estos encabezados.",
        )
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

        # Excel puede exportar el código como número y perder los ceros a la izquierda:
        # se restituye el relleno estándar de SAP (10 caracteres)
        if code.isdigit() and len(code) < 10:
            code = code.zfill(10)

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

    # Guardar CSV actualizado (preservando el formato del archivo original)
    write_flatfile_csv(df_updated, temp_csv_path, csv_layout, csv_encoding)
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
