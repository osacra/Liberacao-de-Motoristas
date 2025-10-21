# core/email_handler.py
import os
import logging
import pythoncom
import subprocess
import gc
import win32com.client as win32
import win32gui
import win32con
from core.alerts import alerta_erro
from config.settings import IMAGEM_EMAIL

def outlook_aberto():
    import psutil
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] and 'OUTLOOK.EXE' in proc.info['name'].upper():
            return True
    return False


def abrir_outlook_minimizado():
    try:
        if not outlook_aberto():
            caminhos_possiveis = [
                r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE",
                r"C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE"
            ]
            caminho_outlook = next((c for c in caminhos_possiveis if os.path.exists(c)), None)
            if caminho_outlook:
                subprocess.Popen([caminho_outlook])
                import time; time.sleep(5)
            else:
                logging.warning("Outlook não encontrado nos caminhos padrões.")

        def enumHandler(hwnd, lParam):
            if win32gui.IsWindowVisible(hwnd) and "Outlook" in win32gui.GetWindowText(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
        win32gui.EnumWindows(enumHandler, None)

    except Exception as e:
        logging.warning(f"Erro ao abrir/minimizar Outlook: {e}")


def enviar_email(destinatarios, assunto, corpo):
    try:
        pythoncom.CoInitialize()
        abrir_outlook_minimizado()

        outlook = win32.gencache.EnsureDispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        mail.To = destinatarios
        mail.Subject = assunto

        if os.path.exists(IMAGEM_EMAIL):
            attachment = mail.Attachments.Add(IMAGEM_EMAIL)
            attachment.PropertyAccessor.SetProperty(
                "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
                "imagemDHL"
            )
            imagem_tag = '<img src="cid:imagemDHL"><br>'
        else:
            logging.warning(f"Imagem não encontrada: {IMAGEM_EMAIL}")
            imagem_tag = ""

        corpo_html = corpo.replace('\n', '<br>')
        complemento_html = (
            "<br>Att."
            "<br>DHL Supply Chain<br>"
            "GLP Guarulhos II – R. Concretex, 800<br>"
            "CEP: 07232-050, Guarulhos<br>"
            "Brasil"
            f"<br>{imagem_tag}"
        )
        mail.HTMLBody = f"{corpo_html}<br>{complemento_html}"
        mail.Send()
        logging.info(f"E-mail enviado para {destinatarios}")

        # Forçar envio imediato
        session = outlook.GetNamespace("MAPI")
        session.SendAndReceive(False)

        del mail
        del outlook
        gc.collect()
        pythoncom.CoUninitialize()
        return True

    except Exception as e:
        logging.error(f"Erro ao enviar e-mail: {e}")
        alerta_erro(f"Falha ao enviar e-mail: {e}")
        return False
