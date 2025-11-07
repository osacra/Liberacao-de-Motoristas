# core/email_handler.py
import os
import logging
import pythoncom
import subprocess
import gc
import shutil
import psutil
import win32com.client as win32
import win32gui
import win32con
from core.alerts import alerta_erro
from config.settings import IMAGEM_EMAIL, DESTINATARIOS


def outlook_aberto():
    """Verifica se o Outlook já está em execução."""
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] and 'OUTLOOK.EXE' in proc.info['name'].upper():
            return True
    return False


def abrir_outlook_minimizado():
    """Abre o Outlook minimizado (caso não esteja aberto)."""
    try:
        if not outlook_aberto():
            # Tenta detectar automaticamente o Outlook em diferentes instalações
            caminhos_possiveis = [
                shutil.which("OUTLOOK.EXE"),  # Caminho automático via PATH do Windows
                r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE",
                r"C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE",
                r"C:\Program Files\Microsoft Office\Office15\OUTLOOK.EXE",
                r"C:\Program Files (x86)\Microsoft Office\Office15\OUTLOOK.EXE",
            ]
            caminhos_possiveis = [c for c in caminhos_possiveis if c and os.path.exists(c)]

            if caminhos_possiveis:
                caminho_outlook = caminhos_possiveis[0]
                subprocess.Popen([caminho_outlook])
                import time; time.sleep(5)
            else:
                logging.warning("⚠️ Outlook não encontrado nos caminhos padrão nem no PATH.")

        # Minimiza a janela se estiver visível
        def enumHandler(hwnd, _):
            if win32gui.IsWindowVisible(hwnd) and "Outlook" in win32gui.GetWindowText(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)

        win32gui.EnumWindows(enumHandler, None)

    except Exception as e:
        logging.warning(f"Erro ao abrir/minimizar Outlook: {e}")


def enviar_email(destinatarios, assunto, corpo):
    """Envia e-mail com imagem no corpo e assinatura padrão."""
    try:
        pythoncom.CoInitialize()
        abrir_outlook_minimizado()

        outlook = win32.gencache.EnsureDispatch('Outlook.Application')
        mail = outlook.CreateItem(0)

        if isinstance(destinatarios, list):
            mail.To = "; ".join(destinatarios)
        else:
            mail.To = str(destinatarios)

        mail.Subject = assunto

        # Imagem no corpo do e-mail
        imagem_tag = ""
        if IMAGEM_EMAIL and os.path.exists(IMAGEM_EMAIL):
            attachment = mail.Attachments.Add(IMAGEM_EMAIL)
            attachment.PropertyAccessor.SetProperty(
                "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
                "imagemDHL"
            )
            imagem_tag = '<img src="cid:imagemDHL"><br>'
        else:
            logging.warning(f"⚠️ Imagem não encontrada: {IMAGEM_EMAIL}")

        # Corpo HTML
        corpo_html = corpo.replace('\n', '<br>')
        complemento_html = (
            "<br>Atenciosamente,"
            "<br><b>DHL Supply Chain</b><br>"
            "GLP Guarulhos II – R. Concretex, 800<br>"
            "CEP: 07232-050, Guarulhos<br>"
            "Brasil"
            f"<br>{imagem_tag}"
        )
        mail.HTMLBody = f"{corpo_html}<br>{complemento_html}"

        mail.Send()
        logging.info(f"✅ E-mail enviado para {destinatarios}")

        # Forçar envio imediato
        session = outlook.GetNamespace("MAPI")
        session.SendAndReceive(False)

        del mail
        del outlook
        gc.collect()
        pythoncom.CoUninitialize()
        return True

    except Exception as e:
        logging.error(f"❌ Erro ao enviar e-mail: {e}")
        alerta_erro(f"Falha ao enviar e-mail: {e}")
        return False
