# core/utils.py
import pyautogui
import time
import logging
import ctypes
import win32gui
import os
import subprocess
import win32con


def outlook_aberto():
    import psutil
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] and 'OUTLOOK.EXE' in proc.info['name'].upper():
            return True
    return False

def abrir_outlook_minimizado():
    try:
        
        if not outlook_aberto():
            caminho_outlook = r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
            if not os.path.exists(caminho_outlook):
                caminho_outlook = r"C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE"
            subprocess.Popen([caminho_outlook])
            time.sleep(5)  

        # Minimiza janela do Outlook
        def enumHandler(hwnd, lParam):
            if win32gui.IsWindowVisible(hwnd):
                if "Outlook" in win32gui.GetWindowText(hwnd):
                    win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
        win32gui.EnumWindows(enumHandler, None)

    except Exception as e:
        logging.warning(f"Erro ao abrir/minimizar Outlook: {e}")


def mexer_mouse():
    try:
        x, y = pyautogui.position()
        screenWidth, screenHeight = pyautogui.size()
        margin = 2

        if (x <= margin and y <= margin) or \
           (x >= screenWidth - margin and y <= margin) or \
           (x <= margin and y >= screenHeight - margin) or \
           (x >= screenWidth - margin and y >= screenHeight - margin):
            pyautogui.moveTo(screenWidth // 2, screenHeight // 2)
            logging.info("Mouse estava em canto, movido para o centro da tela.")
            return

        pyautogui.move(1, 0)
        time.sleep(0.05)
        pyautogui.move(-1, 0)
        logging.debug("Mouse tremido para evitar bloqueio de tela.")

    except Exception as e:
        logging.warning(f"Não foi possível mexer o mouse: {e}")


def esperar_arquivo_liberar(caminho, tentativas=10, intervalo=2):

    for tentativa in range(tentativas):
        try:
            with open(caminho, 'a'):
                return True
        except PermissionError:
            logging.info(f"Arquivo ainda em uso, aguardando... ({tentativa + 1}/{tentativas})")
            time.sleep(intervalo)
    logging.error(f"Timeout ao aguardar liberação do arquivo: {caminho}")
    return False

def obter_saudacao():
    from datetime import datetime
    hora = datetime.now().hour
    if 5 <= hora < 12:
        return "Bom dia"
    elif 12 <= hora < 18:
        return "Boa tarde"
    else:
        return "Boa noite"
    
def bloquear_suspensao_tela():
    """Bloqueia suspensão e desligamento de tela no Windows."""
    ES_CONTINUOUS = 0x80000000
    ES_SYSTEM_REQUIRED = 0x00000001
    ES_DISPLAY_REQUIRED = 0x00000002

    resultado = ctypes.windll.kernel32.SetThreadExecutionState(
        ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED
    )

    if resultado == 0:
        logging.warning("Falha ao aplicar bloqueio de suspensão no sistema.")
    else:
        logging.info("Bloqueio de suspensão e de desligamento de tela habilitado com sucesso.")
    

