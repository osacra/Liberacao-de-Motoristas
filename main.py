# main.py

import warnings
warnings.filterwarnings("ignore", category=UserWarning)

import logging
import time
from config.settings import EXCEL_ONLINE_URL, DOWNLOADED_FILE
from core.driver_setup import criar_driver
from core.excel_handler import fechar_processos_excel_outlook
from core.selenium_excel import abrir_excel_online_com_bypass, baixar_planilha_excel_online
from core.processador import processar_envios
from core.utils import bloquear_suspensao_tela, mexer_mouse, abrir_outlook_minimizado, outlook_aberto
from core.alerts import configurar_logger
import os

# --- Inicializações ---
configurar_logger()
bloquear_suspensao_tela()
fechar_processos_excel_outlook()
driver, wait = criar_driver()

if not outlook_aberto():
    abrir_outlook_minimizado()


# --- Loop principal ---
while True:
    try:
        if abrir_excel_online_com_bypass(driver, EXCEL_ONLINE_URL):
            if baixar_planilha_excel_online(driver):
                if os.path.exists(DOWNLOADED_FILE):
                    processar_envios()
                else:
                    logging.warning("Arquivo não encontrado após tentativa de download. Pulando ciclo.")
            else:
                logging.warning("Falha ao baixar a planilha. Aguardando 30 segundos.")
        else:
            logging.warning("Falha ao abrir Excel Online. Aguardando 30 segundos.")

        mexer_mouse()
        logging.info("Aguardando 10 segundos para próxima verificação...\n")
        time.sleep(10)

    except Exception as e:
        logging.error(f"Erro durante o ciclo: {e}")
        logging.info("Aguardando 30 segundos antes de tentar novamente...")
        time.sleep(30)

    except KeyboardInterrupt:
        logging.info("Execução finalizada manualmente pelo usuário.")
        break

# --- Finalização ---
try:
    driver.quit()
except Exception:
    pass

logging.info("Driver fechado, script finalizado.")
