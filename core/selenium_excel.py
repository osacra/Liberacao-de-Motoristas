# core/selenium_excel.py
import os
import time
import logging
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from core.excel_handler import esperar_arquivo_liberar
from core.alerts import alerta_erro
from config.settings import DOWNLOADED_FILE, EXCEL_ONLINE_URL

def abrir_excel_online_com_bypass(driver, url=EXCEL_ONLINE_URL, tentativas=5):
    for tentativa in range(tentativas):
        try:
            driver.get(url)
            time.sleep(5)
            page_source = driver.page_source.lower()
            title = driver.title.lower()

            if "smart connect" in page_source or "perfil" in page_source or "smart connect" in title:
                logging.warning(f"Smart Connect detectado, recarregando ({tentativa + 1}/{tentativas})")
                time.sleep(10)
                continue

            iframe = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, "//iframe[contains(@id, 'WacFrame_Excel_')]"))
            )
            driver.switch_to.frame(iframe)
            logging.info("Excel Online carregado e bypass aplicado.")
            return True

        except Exception as e:
            logging.error(f"Erro ao abrir Excel Online (tentativa {tentativa + 1}): {e}")
            time.sleep(10)

    alerta_erro("Não foi possível carregar o Excel Online após várias tentativas.")
    return False


def baixar_planilha_excel_online(driver, arquivo_destino=DOWNLOADED_FILE):
    try:
        if os.path.exists(arquivo_destino):
            if not esperar_arquivo_liberar(arquivo_destino):
                return False
            os.remove(arquivo_destino)
            logging.info(f"Arquivo antigo removido: {arquivo_destino}")

        driver.get(EXCEL_ONLINE_URL)
        WebDriverWait(driver, 60).until(lambda d: d.execute_script("return document.readyState") == "complete")
        logging.info("Página carregada (readyState complete).")

        iframe = WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH, "//iframe[contains(@id, 'WacFrame_Excel_')]"))
        )
        driver.switch_to.frame(iframe)
        logging.info("Dentro do iframe do Excel Online.")

        time.sleep(15)  # Espera extra

        try:
            file_button = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(@aria-label,'Ficheiro')]"))
            )
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", file_button)
            time.sleep(1)
            driver.execute_script("arguments[0].click();", file_button)
            logging.info("Menu 'File' aberto com sucesso.")
            time.sleep(2)  # Pequena pausa para o menu carregar
        except TimeoutException:
            logging.error("Botão 'File' não encontrado.")
            driver.switch_to.default_content()
            return False

        # Clicar em File -> Create a Copy / Download
        botoes_possiveis = ["Create a Copy", "Criar uma Cópia"]
        create_copy = None

        for texto_botao in botoes_possiveis:
            try:
                create_copy = WebDriverWait(driver, 15).until(
                    EC.element_to_be_clickable((By.XPATH, f"//span[normalize-space(text())='{texto_botao}']"))
                )
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", create_copy)
                time.sleep(1)
                driver.execute_script("arguments[0].click();", create_copy)
                logging.info(f"Botão '{texto_botao}' clicado.")
                break
            except TimeoutException:
                logging.warning(f"Botão '{texto_botao}' não encontrado, tentando próximo...")

        if create_copy is None:
            logging.error("Botão para criar cópia não encontrado.")
            driver.switch_to.default_content()
            return False

        botoes_download = ["Transferir uma Cópia", "Download a Copy", "Baixar uma Cópia"]
        download_copy = None

        for texto_botao in botoes_download:
            try:
                download_copy = WebDriverWait(driver, 15).until(
                    EC.element_to_be_clickable((By.XPATH, f"//span[normalize-space(text())='{texto_botao}']"))
                )
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", download_copy)
                time.sleep(1)
                driver.execute_script("arguments[0].click();", download_copy)
                logging.info(f"Botão '{texto_botao}' clicado.")
                break
            except TimeoutException:
                logging.warning(f"Botão '{texto_botao}' não encontrado, tentando próximo...")

        if download_copy is None:
            logging.error("Botão de download não encontrado.")
            driver.switch_to.default_content()
            return False

        logging.info("Download solicitado, aguardando arquivo aparecer...")

        for _ in range(60):
            if os.path.exists(arquivo_destino) and esperar_arquivo_liberar(arquivo_destino):
                logging.info("Arquivo baixado com sucesso.")
                driver.switch_to.default_content()
                return True
            time.sleep(1)

        logging.error("Timeout ao aguardar download do arquivo.")
        driver.switch_to.default_content()
        return False

    except Exception as e:
        logging.error(f"Erro ao baixar planilha: {e}")
        try:
            driver.switch_to.default_content()
        except Exception:
            pass
        return False