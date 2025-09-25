import warnings
warnings.filterwarnings("ignore", category=UserWarning)

import logging
import pandas as pd
import os
import time
from datetime import datetime
import win32com.client as win32
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import subprocess
import gc
import pythoncom
import ctypes
from win10toast import ToastNotifier
from selenium.common.exceptions import TimeoutException
import pyautogui
import winsound
import win32gui
import win32con
# --- Configuração do logging ---
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s — %(levelname)s — %(message)s",
    handlers=[
        logging.FileHandler("liberacao_motoristas.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)

logging.info(f"==================== INÍCIO EXECUÇÃO {datetime.now()} ====================")

toaster = ToastNotifier()

# --- Configurações gerais ---
DOWNLOAD_PATH = r"C:\automação"
FILE_NAME = "Liberação de Motoristas e Ajudantes.xlsx"
DOWNLOADED_FILE = os.path.join(DOWNLOAD_PATH, FILE_NAME)
FILE_AUXILIAR = r"C:\automação\Liberação de Motoristas - Auxiliar.xlsx"
EXCEL_ONLINE_URL = (
    "https://dpdhl-my.sharepoint.com/:x:/r/personal/auditoria_shein_dhl_com/_layouts/15/doc2.aspx?sourcedoc=%7B0C4D4A4F-F9F0-42BD-A1B7-B5B3791A8E07%7D&file=Libera%C3%A7%C3%A3o%20de%20Motoristas%20e%20Ajudantes.xlsx&action=edit&mobileredirect=true&wdMsFormsCorrelationId=d6699260-89e5-441f-8de9-b165afbd4e14&wdtf=%20Microsoft.Office.Excel.FMsFormsMetadataInWorkbookMetadata%3Atrue")
DESTINATARIOS = "guarulhos2.gestor@parqueslogisticos.com.br; br.dsc.guarulhos.cftv.shein@dpdhl.onmicrosoft.com; portaria.dhl@sheingroup.com; guarulhos2.adm@parqueslogisticos.com.br; guarulhos2.p1@parqueslogisticos.com.br"
CAMINHO_BASE = r'C:\automação\Cadastro dos Motoristas.xlsx'


# --- Setup Edge com driver local fixo ---
options = Options()
options.add_argument("--headless=new")
options.add_argument("--headless=new")
options.add_argument("--disable-gpu")
options.add_argument("--disable-extensions")
options.add_argument("--disable-popup-blocking")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
prefs = {
    "download.default_directory": DOWNLOAD_PATH,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    
}
options.add_experimental_option("prefs", prefs)

# Caminho fixo para o EdgeDriver
EDGE_DRIVER_PATH = r"C:\automação\driver\msedgedriver.exe"
service = Service(EDGE_DRIVER_PATH)

driver = webdriver.Edge(service=service, options=options)
wait = WebDriverWait(driver, 60)



def mexer_mouse():
    try:
        x, y = pyautogui.position()
        screenWidth, screenHeight = pyautogui.size()
        margin = 2  # tolerância para canto

        # Se o mouse estiver em qualquer canto, mova para o centro
        if (x <= margin and y <= margin) or \
           (x >= screenWidth - margin and y <= margin) or \
           (x <= margin and y >= screenHeight - margin) or \
           (x >= screenWidth - margin and y >= screenHeight - margin):
            center_x, center_y = screenWidth // 2, screenHeight // 2
            pyautogui.moveTo(center_x, center_y)
            logging.info("Mouse estava em canto, movido para o centro da tela para evitar fail-safe.")
            return

        # Se não estiver em canto, apenas treme
        pyautogui.move(1, 0)
        time.sleep(0.05)
        pyautogui.move(-1, 0)
        logging.debug("Mouse tremido para evitar bloqueio de tela.")
    except Exception as e:
        logging.warning(f"Não foi possível mexer o mouse: {e}")


def alerta_erro(mensagem):
    for _ in range(3):
        winsound.Beep(1000, 300)
        time.sleep(0.2)
    toaster.show_toast("Erro na Liberação", mensagem, duration=10)
    logging.error(mensagem)



def fechar_processos_excel_outlook():
    try:
        subprocess.run('taskkill /f /im excel.exe', shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        logging.info("Excel finalizado para liberar arquivo.")
    except Exception as e:
        logging.warning(f"Erro ao finalizar processos Excel: {e}")


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


def abrir_excel_online_com_bypass(driver, url, tentativas=5):
    for tentativa in range(tentativas):
        try:
            driver.get(url)
            time.sleep(5)

            # Tenta detectar o Smart Connect por texto na página
            page_source = driver.page_source.lower()
            title = driver.title.lower()

            if "smart connect" in page_source or "perfil" in page_source or "smart connect" in title:
                logging.warning(f"Smart Connect detectado, recarregando página ({tentativa + 1}/{tentativas})")
                time.sleep(10)
                continue

            # Tenta detectar iframe do Excel Online - garantia extra com XPath
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


def baixar_planilha_excel_online(driver):
    try:
        if os.path.exists(DOWNLOADED_FILE):
            if not esperar_arquivo_liberar(DOWNLOADED_FILE):
                return False
            os.remove(DOWNLOADED_FILE)
            logging.info(f"Arquivo antigo removido: {DOWNLOADED_FILE}")

        driver.get(EXCEL_ONLINE_URL)

        # Aguarda carregamento completo do documento
        WebDriverWait(driver, 60).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
        logging.info("Página carregada (readyState complete).")

        # Aguarda carregar o iframe
        iframe = WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH, "//iframe[contains(@id, 'WacFrame_Excel_')]"))
        )
        driver.switch_to.frame(iframe)
        logging.info("Dentro do iframe do Excel Online.")

        logging.info("Aguardando carregamento adicional do Excel Online (15s)...")
        time.sleep(15)

        # Abrir menu File de forma garantida
        file_menu = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.ID, "FileMenuFlyoutLauncher"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", file_menu)
        time.sleep(1)
        driver.execute_script("arguments[0].click();", file_menu)
        logging.info("Menu 'File' clicado.")

        # Tentar clicar no botão "Create a Copy" ou "Criar uma Cópia"
        botoes_possiveis = ["Create a Copy", "Criar uma Cópia"]
        create_copy = None

        for texto_botao in botoes_possiveis:
            try:
                create_copy = WebDriverWait(driver, 15).until(
                    EC.element_to_be_clickable((By.XPATH, f"//span[normalize-space(text())='{texto_botao}']"))
                )
                if create_copy:
                    logging.info(f"Botão '{texto_botao}' encontrado, clicando.")
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", create_copy)
                    time.sleep(1)
                    driver.execute_script("arguments[0].click();", create_copy)
                    break
            except TimeoutException:
                logging.warning(f"Botão '{texto_botao}' não encontrado, tentando próximo...")

        if create_copy is None:
            logging.error("Botão para criar cópia não encontrado. Abortando download.")
            driver.switch_to.default_content()
            return False

        # Clicar no "Download a Copy" ou "Baixar uma Cópia"
        botoes_download = ["Download a Copy", "Baixar uma Cópia"]
        download_copy = None

        for texto_botao in botoes_download:
            try:
                download_copy = WebDriverWait(driver, 15).until(
                    EC.element_to_be_clickable((By.XPATH, f"//span[normalize-space(text())='{texto_botao}']"))
                )
                if download_copy:
                    logging.info(f"Botão '{texto_botao}' encontrado, clicando.")
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", download_copy)
                    time.sleep(1)
                    driver.execute_script("arguments[0].click();", download_copy)
                    break
            except TimeoutException:
                logging.warning(f"Botão '{texto_botao}' não encontrado, tentando próximo...")

        if download_copy is None:
            logging.error("Botão de download não encontrado. Abortando download.")
            driver.switch_to.default_content()
            return False

        logging.info("Download solicitado, aguardando arquivo aparecer...")

        # Esperar arquivo aparecer e estar liberado
        for _ in range(60):
            if os.path.exists(DOWNLOADED_FILE) and esperar_arquivo_liberar(DOWNLOADED_FILE):
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



def esperar_refresh_concluir(workbook, timeout=60):
    """Aguarda até o Excel concluir o RefreshAll."""
    inicio = time.time()
    while True:
        try:
            if not workbook.Refreshing:
                break
        except AttributeError:
            break 
        if time.time() - inicio > timeout:
            logging.warning("Timeout ao aguardar o RefreshAll finalizar, prosseguindo mesmo assim.")
            break
        time.sleep(1)


def atualizar_planilha_excel(caminho_arquivo):
    try:
        pythoncom.CoInitialize()

        if not esperar_arquivo_liberar(caminho_arquivo, tentativas=10, intervalo=2):
            logging.error(f"Arquivo ainda bloqueado antes de abrir no Excel COM: {caminho_arquivo}")
            alerta_erro(f"Arquivo bloqueado antes de abrir no Excel: {caminho_arquivo}")
            return

        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(caminho_arquivo, ReadOnly=False)
        wb.RefreshAll()
        logging.info("Atualizando planilha, aguardando 15 segundos...")

        for _ in range(15):
            mexer_mouse()
            time.sleep(1)

        caminho_temp = caminho_arquivo.replace(".xlsx", "_temp.xlsx")
        wb.SaveAs(caminho_temp)
        wb.Close(False)
        excel.Quit()
        del wb
        del excel

        # Substitui o arquivo original pelo salvo
        if os.path.exists(caminho_arquivo):
            os.remove(caminho_arquivo)
        os.rename(caminho_temp, caminho_arquivo)

        gc.collect()
        pythoncom.CoUninitialize()
        logging.info("Planilha atualizada e substituída sem bloqueios.")

    except Exception as e:
        logging.error(f"Erro ao atualizar planilha: {e}")
        alerta_erro(f"Erro ao atualizar planilha: {e}")






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

        # Minimiza janela do Outlook (ajuste título se necessário)
        def enumHandler(hwnd, lParam):
            if win32gui.IsWindowVisible(hwnd):
                if "Outlook" in win32gui.GetWindowText(hwnd):
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

        # Prepara caminho da imagem
        imagem_path = r"C:\automação\imagem-dhl.png"
        if os.path.exists(imagem_path):
            attachment = mail.Attachments.Add(imagem_path)
            attachment.PropertyAccessor.SetProperty(
                "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
                "imagemDHL"
            )
            imagem_tag = '<img src="cid:imagemDHL"><br>'
        else:
            logging.warning(f"Imagem não encontrada: {imagem_path}")
            imagem_tag = ""

        # Formata corpo em HTML
        corpo_html = corpo.replace('\n', '<br>')

        complemento_html = (
            "<br>Att."
            "<br>DHL Supply Chain<br>"
            "GLP Guarulhos II – R. Concretex, 800<br>"
            "CEP: 07232-050, Guarulhos<br>"
            "Brasil"
            "<br><b>DHL Supply Chain - Excellence. Simply Delivered</b><br>"
            f"{imagem_tag}"
        )

        mail.HTMLBody = f"{corpo_html}<br>{complemento_html}"

        mail.Send()
        logging.info(f"E-mail enviado para {destinatarios}")

        # >>> FORÇA ENVIO IMEDIATO <<<
        session = outlook.GetNamespace("MAPI")
        session.SendAndReceive(False)
        logging.info("Sincronização forçada: SendAndReceive executado.")

        del mail
        del outlook
        gc.collect()
        pythoncom.CoUninitialize()
        return True

    except Exception as e:
        logging.error(f"Erro ao enviar e-mail: {e}")
        alerta_erro(f"Falha ao enviar e-mail: {e}")
        return False




def obter_saudacao():
    hora = datetime.now().hour
    if 5 <= hora < 12:
        return "Bom dia"
    elif 12 <= hora < 18:
        return "Boa tarde"
    else:
        return "Boa noite"


def processar_envios():
    pythoncom.CoInitialize()
    try:
        fechar_processos_excel_outlook()
        atualizar_planilha_excel(DOWNLOADED_FILE)

        if not esperar_arquivo_liberar(DOWNLOADED_FILE, tentativas=10, intervalo=1):
            logging.error("Arquivo ainda bloqueado após atualização, pulando leitura.")
            return

        df_oficial = pd.read_excel(DOWNLOADED_FILE, dtype={'Id': str})
        df_oficial['Id'] = df_oficial['Id'].astype(str).str.strip()

        try:
            df_aux = pd.read_excel(FILE_AUXILIAR, dtype={'Id': str})
            df_aux['Id'] = df_aux['Id'].astype(str).str.strip()
            logging.info("Planilha auxiliar carregada.")
        except FileNotFoundError:
            df_aux = pd.DataFrame(columns=['Id', 'Data de Envio', 'Nome do Motorista'])
            logging.info("Planilha auxiliar criada, pois não existia.")

        novos_registros = df_oficial[~df_oficial['Id'].isin(df_aux['Id'])].sort_values(by='Id', key=lambda x: x.astype(int))

        if novos_registros.empty:
            logging.info("Nenhum novo registro para envio.")
            return

        primeiro = novos_registros.iloc[0]
        nome_motorista = primeiro['Nome do Motorista']
        cpf = primeiro.get('CPF do Motorista', 'Sem CPF') or 'Sem CPF'
        placa_cavalo = primeiro.get('Placa do Cavalo', 'Sem Placa Cavalo') or 'Sem Placa Cavalo'
        placa_carreta = primeiro.get('Placa da Carreta', '')
        saudacao = obter_saudacao()
        texto_carreta = f"Placa da carreta: {placa_carreta}\n" if pd.notna(placa_carreta) and placa_carreta else ""

        corpo = (
            f"{saudacao}, portaria!\n\n"
            "Segue abaixo a liberação do motorista.\n\n"
            f"Nome: {nome_motorista}\n"
            f"CPF: {cpf}\n"
            f"Placa do cavalo: {placa_cavalo}\n"
            f"{texto_carreta}"
        )

        ajudantes = []
        for i in range(1, 6):
            nome_col = f"Nome do Ajudante {i}"
            cpf_col = f"CPF do Ajudante {i}"
            nome = primeiro.get(nome_col, "")
            cpf = primeiro.get(cpf_col, "")

            if isinstance(nome, str) and nome.strip():
                ajudantes.append((nome.strip(), cpf.strip() if isinstance(cpf, str) else ""))

        if ajudantes:
            if len(ajudantes) == 1:
                nome, cpf = ajudantes[0]
                corpo += f"Ajudante: {nome}\nCPF: {cpf}\n"
            else:
                for idx, (nome, cpf) in enumerate(ajudantes, start=1):
                    corpo += f"Ajudante {idx}: {nome}\nCPF: {cpf}\n\n"

        assunto = "Liberação de Motorista"
        logging.info(f"Enviando e-mail - ID: {primeiro['Id']} | Nome: {nome_motorista}")
        envio_sucesso = enviar_email(DESTINATARIOS, assunto, corpo)

        if envio_sucesso:
            df_aux = pd.concat([df_aux, pd.DataFrame([{
                'Id': primeiro['Id'],
                'Data de Envio': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'Nome do Motorista': nome_motorista
            }])], ignore_index=True)
            df_aux.to_excel(FILE_AUXILIAR, index=False)
            logging.info(f"Registro {primeiro['Id']} salvo na auxiliar.")
        else:
            logging.warning(f"Erro ao enviar e-mail de {primeiro['Id']}, não salvo na auxiliar para tentar novamente no próximo ciclo.")

    except Exception as e:
        logging.error(f"Erro em processar_envios: {e}")
        alerta_erro(f"Erro no processo: {e}")

    finally:
        pythoncom.CoUninitialize()





# Loop Principal
if __name__ == "__main__":
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

    try:
        fechar_processos_excel_outlook()

        if not outlook_aberto():
            abrir_outlook_minimizado()

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
                logging.info(" Aguardando 10 segundos para próxima verificação...\n")
                time.sleep(10)

            except Exception as e:
                logging.error(f"Erro durante o ciclo: {e}")
                logging.info("Aguardando 30 segundos antes de tentar novamente...")
                time.sleep(30)

    except KeyboardInterrupt:
        logging.info("Execução finalizada manualmente pelo usuário.")
    except Exception as e:
        logging.error(f"Erro fatal no loop principal: {e}")
    finally:
        try:
            driver.quit()
        except:
            pass
        logging.info("Driver fechado, script finalizado.")