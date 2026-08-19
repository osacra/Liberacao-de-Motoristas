import os
import json

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.dirname(BASE_DIR)
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")

# Carrega as configurações
with open(CONFIG_PATH, encoding="utf-8") as f:
    config = json.load(f)

# Caminhos das planilhas
DOWNLOAD_PATH = os.path.join(ROOT_DIR, config["arquivos"]["download_path"])
FILE_NAME = config["arquivos"]["file_name"]
DOWNLOADED_FILE = os.path.join(DOWNLOAD_PATH, FILE_NAME)
FILE_AUXILIAR = os.path.join(DOWNLOAD_PATH, config["arquivos"]["file_auxiliar"])
CAMINHO_BASE = os.path.join(DOWNLOAD_PATH, config["arquivos"]["caminho_base"])


# URLs e destinatários
EXCEL_ONLINE_URL = config["urls"]["excel_online"]
DESTINATARIOS = config["emails"]["destinatarios"]

# Nível de log
NIVEL_LOG = config["logging"]["nivel"]

# Driver e opções
EDGE_DRIVER_PATH = os.path.join(ROOT_DIR, config["driver"]["path"])
EDGE_OPTIONS = {
    **config["driver"]["options"],
    "prefs": {
        "download.default_directory": DOWNLOAD_PATH,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True
    },
}

# Imagem usada no corpo do e-mail
IMAGEM_EMAIL = os.path.join(ROOT_DIR, config["email_assets"]["imagem"])

# Timeout padrão do WebDriver
WEBDRIVER_TIMEOUT = 60
