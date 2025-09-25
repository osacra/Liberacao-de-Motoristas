
# Diretórios e arquivos
DOWNLOAD_PATH = r"C:\Projeto Motoristas\planilhas"
FILE_NAME = "Liberação de Motoristas e Ajudantes.xlsx"
DOWNLOADED_FILE = f"{DOWNLOAD_PATH}\\{FILE_NAME}"  
FILE_AUXILIAR = r"C:\Projeto Motoristas\planilhas\Liberação de Motoristas - Auxiliar.xlsx"
CAMINHO_BASE = r"C:\Projeto Motoristas\planilhas\Cadastro dos Motoristas.xlsx"

# URLs e destinatários
EXCEL_ONLINE_URL = (
    "https://dpdhl-my.sharepoint.com/:x:/r/personal/auditoria_shein_dhl_com/_layouts/15/doc2.aspx?"
    "sourcedoc=%7B0C4D4A4F-F9F0-42BD-A1B7-B5B3791A8E07%7D&file=Libera%C3%A7%C3%A3o%20de%20Motoristas%20e%20Ajudantes.xlsx&action=edit&mobileredirect=true&wdMsFormsCorrelationId=d6699260-89e5-441f-8de9-b165afbd4e14&wdtf=%20Microsoft.Office.Excel.FMsFormsMetadataInWorkbookMetadata%3Atrue"
)
DESTINATARIOS = (
    "oarthursacra@gmail.com"
    
)

# Driver e navegador
EDGE_DRIVER_PATH = r"C:\Projeto Motoristas\driver\msedgedriver.exe"
EDGE_OPTIONS = {
    "headless": True,
    "disable_gpu": True,
    "disable_extensions": True,
    "disable_popup_blocking": True,
    "no_sandbox": True,
    "disable_dev_shm_usage": True,
    "prefs": {
        "download.default_directory": DOWNLOAD_PATH,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
    }
}

# Imagem usada no corpo do e-mail
IMAGEM_EMAIL = r"C:\Projeto Motoristas\assets\imagem-dhl.png"

# WebDriver wait
WEBDRIVER_TIMEOUT = 60
