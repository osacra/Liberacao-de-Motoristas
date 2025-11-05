
# Diretórios e arquivos
DOWNLOAD_PATH = r"C:\Users\dwbe01\Downloads\ProjetoMotoristas_GUI\ProjetoMotoristas\planilhas"
FILE_NAME = "Liberação de Motoristas.xlsx"
DOWNLOADED_FILE = f"{DOWNLOAD_PATH}\\{FILE_NAME}"  
FILE_AUXILIAR = r'C:\Users\dwbe01\Downloads\ProjetoMotoristas_GUI\ProjetoMotoristas\planilhas\Liberação de Motoristas - Auxiliar.xlsx'
CAMINHO_BASE = r"C:\Users\dwbe01\Downloads\ProjetoMotoristas_GUI\ProjetoMotoristas\planilhas\Cadastro dos Motoristas1.xlsx"
FILE_OFICIAL = r'C:\Projeto Motoristas\planilhas\Liberação de Motoristas e Ajudantes.xlsx'


# URLs e destinatários

EXCEL_ONLINE_URL = (
    "https://dpdhl-my.sharepoint.com/:x:/r/personal/arthur_mendessacramento_dhl_com/_layouts/15/Doc.aspx?"
    "sourcedoc=%7B70A1618C-C082-4E6D-8DF0-AEA614CFB622%7D&file=Libera%C3%A7%C3%A3o%20de%20Motoristas.xlsx&"
    "action=edit&mobileredirect=true&wdMsFormsCorrelationId=b55083de-31fd-4c5b-a18a-83bcc477edd4&"
    "wdtf=%20Microsoft.Office.Excel.FMsFormsMetadataInWorkbookMetadata%3Atrue"
)

DESTINATARIOS = [
    'oarthursacra@gmail.com',
]

NIVEL_LOG = 'INFO'

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
