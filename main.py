# main.py

import sys
import warnings
import logging
from core.alerts import configurar_logger
from ui.main_ui import TelaInicial
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import Qt

# Configuração de logging
configurar_logger()
warnings.filterwarnings("ignore", category=UserWarning)

if __name__ == "__main__":
    # Configuração para alta DPI em sistemas compatíveis
    QApplication.setAttribute(Qt.ApplicationAttribute.AA_EnableHighDpiScaling)
    QApplication.setAttribute(Qt.ApplicationAttribute.AA_UseHighDpiPixmaps)
    
    app = QApplication(sys.argv)
    
    janela = TelaInicial()
    janela.show()
    

    app.aboutToQuit.connect(janela.parar_monitoramento)
    
    sys.exit(app.exec())
