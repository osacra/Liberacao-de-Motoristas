import logging
import os
from datetime import datetime
from win10toast import ToastNotifier


toaster = ToastNotifier()

def configurar_logger(log_folder="logs", log_filename="liberacao_motoristas.log"):
    log_path = os.path.join(log_folder, log_filename)

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s — %(levelname)s — %(message)s",
        handlers=[
            logging.FileHandler(log_path, encoding="utf-8"),
            logging.StreamHandler()
        ]
    )

    logging.info(f"==================== INÍCIO EXECUÇÃO {datetime.now()} ====================")

def alerta_erro(mensagem):
    """Alerta sonoro + toast + log"""
    for _ in range(3):
        import winsound
        import time
        winsound.Beep(1000, 300)
        time.sleep(0.2)
    toaster.show_toast("Erro na Liberação", mensagem, duration=10)
    logging.error(mensagem)
