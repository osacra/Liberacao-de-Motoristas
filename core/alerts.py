import logging
from datetime import datetime
from win10toast import ToastNotifier

toaster = ToastNotifier()

def configurar_logger():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s — %(levelname)s — %(message)s",
        handlers=[
            logging.FileHandler("liberacao_motoristas.log", encoding="utf-8"),
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
