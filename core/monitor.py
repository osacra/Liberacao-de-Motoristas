# core/monitor.py

import logging
import time
import os
from PySide6.QtCore import QThread, Signal

# Importações da lógica de automação existente
from config.settings import EXCEL_ONLINE_URL, DOWNLOADED_FILE
from core.driver_setup import criar_driver
from core.excel_handler import fechar_processos_excel_outlook
from core.selenium_excel import abrir_excel_online_com_bypass, baixar_planilha_excel_online
from core.processador import processar_envios
from core.utils import mexer_mouse, outlook_aberto, abrir_outlook_minimizado

# Configuração de logging (necessária para que o logger funcione dentro da thread)
logger = logging.getLogger()
logger.setLevel(logging.INFO)

class MonitoramentoThread(QThread):
    # Sinais para comunicação com a GUI
    status_changed = Signal(str, str) # (componente, status)
    log_message = Signal(str)
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self._is_running = False
        self._is_paused = False
        self.driver = None
        self.wait = None
        self.intervalo_verificacao = 10 # Segundos
        
    def run(self):
        self._is_running = True
        self.status_changed.emit("monitoramento", "Iniciado")
        self.log_message.emit("Monitoramento iniciado.")
        
        # Inicialização do driver e Outlook
        try:
            self.log_message.emit("Inicializando driver e verificando Outlook...")
            fechar_processos_excel_outlook()
            self.driver, self.wait = criar_driver()
            self.status_changed.emit("conexao_api", "Conectado")
            
            if not outlook_aberto():
                abrir_outlook_minimizado()
                
        except Exception as e:
            self.log_message.emit(f"ERRO FATAL na inicialização: {e}")
            self.status_changed.emit("monitoramento", "Parado")
            self.status_changed.emit("conexao_api", "Erro")
            self._is_running = False
            return

        while self._is_running:
            if self._is_paused:
                time.sleep(1)
                continue
            
            try:
                self.log_message.emit("Iniciando ciclo de verificação...")
                self.status_changed.emit("ultima_verificacao", time.strftime("%H:%M:%S"))
                
                if abrir_excel_online_com_bypass(self.driver, EXCEL_ONLINE_URL):
                    self.log_message.emit("Excel Online aberto com sucesso.")
                    if baixar_planilha_excel_online(self.driver):
                        self.log_message.emit("Planilha baixada com sucesso.")
                        if os.path.exists(DOWNLOADED_FILE):
                            self.log_message.emit("Iniciando processamento de envios...")
                            processar_envios()
                            self.log_message.emit("Processamento de envios concluído.")
                        else:
                            self.log_message.emit("AVISO: Arquivo não encontrado após tentativa de download. Pulando ciclo.")
                    else:
                        self.log_message.emit("AVISO: Falha ao baixar a planilha. Tentando novamente no próximo ciclo.")
                else:
                    self.log_message.emit("AVISO: Falha ao abrir Excel Online. Tentando novamente no próximo ciclo.")

                # Ação para evitar bloqueio de tela (apenas se não estiver pausado)
                if not self._is_paused:
                    mexer_mouse()
                
                self.log_message.emit(f"Ciclo concluído. Aguardando {self.intervalo_verificacao} segundos...")
                time.sleep(self.intervalo_verificacao)

            except Exception as e:
                self.log_message.emit(f"ERRO durante o ciclo: {e}")
                self.log_message.emit("Aguardando 30 segundos antes de tentar novamente...")
                time.sleep(30)

        # Finalização
        self.log_message.emit("Finalizando monitoramento...")
        try:
            if self.driver:
                self.driver.quit()
        except Exception:
            pass
        self.driver = None
        self.wait = None
        self.status_changed.emit("monitoramento", "Parado")
        self.status_changed.emit("conexao_api", "Desconectado")
        self.log_message.emit("Monitoramento finalizado.")

    def stop(self):
        self._is_running = False
        self.log_message.emit("Sinal de parada recebido. Finalizando thread...")

    def pause(self):
        self._is_paused = True
        self.status_changed.emit("monitoramento", "Pausado")
        self.log_message.emit("Monitoramento pausado.")

    def resume(self):
        self._is_paused = False
        self.status_changed.emit("monitoramento", "Iniciado")
        self.log_message.emit("Monitoramento retomado.")

    def is_running(self):
        return self._is_running
    
    def is_paused(self):
        return self._is_paused
    
    def verificar_manual(self):
        """Executa um ciclo de verificação manual, sem alterar o estado de monitoramento."""
        if self.is_running() and not self.is_paused():
            self.log_message.emit("AVISO: Verificação manual não pode ser executada enquanto o monitoramento automático está ativo.")
            return
        
        self.log_message.emit("Iniciando Verificação Manual...")
        
        # Se o driver não estiver ativo, tenta inicializar apenas para a verificação
        driver_temp = None
        wait_temp = None
        try:
            if not self.driver:
                self.log_message.emit("Inicializando driver temporário para verificação manual...")
                fechar_processos_excel_outlook()
                driver_temp, wait_temp = criar_driver()
                
            driver_usar = self.driver if self.driver else driver_temp
            
            if abrir_excel_online_com_bypass(driver_usar, EXCEL_ONLINE_URL):
                if baixar_planilha_excel_online(driver_usar):
                    if os.path.exists(DOWNLOADED_FILE):
                        processar_envios()
                        self.log_message.emit("Verificação Manual concluída. Processamento de envios executado.")
                    else:
                        self.log_message.emit("AVISO: Arquivo não encontrado após download na Verificação Manual.")
                else:
                    self.log_message.emit("AVISO: Falha ao baixar a planilha na Verificação Manual.")
            else:
                self.log_message.emit("AVISO: Falha ao abrir Excel Online na Verificação Manual.")
                
        except Exception as e:
            self.log_message.emit(f"ERRO durante a Verificação Manual: {e}")
        finally:
            if driver_temp:
                try:
                    driver_temp.quit()
                except Exception:
                    pass
            self.log_message.emit("Verificação Manual finalizada.")

