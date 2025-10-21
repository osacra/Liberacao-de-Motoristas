import sys
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QTableWidget, QTableWidgetItem, QLineEdit,
    QGroupBox, QTabWidget, QHeaderView, QSizePolicy, QTextEdit, QMessageBox
)
from PySide6.QtCore import Qt, QSize, QTimer
from PySide6.QtGui import QFont, QPixmap, QColor, QPalette

from core.monitor import MonitoramentoThread
import time
import os
import pandas as pd
from config.settings import FILE_AUXILIAR

# --- Constantes de Estilo ---
COR_DHL_AMARELO = "#fecb12"
COR_FUNDO_CINZA = "#F0F0F0"
COR_TEXTO_STATUS_VERMELHO = "#E4002B"
COR_TEXTO_STATUS_VERDE = "#008000"
COR_BOTAO_VERDE = "#4CAF50"
COR_BOTAO_VERDE_HOVER = "#45A049"


class TelaInicial(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Sistema de Liberação de Motoristas v2.0 Desktop")
        self.resize(1000, 700)
        self.setStyleSheet(f"background-color: {COR_FUNDO_CINZA};")

        self.monitor_thread = MonitoramentoThread()
        self.monitor_thread.status_changed.connect(self.atualizar_status)
        self.monitor_thread.log_message.connect(self.adicionar_log)

        self.timer_atualizacao_tabela = QTimer(self)
        self.timer_atualizacao_tabela.timeout.connect(self.atualizar_tabela_liberacoes)
        self.timer_atualizacao_tabela.start(5000)

        # --- Widget central ---
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout_principal = QVBoxLayout(central_widget)
        layout_principal.setContentsMargins(0, 0, 0, 0)
        layout_principal.setSpacing(5)

        # 1. Header (Barra Vermelha)
        self._criar_header(layout_principal)

        # 2. Status do Sistema
        self._criar_status_sistema(layout_principal)

        # 3. Controles
        self._criar_controles(layout_principal)

        # 4. Abas (Liberações, Logs, Configurações)
        self._criar_abas(layout_principal)

        # Espaçador para empurrar tudo para cima
        layout_principal.addStretch(1)

        # Inicializa a tabela com os dados existentes
        self.atualizar_tabela_liberacoes()
        self.atualizar_botoes()

    def _criar_header(self, layout_principal):
        import os
        from PySide6.QtWidgets import QWidget, QLabel, QPushButton, QHBoxLayout, QSizePolicy
        from PySide6.QtGui import QPixmap, QFont
        from PySide6.QtCore import Qt

        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        caminho_logo = os.path.join(base_dir, "assets", "logo.png")

        # Header principal
        header_widget = QWidget()
        header_widget.setStyleSheet(f"background-color: {COR_DHL_AMARELO};")
        header_widget.setFixedHeight(70)

        # Layout horizontal do header
        header_layout = QHBoxLayout(header_widget)
        header_layout.setContentsMargins(20, 0, 20, 0)
        header_layout.setSpacing(10)

        # --- Logo ---
        logo_label = QLabel()
        pixmap_original = QPixmap(caminho_logo)
        if not pixmap_original.isNull():
            altura_logo = 50
            pixmap = pixmap_original.scaledToHeight(altura_logo, Qt.SmoothTransformation)
            logo_label.setPixmap(pixmap)
        logo_label.setAlignment(Qt.AlignVCenter | Qt.AlignLeft)
        logo_label.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)

        # --- Botão Sair ---
        btn_sair = QPushButton("Sair")
        btn_sair.setStyleSheet("""
            QPushButton {
                background-color: #E4002B;
                color: white;
                border: none;
                padding: 8px 20px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #b2001a;
            }
        """)
        btn_sair.setFixedHeight(35)
        btn_sair.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        btn_sair.clicked.connect(self.close)

        # --- Título centralizado ---

        titulo_label = QLabel("Liberação de Motoristas", header_widget)
        titulo_font = QFont("Arial", 18, QFont.Weight.Bold)
        titulo_label.setFont(titulo_font)
        titulo_label.setAlignment(Qt.AlignHCenter)
        titulo_label.setStyleSheet("color: black;")
        titulo_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        # --- Montagem do layout ---

        header_layout.addWidget(logo_label, 0, Qt.AlignVCenter | Qt.AlignLeft)
        header_layout.addWidget(titulo_label, 1, Qt.AlignVCenter | Qt.AlignHCenter)
        header_layout.addWidget(btn_sair, 0, Qt.AlignVCenter | Qt.AlignRight)
        layout_principal.addWidget(header_widget)

    def _criar_status_sistema(self, layout_principal):

        status_group = QGroupBox("Status do Sistema")
        status_group.setStyleSheet(
            "QGroupBox { border: 1px solid silver; margin-top: 6px; } "
            "QGroupBox::title { subcontrol-origin: margin; left: 7px; padding: 0 5px 0 5px; }"
        )

        status_layout = QHBoxLayout(status_group)
        status_layout.setContentsMargins(10, 15, 10, 10)
        status_layout.setSpacing(30)

        # Centraliza horizontalmente todo o layout
        status_layout.setAlignment(Qt.AlignHCenter | Qt.AlignVCenter)

        # Conexão API
        lbl_api_texto = QLabel("Conexão API:")
        self.lbl_conexao_api = QLabel("Desconectado")
        self.lbl_conexao_api.setStyleSheet(f"color: {COR_TEXTO_STATUS_VERMELHO}; font-weight: bold;")
        status_layout.addWidget(lbl_api_texto)
        status_layout.addWidget(self.lbl_conexao_api)

        # Monitoramento
        lbl_monitor_texto = QLabel("Monitoramento:")
        self.lbl_monitoramento = QLabel("Parado")
        self.lbl_monitoramento.setStyleSheet("font-weight: bold;")
        status_layout.addWidget(lbl_monitor_texto)
        status_layout.addWidget(self.lbl_monitoramento)

        # Última Verificação
        lbl_ultima_texto = QLabel("Última Verificação:")
        self.lbl_ultima_verificacao = QLabel("Nunca")
        status_layout.addWidget(lbl_ultima_texto)
        status_layout.addWidget(self.lbl_ultima_verificacao)

        # Total de Liberações
        lbl_total_texto = QLabel("Total de Liberações:")
        self.lbl_total_liberacoes = QLabel("0")
        status_layout.addWidget(lbl_total_texto)
        status_layout.addWidget(self.lbl_total_liberacoes)

        layout_principal.addWidget(status_group)

    def _criar_controles(self, layout_principal):
        controles_group = QGroupBox("Controles")
        controles_group.setStyleSheet(
            "QGroupBox { border: 1px solid silver; margin-top: 6px; } "
            "QGroupBox::title { subcontrol-origin: margin; left: 7px; padding: 0 5px 0 5px; }"
        )

        controles_layout = QHBoxLayout(controles_group)
        controles_layout.setContentsMargins(10, 15, 10, 10)
        controles_layout.setSpacing(20)

        # Estilo dos botões de controle
        btn_style = f"""
            QPushButton {{
                background-color: {COR_TEXTO_STATUS_VERMELHO};
                color: white;
                border: none;
                padding: 10px 20px;
                font-size: 14px;
                border-radius: 4px;
            }}
            QPushButton:hover {{
                background-color: #b2001a;
            }}
        """

        # Criação dos botões
        self.btn_iniciar = QPushButton("Iniciar Monitoramento")
        self.btn_pausar = QPushButton("Pausar")
        self.btn_parar = QPushButton("Parar")
        self.btn_verificacao_manual = QPushButton("Verificação Manual")

        # Aplica estilo e tamanho mínimo
        for btn in [self.btn_iniciar, self.btn_pausar, self.btn_parar, self.btn_verificacao_manual]:
            btn.setStyleSheet(btn_style)
            btn.setMinimumWidth(150)
            controles_layout.addWidget(btn)

        # Centraliza todos os botões horizontalmente
        controles_layout.setAlignment(Qt.AlignHCenter)

        layout_principal.addWidget(controles_group)

        # Conectar botões
        self.btn_iniciar.clicked.connect(self.iniciar_monitoramento)
        self.btn_pausar.clicked.connect(self.pausar_monitoramento)
        self.btn_parar.clicked.connect(self.parar_monitoramento)
        self.btn_verificacao_manual.clicked.connect(self.verificacao_manual)

    def _criar_abas(self, layout_principal):
        self.tab_widget = QTabWidget()
        self.tab_widget.setStyleSheet("QTabWidget::pane { border: 1px solid silver; }")

        # Aba Liberações
        liberacoes_widget = QWidget()
        self._criar_aba_liberacoes(liberacoes_widget)
        self.tab_widget.addTab(liberacoes_widget, "Liberações")

        # Aba Logs
        logs_widget = QWidget()
        self._criar_aba_logs(logs_widget)
        self.tab_widget.addTab(logs_widget, "Logs")

        # Aba Configurações
        configuracoes_widget = QWidget()
        self.tab_widget.addTab(configuracoes_widget, "Configurações")

        layout_principal.addWidget(self.tab_widget)

    def _criar_aba_liberacoes(self, parent_widget):
        layout = QVBoxLayout(parent_widget)
        layout.setContentsMargins(5, 5, 5, 5)

        # Layout de Busca
        busca_layout = QHBoxLayout()
        busca_layout.addWidget(QLabel("Buscar:"))
        self.txt_busca = QLineEdit()
        self.txt_busca.setPlaceholderText("Digite ID, Nome, CPF ou Placa...")
        self.txt_busca.textChanged.connect(self.filtrar_tabela)
        busca_layout.addWidget(self.txt_busca)
        busca_layout.addStretch(1)
        layout.addLayout(busca_layout)

        # Tabela de Liberações
        self.tabela_liberacoes = QTableWidget()
        colunas = ["ID", "Nome", "CPF", "Placa", "Data", "Status"]
        self.tabela_liberacoes.setColumnCount(len(colunas))
        self.tabela_liberacoes.setHorizontalHeaderLabels(colunas)
        self.tabela_liberacoes.verticalHeader().setVisible(False)
        self.tabela_liberacoes.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tabela_liberacoes.setSelectionBehavior(QTableWidget.SelectRows)
        self.tabela_liberacoes.setAlternatingRowColors(True)

        # Ajustar o tamanho das colunas
        header = self.tabela_liberacoes.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)  # ID
        header.setSectionResizeMode(6, QHeaderView.ResizeMode.ResizeToContents)  # Status

        layout.addWidget(self.tabela_liberacoes)

    def _criar_aba_logs(self, parent_widget):
        layout = QVBoxLayout(parent_widget)
        self.txt_logs = QTextEdit()
        self.txt_logs.setReadOnly(True)
        self.txt_logs.setStyleSheet("background-color: black; color: white; font-family: monospace; padding: 5px;")
        self.txt_logs.setText(f"[{time.strftime('%H:%M:%S')}] Interface Gráfica Inicializada.")
        layout.addWidget(self.txt_logs)

    # --- Métodos de Controle ---

    def iniciar_monitoramento(self):
        if not self.monitor_thread.is_running():
            self.monitor_thread.start()
        elif self.monitor_thread.is_paused():
            self.monitor_thread.resume()
        self.atualizar_botoes()

    def pausar_monitoramento(self):
        if self.monitor_thread.is_running() and not self.monitor_thread.is_paused():
            self.monitor_thread.pause()
        self.atualizar_botoes()

    def parar_monitoramento(self):
        if self.monitor_thread.is_running():
            self.monitor_thread.stop()

        self.atualizar_botoes()

    def verificacao_manual(self):
        if self.monitor_thread.is_running() and not self.monitor_thread.is_paused():
            QMessageBox.warning(self, "Aviso",
                                "A Verificação Manual não pode ser executada enquanto o Monitoramento Automático estiver ativo.")
            return

        # Executa a verificação manual em uma thread separada para não travar a GUI
        manual_thread = MonitoramentoThread()
        manual_thread.status_changed.connect(self.atualizar_status)
        manual_thread.log_message.connect(self.adicionar_log)
        manual_thread.verificar_manual()

    def parar_monitoramento_thread(self):
        """Chamado ao fechar a aplicação para garantir que a thread seja encerrada."""
        if self.monitor_thread.is_running():
            self.monitor_thread.stop()
            self.monitor_thread.wait()

    # --- Métodos de Atualização da GUI ---

    def atualizar_status(self, componente, status):
        if componente == "conexao_api":
            self.lbl_conexao_api.setText(status)
            if status == "Conectado":
                self.lbl_conexao_api.setStyleSheet(f"color: {COR_TEXTO_STATUS_VERDE}; font-weight: bold;")
            else:
                self.lbl_conexao_api.setStyleSheet(f"color: {COR_TEXTO_STATUS_VERMELHO}; font-weight: bold;")

        elif componente == "monitoramento":
            self.lbl_monitoramento.setText(status)
            self.atualizar_botoes()

        elif componente == "ultima_verificacao":
            self.lbl_ultima_verificacao.setText(status)

        elif componente == "total_liberacoes":
            self.lbl_total_liberacoes.setText(status)

    def adicionar_log(self, mensagem):
        self.txt_logs.append(f"[{time.strftime('%H:%M:%S')}] {mensagem}")

    def atualizar_botoes(self):
        is_running = self.monitor_thread.is_running()
        is_paused = self.monitor_thread.is_paused()

        self.btn_iniciar.setEnabled(not is_running or is_paused)
        self.btn_pausar.setEnabled(is_running and not is_paused)
        self.btn_parar.setEnabled(is_running)
        self.btn_verificacao_manual.setEnabled(not is_running or is_paused)

        if is_running and not is_paused:
            self.btn_iniciar.setText("Monitorando...")
        elif is_paused:
            self.btn_iniciar.setText("Retomar Monitoramento")
        else:
            self.btn_iniciar.setText("Iniciar Monitoramento")

    def atualizar_tabela_liberacoes(self):
        if not os.path.exists(FILE_AUXILIAR):
            self.tabela_liberacoes.setRowCount(0)
            self.lbl_total_liberacoes.setText("0")
            return

        try:
            # Carrega a planilha auxiliar
            df = pd.read_excel(FILE_AUXILIAR, dtype={'Id': str})
            df['Id'] = df['Id'].astype(str).str.strip()

           
            colunas_necessarias = ['Id', 'Nome do Motorista', 'CPF do Motorista',
                                   'Placa do Cavalo', 'Placa da Carreta',
                                   'Data de Envio', 'Status']
            for col in colunas_necessarias:
                if col not in df.columns:
                    df[col] = "N/A"

            # Renomeia colunas para exibir na GUI
            df = df.rename(columns={
                'Id': 'ID',
                'Nome do Motorista': 'Nome',
                'CPF do Motorista': 'CPF',
                'Placa do Cavalo': 'Placa',
                'Data de Envio': 'Data',
                'Placa da Carreta': 'Placa da Carreta',
                'Status': 'Status'
            })

            # Reordena colunas para a tabela
            df = df[['ID', 'Nome', 'CPF', 'Placa', 'Placa da Carreta', 'Data', 'Status']]

            # Aplicar filtro de busca
            termo_busca = self.txt_busca.text().strip().lower()
            if termo_busca:
                df = df[df.apply(lambda row: row.astype(str).str.lower().str.contains(termo_busca).any(), axis=1)]

            # Atualiza tabela
            self.tabela_liberacoes.setRowCount(len(df))
            self.lbl_total_liberacoes.setText(str(len(df)))

            for i, row in df.iterrows():
                for j, col_name in enumerate(df.columns):
                    valor = str(row[col_name])
                    item = QTableWidgetItem(valor)

                    # Estilo para a coluna Status
                    if col_name == "Status":
                        if valor.lower() == "liberado":
                            item.setForeground(QColor(COR_TEXTO_STATUS_VERDE))
                        elif valor.lower() == "pendente":
                            item.setForeground(QColor("#FFA500"))
                        elif valor.lower() == "erro":
                            item.setForeground(QColor(COR_TEXTO_STATUS_VERMELHO))

                    self.tabela_liberacoes.setItem(i, j, item)

        except Exception as e:
            self.adicionar_log(f"ERRO ao ler planilha auxiliar: {e}")
            self.tabela_liberacoes.setRowCount(0)

    def filtrar_tabela(self):
        # A filtragem é feita chamando a função de atualização da tabela
        self.atualizar_tabela_liberacoes()

    def _mostrar_mensagem(self, texto):
        # Função de placeholder para mostrar que os botões estão conectados
        QMessageBox.information(self, "Ação", texto)


if __name__ == "__main__":
    # Configuração para alta DPI em sistemas compatíveis
    QApplication.setAttribute(Qt.ApplicationAttribute.AA_EnableHighDpiScaling)
    QApplication.setAttribute(Qt.ApplicationAttribute.AA_UseHighDpiPixmaps)

    app = QApplication(sys.argv)

    # Configuração de paleta para um visual mais "desktop"
    palette = app.palette()
    palette.setColor(QPalette.Window, QColor(COR_FUNDO_CINZA))
    palette.setColor(QPalette.Base, Qt.white)
    palette.setColor(QPalette.AlternateBase, QColor("#EAEAEA"))
    app.setPalette(palette)

    janela = TelaInicial()
    janela.show()

    # Conecta o sinal aboutToQuit para garantir que a thread de monitoramento seja parada
    app.aboutToQuit.connect(janela.parar_monitoramento_thread)

    sys.exit(app.exec())

