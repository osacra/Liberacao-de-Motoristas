import sys
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QTableWidget, QTableWidgetItem, QLineEdit,
    QGroupBox, QTabWidget, QHeaderView, QSizePolicy, QTextEdit, QMessageBox, QComboBox, QFileDialog
)
from PySide6.QtCore import Qt, QSize, QTimer
from PySide6.QtGui import QFont, QPixmap, QColor, QPalette

from core.monitor import MonitoramentoThread
import time
import os
import pandas as pd
from config.settings import FILE_AUXILIAR
from datetime import date
from PySide6.QtGui import QIcon

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
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        self.setWindowIcon(QIcon(os.path.join(base_dir, "assets", "app.ico")))
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

        self._criar_header(layout_principal)
        self._criar_status_sistema(layout_principal)
        self._criar_controles(layout_principal)
        self._criar_abas(layout_principal)

        layout_principal.addStretch(1)

        self.atualizar_tabela_liberacoes()
        self.atualizar_botoes()

    # ==============================================================
    # HEADER
    # ==============================================================

    def _criar_header(self, layout_principal):
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        caminho_logo = os.path.join(base_dir, "assets", "logo.png")

        header_widget = QWidget()
        header_widget.setStyleSheet(f"background-color: {COR_DHL_AMARELO};")
        header_widget.setFixedHeight(70)

        header_layout = QHBoxLayout(header_widget)
        header_layout.setContentsMargins(20, 0, 20, 0)
        header_layout.setSpacing(10)

        logo_label = QLabel()
        pixmap_original = QPixmap(caminho_logo)
        if not pixmap_original.isNull():
            pixmap = pixmap_original.scaledToHeight(50, Qt.SmoothTransformation)
            logo_label.setPixmap(pixmap)

        titulo_label = QLabel("Liberação de Motoristas")
        titulo_label.setFont(QFont("Arial", 18, QFont.Weight.Bold))
        titulo_label.setAlignment(Qt.AlignCenter)

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
        btn_sair.clicked.connect(self.close)


        header_layout.addWidget(logo_label, 0, Qt.AlignVCenter)
        header_layout.addStretch(1)
        header_layout.addWidget(titulo_label, 0, Qt.AlignHCenter | Qt.AlignVCenter)
        header_layout.addStretch(1)
        header_layout.addWidget(btn_sair, 0, Qt.AlignVCenter)

        layout_principal.addWidget(header_widget)

    # ==============================================================
    # STATUS
    # ==============================================================

    def _criar_status_sistema(self, layout_principal):
        status_group = QGroupBox("Status do Sistema")
        status_layout = QHBoxLayout(status_group)
        status_layout.setAlignment(Qt.AlignHCenter)
        status_layout.setSpacing(30)

        self.lbl_conexao_api = QLabel("Desconectado")
        self.lbl_conexao_api.setStyleSheet(f"color: {COR_TEXTO_STATUS_VERMELHO}; font-weight: bold;")

        self.lbl_monitoramento = QLabel("Parado")
        self.lbl_monitoramento.setStyleSheet("font-weight: bold;")

        self.lbl_ultima_verificacao = QLabel("Nunca")
        self.lbl_total_liberacoes = QLabel("0")

        status_layout.addWidget(QLabel("Conexão API:"))
        status_layout.addWidget(self.lbl_conexao_api)
        status_layout.addWidget(QLabel("Monitoramento:"))
        status_layout.addWidget(self.lbl_monitoramento)
        status_layout.addWidget(QLabel("Última Verificação:"))
        status_layout.addWidget(self.lbl_ultima_verificacao)

        # ALTERAÇÃO: Especificando "Total de Liberações (Hoje)"
        status_layout.addWidget(QLabel("Total de Liberações (Hoje):"))
        status_layout.addWidget(self.lbl_total_liberacoes)

        layout_principal.addWidget(status_group)

    # ==============================================================
    # CONTROLES
    # ==============================================================
    def _criar_controles(self, layout_principal):
        controles_group = QGroupBox("Controles")
        layout = QHBoxLayout(controles_group)
        layout.setAlignment(Qt.AlignHCenter)
        layout.setSpacing(20)
        btn_style = f"""
            QPushButton {{
                background-color: {COR_TEXTO_STATUS_VERMELHO};
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 4px;
            }}
            QPushButton:hover {{ background-color: #b2001a; }}
        """
        self.btn_iniciar = QPushButton("Iniciar Monitoramento")
        self.btn_parar = QPushButton("Parar")
        for btn in [self.btn_iniciar, self.btn_parar]:
            btn.setStyleSheet(btn_style)
            layout.addWidget(btn)
        self.btn_iniciar.clicked.connect(self.iniciar_monitoramento)
        self.btn_parar.clicked.connect(self.parar_monitoramento)
        layout_principal.addWidget(controles_group)

    # ==============================================================
    # ABAS
    # ==============================================================
    def _criar_abas(self, layout_principal):
        self.tab_widget = QTabWidget()
        self.tab_widget.setStyleSheet("QTabWidget::pane { border: 1px solid silver; }")

        liberacoes_widget = QWidget()
        self._criar_aba_liberacoes(liberacoes_widget)
        self.tab_widget.addTab(liberacoes_widget, "Liberações")

        logs_widget = QWidget()
        self._criar_aba_logs(logs_widget)
        self.tab_widget.addTab(logs_widget, "Logs")

        config_widget = QWidget()
        self._criar_aba_configuracoes(config_widget)
        self.tab_widget.addTab(config_widget, "Configurações")

        layout_principal.addWidget(self.tab_widget)

    # ==============================================================
    # ABA LIBERAÇÕES
    # ==============================================================
    def _criar_aba_liberacoes(self, parent_widget):
        layout = QVBoxLayout(parent_widget)
        busca_layout = QHBoxLayout()
        busca_layout.addWidget(QLabel("Buscar:"))
        self.txt_busca = QLineEdit()
        self.txt_busca.setPlaceholderText("Digite Nome, CPF ou Placa...")
        self.txt_busca.textChanged.connect(self.filtrar_tabela)
        busca_layout.addWidget(self.txt_busca)
        layout.addLayout(busca_layout)

        self.tabela_liberacoes = QTableWidget()


        colunas_gui = ["ID", "Nome", "CPF", "Placa do Cavalo", "Placa da Carreta", "Data"]

        self.tabela_liberacoes.setColumnCount(len(colunas_gui))
        self.tabela_liberacoes.setHorizontalHeaderLabels(colunas_gui)
        self.tabela_liberacoes.verticalHeader().setVisible(False)
        self.tabela_liberacoes.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tabela_liberacoes.setSelectionBehavior(QTableWidget.SelectRows)
        self.tabela_liberacoes.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.tabela_liberacoes)


        self.tabela_liberacoes.setMinimumHeight(300)
        self.tabela_liberacoes.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

    # ==============================================================
    # ABA LOGS
    # ==============================================================
    def _criar_aba_logs(self, parent_widget):
        layout = QVBoxLayout(parent_widget)
        self.txt_logs = QTextEdit()
        self.txt_logs.setReadOnly(True)
        self.txt_logs.setStyleSheet("background-color: black; color: white; font-family: monospace; padding: 5px;")
        self.txt_logs.setText(f"[{time.strftime('%H:%M:%S')}] Interface inicializada.")
        layout.addWidget(self.txt_logs)

    # ==============================================================
    # ABA CONFIGURAÇÕES
    # ==============================================================

    def _criar_aba_configuracoes(self, parent_widget):
        layout = QVBoxLayout(parent_widget)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)  # espaço mínimo entre widgets
        layout.setSizeConstraint(QVBoxLayout.SetDefaultConstraint)  # não força esticar verticalmente

        import config.settings as settings

        # --- Emails destinatários ---
        lbl_emails = QLabel("Emails destinatários (separados por vírgula):")
        self.txt_emails = QLineEdit()
        self.txt_emails.setText(", ".join(settings.DESTINATARIOS))
        self.txt_emails.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)



        # --- Caminho da planilha auxiliar ---
        lbl_file_auxiliar = QLabel("Caminho da planilha auxiliar:")
        self.txt_file_auxiliar = QLineEdit()
        self.txt_file_auxiliar.setText(settings.FILE_AUXILIAR)
        self.txt_file_auxiliar.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)


        # --- Botão Salvar ---
        btn_salvar = QPushButton("Salvar")
        btn_salvar.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #45A049;
            }
        """)
        btn_salvar.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        btn_salvar.clicked.connect(self._salvar_configuracoes)

        # --- Adicionando widgets ao layout ---
        layout.addWidget(lbl_emails)
        layout.addWidget(self.txt_emails)

        layout.addWidget(lbl_file_auxiliar)
        layout.addWidget(self.txt_file_auxiliar)
        layout.addWidget(btn_salvar)

    def _salvar_configuracoes(self):
        import config.settings as settings
        import os
        from PySide6.QtWidgets import QMessageBox


        novos_emails = [e.strip() for e in self.txt_emails.text().split(",") if e.strip()]

        file_auxiliar = self.txt_file_auxiliar.text().strip()


        if not novos_emails  or not file_auxiliar:
            QMessageBox.warning(self, "Aviso", "Preencha todos os campos antes de salvar.")
            return


        settings.DESTINATARIOS = novos_emails
        settings.FILE_AUXILIAR = file_auxiliar

        arquivo_settings = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "config", "settings.py")

        try:
            with open(arquivo_settings, "r", encoding="utf-8") as f:
                linhas = f.readlines()

            with open(arquivo_settings, "w", encoding="utf-8") as f:
                i = 0
                while i < len(linhas):
                    linha = linhas[i]
                    if linha.strip().startswith("DESTINATARIOS"):
                        # Ignora todas as linhas até encontrar o fechamento da lista
                        while i < len(linhas) and "]" not in linhas[i]:
                            i += 1
                        i += 1

                        f.write("DESTINATARIOS = [\n")
                        for email in novos_emails:
                            f.write(f"    '{email}',\n")
                        f.write("]\n")
                    elif linha.strip().startswith("FILE_AUXILIAR"):
                        f.write(f"FILE_AUXILIAR = r'{file_auxiliar}'\n")
                        i += 1
                    else:
                        f.write(linha)
                        i += 1

            QMessageBox.information(self, "Sucesso", "Configurações salvas com sucesso!")

        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Não foi possível salvar as configurações:\n{e}")

    def _selecionar_arquivo(self, line_edit):
        caminho, _ = QFileDialog.getOpenFileName(self, "Selecionar planilha", "", "Planilhas (*.xlsx *.xls)")
        if caminho:
            line_edit.setText(caminho)

    # ==============================================================
    # LÓGICA PRINCIPAL
    # ==============================================================

    def iniciar_monitoramento(self):
        if not self.monitor_thread.is_running():
            self.monitor_thread.start()
        self.atualizar_botoes()

    def parar_monitoramento(self):
        if self.monitor_thread.is_running():
            self.monitor_thread.stop()
        self.atualizar_botoes()

    def atualizar_status(self, componente, status):
        if componente == "conexao_api":
            self.lbl_conexao_api.setText(status)
            cor = COR_TEXTO_STATUS_VERDE if status == "Conectado" else COR_TEXTO_STATUS_VERMELHO
            self.lbl_conexao_api.setStyleSheet(f"color: {cor}; font-weight: bold;")
        elif componente == "monitoramento":
            self.lbl_monitoramento.setText(status)
            self.atualizar_botoes()
        elif componente == "ultima_verificacao":
            self.lbl_ultima_verificacao.setText(status)
        elif componente == "total_liberacoes":
            self.lbl_total_liberacoes.setText(status)

    def adicionar_log(self, msg):
        self.txt_logs.append(f"[{time.strftime('%H:%M:%S')}] {msg}")

    def atualizar_botoes(self):
        is_running = self.monitor_thread.is_running()
        self.btn_iniciar.setEnabled(not is_running)
        self.btn_parar.setEnabled(is_running)
        self.btn_iniciar.setText("Monitorando..." if is_running else "Iniciar Monitoramento")

    def atualizar_tabela_liberacoes(self):
        if not os.path.exists(FILE_AUXILIAR):
            self.tabela_liberacoes.setRowCount(0)
            self.lbl_total_liberacoes.setText("0")
            return

        try:
            df = pd.read_excel(FILE_AUXILIAR, dtype={'Id': str})
            df['Id'] = df['Id'].astype(str).str.strip()

            colunas_origem = [
                'Id', 'Nome do Motorista', 'CPF do Motorista',
                'Placa do Cavalo', 'Placa da Carreta', 'Data de Envio'
            ]
            for c in colunas_origem:
                if c not in df.columns:
                    df[c] = "N/A"

            # Conversão de datas
            if not df['Data de Envio'].dropna().empty:
                amostra = str(df['Data de Envio'].dropna().astype(str).iloc[0])
                dayfirst = "-" not in amostra[:10]
            else:
                dayfirst = True
            df['Data de Envio'] = pd.to_datetime(df['Data de Envio'], errors='coerce', dayfirst=dayfirst)

            # Renomear colunas
            df = df.rename(columns={
                'Id': 'ID',
                'Nome do Motorista': 'Nome',
                'CPF do Motorista': 'CPF',
                'Placa do Cavalo': 'Placa do Cavalo',
                'Placa da Carreta': 'Placa da Carreta',
                'Data de Envio': 'Data'
            })

            # Filtro de busca
            termo = self.txt_busca.text().strip().lower()
            if termo:
                df = df[df.apply(lambda r: r.astype(str).str.lower().str.contains(termo).any(), axis=1)]

            colunas_visiveis = ["ID", "Nome", "CPF", "Placa do Cavalo", "Placa da Carreta", "Data"]
            df = df[colunas_visiveis]

            self.tabela_liberacoes.setColumnCount(len(colunas_visiveis))
            self.tabela_liberacoes.setHorizontalHeaderLabels(colunas_visiveis)
            self.tabela_liberacoes.verticalHeader().setVisible(False)
            self.tabela_liberacoes.setRowCount(len(df))

            # Total de liberações de hoje
            df_hoje = df[pd.to_datetime(df["Data"]).dt.date == date.today()]
            self.lbl_total_liberacoes.setText(str(len(df_hoje)))

            # Preenchimento da tabela
            for i, (_, row) in enumerate(df.iterrows()):
                for j, col in enumerate(colunas_visiveis):
                    valor = row[col]
                    if col == "Data" and pd.notna(valor):
                        valor = valor.strftime("%d/%m/%Y %H:%M:%S")
                    item = QTableWidgetItem(str(valor))
                    self.tabela_liberacoes.setItem(i, j, item)

        except Exception as e:
            self.adicionar_log(f"ERRO: {e}")
            self.tabela_liberacoes.setRowCount(0)

    def filtrar_tabela(self):
        self.atualizar_tabela_liberacoes()


if __name__ == "__main__":
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps)
    app = QApplication(sys.argv)

    # Paleta de cores
    palette = app.palette()
    palette.setColor(QPalette.Window, QColor(COR_FUNDO_CINZA))
    palette.setColor(QPalette.Base, Qt.white)
    palette.setColor(QPalette.AlternateBase, QColor("#EAEAEA"))
    app.setPalette(palette)

    # Janela principal
    janela = TelaInicial()
    janela.show()

    # Garantir parada da thread ao sair
    app.aboutToQuit.connect(janela.parar_monitoramento)

    sys.exit(app.exec())