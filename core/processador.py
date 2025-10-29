# core/processador.py

import pandas as pd
import logging
from datetime import datetime
import pythoncom
from config.settings import DOWNLOADED_FILE, FILE_AUXILIAR, DESTINATARIOS
from core.utils import mexer_mouse, esperar_arquivo_liberar, obter_saudacao
from core.email_handler import enviar_email, abrir_outlook_minimizado
from core.excel_handler import atualizar_planilha_excel, fechar_processos_excel_outlook, padronizar_dados


def processar_envios():
    pythoncom.CoInitialize()
    try:
        # Fecha Excel/Outlook antes de abrir planilhas
        fechar_processos_excel_outlook()

        # Atualiza planilha principal
        atualizar_planilha_excel(DOWNLOADED_FILE)

        # Espera a planilha liberar
        if not esperar_arquivo_liberar(DOWNLOADED_FILE, tentativas=10, intervalo=1):
            logging.error("Arquivo ainda bloqueado após atualização, pulando leitura.")
            return

        # Carrega planilha oficial
        df_oficial = pd.read_excel(DOWNLOADED_FILE, dtype=str)
        df_oficial = padronizar_dados(df_oficial)

        # Carrega planilha auxiliar
        try:
            df_aux = pd.read_excel(FILE_AUXILIAR, dtype=str)
            df_aux = padronizar_dados(df_aux)
            logging.info("Planilha auxiliar carregada.")
        except FileNotFoundError:
            df_aux = pd.DataFrame(columns=['Id', 'Data de Envio', 'Nome do Motorista','CPF do Motorista', 'Placa do Cavalo', 'Placa da Carreta', 'Nome do Ajudante 1', 'CPF do Ajudante 1', 'Nome do Ajudante 2', 'CPF do Ajudante 2'])
            logging.info("Planilha auxiliar criada, pois não existia.")

        # Identifica novos registros
        novos_registros = df_oficial[~df_oficial['Id'].isin(df_aux['Id'])].sort_values(
            by='Id', key=lambda x: x.astype(int)
        )

        if novos_registros.empty:
            logging.info("Nenhum novo registro para envio.")
            return

        # Pega o primeiro registro
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

        # Adiciona ajudantes
        ajudantes = []
        for i in range(1, 6):
            nome_col = f"Nome do Ajudante {i}"
            cpf_col = f"CPF do Ajudante {i}"
            nome = primeiro.get(nome_col, "")
            cpf_aux = primeiro.get(cpf_col, "")
            if isinstance(nome, str) and nome.strip():
                ajudantes.append((nome.strip(), cpf_aux.strip() if isinstance(cpf_aux, str) else ""))

        if ajudantes:
            if len(ajudantes) == 1:
                nome, cpf_aux = ajudantes[0]
                corpo += f"Ajudante: {nome}\nCPF: {cpf_aux}\n"
            else:
                for idx, (nome, cpf_aux) in enumerate(ajudantes, start=1):
                    corpo += f"Ajudante {idx}: {nome}\nCPF: {cpf_aux}\n\n"

        assunto = "Liberação de Motorista"

        # Garante Outlook aberto
        abrir_outlook_minimizado()

        logging.info(f"Enviando e-mail - ID: {primeiro['Id']} | Nome: {nome_motorista}")
        envio_sucesso = enviar_email(DESTINATARIOS, assunto, corpo)

        if envio_sucesso:
            # Monta dicionário com dados principais
            dados = {
                'Id': primeiro['Id'],
                'Data de Envio': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'Nome do Motorista': nome_motorista,
                'CPF do Motorista': cpf,
                'Placa do Cavalo': placa_cavalo,
                'Placa da Carreta': placa_carreta if placa_carreta else 'Sem carreta',
            }

            # Adiciona ajudantes (apenas até 2)
            for i in range(1, 3):
                nome_col = f"Nome do Ajudante {i}"
                cpf_col = f"CPF do Ajudante {i}"
                nome = primeiro.get(nome_col, "")
                cpf_a = primeiro.get(cpf_col, "")
                dados[f"Nome do Ajudante {i}"] = nome if isinstance(nome, str) else ""
                dados[f"CPF do Ajudante {i}"] = cpf_a if isinstance(cpf_a, str) else ""

            # Adiciona linha na planilha auxiliar
            df_aux = pd.concat([df_aux, pd.DataFrame([dados])], ignore_index=True)
            df_aux.to_excel(FILE_AUXILIAR, index=False)
            logging.info(f"Registro {primeiro['Id']} salvo na auxiliar com dados completos.")

        else:
            logging.warning(f"Erro ao enviar e-mail de {primeiro['Id']}, não salvo na auxiliar para tentar novamente no próximo ciclo.")

        # Mexe o mouse após envio para evitar bloqueio de tela
        mexer_mouse()

    except Exception as e:
        logging.error(f"Erro em processar_envios: {e}")
    finally:
        pythoncom.CoUninitialize()
