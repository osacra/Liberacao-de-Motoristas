import pandas as pd
import logging
from datetime import datetime
import pythoncom
from config.settings import DOWNLOADED_FILE, FILE_AUXILIAR, DESTINATARIOS, CAMINHO_BASE
from core.utils import mexer_mouse, esperar_arquivo_liberar, obter_saudacao
from core.email_handler import enviar_email, abrir_outlook_minimizado
from core.excel_handler import atualizar_planilha_excel, fechar_processos_excel_outlook, padronizar_dados
from core.excel_handler import atualizar_base_motoristas, checar_cpf


def motorista_nao_autorizado(nome_motorista: str, cpf_motorista: str):
    try:
        saudacao = obter_saudacao()
        assunto = "Motorista não autorizado na base"
        corpo = (
            f"{saudacao}, portaria!\n\n"
            "O motorista abaixo NÃO foi liberado pois o CPF informado não consta na base de autorizados.\n\n"
            f"Nome: {nome_motorista}\n"
            f"CPF: {cpf_motorista}\n\n"
            "Favor verificar se o cadastro está atualizado ou se há necessidade de inclusão na base."
        )

        abrir_outlook_minimizado()
        sucesso = enviar_email(DESTINATARIOS, assunto, corpo)

        if sucesso:
            logging.info(f"E-mail de motorista não liberado enviado para {DESTINATARIOS}: {nome_motorista} ({cpf_motorista})")
        else:
            logging.warning(f"Falha ao enviar e-mail de motorista não liberado: {nome_motorista} ({cpf_motorista})")

    except Exception as e:
        logging.error(f"Erro ao enviar notificação de motorista não liberado: {e}")


def processar_envios():
    pythoncom.CoInitialize()
    try:
        fechar_processos_excel_outlook()
        atualizar_planilha_excel(DOWNLOADED_FILE)

        if not esperar_arquivo_liberar(DOWNLOADED_FILE, tentativas=10, intervalo=1):
            logging.error("Arquivo ainda bloqueado após atualização, pulando leitura.")
            return

        df_oficial = pd.read_excel(DOWNLOADED_FILE, dtype=str)
        df_oficial = padronizar_dados(df_oficial)

        try:
            df_aux = pd.read_excel(FILE_AUXILIAR, dtype=str)
            df_aux = padronizar_dados(df_aux)
            if 'Status' not in df_aux.columns:
                df_aux['Status'] = ''
            logging.info("Planilha auxiliar carregada.")
        except FileNotFoundError:
            df_aux = pd.DataFrame(columns=[
                'Id', 'Data de Envio', 'Nome do Motorista', 'CPF do Motorista',
                'Placa do Cavalo', 'Placa da Carreta',
                'Nome do Ajudante 1', 'CPF do Ajudante 1',
                'Nome do Ajudante 2', 'CPF do Ajudante 2',
                'Status'
            ])
            logging.info("Planilha auxiliar criada, pois não existia.")

        #  CASO 1 — Cadastro de Motorista
        if "O que você deseja fazer?" in df_oficial.columns:
            novos_cadastros = df_oficial[
                (df_oficial["O que você deseja fazer?"].str.strip().str.lower() == "cadastro de motorista")
                & (~df_oficial["Id"].isin(df_aux["Id"]))
            ]

            if not novos_cadastros.empty:
                logging.info(f"{len(novos_cadastros)} novos cadastros encontrados.")

                for _, linha in novos_cadastros.iterrows():
                    nome = str(linha.get('Nome do Novo Motorista', '')).strip().title()
                    cpf = str(linha.get('CPF do Novo Motorista', '')).strip()

                    if not cpf or not nome:
                        logging.warning("Cadastro ignorado por falta de nome ou CPF.")
                        continue

                    existe = checar_cpf(cpf, CAMINHO_BASE)
                    atualizar_base_motoristas(pd.DataFrame([linha]))

                    status = "Cadastro Atualizado" if existe else "Novo Cadastro"

                    dados = {
                        'Id': str(linha['Id']).strip(),
                        'Data de Envio': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        'Nome do Motorista': nome,
                        'CPF do Motorista': cpf,
                        'Placa do Cavalo': '',
                        'Placa da Carreta': '',
                        'Nome do Ajudante 1': '',
                        'CPF do Ajudante 1': '',
                        'Nome do Ajudante 2': '',
                        'CPF do Ajudante 2': '',
                        'Status': status
                    }

                    dados_df = pd.DataFrame([dados])
                    dados_df = padronizar_dados(dados_df)  # 🔹 garante padronização antes de salvar
                    df_aux = pd.concat([df_aux, dados_df], ignore_index=True)
                    logging.info(f"{status} processado para {nome} ({cpf}).")

                df_aux.to_excel(FILE_AUXILIAR, index=False)
                logging.info("Cadastros processados e salvos na planilha auxiliar.")
                return  # encerra aqui, não envia e-mail

        #  CASO 2 — Liberação de Motorista
        novos_registros = df_oficial[
            (df_oficial["O que você deseja fazer?"].str.strip().str.lower() == "liberação de motorista")
            & (~df_oficial['Id'].isin(df_aux['Id']))
        ].sort_values(by='Id', key=lambda x: x.astype(int))

        if novos_registros.empty:
            logging.info("Nenhum novo registro para envio.")
            return

        primeiro = novos_registros.iloc[0]
        nome_motorista = str(primeiro.at['Nome do Motorista']).strip()
        cpf = str(primeiro.at['CPF do Motorista']).strip()
        placa_cavalo = str(primeiro.at['Placa do Cavalo']).strip() or 'Sem Placa Cavalo'
        saudacao = obter_saudacao()

        texto_carreta = ""
        if 'O veículo possui carreta?' in primeiro and str(primeiro['O veículo possui carreta?']).strip().lower() == "sim":
            placa_carreta = str(primeiro.get('Placa da Carreta', '')).strip()
            if placa_carreta and placa_carreta.lower() not in ['nan', 'none', '']:
                texto_carreta = f"Placa da carreta: {placa_carreta}\n"
        else:
            placa_carreta = ""

        # Verifica se o motorista está na base
        if not checar_cpf(cpf, CAMINHO_BASE):
            logging.warning(f"Motorista {nome_motorista} ({cpf}) não encontrado na base — liberação bloqueada.")
            motorista_nao_autorizado(nome_motorista, cpf)
            status_registro = 'Não Liberado'
        else:
            corpo = (
                f"{saudacao}, portaria!\n\n"
                "Segue abaixo a liberação do motorista.\n\n"
                f"Nome: {nome_motorista}\n"
                f"CPF: {cpf}\n"
                f"Placa do cavalo: {placa_cavalo}\n"
                f"{texto_carreta}"
            )

            ajudantes = []
            for i in range(1, 6):
                nome_col = f"Nome do Ajudante {i}"
                cpf_col = f"CPF do Ajudante {i}"
                nome_a = str(primeiro.get(nome_col, '')).strip()
                cpf_a = str(primeiro.get(cpf_col, '')).strip()
                if nome_a and nome_a.lower() not in ['nan', 'none', '']:
                    ajudantes.append((nome_a, cpf_a))

            if ajudantes:
                corpo += "\n"
                for idx, (nome_a, cpf_a) in enumerate(ajudantes, start=1):
                    corpo += f"Ajudante {idx}: {nome_a}"
                    if cpf_a:
                        corpo += f"\nCPF: {cpf_a}"
                    corpo += "\n"

            abrir_outlook_minimizado()
            envio_sucesso = enviar_email(DESTINATARIOS, "Liberação de Motorista", corpo)
            status_registro = 'Liberado' if envio_sucesso else 'Não Liberado'


        dados = {
            'Id': str(primeiro.at['Id']).strip(),
            'Data de Envio': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'Nome do Motorista': nome_motorista,
            'CPF do Motorista': cpf,
            'Placa do Cavalo': placa_cavalo,
            'Placa da Carreta': placa_carreta if placa_carreta else 'Sem carreta',
            'Status': status_registro
        }

        for i in range(1, 3):
            nome_col = f"Nome do Ajudante {i}"
            cpf_col = f"CPF do Ajudante {i}"
            dados[nome_col] = str(primeiro.at.get(nome_col, '')).strip()
            dados[cpf_col] = str(primeiro.at.get(cpf_col, '')).strip()

        dados_df = pd.DataFrame([dados])
        dados_df = padronizar_dados(dados_df)  # 🔹 garante padronização
        df_aux = pd.concat([df_aux, dados_df], ignore_index=True)
        df_aux.to_excel(FILE_AUXILIAR, index=False)
        logging.info(f"Registro {primeiro.at['Id']} salvo na auxiliar com status: {status_registro}.")

        mexer_mouse()

    except Exception as e:
        logging.error(f"Erro em processar_envios: {e}")
    finally:
        pythoncom.CoUninitialize()
