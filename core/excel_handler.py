#core/excel_handler.py
import os
import time
import logging
import gc
import pythoncom
import win32com.client as win32
from core.alerts import alerta_erro
import pandas as pd
import re
from config.settings import CAMINHO_BASE



def checar_cpf(cpf_motorista, CAMINHO_BASE: str) -> bool:
    try:
        if isinstance(cpf_motorista, pd.Series):
            cpf_motorista = cpf_motorista.iloc[0]  # Pega o primeiro valor se for Series

        cpf_motorista = str(cpf_motorista).strip()
        base = pd.read_excel(CAMINHO_BASE, dtype=str)

        if "CPF do Motorista" not in base.columns:
            raise KeyError("A base não contém uma coluna chamada 'CPF do Motorista'.")

        return cpf_motorista in base["CPF do Motorista"].astype(str).str.strip().values

    except Exception as e:
        print(f"Erro ao verificar CPF na base: {e}")
        return False

def fechar_processos_excel_outlook():
    import subprocess
    try:
        subprocess.run('taskkill /f /im excel.exe', shell=True,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        logging.info("Excel finalizado para liberar arquivo.")
    except Exception as e:
        logging.warning(f"Erro ao finalizar processos Excel: {e}")


def esperar_arquivo_liberar(caminho, tentativas=10, intervalo=2):
    for tentativa in range(tentativas):
        try:
            with open(caminho, 'a'):
                return True
        except PermissionError:
            logging.info(f"Arquivo ainda em uso, aguardando... ({tentativa + 1}/{tentativas})")
            time.sleep(intervalo)
    logging.error(f"Timeout ao aguardar liberação do arquivo: {caminho}")
    return False


def esperar_refresh_concluir(workbook, timeout=60):
    inicio = time.time()
    while True:
        try:
            if not workbook.Refreshing:
                break
        except AttributeError:
            break
        if time.time() - inicio > timeout:
            logging.warning("Timeout ao aguardar o RefreshAll finalizar, prosseguindo mesmo assim.")
            break
        time.sleep(1)


def atualizar_planilha_excel(caminho_arquivo):

    try:
        pythoncom.CoInitialize()

        if not esperar_arquivo_liberar(caminho_arquivo, tentativas=10, intervalo=2):
            logging.error(f"Arquivo bloqueado antes de abrir no Excel COM: {caminho_arquivo}")
            alerta_erro(f"Arquivo bloqueado antes de abrir no Excel: {caminho_arquivo}")
            return

        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(caminho_arquivo, ReadOnly=False)
        wb.RefreshAll()
        logging.info("Atualizando planilha, aguardando 15 segundos...")

        for _ in range(15):
            time.sleep(1)

        caminho_temp = caminho_arquivo.replace(".xlsx", "_temp.xlsx")
        wb.SaveAs(caminho_temp)
        wb.Close(False)
        excel.Quit()
        del wb
        del excel

        # Substitui o arquivo original pelo salvo
        if os.path.exists(caminho_arquivo):
            os.remove(caminho_arquivo)
        os.rename(caminho_temp, caminho_arquivo)

        gc.collect()
        pythoncom.CoUninitialize()
        logging.info("Planilha atualizada e substituída sem bloqueios.")

    except Exception as e:
        logging.error(f"Erro ao atualizar planilha: {e}")
        alerta_erro(f"Erro ao atualizar planilha: {e}")


def padronizar_cpf(cpf: str) -> str:

    if not isinstance(cpf, str):
        return ''

    numeros = re.sub(r'\D', '', cpf)

    return f"{numeros[:3]}.{numeros[3:6]}.{numeros[6:9]}-{numeros[9:]}"


def padronizar_dados(df: pd.DataFrame) -> pd.DataFrame:

    df = df.copy()


    def formatar_cpf(cpf: str) -> str:
        if not isinstance(cpf, str):
            return ''
        numeros = re.sub(r'\D', '', cpf)
        if len(numeros) != 11:
            return ''
        return f"{numeros[:3]}.{numeros[3:6]}.{numeros[6:9]}-{numeros[9:]}"

    # ID
    if 'Id' in df.columns:
        df['Id'] = df['Id'].astype(str).str.strip()

    # Nome do Motorista
    if 'Nome do Motorista' in df.columns:
        df['Nome do Motorista'] = df['Nome do Motorista'].astype(str).str.strip().str.title()

    # CPF do Motorista
    if 'CPF do Motorista' in df.columns:
        df['CPF do Motorista'] = df['CPF do Motorista'].apply(formatar_cpf)

    # Placa do Cavalo
    if 'Placa do Cavalo' in df.columns:
        df['Placa do Cavalo'] = df['Placa do Cavalo'].astype(str).str.replace(r'\s+', '', regex=True).str.upper()

    # Placa da Carreta
    if 'Placa da Carreta' in df.columns:
        df['Placa da Carreta'] = df['Placa da Carreta'].astype(str).str.replace(r'\s+', '', regex=True).str.upper()

    # Ajudantes
    for i in range(1, 6):
        nome_col = f"Nome do Ajudante {i}"
        cpf_col = f"CPF do Ajudante {i}"

        if nome_col in df.columns:
            df[nome_col] = df[nome_col].astype(str).str.strip().str.title().replace({'': ''})
        if cpf_col in df.columns:
            df[cpf_col] = df[cpf_col].apply(padronizar_cpf)

    return df


def atualizar_base_motoristas(df_forms: pd.DataFrame):


    # Garante que a base exista e contenha as colunas necessárias
    if not os.path.exists(CAMINHO_BASE):
        base_df = pd.DataFrame(columns=["Nome do Motorista", "CPF do Motorista", "E-mail do Solicitante"])
    else:
        base_df = pd.read_excel(CAMINHO_BASE, dtype=str)
        for col in ["Nome do Motorista", "CPF do Motorista", "E-mail do Solicitante"]:
            if col not in base_df.columns:
                base_df[col] = ""

    # Padroniza os dados do Forms para comparação
    df_forms = df_forms.copy()
    df_forms["Nome do Novo Motorista"] = df_forms["Nome do Novo Motorista"].astype(str).str.strip().str.title()
    df_forms["CPF do Novo Motorista"] = df_forms["CPF do Novo Motorista"].apply(padronizar_cpf)
    df_forms["Insira seu e-mail"] = df_forms["Insira seu e-mail"].astype(str).str.strip()


    cadastros = df_forms[df_forms["O que você deseja fazer?"].str.lower().str.contains("cadastro")]

    if cadastros.empty:
        return

    for _, linha in cadastros.iterrows():
        nome = linha["Nome do Novo Motorista"]
        cpf = linha["CPF do Novo Motorista"]
        email = linha["Insira seu e-mail"]

        if not cpf or not nome:
            continue


        existe = checar_cpf(cpf, CAMINHO_BASE)

        if existe:
            # Atualiza nome e e-mail onde o CPF for igual
            base_df.loc[
                base_df["CPF do Motorista"].astype(str).str.strip() == cpf,
                ["Nome do Motorista", "E-mail do Solicitante"]
            ] = [nome, email]
        else:
            # Adiciona novo motorista
            nova_linha = pd.DataFrame({
                "Nome do Motorista": [nome],
                "CPF do Motorista": [cpf],
                "E-mail do Solicitante": [email]
            })
            base_df = pd.concat([base_df, nova_linha], ignore_index=True)


    try:
        base_df.to_excel(CAMINHO_BASE, index=False)
        print("✅ Base de motoristas atualizada com sucesso.")
    except Exception as e:
        print(f"⚠️ Erro ao salvar a base de motoristas: {e}")

