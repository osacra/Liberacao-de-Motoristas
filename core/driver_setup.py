# core/driver_setup.py
import os
import logging
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from config import settings


def criar_driver():

    options = Options()

    # Configurações básicas
    if settings.EDGE_OPTIONS.get("headless", False):
        options.add_argument("--headless=new")
    if settings.EDGE_OPTIONS.get("disable_gpu"):
        options.add_argument("--disable-gpu")
    if settings.EDGE_OPTIONS.get("disable_extensions"):
        options.add_argument("--disable-extensions")
    if settings.EDGE_OPTIONS.get("disable_popup_blocking"):
        options.add_argument("--disable-popup-blocking")
    if settings.EDGE_OPTIONS.get("no_sandbox"):
        options.add_argument("--no-sandbox")
    if settings.EDGE_OPTIONS.get("disable_dev_shm_usage"):
        options.add_argument("--disable-dev-shm-usage")

    # --- Ajuste do diretório de download ---
    prefs = settings.EDGE_OPTIONS.get("prefs", {}).copy()
    download_dir = prefs.get("download.default_directory")

    if not download_dir:
        logging.warning("Nenhum diretório de download definido em EDGE_OPTIONS. Usando diretório atual.")
        download_dir = os.getcwd()

    # Garante caminho absoluto e criação do diretório
    download_dir = os.path.abspath(download_dir)
    os.makedirs(download_dir, exist_ok=True)
    prefs["download.default_directory"] = download_dir

    options.add_experimental_option("prefs", prefs)

    logging.info(f"Diretório de download configurado: {download_dir}")
    logging.info(f"EdgeDriver: {settings.EDGE_DRIVER_PATH}")

    # Criação do serviço e do driver
    service = Service(settings.EDGE_DRIVER_PATH)
    driver = webdriver.Edge(service=service, options=options)

    # WebDriverWait configurado
    wait = WebDriverWait(driver, settings.WEBDRIVER_TIMEOUT)

    logging.info("Driver Edge inicializado com sucesso.")
    return driver, wait
