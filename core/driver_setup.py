# core/driver_setup.py
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from config import settings

def criar_driver():
    options = Options()
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

    prefs = settings.EDGE_OPTIONS.get("prefs", {})
    options.add_experimental_option("prefs", prefs)

    service = Service(settings.EDGE_DRIVER_PATH)
    driver = webdriver.Edge(service=service, options=options)
    wait = WebDriverWait(driver, settings.WEBDRIVER_TIMEOUT)

    return driver, wait
