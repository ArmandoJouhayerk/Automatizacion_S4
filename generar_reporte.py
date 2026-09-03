import os
from playwright.sync_api import sync_playwright
from config import Config
from logger import Logger
from outlook_client import OutlookClient
from kibana_client import KibanaClient

# configuracion de login y descarga 
USUARIO = Config.USUARIO
PASSWORD = Config.PASSWORD

URL_KIBANA = Config.URL_KIBANA

CARPETA_DESCARGAS = Config.CARPETA_DESCARGAS

ARCHIVO_LOG = Config.ARCHIVO_LOG

os.makedirs(CARPETA_DESCARGAS, exist_ok=True)

# Instancia 
logger = Logger(ARCHIVO_LOG)
outlook_client = OutlookClient(logger) 
kibana_client = KibanaClient(logger, USUARIO, PASSWORD, URL_KIBANA, CARPETA_DESCARGAS)

try:

    logger.escribir_log("Inicio de ejecución")

    with sync_playwright() as p:

        browser = p.chromium.launch(
            channel="msedge",
            headless=False
        )

        context = browser.new_context(
            ignore_https_errors=True,
            accept_downloads=True
        )

        page = context.new_page()

        # login
        kibana_client.login(page)

        # Seleccion de espacio Default
        kibana_client.seleccionar_espacio(page)

        # Abriendo Analytics>Dashboards
        kibana_client.abrir_dashboard(page)

        # Filtro today
        kibana_client.aplicar_filtro_today(page)

        # Abrir Share
        kibana_client.abrir_share(page)

        # Abrir Export
        kibana_client.abrir_export(page)

        archivo_pdf = kibana_client.exportar_pdf(page)

        outlook_client.enviar_reporte(
            archivo_pdf
        )

        logger.escribir_log(
            "Cerrando navegador"
        )

        context.close()
        browser.close()

        logger.escribir_log(
            "Fin de ejecución"
        )

        import sys

        sys.stdout.flush()
        os._exit(0)

except Exception as e:

    logger.escribir_log(
        f"ERROR GENERAL: {str(e)}"
    )

    raise