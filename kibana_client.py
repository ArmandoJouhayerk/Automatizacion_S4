from playwright.sync_api import sync_playwright
from datetime import datetime
import os


class KibanaClient:

    def __init__(
        self,
        logger,
        usuario,
        password,
        url_kibana,
        carpeta_descargas
    ):

        self.logger = logger
        self.usuario = usuario
        self.password = password
        self.url_kibana = url_kibana
        self.carpeta_descargas = carpeta_descargas

    def obtener_ruta_pdf(self):

        fecha_archivo = datetime.now().strftime(
            "%d-%m-%Y_%H-%M"
        )

        return os.path.join(
            self.carpeta_descargas,
            f"S4 SERVICIO ADMINISTRADO DE CONECTIVIDAD v4_{fecha_archivo}.pdf"
        )
    def login(self,page):

        self.logger.escribir_log(
            "Abriendo Kibana"
        )

        page.goto(
            self.url_kibana
        )

        page.wait_for_timeout(3000)

        self.logger.escribir_log(
            "Ingresando credenciales"
        )

        page.locator(
            "input[type='text']"
        ).fill(
            self.usuario
        )

        page.locator(
            "input[type='password']"
        ).fill(
            self.password
        )

        page.locator(
            "button"
        ).last.click()

        page.wait_for_timeout(5000)