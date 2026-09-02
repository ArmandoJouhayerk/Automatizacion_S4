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

