import time
import win32com.client
from datetime import datetime


class OutlookClient:

    def __init__(self, logger):

        self.logger = logger

    # Función para enviar el reporte por correo electrónico
    def enviar_reporte(
        self,
        archivo_pdf
    ):

        self.logger.escribir_log(
            "Iniciando Outlook"
        )

        outlook = win32com.client.Dispatch(
            "Outlook.Application"
        )

        self.logger.escribir_log(
            "Outlook iniciado"
        )

        correo = None

        for intento in range(5):

            try:

                correo = outlook.CreateItem(0)

                self.logger.escribir_log(
                    f"Correo creado en intento {intento + 1}"
                )

                break

            except Exception as error:

                self.logger.escribir_log(
                    f"CreateItem falló intento {intento + 1}: {error}"
                )

                time.sleep(5)

        if correo is None:

            raise Exception(
                "No fue posible crear el correo en Outlook"
            )

        # Remplazar por las direcciones de correo deseadas
        correo.To = (
            "tu_correo@ejemplo.com"
        )

        # Remplazar por las direcciones de correo deseadas
        correo.CC = (
            "tu_correo@ejemplo.com"
        )

        correo.Subject = (
            "S4 SERVICIO ADMINISTRADO DE CONECTIVIDAD v4 - "
            + datetime.now().strftime("%d-%m-%Y")
        )

        correo.Body = """
    Buenos días, equipo:

    Se comparte el reporte S4 Servicio Administrado de Conectividad, el cual presenta una operación estable y sin incidencias al    momento de la validación.

    Quedo atento a cualquier comentario.

    Saludos.
"""

        correo.Attachments.Add(
            archivo_pdf
        )

        self.logger.escribir_log(
            "Reporte agregado al correo"
        )

        correo.Send()

        self.logger.escribir_log(
            "Correo enviado correctamente"
        )

        correo = None
        outlook = None
