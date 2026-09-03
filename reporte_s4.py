class ReporteS4:

    def __init__(
        self,
        kibana_client,
        outlook_client,
        logger
    ):

        self.kibana_client = kibana_client
        self.outlook_client = outlook_client
        self.logger = logger

    def ejecutar(self):

        self.logger.escribir_log(
            "Iniciando flujo Reporte S4"
        )

        archivo_pdf = (
            self.kibana_client.ejecutar()
        )

        self.outlook_client.enviar_reporte(
            archivo_pdf
        )

        self.logger.escribir_log(
            "Flujo Reporte S4 finalizado"
        )