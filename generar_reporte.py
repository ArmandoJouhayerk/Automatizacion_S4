from playwright.sync_api import sync_playwright
from datetime import datetime
import os
import time
import win32com.client

# inicio de refactorizacion a POO

# configuracion de login y descarga
USUARIO = "monitor"
PASSWORD = "CONTRASEÑA"

URL_KIBANA = "https://172.16.17.55:5601"

CARPETA_DESCARGAS = (
    r"C:\Users\jose.mendez\OneDrive - Scontinuidad Latam SA de CV\Documentos\S4Automation\Reportes_S4"
)

ARCHIVO_LOG = (
    r"C:\Users\jose.mendez\OneDrive - Scontinuidad Latam SA de CV\Documentos\S4Automation\Reporte_S4.log"
)

os.makedirs(CARPETA_DESCARGAS, exist_ok=True)


def escribir_log(mensaje):
    with open(
        ARCHIVO_LOG,
        "a",
        encoding="utf-8"
    ) as log:
        fecha_log = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
        log.write(f"[{fecha_log}] {mensaje}\n")


try:

    escribir_log("Inicio de ejecución")

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
        escribir_log("Abriendo Kibana")
        page.goto(URL_KIBANA)

        page.wait_for_timeout(3000)

        escribir_log("Ingresando credenciales")
        page.locator("input[type='text']").fill(USUARIO)
        page.locator("input[type='password']").fill(PASSWORD)
        page.locator("button").last.click()

        # seleccion de space
        page.wait_for_timeout(5000)

        escribir_log("Seleccionando espacio Default")
        page.get_by_role(
            "link",
            name="Default"
        ).click()

        # Ir a dashboards
        page.wait_for_timeout(5000)

        escribir_log("Abriendo lista de Dashboards")

        page.goto(
            "https://172.16.17.55:5601/app/dashboards#/list",
            wait_until="networkidle"
        )

        page.wait_for_timeout(5000)

        escribir_log("Abriendo Dashboard S4")

        page.get_by_text(
            "S4 SERVICIO ADMINISTRADO DE CONECTIVIDAD",
            exact=False
        ).first.click()

        page.wait_for_timeout(10000)

        # Filtro today
        escribir_log("Aplicando filtro Today")

        page.locator(
            "[aria-label='Date quick select']"
        ).click()

        page.wait_for_timeout(2000)

        page.get_by_text(
            "Today",
            exact=True
        ).click()

        page.wait_for_timeout(10000)

        # abri share
        escribir_log("Abriendo Share")

        page.get_by_text(
            "Share",
            exact=True
        ).click()

        page.wait_for_timeout(2000)

        # Exportar
        escribir_log("Abriendo Export")

        page.get_by_text(
            "Export",
            exact=True
        ).click()

        page.wait_for_timeout(2000)

        # Generar PDF
        escribir_log("Solicitando PDF")

        page.get_by_text(
            "Export file",
            exact=True
        ).click()

        escribir_log("Esperando generación del PDF")

        page.get_by_text(
            "Download report",
            exact=True
        ).wait_for(timeout=300000)

        escribir_log("Reporte listo")

        enlace = page.locator(
            "a:has-text('Download report')"
        )

        with page.expect_download(timeout=300000) as download_info:
            enlace.click()

        descarga = download_info.value

        fecha_archivo = datetime.now().strftime(
            "%d-%m-%Y_%H-%M"
        )

        # descarga y nombramiento de reporte
        archivo_pdf = os.path.join(
            CARPETA_DESCARGAS,
            f"S4 SERVICIO ADMINISTRADO DE CONECTIVIDAD v4_{fecha_archivo}.pdf"
        )

        descarga.save_as(archivo_pdf)

        escribir_log(
            f"PDF guardado: {archivo_pdf}"
        )

        escribir_log(
            f"Tamaño PDF: {os.path.getsize(archivo_pdf)} bytes"
        )

        # Envio de correo con el reporte adjunto
        escribir_log("Iniciando Outlook")

        outlook = win32com.client.Dispatch(
            "Outlook.Application"
        )

        escribir_log("Outlook iniciado")

        correo = None

        for intento in range(5):

            try:

                correo = outlook.CreateItem(0)

                escribir_log(
                    f"Correo creado en intento {intento + 1}"
                )

                break

            except Exception as error:

                escribir_log(
                    f"CreateItem falló intento {intento + 1}: {error}"
                )

                time.sleep(5)

        if correo is None:

            raise Exception(
                "No fue posible crear el correo en Outlook"
            )

        correo.To = (
            "ejemplo@mail.com;"
        )

        correo.CC = (
            "ejemplo@mail.com"

        )

        correo.Subject = (
            "S4 SERVICIO ADMINISTRADO DE CONECTIVIDAD v4 - "
            + datetime.now().strftime("%d-%m-%Y")
        )

        correo.Body = """Buenos días, equipo:

            Se comparte el reporte S4 Servicio Administrado de Conectividad, el cual presenta una operación estable y sin incidencias al momento de la validación.

        Quedo atento a cualquier comentario.

        Saludos.
        """

        correo.Attachments.Add(
            archivo_pdf
        )

        escribir_log(
            "Reporte agregado al correo"
        )

        correo.Display()

        escribir_log(
            "Correo enviado correctamente"
        )

        correo = None
        outlook = None


        escribir_log(
            "Cerrando navegador"
        )

        context.close()
        browser.close()

        escribir_log(
            "Fin de ejecución"
        )

        import sys

        sys.stdout.flush()
        os._exit(0)

except Exception as e:

    escribir_log(
        f"ERROR GENERAL: {str(e)}"
    )

    raise