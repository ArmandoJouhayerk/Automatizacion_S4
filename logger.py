from datetime import datetime

class Logger:
    def __init__(self, log_file):
        self.log_file = log_file

    def escribir_log(self, mensaje):
        with open(self.log_file, "a", encoding="utf-8") as log:
            fecha_log = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
            log.write(f"[{fecha_log}] {mensaje}\n")