import datetime

class Log:
    def mensagem_de_log_completa(self, mensagem, caminho):
        with open(caminho, 'a') as log:
            momento = datetime.datetime.now()
            log.write(
                f"{mensagem} - {momento.strftime('%d/%m/%Y %H:%M:%S')}\n"
            )
            print(f"{mensagem} - {momento.strftime('%d/%m/%Y %H:%M:%S')}")

    def mensagem_de_log_sem_data(self, mensagem, caminho):
        with open(caminho, 'a') as log:
            momento = datetime.datetime.now()
            log.write(f"{mensagem} - {momento.strftime('%H:%M:%S')}\n")
            print(f"{mensagem} - {momento.strftime('%H:%M:%S')}")

    def mensagem_de_log_simples(self, mensagem, caminho):
        with open(caminho, 'a') as log:
            log.write(f"{mensagem}\n")
            print(f"{mensagem}")