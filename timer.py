import schedule
import time
import subprocess
from datetime import datetime

def abrir_aplicativo():
    try:
        subprocess.Popen(["python", "Envio de emails.py"])
    except FileNotFoundError as e:
        print("O arquivo não foi encontrado erro: " + str(e))
    
def hora_atual_convertida():
    hora_atual = datetime.now()
    hora_convertida = hora_atual.strftime("%H:%M:%S")
    return hora_convertida

def executar_hora_definida():
    horario_definido = "12:30"
    schedule.every().day.at(horario_definido).do(abrir_aplicativo)

    # Adiciona o trabalho agendado para sábado e domingo
    # schedule.every().saturday.at(horario_definido).do(abrir_aplicativo)
    # schedule.every().sunday.at(horario_definido).do(abrir_aplicativo)
    
executar_hora_definida()

resultado_func = hora_atual_convertida()

print("Iniciando tarefa ás: " + resultado_func)

time.sleep(2) 

while True:
    schedule.run_pending()
    time.sleep(1)  