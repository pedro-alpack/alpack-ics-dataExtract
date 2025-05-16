import time
import datetime
import subprocess
import os

# Caminho base dos arquivos
base_path = os.path.dirname(os.path.abspath(__file__))

# Flags de controle
ja_rodou_historico = False

while True:
    agora = datetime.datetime.now()
    hora = agora.time()

    # Entre 07:44 e 17:50 – rodar vendasExtract.py a cada minuto
    if datetime.time(7, 44) <= hora <= datetime.time(17, 50):
        subprocess.run(['python', os.path.join(base_path, 'vendasExtract.py')])
        time.sleep(60)  # Espera 1 minuto
        ja_rodou_historico = False  # Garante que o histórico possa rodar às 18h

    # Às 18:00 – rodar historicoVendasExtract.py, importarProdutos.py e topProdutos.py apenas uma vez 
    elif hora.hour == 18 and hora.minute == 0 and not ja_rodou_historico:
        subprocess.run(['python', os.path.join(base_path, 'historicoVendasExtract.py')])
        time.sleep(120)
        subprocess.run(['python', os.path.join(base_path, 'importarProdutos.py')])
        time.sleep(60)
        subprocess.run(['python', os.path.join(base_path, 'topProdutos.py')])
        ja_rodou_historico = True
        time.sleep(60)

    else:
        time.sleep(10)  # Verifica a hora a cada 10s fora do intervalo útil
