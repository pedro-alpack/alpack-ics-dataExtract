# NOTA → Para o total entendimento dessa automação é necessário ler o conjunto de 'funções atalho' que simplificam certas ações 

# Imports
import calendar
from datetime import datetime, date
from dotenv import load_dotenv
import os
import pandas as pd
import psycopg
import pyautogui
import subprocess
import time

 # Carrega variáveis do .env
load_dotenv() 

db_config = {
    'host': os.getenv('DB_HOST'),
    'dbname': os.getenv('DB_NAME'),
    'user': os.getenv('DB_USER'),
    'password': os.getenv('DB_PASSWORD'),
    'port': int(os.getenv('DB_PORT'))
}

def get_db_connection():
    return psycopg.connect(**db_config)


# 🚨 Funções atalho

# Simplificação do delay entre as execuções
def wait(interval):
    time.sleep(interval)
    
# Simplificação para clicar com o botão esquerdo em imagens na tela
def leftClickAt(img):
    path = os.path.join(os.path.dirname(__file__), 'assets', img)
    location = pyautogui.locateCenterOnScreen(path, confidence=0.7)
    if location:
        pyautogui.click(location)
    else:
        print(f"Imagem '{img}' não encontrada na tela :(")
        
# Simplificação para clicar duas vezes com o botão esquerdo em imagens na tela
def leftDoubleClickAt(img):
    path = os.path.join(os.path.dirname(__file__), 'assets', img)
    location = pyautogui.locateCenterOnScreen(path, confidence=0.7)
    if location:
        pyautogui.doubleClick(location)
    else:
        print(f"Imagem '{img}' não encontrada na tela :(")
        
# Função para esperar com que uma imagem exista na tela
def exists(img):
    path = os.path.join(os.path.dirname(__file__), 'assets', img)
    
    while True:
        try:
            location = pyautogui.locateCenterOnScreen(path, confidence=0.9)
            if location:
                print(f"Imagem '{img}' encontrada!")
                break
            else:
                print(f"Aguardando imagem '{img}' aparecer...")
                wait(0.5)
        except pyautogui.ImageNotFoundException:
            print(f"Imagem '{img}' ainda não encontrada (exceção).")
            wait(0.5)
            
# Função para esperar com que a janela do excel abra
def waitForExcelWindow(title_contains='Excel'):
    while True:
        windows = pyautogui.getWindowsWithTitle(title_contains)
        if windows:
            print(f"Janela do Excel encontrada: '{windows[0].title}'")
            break
        print("Aguardando o Excel abrir...")
        wait(1)

# Data atual
hoje = date.today()
ano = hoje.year
mes = hoje.month

# Primeiro dia do mês
primeiro_dia = date(ano, mes, 1)

# Último dia do mês
ultimo_dia_num = calendar.monthrange(ano, mes)[1]
ultimo_dia = date(ano, mes, ultimo_dia_num)

# Datas formatadas para strings
primeiro_dia_str = primeiro_dia.strftime('%d%m%Y')
ultimo_dia_str = ultimo_dia.strftime('%d%m%Y')

# ⚙ Automação
def exportFromSystem():
    pyautogui.hotkey('winleft', 'd')
    wait(3)
    leftDoubleClickAt('programIcon.png')
    exists('loginIcon.png')
    wait(3)
    pyautogui.write('1234')
    pyautogui.press('enter')
    wait(1)
    exists('homeIcon.png')
    wait(1)
    leftClickAt('homeIcon.png')
    wait(1)
    pyautogui.press('t')
    wait(0.5)
    pyautogui.press('enter')
    exists('excelIcon.png')
    wait(1)
    leftClickAt('filterIcon.png')
    wait(0.5)
    leftClickAt('filtrarPedidos.png')
    wait(1)
    for i in range(4):
        pyautogui.press('tab')
    pyautogui.write(primeiro_dia_str)
    pyautogui.press('tab')
    pyautogui.write(ultimo_dia_str)
    wait(0.5)
    leftClickAt('aplicarFiltro.png')
    wait(5)
    leftClickAt('excelIcon.png')
    waitForExcelWindow()
    wait(5)
    largura, altura = pyautogui.size()

    # Calcula o ponto central
    centro_x = largura // 2
    centro_y = altura // 2

    # Move o cursor até o centro e clica
    pyautogui.click(centro_x, centro_y)
    exists('salvar_ico.png')
    leftClickAt('salvar_ico.png')
    for i in range(4):
        pyautogui.press('enter')
        wait(0.5)
    wait(2)
    pyautogui.press('enter')
    wait(2)
    pyautogui.hotkey('alt', 'f4')
    wait(1)
    pyautogui.hotkey('alt', 'f4')
    wait(1)
    pyautogui.press('enter')
    
# Função para carregar os dados do Excel
def load_excel():
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    file_path = os.path.join(desktop, "Pasta1.xlsx")
    df = pd.read_excel(file_path, engine="openpyxl")  # sem dtype=str!

    # Usa a coluna I (índice 8) se houver valor, senão usa a coluna A (índice 0)
    df["data"] = df.apply(
        lambda row: row.iloc[8] if pd.notnull(row.iloc[8]) else row.iloc[0],
        axis=1
    )

    df["data"] = pd.to_datetime(df["data"], dayfirst=True).dt.date

    print("VALORES ORIGINAIS DO EXCEL:")
    print(df["Total Pedido"].head(10).tolist())

    # Conversão segura
    df["valor_vendido"] = df["Total Pedido"].apply(lambda x: round(float(x), 2))

    print("VALORES CONVERTIDOS:")
    print(df[["Total Pedido", "valor_vendido"]].head(10))

    df["vendedor"] = df["Vendedor 1"]
    df["statuspedido"] = df["Status do Pedido"]
    df["cliente"] = df["Cliente"]
    df["regiao"] = df["Rota"]
    df["vendaID"] = df["Id"].astype(int)

    return df[["data", "vendedor", "valor_vendido", "statuspedido", "cliente", "regiao", "vendaID"]], file_path



# Função para inserir os dados no db
def sync_to_postgres(df):
    try:
        with get_db_connection() as conn:
            with conn.cursor() as cursor:
                # Remove todos os registros da tabela
                cursor.execute("DELETE FROM vendas")

                # Reseta a sequência do campo id para 1
                cursor.execute("SELECT setval('vendas_id_seq', 1, false);")

                for _, row in df.iterrows():
                    if abs(row["valor_vendido"]) >= 10**8:
                        print(f"Valor muito alto ignorado: {row['valor_vendido']} (vendaID={row['vendaID']})")
                        continue

                    cursor.execute(
                        """
                        INSERT INTO vendas
                          (data, vendedor, valor_vendido, "statusPedido", cliente, regiao, "vendaID")
                        VALUES (%s, %s, %s, %s, %s, %s, %s)
                        """,
                        (row["data"], row["vendedor"], row["valor_vendido"],
                         row["statuspedido"], row["cliente"], row["regiao"], row["vendaID"])
                    )

            conn.commit()
            print("✔ Dados inseridos com sucesso na tabela 'vendas'.")
    except Exception as e:
        print(f"Erro ao inserir dados: {e}")



exportFromSystem()

wait(4)

if __name__ == "__main__":
    df, excel_path = load_excel()
    sync_to_postgres(df)

    try:
        os.remove(excel_path)
        print(f"Arquivo {excel_path} removido com sucesso.")
        subprocess.run(["python", "ligacoesDownload.py"]) # Executa a próxima automação da sequência
    except OSError as e:
        print(f"Erro ao remover o arquivo: {e}")

