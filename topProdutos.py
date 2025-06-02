# NOTA → Para o total entendimento dessa automação é necessário ler o conjunto de 'funções atalho' que simplificam certas ações 

# Imports
from datetime import datetime
from dotenv import load_dotenv
import os
import pandas as pd
import psycopg2
from psycopg2 import sql
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
    return psycopg2.connect(**db_config)

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


# Primeiro e último dia do ano
ano_atual = datetime.now().year
primeiro_dia = f"0101{ano_atual}"
ultimo_dia = f"3112{ano_atual}"

# ⚙ Automação
def exportFromSystem():
    pyautogui.hotkey('winleft', 'd')
    wait(3)
    leftDoubleClickAt('aresIcon.png')
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
    pyautogui.press('enter')
    exists('excelIcon.png')
    wait(1)
    leftClickAt('pedidoVendaGerencial.png')
    wait(0.5)
    for i in range(2):
        pyautogui.press('tab')
    leftClickAt('selectIcon.png')
    pyautogui.write('pr')
    pyautogui.press('enter')
    leftClickAt('selectIcon.png')
    pyautogui.write('s')
    pyautogui.press('enter')
    for i in range(2):
        pyautogui.press('tab')
    pyautogui.write(primeiro_dia)
    pyautogui.press('tab')
    pyautogui.write(ultimo_dia)
    wait(0.5)
    leftClickAt('excelBtn.png')
    waitForExcelWindow()
    wait(1)
    pyautogui.hotkey('ctrl', 'b')
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
    pyautogui.hotkey('alt', 'f4')
    wait(1)
    pyautogui.press('enter')
    
# Função para carregar os dados do Excel
def load_excel():
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    file_path = os.path.join(desktop, "Pasta1.xlsx")

    # Lê colunas A (0), B (1), D (3) e K (10), a partir da linha 2
    df = pd.read_excel(file_path, usecols=[0, 1, 3, 10], skiprows=1)
    df.columns = ['prodid', 'prodname', 'qtdvend', 'valvend']

    # Substitui valores vazios em 'prodname'
    df['prodname'] = df['prodname'].fillna('produto sem nome')

    return df, file_path

    


# Função para inserir os dados no db
def sync_to_postgres(df):
    conn = get_db_connection()
    cursor = conn.cursor()

    try:
        # Remove todos os registros da tabela topvendas
        cursor.execute("DELETE FROM topvendas;")
        
        # Reseta a sequência do campo id para 1
        cursor.execute("SELECT setval('topvendas_id_seq', 1, false);")
        
        # Insere os dados da planilha
        for index, row in df.iterrows():
            cursor.execute(
                """
                INSERT INTO topvendas (prodid, prodname, qtdvend, valvend)
                VALUES (%s, %s, %s, %s);
                """,
                (
                    int(row['prodid']),
                    row['prodname'],
                    float(row['qtdvend']) if not pd.isna(row['qtdvend']) else 0,
                    float(row['valvend']) if not pd.isna(row['valvend']) else 0.00
                )
            )

        conn.commit()
        print("Dados inseridos com sucesso na tabela topvendas.")

    except Exception as e:
        conn.rollback()
        print(f"Erro ao inserir dados: {e}")
    finally:
        cursor.close()
        conn.close()


    
exportFromSystem()

wait(4)

if __name__ == "__main__":
    df, excel_path = load_excel()
    sync_to_postgres(df)

    try:
        os.remove(excel_path)
        print(f"Arquivo {excel_path} removido com sucesso.")
    except OSError as e:
        print(f"Erro ao remover o arquivo: {e}")

