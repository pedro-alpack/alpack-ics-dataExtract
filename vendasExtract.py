# NOTA → Para o total entendimento dessa automação é necessário ler o conjunto de 'funções atalho' que simplificam certas ações 

import pyautogui
import time
from datetime import datetime
import os


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
        
# Simplificação para clicar com o botão direito em imagens na tela
def rightClickAt(img):
    path = os.path.join(os.path.dirname(__file__), 'assets', img)
    location = pyautogui.locateCenterOnScreen(path, confidence=0.7)
    if location:
        pyautogui.click(location, button='right')
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

# Dia de hoje
data_formatada = datetime.now().strftime('%d%m%Y')

# Automação
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
    pyautogui.write(data_formatada)
    pyautogui.press('tab')
    pyautogui.write(data_formatada)
    wait(0.5)
    leftClickAt('aplicarFiltro.png')
    wait(5)
    leftClickAt('excelIcon.png')
    waitForExcelWindow()
    wait(1)
    pyautogui.hotkey('ctrl', 'b')
    for i in range(4):
        pyautogui.press('enter')
    wait(1)
    pyautogui.press('enter')
    wait(0.2)
    pyautogui.hotkey('alt', 'f4')
    wait(0.2)
    pyautogui.hotkey('alt', 'f4')
    pyautogui.press('enter')
    
exportFromSystem()
