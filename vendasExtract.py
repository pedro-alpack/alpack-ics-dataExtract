# NOTA → Para o total entendimento dessa automação é necessário ler o conjunto de 'funções atalho' que simplificam certas ações 

import pyautogui
import time
import os


# 🚨 Funções atalho

# Simula o WIN + D (minimizar todas as janelas)
def winD():
    pyautogui.keyDown('winleft')
    pyautogui.press('d')
    pyautogui.keyUp('winleft') 

# Simplificação do delay entre as execuções
def wait(interval):
    time.sleep(interval)
    
# Simplificação para clicar com o botão esquerdo em imagens na tela
def leftClickAt(img):
    path = os.path.join(os.path.dirname(__file__), 'assets', img)
    location = pyautogui.locateCenterOnScreen(path)
    if location:
        pyautogui.click(location)
    else:
        print(f"Imagem '{img}' não encontrada na tela :(")
        
# Simplificação para clicar duas vezes com o botão esquerdo em imagens na tela
def leftDoubleClickAt(img):
    path = os.path.join(os.path.dirname(__file__), 'assets', img)
    location = pyautogui.locateCenterOnScreen(path)
    if location:
        pyautogui.doubleClick(location)
    else:
        print(f"Imagem '{img}' não encontrada na tela :(")
        
# Simplificação para clicar com o botão direito em imagens na tela
def rightClickAt(img):
    path = os.path.join(os.path.dirname(__file__), 'assets', img)
    location = pyautogui.locateCenterOnScreen(path)
    if location:
        pyautogui.click(location, button='right')
    else:
        print(f"Imagem '{img}' não encontrada na tela :(")
        
# Função para esperar com que uma imagem exista na tela
def exists(img):
    path = os.path.join(os.path.dirname(__file__), 'assets', img)
    
    while True:
        location = pyautogui.locateCenterOnScreen(path)
        if location:
            print(f"Imagem '{img}' encontrada!")
            break
        else:
            print(f"Aguardando imagem '{img}' aparecer...")
            time.sleep(0.5)


# Automação
def exportFromSystem():
    winD()
    wait(1)
    leftDoubleClickAt('aresIcon.png')
    exists('loginIcon.pbg')
    wait(1)
    pyautogui.write('1234')
    pyautogui.press('enter')
    
    
exportFromSystem()
