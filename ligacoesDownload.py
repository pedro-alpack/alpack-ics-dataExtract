# NOTA → Para o total entendimento dessa automação é necessário ler o conjunto de 'funções atalho' que simplificam certas ações 

# Imports
import os
import pyautogui
import subprocess
import time

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
    location = pyautogui.locateCenterOnScreen(path, confidence=0.9)
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
  
# ⚙ Automação
def downloadFile():
    pyautogui.hotkey('winleft', 'd')
    wait(3)
    exists('chrome.png')
    leftDoubleClickAt('chrome.png')
    wait(1)
    try:     
        path = os.path.join(os.path.dirname(__file__), 'assets', 'fullscreen.png')
        location = pyautogui.locateCenterOnScreen(path, confidence=0.9)
        if location:
            pyautogui.click(location)
    except:
        print('Chrome já está em tela cheia, continuando...')
    wait(3)
    exists('chromeOpen.png')
    exists('gmailIcon.png')
    leftClickAt('gmailIcon.png')
    exists('composebtn.png')
    wait(1)
    leftClickAt('sent.png')
    wait(1)
    pyautogui.click(475, 329)
    exists('file.png')
    filepath = os.path.join(os.path.dirname(__file__), 'assets', 'file.png')
    location = pyautogui.locateCenterOnScreen(filepath, confidence=0.9)

    if location:
        x, y = location
        pyautogui.click(x, y - 60)
    else:
        print("Imagem 'file.png' não encontrada na tela :(")
    exists('fileOpen.png')
    wait(1)
    leftClickAt('downloadBtn.png')
    wait(5)
    leftClickAt('closeChrome.png')
    
downloadFile()
subprocess.run(["python", "ligacoesExtract.py"]) # Executa a próxima automação da sequência
