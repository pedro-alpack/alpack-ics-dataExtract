import os
import psycopg2
from docx import Document
from tkinter import Tk, filedialog
from dotenv import load_dotenv

# Carrega variáveis do .env
load_dotenv()

db_config = {
    'host': os.getenv('DB_HOST'),
    'dbname': os.getenv('DB_NAME'),
    'user': os.getenv('DB_USER'),
    'password': os.getenv('DB_PASSWORD'),
    'port': int(os.getenv('DB_PORT'))
}

def selecionar_arquivo_word():
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        filetypes=[("Documentos Word", "*.docx")],
        title="Selecione o arquivo Word"
    )
    return file_path

def limpar_frase(linha):
    linha = linha.strip().strip('"“”')  # Remove aspas e espaços extras
    if "–" in linha:
        partes = linha.split("–", 1)
    elif "-" in linha:
        partes = linha.split("-", 1)
    else:
        return linha  # Sem autor
    frase = partes[0].strip()
    autor = partes[1].strip()
    return f"{frase} - {autor}"

def extrair_frases(doc_path):
    doc = Document(doc_path)
    frases = []
    for para in doc.paragraphs:
        texto = para.text.strip()
        if texto:
            frases.append(limpar_frase(texto))
    return frases

def inserir_frases_no_banco(frases):
    try:
        conn = psycopg2.connect(**db_config)
        cur = conn.cursor()
        for frase in frases:
            cur.execute(
                "INSERT INTO frases (text, display) VALUES (%s, %s)",
                (frase, True)
            )
        conn.commit()
        cur.close()
        conn.close()
        print(f"{len(frases)} frases inseridas com sucesso.")
    except Exception as e:
        print("Erro ao inserir no banco:", e)

def main():
    caminho_doc = selecionar_arquivo_word()
    if not caminho_doc:
        print("Nenhum arquivo selecionado.")
        return
    frases = extrair_frases(caminho_doc)
    if frases:
        inserir_frases_no_banco(frases)
    else:
        print("Nenhuma frase válida encontrada.")

if __name__ == "__main__":
    main()
