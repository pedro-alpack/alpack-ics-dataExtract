# Imports
from datetime import datetime
from dotenv import load_dotenv
import fitz
import os
from pprint import pprint
import psycopg2
import re

# Carrega variáveis do .env
load_dotenv()  

db_config = {
    'host': os.getenv('DB_HOST'),
    'dbname': os.getenv('DB_NAME'),
    'user': os.getenv('DB_USER'),
    'password': os.getenv('DB_PASSWORD'),
    'port': int(os.getenv('DB_PORT'))
}

# Função para ler e extrair as informações do relatório gerado por outra automação (.pdf)
def extrair_dados_pdf(caminho_pdf):
    doc = fitz.open(caminho_pdf)
    registros = {}
    ramal_atual = None
    lig_index = 0
    captura = False
    buffer = []

    for page in doc:
        linhas = page.get_text("text").split('\n')
        for linha in linhas:
            linha = linha.strip()

            # Detecta RAMAL com regex (ex: "RAMAL  4254:")
            ramal_match = re.match(r'^RAMAL\s+(\d+):', linha)
            if ramal_match:
                ramal_atual = ramal_match.group(1)
                registros[ramal_atual] = {}
                lig_index = 0
                captura = True
                print(f"🔍 Iniciando captura do RAMAL: {ramal_atual}")
                continue

            # Detecta fim do bloco
            if "Total Duração:" in linha:
                print(f"✅ Finalizando captura do RAMAL: {ramal_atual}")
                captura = False
                buffer.clear()
                continue

            # Captura ligações em blocos de 4 linhas
            if captura:
                if linha.startswith(("I27f", "I27e", "I27d")):
                    buffer = [linha]
                elif buffer:
                    buffer.append(linha)
                    if len(buffer) == 4:
                        try:
                            data_hora = buffer[1].split()
                            if len(data_hora) == 2:
                                data, horario = data_hora
                                duracao = buffer[2]
                                from datetime import time

                                hora_minima = time(7, 40)
                                hora_maxima = time(17, 50)

                                hora_ligacao = datetime.strptime(horario, "%H:%M:%S").time()
                                if hora_minima <= hora_ligacao <= hora_maxima:
                                    registros[ramal_atual][f'lig{lig_index}'] = {
                                        'data': data,
                                        'horario': horario,
                                        'duracao': duracao,
                                        'ramal': ramal_atual
                                    }
                                    print(f"→ {ramal_atual} - lig{lig_index}: {data} {horario} {duracao}")
                                    lig_index += 1
                                else:
                                    print(f"⏹ Ignorada ligação fora do horário: {data} {horario}")

                        except Exception as e:
                            print(f"⚠️ Erro ao processar ligação: {buffer} → {e}")
                        finally:
                            buffer.clear()

    return registros

# Função para ordenar pelas chaves 'ligX' com base no número X
def ordenar_por_indices_renumerando(registros):
    registros_ordenados = {}
    for ramal, ligacoes in registros.items():
        
        ordenado = sorted(
            ligacoes.items(),
            key=lambda x: int(x[0].replace("lig", ""))
        )
        novo = {f'lig{i}': lig for i, (_, lig) in enumerate(ordenado)}
        registros_ordenados[ramal] = novo
    return registros_ordenados

caminho_pdf = os.path.expanduser("~/Downloads/Relatórios.pdf")
resultado = extrair_dados_pdf(caminho_pdf)
resultado = ordenar_por_indices_renumerando(resultado)

pprint(resultado)

try:
    conn = psycopg2.connect(**db_config)
    cur = conn.cursor()
    
    for ramal, ligacoes in resultado.items():
        for lig_id, info in ligacoes.items():
            data_sql = datetime.strptime(info['data'], "%d/%m/%Y").date()
            horario_sql = datetime.strptime(info['horario'], "%H:%M:%S").time()
            duracao_sql = datetime.strptime(info['duracao'], "%H:%M:%S").time()
            ramal_int = int(info['ramal'])

            # Verificação antes de inserir
            cur.execute("""
                SELECT 1 FROM ligacoes
                WHERE data = %s AND horario = %s AND ramal = %s AND duracao = %s AND ligID = %s
                LIMIT 1
            """, (data_sql, horario_sql, ramal_int, duracao_sql, lig_id))
            
            if cur.fetchone():
                print(f"⏩ Ignorado duplicado: {data_sql} {horario_sql} ramal {ramal_int} ligID {lig_id}")
                continue  # já existe

            cur.execute("""
                INSERT INTO ligacoes (data, horario, ramal, duracao, ligID)
                VALUES (%s, %s, %s, %s, %s)
            """, (data_sql, horario_sql, ramal_int, duracao_sql, lig_id))

    conn.commit()
    print("✔ Inserções finalizadas!")
    
    # Remove o PDF após sucesso
    if os.path.exists(caminho_pdf):
        os.remove(caminho_pdf)
        print(f"🗑 Arquivo PDF removido: {caminho_pdf}")

except Exception as e:
    print("Erro ao inserir no banco:", repr(e))

finally:
    if 'cur' in locals(): cur.close()
    if 'conn' in locals(): conn.close()

