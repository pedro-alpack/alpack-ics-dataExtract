import fitz  # PyMuPDF
import re

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
                                registros[ramal_atual][f'lig{lig_index}'] = {
                                    'data': data,
                                    'horario': horario,
                                    'duracao': duracao,
                                    'ramal': ramal_atual
                                }
                                print(f"→ {ramal_atual} - lig{lig_index}: {data} {horario} {duracao}")
                                lig_index += 1
                        except Exception as e:
                            print(f"⚠️ Erro ao processar ligação: {buffer} → {e}")
                        finally:
                            buffer.clear()

    return registros

def ordenar_por_indices_renumerando(registros):
    registros_ordenados = {}
    for ramal, ligacoes in registros.items():
        # Ordena pelas chaves 'ligX' com base no número X
        ordenado = sorted(
            ligacoes.items(),
            key=lambda x: int(x[0].replace("lig", ""))
        )
        # Recria o dicionário com lig0, lig1, lig2... renumerados
        novo = {f'lig{i}': lig for i, (_, lig) in enumerate(ordenado)}
        registros_ordenados[ramal] = novo
    return registros_ordenados

# Caminho do PDF
caminho_pdf = 'C:/Users/tialp/Desktop/Relatórios.pdf'
resultado = extrair_dados_pdf(caminho_pdf)
resultado = ordenar_por_indices_renumerando(resultado)

from pprint import pprint
pprint(resultado)
