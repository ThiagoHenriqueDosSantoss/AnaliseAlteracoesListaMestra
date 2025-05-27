import pandas as pd
import re
from openpyxl.reader.excel import load_workbook
from tqdm import tqdm
import os
"""
    Este script automatiza a análise do relatório MPOP218 (Histórico de alterações da Lista Mestra),
    realizando previamente os calculos necessários com base nos critérios fornecidos pelo relatório exportado
    do ERP Senior.
"""
tqdm.pandas()
def simplificar_descricao(descricao):
    if pd.isna(descricao):
        return ""
    return re.sub(r" de .+? para .+", "", str(descricao))

def gerar_nome_saida(base="HALM", extensao=".xlsx"):
    contador = 1
    while os.path.exists(f"{base}{contador}{extensao}"):
        contador += 1
    return f"{base}{contador}{extensao}"

# Caminho do arquivo original exportado do ERP
arquivoEntrada = "C:/Users/thiago.santos/Desktop/MPOP218.xlsx"

df = pd.read_excel(arquivoEntrada, skiprows=2)

print("Colunas disponíveis:", df.columns.tolist())

df = df.dropna(axis=1, how='all')
df = df.dropna(axis=0, how='all')

if 'Alteração' in df.columns:
    df['Alteração Simplificada'] = df['Alteração'].progress_apply(simplificar_descricao)
else:
    print("Coluna 'Alteração' não encontrada!")


arquivoSaida = gerar_nome_saida()
df.to_excel(arquivoSaida, index=False)
print("Análise Completa!")