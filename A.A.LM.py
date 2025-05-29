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
def contagem_QPA(df):
    criterio = "Quantidade prevista alterada"
    return df['Alteração Simplificada'].value_counts().get(criterio, 0)
def contagem_DPA(df):
    criterio = "Data do lote alterada alterada"
    return df['Alteração Simplificada'].value_counts().get(criterio, 0)
def contagem_FA(df):
    criterio = "Fase alterada"
    return df['Alteração Simplificada'].value_counts().get(criterio, 0)
def contagem_INC(df):
    criterio = "Inclusão de item"
    return df['Alteração Simplificada'].value_counts().get(criterio, 0)
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

criterios = {
    "Quantidade prevista alterada": contagem_QPA(df),
    "Data do lote alterada alterada": contagem_DPA(df),
    "Fase alterada": contagem_FA(df),
    "Inclusão de item": contagem_INC(df)
}

df_cabecalho = pd.DataFrame([criterios])
df = pd.concat([df,df_cabecalho], ignore_index=True)

arquivoSaida = gerar_nome_saida()
df.to_excel(arquivoSaida, index=False)
print("Análise Completa!")
"""
ufvs = [
    "UFV ACOPIARA I E II",
    "UFV ALTA FLORESTA I A V",
    "UFV ALTO CARACOL",
    "UFV ALTO DO RODRIGUES I - SANTA CRUZ",
    "UFV ALTO DO RODRIGUES II E III",
    "UFV ALUMINIO I A V",
    "UFV AMPERE",
    "UFV ANDRADINA",
    "UFV ANGICOS I",
    "UFV ANTA I, II E III",
    "UFV ARAPOTI I, II E III",
    "UFV ARROIO DOS RATOS I",
    "UFV ARROIO DOS RATOS II",
    "UFV ATERRO SANITÁRIO",
    "UFV BAGÉ I A V",
    "UFV BARRA DO CHOÇA V E VI",
    "UFV BATALHA I",
    "UFV BATALHA II",
    "UFV BATURITÉ I",
    "UFV BATURITÉ II",
    "UFV BILAC I",
    "UFV BRUMADO I, II E III",
    "UFV BRUMADO IV, V E VI",
    "UFV CÁCERES I A V",
    "UFV CACHOEIRA DO SUL I",
    "UFV CAFEZAL",
    "UFV CAMAQUÃ I A V",
    "UFV CAMPO GRANDE I",
    "UFV CAMPO MOURÃO I, II E III",
    "UFV CAMPO MOURÃO IV, V E VI",
    "UFV CAMPO MOURÃO VII, VIII E IX",
    "UFV CANDELARIA I E II",
    "UFV CAPANEMA V",
    "UFV CARRILHO",
    "UFV CASCAVEL I",
    "UFV CASSILANDIA I",
    "UFV CASSILANDIA II",
    "UFV CASSILANDIA III",
    "UFV CERQUEIRA CÉSAR I",
    "UFV CERQUEIRA CÉSAR II E III",
    "UFV CERQUEIRA CÉSAR IV E V",
    "UFV COLIDER I A V",
    "UFV COXIM",
    "UFV FRANCISCO BELTRÃO I, II E III",
    "UFV FRANCISCO BELTRÃO IV, V E VI",
    "UFV GETULINA",
    "UFV GUAIMBÊ",
    "UFV IBIRAPUÃ I",
    "UFV IBIRAPUÃ II",
    "UFV INOCÊNCIA",
    "UFV IPORÃ I",
    "UFV IRARÁ I",
    "UFV ITAPAJÉ I",
    "UFV IVINHEMA",
    "UFV JACOBINA I A V",
    "UFV JACOBINA II (LARANJEIRAS)",
    "UFV JANIO QUADROS",
    "UFV JUCURUTU I",
    "UFV JUINA I",
    "UFV JUINA II",
    "UFV LAGOINHA DO PIAUI I",
    "UFV LARANJEIRAS DO SUL I",
    "UFV LARANJEIRAS DO SUL II",
    "UFV LEME I, II E III",
    "UFV MÃE DO RIO I A V",
    "UFV MARECHAL",
    "UFV MAURITI I E II",
    "UFV MIRACEMA I A V",
    "UFV MIRANDOPOLIS",
    "UFV MODELO I A V",
    "UFV MORADA NOVA",
    "UFV NOVA CRIXÁS I A V",
    "UFV PALMITOS I A V",
    "UFV PEROBAL",
    "UFV PIANCÓ",
    "UFV PONTES E LACERDA",
    "UFV PORANGATU IV, V E VI",
    "UFV PORTO NACIONAL",
    "UFV REALEZA I, II E III",
    "UFV RIACHO DA CRUZ",
    "UFV RONDON I, II E III",
    "UFV SANTANESIA",
    "UFV SANTO ANTONIO DO SUDOESTE I E II",
    "UFV SÃO CAETANO I",
    "UFV SÃO CAETANO II",
    "UFV SÃO CAETANO III",
    "UFV SÃO FRANCISCO I",
    "UFV SÃO GABRIEL I",
    "UFV SÃO LOURENÇO DO OESTE III - ANDRE PERAZOLI",
    "UFV SÃO LOURENÇO DO SUL I",
    "UFV SÃO LOURENÇO DO SUL II",
    "UFV SÃO ROQUE I",
    "UFV SÃO ROQUE II",
    "UFV SAUDADES DO IGUAÇU I",
    "UFV SIDROLÂNDIA I",
    "UFV TAIUVA I",
    "UFV TAQUARA I, II e III",
    "UFV TAQUARA IV",
    "UFV TERRA DE AREIA I",
    "UFV TERRA DE AREIA II",
    "UFV TERRA DE AREIA III",
    "UFV TERRA ROXA II",
    "UFV TOMÉ I E II",
    "UFV TUPIRAMA I A V",
    "UFV UMUARAMA",
    "UFV URUAÇU",
    "UFV VILA PAVÃO",
    "UFV VITORIA DA CONQUISTA I A V"
]
"""
