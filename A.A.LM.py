import time

import pandas as pd
import re
from tqdm import tqdm
from datetime import datetime
import os

"""
    Este script automatiza a análise do relatório MPOP218 (Histórico de alterações da Lista Mestra),
    realizando previamente os calculos necessários com base nos critérios fornecidos pelo relatório exportado
    do ERP Senior.
"""
def main():
    tqdm.pandas()

    def valida_data(df, dataIni, dataFim):
        if 'Data' not in df.columns:
            time.sleep(3)
            print("Coluna 'Data' não encontrada!")
            time.sleep(3)
            print("Colunas disponíveis:", df.columns.tolist())
            exit()

        try:
            df['Data'] = pd.to_datetime(df['Data'], dayfirst=True, errors='coerce')
        except Exception as e:
            time.sleep(3)
            print(f"Erro ao converter datas: {e}")
            exit()

        df = df.dropna(subset=['Data'])
        dfData_filtrado = df[(df['Data'] >= dataIni) & (df['Data'] <= dataFim)]

        if dfData_filtrado.empty:
            time.sleep(3)
            print("❌ Nenhum registro encontrado dentro do período informado!")
            exit()
        time.sleep(3)
        print(f"✅ {len(dfData_filtrado)} registro(s) dentro do período de {dataIni.date()} até {dataFim.date()} encontrado(s).")
        return dfData_filtrado
    def validar_grupo_usina(df, numeroGrupo):
        if 'Usina' not in df.columns:
            print("Coluna 'Usina' não encontrada no Excel!")
            exit()

        df_filtrado = df[df['Usina'].astype(str).str.contains(numeroGrupo)]

        if df_filtrado.empty:
            print(f"Nenhuma usina com grupo '{numeroGrupo}' encontrada na coluna 'Usina'.")
            exit()
        time.sleep(3)
        print(f"Usina(s) com grupo '{numeroGrupo}' encontrada(s)!")
        return df_filtrado


    def simplificar_descricao(descricao):
        if pd.isna(descricao):
            return ""
        return re.sub(r" de .+? para .+", "", str(descricao))
    def contagem_QPA(df):
        criterio = "Quantidade prevista alterada"
        return df['Alteração Simplificada'].value_counts().get(criterio, 0)
    def contagem_DPA(df):
        criterio = "Data do lote alterada"
        return df['Alteração Simplificada'].value_counts().get(criterio, 0)
    def contagem_FA(df):
        criterio = "Fase alterada"
        return df['Alteração Simplificada'].value_counts().get(criterio, 0)
    def contagem_INC(df):
        criterio = "Inclusão de item"
        return df['Alteração Simplificada'].value_counts().get(criterio, 0)
    def gerar_nome_saida(caminhoSaida,base="HALM", extensao=".xlsx"):
        contador = 0
        os.makedirs(caminhoSaida, exist_ok=True)
        while True:
            contador += 1
            caminho_completo = os.path.join(caminhoSaida, f"{base}{contador}{extensao}")
            if not os.path.exists(caminho_completo):
                return caminho_completo

    def gerar_analise():
        while True:
            try:
                numeroGrupo = input("Digite o número do grupo da usina para análise: ")
                time.sleep(1)
                dataIni = input("Digite a data inicial (no formato dd/mm/aaaa): ").strip()
                time.sleep(1)
                dataFim = input("Digite a data final (no formato dd/mm/aaaa): ").strip()

                try:
                    dataIni = datetime.strptime(dataIni, "%d/%m/%Y")
                    dataFim = datetime.strptime(dataFim, "%d/%m/%Y")
                except ValueError:
                    time.sleep(3)
                    print("Formato de data inválido. Use dd/mm/aaaa.")

                # Caminho do arquivo original exportado do ERP
                arquivoEntrada = "C:/Users/thiago.santos/PycharmProjects/pythonProject1/AnaliseAlteracoesListaMestra/MPOP218.xlsx"

                if not os.path.exists(arquivoEntrada):
                    time.sleep(2)
                    print(f"❌ Arquivo não encontrado: {arquivoEntrada}")
                    break

                df = pd.read_excel(arquivoEntrada, skiprows=2)
                df = validar_grupo_usina(df, numeroGrupo)
                df = valida_data(df, dataIni, dataFim)

                df['Alteração Simplificada'] = df['Alteração'].apply(simplificar_descricao)

                df.reset_index(drop=True, inplace=True)
                df["Quantidade prevista alterada"] = ""
                df.at[0, "Quantidade prevista alterada"] = contagem_QPA(df)

                df["Data do lote alterada"] = ""
                df.at[0, "Data do lote alterada"] = contagem_DPA(df)

                df["Fase alterada"] = ""
                df.at[0, "Fase alterada"] = contagem_FA(df)

                df["Inclusão de item"] = ""
                df.at[0, "Inclusão de item"] = contagem_INC(df)

                caminhoSaida = "C:/Users/thiago.santos/PycharmProjects/pythonProject1/AnaliseAlteracoesListaMestra"
                arquivoSaida = gerar_nome_saida(caminhoSaida)
                df.to_excel(arquivoSaida, index=False)

                print("✅ Análise Completa!", f"Arquivo gerado: {arquivoSaida}")
                time.sleep(2)
                print("E N C E R R A N D O . . .")
                time.sleep(2)

                resposta = input("Deseja realizar outra análise? (s/n): ").strip().lower()
                if resposta != 's':
                    print("Programa finalizado.")
                    break

            except Exception as e:
                print("Erro durante a execução da análise:", str(e))
                resposta = input("Deseja tentar novamente? (s/n): ").strip().lower()
                if resposta != 's':
                    print("Programa finalizado.")
                    break
    gerar_analise()

if __name__ == "__main__":
    main()
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
    "UFV ANTA I, II E III",+6
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