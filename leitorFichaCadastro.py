import pandas as pd
import numpy as np
import os
from dotenv import load_dotenv
from sqlalchemy import create_engine
from tkinter import filedialog as fd
import warnings



class AcessarBanco():

    def __init__(self):

        load_dotenv()
        DB_HOST     = os.getenv("DB_HOST")
        DB_NAME     = os.getenv("DB_NAME")
        DB_USER     = os.getenv("DB_USER")
        DB_PASSWORD = os.getenv("DB_PASSWORD")
        DB_PORT     = os.getenv("DB_PORT")

        try:
            connection_url = f"postgresql+psycopg2://{DB_USER}:{DB_PASSWORD}@{DB_HOST}:{DB_PORT}/{DB_NAME}"
            self.engine = create_engine(connection_url)
            print("Conexão ao banco de dados realizada com sucesso!")

        except Exception as e:
            print("Erro ao conectar ao banco de dados:", e)

    def buscarProdutos(self):

        query = '''
            select
            pd.gtin_principal as barra,
            pd.produto_key as codigo
            from produto_d pd 
            where embalagem_key = 'UN' and (status = 1 or status = 2)
            '''

        try:
            # Executa a query e guarda os resultados em um DataFrame
            dfProdutos = pd.read_sql_query(query, self.engine)
            dfProdutos['codigo'] = dfProdutos['codigo'].astype(str)
            dfProdutos['barra'] = dfProdutos['barra'].astype(str)
            print("Consulta de Codigos executada com sucesso!")

        except Exception as e:
            print("Erro ao executar a consulta SQL:", e)
            return None
        
        return dfProdutos


def meltDataframe(dataFrame):

    colunasEncontradas = []

    codigos = dataFrame.filter(items=['codigo']).copy()
    dataFrame.drop(columns=['codigo'], inplace=True)
    # Remove a coluna 'arquivo_origem' antes de converter para int
    if 'arquivo_origem' in dataFrame.columns:
        cols_sem_arquivo = [col for col in dataFrame.columns if col != 'arquivo_origem']
        dataFrame[cols_sem_arquivo] = dataFrame[cols_sem_arquivo].astype('int')
        dataFrame.insert(0, 'codigo', codigos)
    else:
        dataFrame.columns = dataFrame.columns.astype('int')
        dataFrame.insert(0, 'codigo', codigos)

    listaLojas=[1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,201,202,203,204,205,206,207,208,209,210,602,301,302,303,304,305,306,307,308,309,310]
    colunasPlanilha = list(dataFrame.columns)

    for coluna in colunasPlanilha:

        if coluna in listaLojas:
            colunasEncontradas.append(coluna)

    for loja in colunasEncontradas:
        colunasPlanilha.remove(loja)

    dataFrameArrumado = dataFrame.melt(
        id_vars=colunasPlanilha, value_vars=colunasEncontradas, var_name="Loja", value_name="Minimo")

    dfAvoid = dataFrameArrumado[dataFrameArrumado['Minimo'] == 0]
    dfMinimo = dataFrameArrumado[dataFrameArrumado['Minimo'] > 0]

    dfMinimo.to_excel(excel_writer="banco minimo.xlsx", index = False)
    dfAvoid.to_excel(excel_writer="banco avoid.xlsx", index = False)
    

def prepararArquivo(arquivo, dataFrameParametro):

    colunas = ['EAN 13 (UNIDADE)',1,2,4,6,7,8,10,11,12,14,19,20,301,303,304,305,307,308]
    dataFrameProdutos = pd.read_excel(arquivo, skiprows=7, dtype={'EAN 13 (UNIDADE)': str})
    dataFrameProdutos = dataFrameProdutos.filter(items=colunas)

    # Adiciona coluna com o nome do arquivo de origem
    dataFrameProdutos['arquivo_origem'] = os.path.basename(arquivo)

    dataFrameParametro = pd.concat([dataFrameParametro, dataFrameProdutos])
    dataFrameParametro['EAN 13 (UNIDADE)'] = dataFrameParametro['EAN 13 (UNIDADE)'].astype(str)
    dataFrameParametro.dropna(subset=['EAN 13 (UNIDADE)'], inplace=True)

    return dataFrameParametro

if __name__ == "__main__":
    warnings.simplefilter(action='ignore', category=UserWarning)

    arquivos = fd.askopenfilenames()
    dataFrameParametro = pd.DataFrame()
    banco = AcessarBanco()

    for arquivo in arquivos:

        dataFrameParametro = prepararArquivo(arquivo, dataFrameParametro)

    dataFrameParametro.rename(columns={'EAN 13 (UNIDADE)': 'barra'}, inplace=True)
    dataFrameParametro['barra'] = dataFrameParametro['barra'].astype(str)
    # Remove todos os caracteres que não são dígitos da coluna 'barra'
    dataFrameParametro['barra'] = dataFrameParametro['barra'].str.replace(r'\D', '', regex=True)

    dfbanco = banco.buscarProdutos()
    if dfbanco is not None:
        dfbanco = dfbanco[dfbanco['barra'].isin(dataFrameParametro['barra'])]
        dfbanco = dfbanco[dfbanco['barra'].astype(str) != 'nan']
        dataFrameParametro = dataFrameParametro.dropna(subset=['barra'])
        dataFrameParametro = pd.merge(dataFrameParametro, dfbanco, on='barra', how='left')
        dataFrameErro = dataFrameParametro[dataFrameParametro['codigo'].isna()].copy()
        # Remove linhas em que barra for vazio
        dataFrameErro = dataFrameErro[dataFrameErro['barra'].str.strip() != '']
        dataFrameErro.dropna(subset=['barra'])
        # Adiciona coluna com o nome do arquivo de origem, se não existir
        if 'arquivo_origem' not in dataFrameErro.columns:
            # Tenta buscar do DataFrame original
            if 'arquivo_origem' in dataFrameParametro.columns:
                dataFrameErro['arquivo_origem'] = dataFrameParametro.loc[dataFrameErro.index, 'arquivo_origem']
            else:
                dataFrameErro['arquivo_origem'] = 'desconhecido'
        if not dataFrameErro.empty:
            dataFrameErro.to_excel(excel_writer='erro parametro.xlsx', index=False)

        dataFrameParametro.dropna(subset=['codigo'], inplace=True)
        dataFrameParametro.drop(columns=['barra'], inplace=True)
        meltDataframe(dataFrameParametro)

    else:
        print("Erro ao buscar produtos do banco de dados.")

