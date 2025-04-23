import pandas as pd
import re
import os
from dotenv import load_dotenv
from sqlalchemy import create_engine
from tkinter import filedialog as fd

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
            pd.produto_key as codigo,
            pd.descricao as descricao
            from produto_d pd 
            where embalagem_key = 'UN' and (status = 1 or status = 2)
            '''

        try:
            # Executa a query e guarda os resultados em um DataFrame
            dfProdutos = pd.read_sql_query(query, self.engine)
            dfProdutos['codigo'] = dfProdutos['codigo'].astype(str)
            print("Consulta de Codigos executada com sucesso!")

        except Exception as e:
            print("Erro ao executar a consulta SQL:", e)
            return None
        
        return dfProdutos


def meltDataframe(dataFrame):
    
    colunasEncontradas = []
    
    codigos = dataFrame.filter(items=['codigo']).copy()
    dataFrame.drop(columns=['codigo'], inplace=True)
    dataFrame.columns = dataFrame.columns.astype(int)
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

    dataFrameArrumado.to_excel(excel_writer="Banco Parametrizacao.xlsx", index = False)


def prepararArquivo(arquivo, dataFrameFinal):

    dataFrameProdutos = pd.read_excel(arquivo, skiprows=7)
    dataFrameProdutos.columns = dataFrameProdutos.columns.astype(str)
    try:
        dataFrameFiltrado = dataFrameProdutos[
            ['EAN 13 (UNIDADE)','1','2','3','4','6','7','8','10','11','12','14','19','20','301','303','304','305','306','307','308']]

    except KeyError as e:
        colunasFaltantes = re.findall(r'(\d+)', str(e))
        for loja in colunasFaltantes:
            dataFrameProdutos.insert(0, loja, 0)

        dataFrameFiltrado = dataFrameProdutos[
            ['EAN 13 (UNIDADE)','1','2','3','4','6','7','8','10','11','12','14','19','20','301','303','304','305','306','307','308']]

    dataFrameFiltrado.dropna(subset=['EAN 13 (UNIDADE)'], inplace=True)
    dataFrameFinal = pd.concat([dataFrameFinal, dataFrameFiltrado])
    return dataFrameFinal

if __name__ == "__main__":

    arquivos = fd.askopenfilenames()
    dataFrameParametro = pd.DataFrame()

    for arquivo in arquivos:

        dataFrameParametro = prepararArquivo(arquivo, dataFrameParametro)

    linhasParaRemover = dataFrameParametro[dataFrameParametro.columns[1:]].isna().all(axis=1)
    dataFrameParametroFiltrado = dataFrameParametro[~linhasParaRemover]
    dataFrameParametroFiltrado.fillna(0, inplace=True)
    dataFrameParametroFiltrado.rename(columns={'EAN 13 (UNIDADE)': 'barra'}, inplace=True)

    dfProdutos = AcessarBanco().buscarProdutos()
    
    dataFrameFinal = pd.merge(dataFrameParametroFiltrado, dfProdutos, on='barra', how='left')
    dataFrameErros = dataFrameFinal[dataFrameFinal['codigo'].isna()].copy()
    dataFrameFinal.dropna(subset=['codigo'], inplace=True)
    dataFrameFinal.drop(columns=['descricao','barra'], inplace=True)
    
    novasColunas = ['codigo'] + [col for col in dataFrameFinal.columns if col != 'codigo']
    dataFrameFinal = dataFrameFinal[novasColunas]
    
    if not dataFrameErros.empty:
        dataFrameErros.to_excel('Erros Parametrizacao.xlsx', index=False)
    
    meltDataframe(dataFrameFinal)
