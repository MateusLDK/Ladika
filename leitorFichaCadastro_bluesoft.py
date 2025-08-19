import pandas as pd
import requests
import os
from tkinter import filedialog as fd
import warnings
from dotenv import load_dotenv
from tqdm import tqdm

def requisitarAutenticacao():

    load_dotenv(".env")

    auth_url = f"https://erp.bluesoft.com.br/minipreco/oauth2/token"

    auth_response = requests.post(auth_url, data={
        'grant_type': 'client_credentials',
        'scope': 'switch.write',
        'client_id': os.getenv('client_id'),
        'client_secret': os.getenv('client_secret')
    })

    if auth_response.status_code == 200:
        token = auth_response.json().get('access_token')
        #print("\nToken obtido!")

        return token

    else:
        print(f"Erro na autenticação: {auth_response.status_code}")
        print(auth_response.text)


def consultar_produto_bluesoft(gtin, token):
    url = f"https://erp.bluesoft.com.br/minipreco/api/comercial/produtos/gtin/{gtin}"
    headers = {
        "Accept": "application/json",
        "Authorization": f"Bearer {token}"
    }
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        data = response.json()
        produto_key = data.get('produtoKey') or data.get('produto_key')
        return produto_key
    else:
        return None

def meltDataframe(dataFrame):
    colunasEncontradas = []
    codigos = dataFrame.filter(items=['codigo']).copy()
    dataFrame.drop(columns=['codigo'], inplace=True)
    if 'arquivo_origem' in dataFrame.columns:
        cols_sem_arquivo = [col for col in dataFrame.columns if col != 'arquivo_origem']
        dataFrame[cols_sem_arquivo] = dataFrame[cols_sem_arquivo].astype('Int64')
        dataFrame.insert(0, 'codigo', codigos)
    else:
        dataFrame.columns = dataFrame.columns.astype('Int64')
        dataFrame.insert(0, 'codigo', codigos)
    listaLojas = [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,201,202,203,204,205,206,207,208,209,210,602,301,302,303,304,305,306,307,308,309,310]
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
    dfMinimo.to_excel(excel_writer="banco minimo.xlsx", index=False)
    dfAvoid.to_excel(excel_writer="banco avoid.xlsx", index=False)

def prepararArquivo(arquivo, dataFrameParametro):
    colunas = ['EAN 13 (UNIDADE)',1,2,4,6,7,8,10,11,12,14,19,20,301,303,304,305,307,308]
    dataFrameProdutos = pd.read_excel(arquivo, skiprows=7, dtype={'EAN 13 (UNIDADE)': str})
    dataFrameProdutos = dataFrameProdutos.filter(items=colunas)
    dataFrameProdutos['arquivo_origem'] = os.path.basename(arquivo)
    dataFrameParametro = pd.concat([dataFrameParametro, dataFrameProdutos])
    dataFrameParametro['EAN 13 (UNIDADE)'] = dataFrameParametro['EAN 13 (UNIDADE)'].astype(str)
    dataFrameParametro.dropna(subset=['EAN 13 (UNIDADE)'], inplace=True)
    return dataFrameParametro

if __name__ == "__main__":
    warnings.simplefilter(action='ignore', category=UserWarning)
    # Usa a função para requisitar o token
    token = requisitarAutenticacao()
    if not token:
        print("Não foi possível obter o token de autenticação. Encerrando.")
        exit(1)
    arquivos = fd.askopenfilenames()
    dataFrameParametro = pd.DataFrame()
    for arquivo in arquivos:
        dataFrameParametro = prepararArquivo(arquivo, dataFrameParametro)
    dataFrameParametro.rename(columns={'EAN 13 (UNIDADE)': 'barra'}, inplace=True)
    dataFrameParametro['barra'] = dataFrameParametro['barra'].astype(str)
    dataFrameParametro['barra'] = dataFrameParametro['barra'].str.replace(r'\D', '', regex=True)
    # Consulta a API para cada barra e adiciona produtoKey
    barras_unicas = dataFrameParametro['barra'].unique()
    barra_to_key = {}
    print(f"Consultando {len(barras_unicas)} barras na API Bluesoft...")
    for barra in tqdm(barras_unicas, desc="Consultando API Bluesoft"):
        produto_key = consultar_produto_bluesoft(barra, token)
        barra_to_key[barra] = produto_key
    dataFrameParametro['codigo'] = dataFrameParametro['barra'].map(barra_to_key)
    # Erros: barras sem produtoKey
    dataFrameErro = dataFrameParametro[dataFrameParametro['codigo'].isna()].copy()
    dataFrameErro = dataFrameErro[dataFrameErro['barra'].str.strip() != '']
    dataFrameErro.dropna(subset=['barra'], inplace=True)
    if 'arquivo_origem' not in dataFrameErro.columns:
        if 'arquivo_origem' in dataFrameParametro.columns:
            dataFrameErro['arquivo_origem'] = dataFrameParametro.loc[dataFrameErro.index, 'arquivo_origem']
        else:
            dataFrameErro['arquivo_origem'] = 'desconhecido'
    if not dataFrameErro.empty:
        dataFrameErro.to_excel(excel_writer='erro parametro.xlsx', index=False)
    dataFrameParametro.dropna(subset=['codigo'], inplace=True)
    print(f"Taxa de sucesso: {len(dataFrameParametro['codigo'])}/{len(barras_unicas)}")
    dataFrameParametro.drop(columns=['barra'], inplace=True)
    meltDataframe(dataFrameParametro)
