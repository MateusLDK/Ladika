import pandas as pd
import re
import os
from dotenv import load_dotenv
from sqlalchemy import create_engine
from tkinter import filedialog as fd


def prepararArquivo(arquivo, dataFrameFinal):

    dataFrameProdutos = pd.read_excel(arquivo, skiprows=1)

    try:
        dataFrameFiltrado = dataFrameProdutos[['Código interno','Loja','Parametrização']]
    except KeyError:
        dataFrameProdutos = pd.read_excel(arquivo, skiprows=0)
        dataFrameFiltrado = dataFrameProdutos[['Código interno','Loja','Parametrização']]

    dataFrameProdutos.columns = dataFrameProdutos.columns.astype(str)

    dataFrameFiltrado.dropna(subset=['Código interno'], inplace=True)
    dataFrameFinal = pd.concat([dataFrameFinal, dataFrameFiltrado])
    return dataFrameFinal


if __name__ == "__main__":

    arquivos = fd.askopenfilenames()
    dataFrameFinal = pd.DataFrame()

    for arquivo in arquivos:

        dataFrameFinal = prepararArquivo(arquivo, dataFrameFinal)

dataFrameFinal.to_excel(excel_writer="Planilha unificada.xlsx", index = False)
