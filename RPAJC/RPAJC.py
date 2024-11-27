import pandas as pd
from tkinter import filedialog as fd


def arrumarTabela():

    colunasValidas = ["Código Verba", "CNPJ", "Nº CC", "Quantidade", "Valor"]

    print('Iniciando processo...')
    while True:

        print('Selecione um arquivo .XLSX')
        arquivoExcel = fd.askopenfilename()
        try: 
            dataFrame = pd.read_excel(arquivoExcel, usecols=colunasValidas, decimal=',')
            break

        except ValueError as err:
            print('Erro ao ler o arquivo!')
            print(err)

    print('Arquivo lido com sucesso!')
    print('Arrumando tabela.')

    dataFrame['Quantidade']  = dataFrame['Quantidade'].astype(float)
    dataFrame['Valor']       = dataFrame['Valor'].astype(float)

    dataFrameAgrupado = dataFrame.groupby(
        ["Código Verba", "CNPJ", "Nº CC"], as_index=False).agg({'Quantidade': 'sum','Valor': 'sum'})

    dataFrameAgrupado['Valor'] = dataFrameAgrupado['Valor'].round(decimals = 2)
    dataFrameAgrupado['Valor'] = dataFrameAgrupado['Valor'].apply(str)
    dataFrameAgrupado['Valor'] = dataFrameAgrupado['Valor'].apply(lambda x: x.replace('.' , ','))

    dataFrameAgrupado['Quantidade'] = dataFrameAgrupado['Quantidade'].round(decimals = 2)
    dataFrameAgrupado['Quantidade'] = dataFrameAgrupado['Quantidade'].apply(str)
    dataFrameAgrupado['Quantidade'] = dataFrameAgrupado['Quantidade'].apply(lambda x: x.replace('.' , ','))

    dataFrameAgrupado.insert(0,"CPF", "")
    dataFrameAgrupado.insert(1,"Tabela", numeroTabela)

    dataFrameAgrupado.rename(
        columns={'Código Verba': 'Evento', 'Quantidade': 'Referencia', 'CNPJ':'CNPJ Loja'}, inplace=True)

    dataFrameFormatado = dataFrameAgrupado.loc[:,['CPF','Tabela', 'Evento', 'Referencia', 'Valor', 'Nº CC', 'CNPJ Loja']]

    dataFrameFormatado.to_csv('planilhaFormatada.csv', sep = ';', index = False, header = False)


if __name__ == "__main__":

    print('Inicializando...')
    
    while True:

        numeroTabela = input('Digite o numero da tabela: ')

        if (numeroTabela.isnumeric() is True) and (len(numeroTabela) < 5):
            break
        else:
            print("Número de tabela inválido. tente novamente.")

    arrumarTabela()

    print('Processo concluído!')
    input("Pressione Enter para Encerrar...")
