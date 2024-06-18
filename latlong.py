from geopy.geocoders import Nominatim
import pandas as pd
from tkinter import filedialog as fd
import requests

filePath = fd.askopenfilename()

dfEnderecos = pd.read_excel(filePath)
dfEnderecos.insert(0,"Bairro"," ")

i=0
for cep in dfEnderecos["CEP ENTREGA"]:

    if len(str(cep)) == 8:

        link = f'https://viacep.com.br/ws/{cep}/json/'

        requisicao = requests.get(link)
        dic_requisicao = requisicao.json()
        try:
            bairro = dic_requisicao['bairro']
            print(bairro)
            dfEnderecos.at[i,'BAIRRO'] = bairro
        except KeyError:
            pass
        
        i+=1

    else:
        print("CEP Inválido")

dfEnderecos.to_excel('planilhaBravo2.xlsx')