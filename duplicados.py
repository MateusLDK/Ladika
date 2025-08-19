import pandas as pd
from tkinter import filedialog as fd

arquivo= fd.askopenfilename()

df = pd.read_excel(arquivo)

df_uniq = df.drop_duplicates(subset=['Codigo', 'Loja'], keep='last')

# Salva o resultado
salvo = arquivo.replace('.xlsx', '_sem_duplicados.xlsx')
df_uniq.to_excel(salvo, index=False)
print(f'Arquivo salvo como {salvo}')
