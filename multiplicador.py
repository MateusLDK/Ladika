import pandas as pd

from tkinter import filedialog as fd

folderPath = fd.askopenfilename()

df = pd.read_excel(folderPath)


# Estrutura de Tamanho e Curva
tamanho_curva = {
    'Tamanho': ['G', 'G', 'G', 'M', 'M', 'M', 'P', 'P', 'P', 'PP', 'PP', 'PP'],
    'Curva': ['A', 'B', 'C', 'A', 'B', 'C', 'A', 'B', 'C', 'A', 'B', 'C']
}

# Criar DataFrame da estrutura
df_tamanho_curva = pd.DataFrame(tamanho_curva)
# Repetir a estrutura de Tamanho e Curva para cada linha do DataFrame original


df_tamanho_curva_repetido = pd.concat([df_tamanho_curva] * len(df), ignore_index=True)
print(len(df_tamanho_curva_repetido['Tamanho']))

# Replicar cada linha do DataFrame original 11 vezes
df_replicado = df.loc[df.index.repeat(12)].reset_index(drop=True)
print(len(df_replicado['Departamento']))

# Concatenar o DataFrame replicado com a estrutura de Tamanho e Curva
df_final = pd.concat([df_replicado, df_tamanho_curva_repetido], axis=1)

# Exibir o resultado
#print(df_final)
df_final.to_excel(excel_writer="arquivo formato onebeat.xlsx", index = False)