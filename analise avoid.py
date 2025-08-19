import pandas as pd

colunas = ['Codigo', 'Loja']

print("iniciando leitura dos arquivos...")
dfAvoid = pd.read_csv('avoid.csv', sep=',', encoding='latin1' )
dfSaldo = pd.read_csv('estoque.csv', sep=';', usecols=['produto_key','loja_key'])
dfSaldo.rename(columns={'produto_key': 'Codigo', 'loja_key': 'Loja'}, inplace=True)
print("leitura dos arquivos concluída.")

merged = dfAvoid.merge(
    dfSaldo,
    on=['Codigo', 'Loja'],
    how='left',
    indicator=True
)
dfAvoid_filtrado = merged[merged['_merge'] == 'left_only'].drop(columns=['_merge'])

dfAvoid_filtrado.to_csv('avoid_filtrado.csv', index=False)
