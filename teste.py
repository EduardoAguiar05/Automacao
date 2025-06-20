import pandas as pd

# Ler a planilha de origem
df = pd.read_excel("relatório - notas(novo).xlsx")

# Mostrar informações sobre o DataFrame
print("Informações do DataFrame:")
print(df.info())

print("\nPrimeiras 5 linhas:")
print(df.head())

print("\nVerificando valores nulos:")
print(df.isnull().sum())

print("\nTipos de dados das colunas:")
print(df.dtypes)

# Verificar especificamente as colunas numéricas
colunas_numericas = ['CÓDIGO DO ITEM', 'NÚMERO DO DOCUMENTO', 'NÚMERO PROJETO']
print("\nValores únicos nas colunas numéricas:")
for coluna in colunas_numericas:
    if coluna in df.columns:
        print(f"\n{coluna}:")
        print(df[coluna].unique()) 