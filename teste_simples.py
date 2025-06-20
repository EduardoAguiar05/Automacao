import pandas as pd

try:
    print("Tentando ler a planilha...")
    df = pd.read_excel("relatório - notas(novo).xlsx")
    print("Planilha lida com sucesso!")
    print(f"Total de linhas: {len(df)}")
    print(f"Colunas encontradas: {df.columns.tolist()}")
except Exception as e:
    print(f"Erro ao ler planilha: {str(e)}") 