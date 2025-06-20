import pandas as pd

try:
    print("Lendo planilha...")
    df = pd.read_excel("relatório - notas(novo).xlsx")
    print(f"Total de linhas: {len(df)}")
    
    print("\nÚltimas 10 linhas da planilha:")
    ultimas_linhas = df.tail(10)
    for idx, linha in ultimas_linhas.iterrows():
        print(f"\nLinha {idx}:")
        print(f"Código: {linha['CÓDIGO DO ITEM']}")
        print(f"Descrição: {linha['DESCRIÇÃO DO ITEM']}")
        print(f"NF: {linha['NÚMERO DO DOCUMENTO']}")
        print(f"Projeto: {linha['NÚMERO PROJETO']}")
        
except Exception as e:
    print(f"Erro ao ler planilha: {str(e)}") 