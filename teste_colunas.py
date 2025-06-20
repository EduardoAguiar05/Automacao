import pandas as pd

try:
    print("Lendo planilha...")
    df = pd.read_excel("relatório - notas(novo).xlsx")
    
    colunas = ['CÓDIGO DO ITEM', 'NÚMERO DO DOCUMENTO', 'NÚMERO PROJETO', 'DESCRIÇÃO DO ITEM']
    
    print("\nTestando cada coluna:")
    for coluna in colunas:
        print(f"\nColuna: {coluna}")
        print(f"Tipo de dados: {df[coluna].dtype}")
        print(f"Tem valores nulos? {df[coluna].isnull().any()}")
        print(f"Primeiros 5 valores: {df[coluna].head().tolist()}")
        
        if df[coluna].dtype != 'object':  # Se não for texto
            try:
                print("Tentando converter para inteiro...")
                valores_int = df[coluna].fillna(0).astype(int)
                print("Conversão para inteiro bem sucedida!")
            except Exception as e:
                print(f"Erro na conversão: {str(e)}")
                print("Valores problemáticos:")
                print(df[~df[coluna].fillna(0).astype(str).str.match(r'^-?\d+$')][coluna].unique())
        
except Exception as e:
    print(f"Erro: {str(e)}") 