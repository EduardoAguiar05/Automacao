import pandas as pd

try:
    print("Lendo planilha 'itens de escalada.xlsx'...")
    df = pd.read_excel("itens de escalada.xlsx")
    print(f"Total de linhas: {len(df)}")
    
    # Verificar se a linha existe
    linha_alvo = 42165
    if linha_alvo >= len(df):
        print(f"\nA linha {linha_alvo} não existe na planilha!")
        print(f"A planilha tem apenas {len(df)} linhas.")
        
        print("\nMostrando as últimas 5 linhas da planilha:")
        ultimas_linhas = df.tail()
        for idx, linha in ultimas_linhas.iterrows():
            print(f"\nLinha {idx}:")
            for coluna in df.columns:
                print(f"{coluna}: {linha[coluna]}")
    else:
        print(f"\nAnalisando linha {linha_alvo}:")
        linha = df.iloc[linha_alvo]
        print("\nValores da linha:")
        for coluna in df.columns:
            valor = linha[coluna]
            tipo = type(valor)
            print(f"{coluna}: {valor} (tipo: {tipo})")
            
        print("\nTentando converter valores:")
        try:
            codigo = int(linha['CÓDIGO DO ITEM'])
            nota = int(linha['NÚMERO DO DOCUMENTO'])
            projeto = int(linha['NÚMERO PROJETO'])
            print("Conversões bem sucedidas:")
            print(f"Código: {codigo}")
            print(f"Nota: {nota}")
            print(f"Projeto: {projeto}")
        except Exception as e:
            print(f"Erro na conversão: {str(e)}")
            
except Exception as e:
    print(f"Erro ao ler planilha: {str(e)}") 