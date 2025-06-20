import pandas as pd

def processar_linha(linha):
    try:
        codigo = int(linha['CÓDIGO DO ITEM'])
        nota_fiscal = int(linha['NÚMERO DO DOCUMENTO'])
        projeto = int(linha['NÚMERO PROJETO'])
        descricao = linha['DESCRIÇÃO DO ITEM']
        
        print(f"Linha processada com sucesso:")
        print(f"  Código: {codigo} (tipo: {type(codigo)})")
        print(f"  NF: {nota_fiscal} (tipo: {type(nota_fiscal)})")
        print(f"  Projeto: {projeto} (tipo: {type(projeto)})")
        print(f"  Descrição: {descricao}")
        return True
    except Exception as e:
        print(f"Erro ao processar linha: {str(e)}")
        print("Valores da linha:")
        print(f"  CÓDIGO DO ITEM: {linha['CÓDIGO DO ITEM']} (tipo: {type(linha['CÓDIGO DO ITEM'])})")
        print(f"  NÚMERO DO DOCUMENTO: {linha['NÚMERO DO DOCUMENTO']} (tipo: {type(linha['NÚMERO DO DOCUMENTO'])})")
        print(f"  NÚMERO PROJETO: {linha['NÚMERO PROJETO']} (tipo: {type(linha['NÚMERO PROJETO'])})")
        return False

try:
    print("Lendo planilha...")
    df = pd.read_excel("relatório - notas(novo).xlsx")
    print(f"Total de linhas: {len(df)}")
    
    print("\nTestando primeiras 5 linhas:")
    for idx, linha in df.head().iterrows():
        print(f"\nProcessando linha {idx}:")
        processar_linha(linha)
        
except Exception as e:
    print(f"Erro: {str(e)}") 