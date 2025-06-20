import pandas as pd

# Altere para o nome correto da sua planilha e aba
arquivo = "relatório - notas(novo).xlsx"
linha_alvo = 41989  # pandas é zero-based, então linha 41990 no Excel é 41989 no pandas

# Lê a planilha
df = pd.read_excel(arquivo, sheet_name="notas_com_soli")

print("Nomes das colunas lidos pelo pandas:")
for col in df.columns:
    print(f"- '{col}'")

print(f"\nValores da linha {linha_alvo + 1} (índice pandas {linha_alvo}):")
linha = df.iloc[linha_alvo]
for col in df.columns:
    print(f"{col}: {linha[col]} (tipo: {type(linha[col])})") 