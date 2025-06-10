import pandas as pd
import os

def buscar_e_copiar_linhas(arquivo_origem, arquivo_destino, palavras_chave, palavras_excluir):
    """
    Função para buscar linhas com múltiplas palavras-chave específicas e adicionar em uma planilha existente
    Copia todas as colunas disponíveis das linhas que contêm as palavras-chave, excluindo linhas com palavras indesejadas
    
    Parâmetros:
    arquivo_origem: caminho do arquivo Excel de origem
    arquivo_destino: caminho do arquivo Excel de destino
    palavras_chave: lista de palavras que serão buscadas
    palavras_excluir: lista de palavras que devem ser excluídas
    """
    try:
        print(f"Tentando abrir o arquivo: {arquivo_origem}")
        # Primeiro vamos ler o arquivo para ver quantas colunas ele tem
        df_origem = pd.read_excel(arquivo_origem)
        print(f"Arquivo aberto com sucesso. Total de colunas: {len(df_origem.columns)}")
        
        # Dicionário para guardar quais palavras-chave encontraram cada linha
        linhas_por_palavra = {}
        
        # Para cada palavra-chave, procura em todas as colunas
        for palavra in palavras_chave:
            mascara_palavra = pd.Series(False, index=df_origem.index)
            for coluna in df_origem.columns:
                # Converte a coluna para string e adiciona espaços no início e fim
                valores_coluna = " " + df_origem[coluna].astype(str).str.upper() + " "
                # Adiciona espaços na palavra-chave para busca exata
                palavra_busca = f" {palavra} "
                mascara = valores_coluna.str.contains(palavra_busca, case=False, na=False)
                mascara_palavra = mascara_palavra | mascara
            
            # Exclui linhas que contêm palavras indesejadas
            for palavra_excluir in palavras_excluir:
                for coluna in df_origem.columns:
                    valores_coluna = " " + df_origem[coluna].astype(str).str.upper() + " "
                    palavra_excluir = f" {palavra_excluir} "
                    mascara_excluir = valores_coluna.str.contains(palavra_excluir, case=False, na=False)
                    mascara_palavra = mascara_palavra & ~mascara_excluir
            
            # Guarda as linhas encontradas para esta palavra
            linhas_encontradas_palavra = df_origem[mascara_palavra]
            if len(linhas_encontradas_palavra) > 0:
                print(f"\nPalavra '{palavra}' encontrou {len(linhas_encontradas_palavra)} linhas:")
                for idx, linha in linhas_encontradas_palavra.iterrows():
                    descricao = str(linha.iloc[0])
                    if len(descricao) > 100:
                        descricao = descricao[:100] + "..."
                    print(f"- {descricao}")
                linhas_por_palavra[palavra] = linhas_encontradas_palavra
        
        # Combina todas as linhas encontradas
        if len(linhas_por_palavra) > 0:
            linhas_encontradas = pd.concat(linhas_por_palavra.values()).drop_duplicates()
        else:
            print("Nenhuma linha encontrada com as palavras-chave especificadas")
            return
            
        # Verifica se o arquivo de destino já existe
        if os.path.exists(arquivo_destino):
            print(f"\nArquivo de destino encontrado: {arquivo_destino}")
            # Se existe, lê o arquivo existente
            df_existente = pd.read_excel(arquivo_destino)
            # Concatena as novas linhas com as existentes
            df_final = pd.concat([df_existente, linhas_encontradas], ignore_index=True)
        else:
            print(f"\nCriando novo arquivo: {arquivo_destino}")
            # Se não existe, usa apenas as linhas encontradas
            df_final = linhas_encontradas
        
        # Remove duplicatas baseado em todas as colunas
        df_final = df_final.drop_duplicates()
        
        # Salva o resultado final no arquivo Excel
        df_final.to_excel(arquivo_destino, index=False)
        print(f"\nForam encontradas {len(linhas_encontradas)} novas linhas com as palavras-chave")
        print(f"Após remover duplicatas, ficaram {len(df_final)} linhas")
        print(f"As linhas foram salvas no arquivo: {arquivo_destino}")
        
    except Exception as e:
        print(f"Ocorreu um erro: {str(e)}")
        print(f"Diretório atual: {os.getcwd()}")
        print("Arquivos na pasta:")
        for arquivo in os.listdir():
            print(f"- {arquivo}")

# Exemplo de uso
if __name__ == "__main__":
    # Configurações com os nomes exatos dos arquivos
    arquivo_origem = "relatório - notas(novo)"  # Nome do arquivo de origem
    arquivo_destino = "itens de escalada"  # Nome do arquivo de destino
    
    # Adiciona a extensão .xlsx aos arquivos
    arquivo_origem = arquivo_origem + ".xlsx"
    arquivo_destino = arquivo_destino + ".xlsx"
    
    print(f"Iniciando processamento...")
    print(f"Arquivo de origem: {arquivo_origem}")
    print(f"Arquivo de destino: {arquivo_destino}")
    
    # Lista de palavras-chave para buscar
    palavras_chave = [
        "CINTO DE SEGURANÇA",
        "CAPACETE ALPINISTA",
        "DESCENSOR",
        "CORTADOR DE FIOS",
        "MOSQUETÃO",
        "ASCENSOR",
        "TRAVAQUEDA",
        "PROTETOR DE CORDA",
        "ESTRIBO",
        "CORDA",
        "RIG DESCENSOR",
        "POLIA",
        "VARA"
    ]
    
    # Lista de palavras que devem ser excluídas (relacionadas à refrigeração)
    palavras_excluir = [
        "REFRIGERAÇÃO",
        "REFRIGERANTE",
        "TUBO DE COBRE",
        "TUBO ESPONJOSO"
    ]
    
    # Executa a busca e cópia
    buscar_e_copiar_linhas(arquivo_origem, arquivo_destino, palavras_chave, palavras_excluir) 