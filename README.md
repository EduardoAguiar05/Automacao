# Automação de Busca em Planilhas Excel

Este projeto contém um script Python que automatiza a busca de itens específicos em uma planilha Excel e copia as linhas encontradas para outra planilha.

## Funcionalidades

- Busca itens específicos em uma planilha Excel usando palavras-chave
- Ignora itens indesejados através de uma lista de exclusão
- Copia as linhas encontradas para uma nova planilha
- Remove duplicatas automaticamente
- Preserva todas as colunas dos itens encontrados

## Como usar

1. Instale as dependências necessárias:
```bash
pip install pandas openpyxl
```

2. Coloque os arquivos na mesma pasta:
   - `automacao_excel.py` (o script Python)
   - `relatório - notas(novo).xlsx` (planilha de origem)
   - `itens de escalada.xlsx` (planilha de destino)

3. Execute o script:
```bash
python automacao_excel.py
```

## Configuração

No arquivo `automacao_excel.py`, você pode configurar:

1. Palavras-chave para busca (lista `palavras_chave`)
2. Palavras para excluir (lista `palavras_excluir`)
3. Nomes dos arquivos de origem e destino

## Requisitos

- Python 3.x
- pandas
- openpyxl 