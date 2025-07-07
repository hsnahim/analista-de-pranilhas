import pandas as pd

# Carrega o arquivo Excel
xl_file = pd.ExcelFile('CONTROLE ESTAÇÃO.ATUAL.xlsx')

# Lê todas as abas da planilha para um dicionário
dfs = {sheet_name: xl_file.parse(sheet_name) for sheet_name in xl_file.sheet_names}

# Itera e imprime o conteúdo de cada aba
for sheet_name, df in dfs.items():
    print(f"\n--- ABA: {sheet_name} ---")
    print(df)
