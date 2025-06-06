import pandas as pd

# Carregar as planilhas
caminho_planilha_36 = r'C:/Users/mateusibanhes/Desktop/a MATEUS - retroativos DPP/ausentes_siape_por_siape.xlsx'
caminho_planilha_671 = r'C:/Users/mateusibanhes/Desktop/a MATEUS - retroativos DPP/TODOS - PROGRESSÃO MÉRITO.ods'

# Carregar os dados
df_36 = pd.read_excel(caminho_planilha_36)
df_671 = pd.read_excel(caminho_planilha_671)

# Pegar o nome da segunda coluna (coluna A2) de cada dataframe
nome_coluna_36 = df_36.columns[1]  # Índice 1 para a segunda coluna
nome_coluna_671 = df_671.columns[1]

# Obter valores únicos
valores_36 = set(df_36[nome_coluna_36].unique())
valores_671 = set(df_671[nome_coluna_671].unique())

# Encontrar os 7 valores que estão apenas na planilha de 36
valores_apenas_36 = valores_36 - valores_671

# Filtrar o dataframe original para pegar apenas esses 7 registros
df_apenas_36 = df_36[df_36[nome_coluna_36].isin(valores_apenas_36)]

# Salvar os 7 registros em um arquivo separado
caminho_apenas_36 = r'C:/Ler arquivo/apenas_na_36.xlsx'
df_apenas_36.to_excel(caminho_apenas_36, index=False)

print(f"Os 7 valores que estão apenas na planilha de 36 foram salvos em:")
print(caminho_apenas_36)
print("\nConteúdo salvo:")
print(df_apenas_36)