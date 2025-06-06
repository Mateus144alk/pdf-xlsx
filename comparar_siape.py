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

print(f"Usando coluna '{nome_coluna_36}' da planilha de 36 registros")
print(f"Usando coluna '{nome_coluna_671}' da planilha de 671 registros")

# Obter valores únicos de cada planilha
valores_36 = set(df_36[nome_coluna_36].unique())
valores_671 = set(df_671[nome_coluna_671].unique())

# Encontrar diferenças
diferencas_671 = valores_671 - valores_36  # Só na de 671
diferencas_36 = valores_36 - valores_671   # Só na de 36

# Filtrar os dataframes originais
df_diferencas_671 = df_671[df_671[nome_coluna_671].isin(diferencas_671)]
df_diferencas_36 = df_36[df_36[nome_coluna_36].isin(diferencas_36)]

# Salvar os resultados
caminho_saida = r'C:/Ler arquivo/diferencas.xlsx'

with pd.ExcelWriter(caminho_saida) as writer:
    df_diferencas_671.to_excel(writer, sheet_name='Só_na_671', index=False)
    df_diferencas_36.to_excel(writer, sheet_name='Só_na_36', index=False)

print("\nProcesso concluído. Resultados salvos em:")
print(caminho_saida)
print(f"\nResumo:")
print(f"- Total na planilha de 36: {len(df_36)} registros")
print(f"- Total na planilha de 671: {len(df_671)} registros")
print(f"- Valores únicos na de 36: {len(valores_36)}")
print(f"- Valores únicos na de 671: {len(valores_671)}")
print(f"- Valores só na de 671: {len(diferencas_671)}")
print(f"- Valores só na de 36: {len(diferencas_36)}")