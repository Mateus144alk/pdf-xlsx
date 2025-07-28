import pandas as pd

# Caminho do arquivo de entrada
arquivo_entrada = 'Ferias pagas fl dez e a partir jan25 - Copia.xlsx'  # Altere conforme necessário

# Lê a planilha
df = pd.read_excel(arquivo_entrada)

# Remove espaços extras dos nomes das colunas (caso existam)
df.columns = df.columns.str.strip()

# Converte a coluna 'RENDIM' para número (float), tratando vírgulas e pontos
df['RENDIM'] = df['RENDIM'].astype(str).str.replace('.', '', regex=False)
df['RENDIM'] = df['RENDIM'].str.replace(',', '.', regex=False)
df['RENDIM'] = pd.to_numeric(df['RENDIM'], errors='coerce').fillna(0)

# Agrupa por SIAPE e NOME, somando os rendimentos
df_somado = df.groupby(['SIAPE', 'NOME'], as_index=False)['RENDIM'].sum()

# Ordena por NOME (em ordem alfabética)
df_somado = df_somado.sort_values(by='NOME')

# Exporta para uma nova planilha Excel
arquivo_saida = 'rendimentos.xlsx'
df_somado.to_excel(arquivo_saida, index=False)

print(f'Arquivo salvo como: {arquivo_saida}')
