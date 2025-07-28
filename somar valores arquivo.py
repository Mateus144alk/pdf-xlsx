import pandas as pd

# Caminho do arquivo de entrada
arquivo_entrada = 'Ferias pagas fl dez e a partir jan25 - Copia.xlsx'

# Lê o arquivo Excel
df = pd.read_excel(arquivo_entrada)

# Remove espaços extras nos nomes das colunas
df.columns = df.columns.str.strip()

# Garante que RENDIM seja numérico
df['RENDIM'] = pd.to_numeric(df['RENDIM'], errors='coerce').fillna(0)

# Agrupa por SIAPE e NOME e soma os rendimentos
df_somado = df.groupby(['SIAPE', 'NOME'], as_index=False)['RENDIM'].sum()

# Ordena por nome
df_somado = df_somado.sort_values(by='NOME')

# Cria uma nova coluna com o valor formatado no padrão brasileiro (vírgula decimal)
df_somado['RENDIM_FORMATADO'] = df_somado['RENDIM'].apply(
    lambda x: f'{x:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')
)

# Reorganiza as colunas: SIAPE, NOME, RENDIM, RENDIM_FORMATADO
df_somado = df_somado[['SIAPE', 'NOME', 'RENDIM', 'RENDIM_FORMATADO']]

# Salva no Excel
arquivo_saida = 'rendimentos-ponto-formatado.xlsx'
df_somado.to_excel(arquivo_saida, index=False)

print(f'✅ Arquivo salvo com vírgulas como separador decimal: {arquivo_saida}')
