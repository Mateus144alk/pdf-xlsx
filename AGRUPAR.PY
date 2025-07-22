import pandas as pd

# Lê o arquivo Excel (ajuste o nome conforme necessário)
df = pd.read_excel('seu_arquivo.xlsx')

# Verifica quais colunas existem no DataFrame
print("Colunas disponíveis no arquivo:")
print(df.columns.tolist())

# Supondo que a coluna de valores se chame 'RENDIMENTO' (ajuste conforme seu arquivo)
# Primeiro precisamos extrair os valores numéricos da coluna de texto
def extrair_valor(texto):
    try:
        # Encontra o último conjunto de números com vírgula/separador decimal
        import re
        valor_str = re.findall(r'(\d[\d.,]*\d)', str(texto))[-1]
        return float(valor_str.replace('.', '').replace(',', '.'))
    except:
        return 0.0

# Cria uma coluna com os valores numéricos
df['VALOR_NUMERICO'] = df.iloc[:, -1].apply(extrair_valor)  # Assume que o valor está na última coluna

# Agora agrupamos os dados
df_consolidado = df.groupby(['NOME SERVIDOR', 'SIAPE SERVIDOR', 'CARGO', 'GR CLAS CARGO', 'NÍVEL CARGO', 'SITUAÇÃO VÍNCULO']).agg({
    'RUBRICA': lambda x: ', '.join(x.unique()),  # Lista as rubricas
    'VALOR_NUMERICO': 'sum'  # Soma os valores
}).reset_index()

# Renomeia a coluna de valor somado
df_consolidado = df_consolidado.rename(columns={'VALOR_NUMERICO': 'TOTAL RENDIMENTOS'})

# Formata o valor para o padrão brasileiro
df_consolidado['TOTAL RENDIMENTOS'] = df_consolidado['TOTAL RENDIMENTOS'].apply(
    lambda x: f'R$ {x:,.2f}'.replace('.', '|').replace(',', '.').replace('|', ',')
)

# Salva o resultado
df_consolidado.to_excel('servidores_consolidados_r.xlsx', index=False)

print("✅ Arquivo consolidado gerado com sucesso!")
print(f"Total de servidores consolidados: {len(df_consolidado)}")