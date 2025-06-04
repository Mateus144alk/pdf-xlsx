import pandas as pd

# Caminhos das planilhas
planilha_certa = "C:/Ler arquivo/pdf-xlsx/3 - Pgto Retroativo PROGRESSÕES.ods"
planilha_comparar = "C:/Ler arquivo/pdf-xlsx/datas_extraidas.ods"

# Leitura das planilhas
df_certo = pd.read_excel(planilha_certa, engine="odf", header=1)
df_comparar = pd.read_excel(planilha_comparar, engine="odf")

# Padronizar os nomes das colunas
df_certo.columns = [str(col).strip().upper() for col in df_certo.columns]
df_comparar.columns = [col.strip().upper() for col in df_comparar.columns]

# Encontrar colunas de interesse
col_nome = next((col for col in df_certo.columns if 'NOME' in col.upper()), None)
col_siape = next((col for col in df_certo.columns if 'SIAPE' in col.upper()), None)
if not col_nome or not col_siape:
    raise ValueError("Não foi possível encontrar as colunas NOME e SIAPE na planilha correta.")

# Selecionar e renomear
df_certo = df_certo[[col_nome, col_siape]]
df_certo.columns = ['NOME', 'SIAPE']
df_comparar = df_comparar[['SIAPE', 'NOME', 'DIA', 'MÊS']]

# --- 🔄 PADRONIZAÇÃO ---

def limpar_siape(valor):
    if pd.isna(valor):
        return ''
    return ''.join(filter(str.isdigit, str(valor))).strip()
print(f"🔢 Total na planilha correta: {len(df_certo)}")
print(f"🔍 Total na planilha de comparação: {len(df_comparar)}")

df_certo['SIAPE'] = df_certo['SIAPE'].apply(limpar_siape)
df_comparar['SIAPE'] = df_comparar['SIAPE'].apply(limpar_siape)

df_certo['NOME'] = df_certo['NOME'].astype(str).str.strip().str.upper()
df_comparar['NOME'] = df_comparar['NOME'].astype(str).str.strip().str.upper()

# Remover linhas com SIAPE vazio
df_certo = df_certo[df_certo['SIAPE'] != '']
df_comparar = df_comparar[df_comparar['SIAPE'] != '']

# --- 🔍 COMPARAÇÕES ---
siapes_certo = set(df_certo['SIAPE'])
siapes_comparar = set(df_comparar['SIAPE'])

# SIAPEs que estão em comparar, mas não em certo
siapes_a_mais = siapes_comparar - siapes_certo
registros_a_mais = df_comparar[df_comparar['SIAPE'].isin(siapes_a_mais)]

# SIAPEs com nome diferente
erros = []
for _, row in df_comparar.iterrows():
    siape = row['SIAPE']
    nome = row['NOME']
    if siape in siapes_certo:
        nome_correto = df_certo[df_certo['SIAPE'] == siape]['NOME'].values[0]
        if nome != nome_correto:
            erros.append(row)

df_erros = pd.DataFrame(erros)

# --- 💾 SALVAR RESULTADOS ---
registros_a_mais.to_excel("registros_a_mais.ods", index=False, engine="odf")
df_erros.to_excel("erros_nomes_ou_siape.ods", index=False, engine="odf")

print("✔ Corrigido e arquivos gerados com sucesso:")
print(f"- SIAPEs a mais: {len(registros_a_mais)} registros")
print(f"- Nomes divergentes: {len(df_erros)} registros")
