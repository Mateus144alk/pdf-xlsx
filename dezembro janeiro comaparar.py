import pandas as pd

# Arquivos de entrada
arquivo_janeiro = 'janeiro.xlsx'
arquivo_dezembro = 'dezembro.xlsx'

# Janeiro: lê SIAPE (A), NOME (B), VALOR (D)
df_jan = pd.read_excel(
    arquivo_janeiro,
    usecols="A,B,D",
    dtype={"A": str},  # SIAPE como string
    header=0
)

# Dezembro: mesma estrutura
df_dez = pd.read_excel(
    arquivo_dezembro,
    usecols="A,B,D",
    dtype={"A": str},  # SIAPE como string
    header=0
)

# Renomeia as colunas
df_jan.columns = ["SIAPE", "NOME", "VALOR_JANEIRO"]
df_dez.columns = ["SIAPE", "NOME", "VALOR_DEZEMBRO"]

# Função para limpar e converter valores
def clean_currency(value):
    if isinstance(value, str):
        return float(value.replace(".", "").replace(",", "."))
    return float(value)

# Converte os valores
df_jan["VALOR_JANEIRO"] = df_jan["VALOR_JANEIRO"].apply(clean_currency)
df_dez["VALOR_DEZEMBRO"] = df_dez["VALOR_DEZEMBRO"].apply(clean_currency)

# Faz o merge completo (outer) para manter todos os registros de ambos os meses
df_final = pd.merge(
    df_jan,
    df_dez,
    on=["SIAPE", "NOME"],
    how="outer",
    suffixes=('', '_y')
)

# Remove colunas duplicadas se houver
df_final = df_final.loc[:,~df_final.columns.duplicated()]

# Ordena as colunas como solicitado
df_final = df_final[["SIAPE", "NOME", "VALOR_JANEIRO", "VALOR_DEZEMBRO"]]

# Salva o resultado
df_final.to_excel("comparativo_jan_dez.xlsx", index=False, float_format="%.2f")

print("✅ Comparativo gerado com sucesso: comparativo_jan_dez.xlsx")
print("Total de registros em Janeiro:", len(df_jan))
print("Total de registros em Dezembro:", len(df_dez))
print("Total no arquivo final:", len(df_final))