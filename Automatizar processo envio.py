import pandas as pd
from datetime import datetime

rubricas_df = pd.read_excel("rubricas.xlsx")  # MatSiape, DVSiape, Valor Atual, Valor Antigo, Rubrica, etc.
retro_df = pd.read_excel("retroativos.xlsx")  # SIAPE, Nome, Data Retroativa
extrator_df = pd.read_excel("extrator_siape.xlsx")  # SIAPE, MatriculaOrigem, DV

# 2. Calcular valor da diferença (Atual - Antigo)
rubricas_df["ValorDif"] = rubricas_df["ValorAtual"] - rubricas_df["ValorAntigo"]

# 3. Juntar com a data retroativa
base_df = rubricas_df.merge(retro_df, left_on="MatSiape", right_on="SIAPE", how="left")

# 4. Calcular valor retroativo proporcional por dias
def calcular_proporcional(row, data_folha):
    try:
        dias = (data_folha - pd.to_datetime(row["Data Retroativa"])).days
        meses = dias / 30
        return round(row["ValorDif"] * meses, 2)
    except:
        return 0

data_folha_pgto = datetime(2025, 6, 1)  # altere conforme necessário
base_df["ValorRetroativo"] = base_df.apply(lambda x: calcular_proporcional(x, data_folha_pgto), axis=1)

# 5. Adicionar dados do extrator para compor MatriculaOrigem e DVSiape
base_df = base_df.merge(extrator_df, left_on="MatSiape", right_on="SIAPE", how="left", suffixes=("", "_extrator"))

# 6. Gerar DataFrame final para carga
carga_batch = pd.DataFrame({
    "MatSiape": base_df["MatSiape"],
    "DVSiape": base_df["DVSiape"],
    "Comando": 4,  # padrão
    "RendimentoDesconto": 1,  # padrão
    "Rubrica": base_df["Rubrica"],
    "Sequencia": 6,  # ou input do usuário
    "Valor": base_df["ValorRetroativo"],
    "MatriculaOrigem": base_df["MatriculaOrigem"]
})

# 7. Exportar arquivo CSV para carga batch
carga_batch.to_csv("carga_batch.csv", index=False, sep=";", encoding="utf-8")

# 8. Exportar também os dados de valores retroativos
base_df[["MatSiape", "Nome", "Data Retroativa", "ValorRetroativo"]].to_excel("valores_retroativos.xlsx", index=False)

print("✅ Arquivo CSV para carga gerado com sucesso!")
