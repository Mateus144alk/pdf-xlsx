import pandas as pd
def comparar_planilhas(caminho_abril, caminho_maio, caminho_saida):
    df_abr = pd.read_excel(caminho_abril)
    df_mai = pd.read_excel(caminho_maio)

    # Garante que só os dados numéricos sejam usados na diferença
    colunas_excluir = ["Nome", "Matrícula"]
    colunas_valores = [col for col in df_abr.columns if col not in colunas_excluir]

    # Corrige tipos
    for col in colunas_valores:
        df_abr[col] = pd.to_numeric(df_abr[col], errors="coerce").fillna(0)
        df_mai[col] = pd.to_numeric(df_mai[col], errors="coerce").fillna(0)

    # Merge por matrícula
    df_comparado = pd.merge(df_abr, df_mai, on="Matrícula", suffixes=(" (Abr)", " (Mai)"))

    # Pega nome do mês mais recente (ou qualquer um)
    df_comparado["Nome"] = df_comparado["Nome (Mai)"].combine_first(df_comparado["Nome (Abr)"])

    # Calcula as diferenças
    for col in colunas_valores:
        col_abr = f"{col} (Abr)"
        col_mai = f"{col} (Mai)"
        df_comparado[f"{col} (Diferença)"] = df_comparado[col_mai] - df_comparado[col_abr]

    # Reorganiza as colunas
    ordem = ["Matrícula", "Nome"]
    for col in colunas_valores:
        ordem.extend([
            f"{col} (Abr)", f"{col} (Mai)", f"{col} (Diferença)"
        ])

    df_final = df_comparado[ordem]
    df_final.to_excel(caminho_saida, index=False)
    print(f"\n✅ Comparação gerada com sucesso em: {caminho_saida}")





comparar_planilhas(
    caminho_abril="C:/Ler arquivo/pdf-xlsx/pdf-abril/planilha_abril.xlsx",
    caminho_maio="C:/Ler arquivo/pdf-xlsx/pdf-maio/planilha_maio.xlsx",
    caminho_saida="C:/Ler arquivo/pdf-xlsx/comparativo_abril_maio.xlsx"
)
