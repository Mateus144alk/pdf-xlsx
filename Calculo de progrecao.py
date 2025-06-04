import pandas as pd
import pdfplumber
import os
import re
from tkinter import filedialog, Tk, messagebox


def extrair_texto_pdf(caminho_pdf):
    texto_total = ""
    with pdfplumber.open(caminho_pdf) as pdf:
        for page in pdf.pages:
            texto = page.extract_text()
            if texto:
                texto_total += texto + "\n"
    return texto_total


def comparar_planilha_com_pdf(planilha, pdfs, coluna_nome="B", coluna_siape="A", linha_inicial=4):
    df = pd.read_excel(planilha, dtype=str)
    nomes_planilha = df.iloc[linha_inicial-1:, 0].fillna("").str.upper().str.strip()   # Coluna A = Nome
    siapes_planilha = df.iloc[linha_inicial-1:, 1].fillna("").astype(str).str.strip() # Coluna B = SIAPE

    entradas_planilha = list(zip(siapes_planilha, nomes_planilha))

    texto_pdf = ""
    for pdf in pdfs:
        texto_pdf += extrair_texto_pdf(pdf).upper()

    # ➕ extrai a data base do texto
    dia, mes = extrair_datas_flexivel(texto_pdf)
    print("Data encontrada:", dia, mes)
    relatorio_nao_encontrados = []
    resultados = []

    for siape, nome in entradas_planilha:
        if not siape or not nome:
            continue
        encontrado = siape in texto_pdf or nome in texto_pdf
        if encontrado:
            resultados.append({"SIAPE": siape, "Nome": nome, "Dia": dia, "Mês": mes})
        else:
            relatorio_nao_encontrados.append(f"Não encontrado: {siape} - {nome}")

    if resultados:
        df_resultado = pd.DataFrame(resultados)
        df_resultado.to_excel("resultado_comparacao.ods", index=False)
        print("✅ Arquivo 'resultado_comparacao.ods' gerado com sucesso!")
    else:
        print("⚠️ Nenhum resultado encontrado.")

    if relatorio_nao_encontrados:
        with open("nao_encontrados.txt", "w", encoding="utf-8") as f:
            f.write("\n".join(relatorio_nao_encontrados))
        print("📄 Lista de não encontrados salva em 'nao_encontrados.txt'")


def extrair_datas_flexivel(texto_pdf):
    meses = {
        "JAN": "01", "FEV": "02", "MAR": "03", "ABR": "04",
        "MAI": "05", "JUN": "06", "JUL": "07", "AGO": "08",
        "SET": "09", "OUT": "10", "NOV": "11", "DEZ": "12"
    }

    linhas = texto_pdf.upper().splitlines()
    datas = []

    for i in range(len(linhas)):
        linha = linhas[i].strip()

        # Padrão 1: tudo junto em uma linha
        match1 = re.search(r"(\d{1,2})[ ]*(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)/?20\d{2}", linha)
        if match1:
            dia = match1.group(1).zfill(2)
            mes = meses.get(match1.group(2), "")
            datas.append((dia, mes))
            continue

        # Padrão 2: número separado e linha seguinte com MM/YYYY
        if re.fullmatch(r"\d{1,2}", linha) and i + 1 < len(linhas):
            proxima = linhas[i+1].strip()
            match2 = re.fullmatch(r"(\d{2})/20\d{2}", proxima)
            if match2:
                dia = linha.zfill(2)
                mes = match2.group(1)
                datas.append((dia, mes))
                continue

        # Padrão 3: 3 linhas (dia, mês abreviado, ano)
        if re.fullmatch(r"\d{1,2}", linha) and i + 2 < len(linhas):
            mes_linha = linhas[i+1].strip().replace("/", "")
            ano_linha = linhas[i+2].strip()
            if mes_linha in meses and re.search(r"20\d{2}", ano_linha):
                dia = linha.zfill(2)
                mes = meses[mes_linha]
                datas.append((dia, mes))
                continue

    return datas[0] if datas else ("", "")






# --- Execução simples via Tkinter ---
if __name__ == "__main__":
    Tk().withdraw()
    caminho_planilha = filedialog.askopenfilename(title="Selecione a planilha Excel", filetypes=[("Excel", "*.ods")])
    arquivos_pdfs = filedialog.askopenfilenames(title="Selecione os arquivos PDF", filetypes=[("PDFs", "*.pdf")])

    if not caminho_planilha or not arquivos_pdfs:
        messagebox.showwarning("Aviso", "É necessário selecionar uma planilha e pelo menos um PDF.")
    else:
        comparar_planilha_com_pdf(caminho_planilha, arquivos_pdfs)