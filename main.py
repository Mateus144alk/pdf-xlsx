import pdfplumber
import pandas as pd
import os

import re

def extrair_dados_pdf(path_pdf):
    dados = []
    tipo_rubrica = ""

    with pdfplumber.open(path_pdf) as pdf:
        for page in pdf.pages:
            texto = page.extract_text()
            if not texto:
                continue

            if "RUBRICA:" in texto:
                for linha in texto.splitlines():
                    if linha.strip().startswith("RUBRICA:"):
                        tipo_rubrica = linha.split("RUBRICA:")[1].strip()
                        break

            for linha in texto.splitlines():
                match = re.search(r"(.+?)\s+EST\d{2}\s+(\d{6,7})\s+R\s+0\s+([\d.,]+)\s+N", linha)
                if match:
                    nome = match.group(1).strip()
                    matricula = match.group(2)
                    valor = float(match.group(3).replace(".", "").replace(",", "."))
                    dados.append((nome, matricula, tipo_rubrica, valor))

    return dados


def consolidar_dados(lista_arquivos_pdf, caminho_saida):
    registros = []

    for arquivo in lista_arquivos_pdf:
        registros.extend(extrair_dados_pdf(arquivo))

    # Primeiro, coletar todas as rubricas únicas
    tipos_adicionais = sorted(set(rubrica for _, _, rubrica, _ in registros if "BASICO" not in rubrica.upper()))

    consolidado = {}
    for nome, matricula, rubrica, valor in registros:
        if matricula not in consolidado:
            consolidado[matricula] = {
                "Nome": nome,
                "Matrícula": matricula,
                "Salário Básico": 0,
                **{tipo: 0 for tipo in tipos_adicionais}
            }
        if "BASICO" in rubrica.upper():
            consolidado[matricula]["Salário Básico"] = valor
        else:
            consolidado[matricula][rubrica] = valor

    df = pd.DataFrame(consolidado.values())
    # Renomear colunas para remover " - LEI 11.091/05 AT"
    df.rename(columns={
    col: col.replace(" - LEI 11.091/05 AT", "") 
    for col in df.columns if " - LEI 11.091/05 AT" in col
    }, inplace=True)
    df.to_excel(caminho_saida, index=False)
    print(f"\n✅ Planilha gerada com adicionais em colunas: {caminho_saida}")


import glob

if __name__ == "__main__":
    pasta_pdf = r"C:/Ler arquivo/pdf-xlsx/pdfs"
    
    arquivos_pdf = glob.glob(os.path.join(pasta_pdf, "*.pdf"))

    saida_excel = os.path.join(pasta_pdf, "planilha_salarios_abril.xlsx")

    consolidar_dados(arquivos_pdf, saida_excel)
