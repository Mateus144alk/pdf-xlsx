from odf.opendocument import load
from odf.table import Table, TableRow, TableCell
from odf.text import P
import pandas as pd
import re
from tkinter import filedialog, Tk

def extrair_dados_odt(caminho_odt):
    doc = load(caminho_odt)
    registros = []

    for tabela in doc.getElementsByType(Table):
        for linha in tabela.getElementsByType(TableRow):
            celulas = linha.getElementsByType(TableCell)
            textos = []
            for celula in celulas:
                textos_na_celula = celula.getElementsByType(P)
                conteudo = " ".join([
                    str(node.data).strip()
                    for t in textos_na_celula
                    for node in t.childNodes
                    if hasattr(node, "data")
                ])
                textos.append(conteudo)

            linha_texto = " ".join(textos).strip().upper()

            match_data = re.search(
                r"(\d{1,2})\s+((?:JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ|\d{2})/\d{4})",
                linha_texto
            )
            if match_data:
                dia = match_data.group(1)
                mes_raw = match_data.group(2)

                if re.match(r"\d{2}/\d{4}", mes_raw):
                    mes = mes_raw[:2]
                else:
                    mes_ext = mes_raw[:3]
                    meses = {
                        "JAN": "01", "FEV": "02", "MAR": "03", "ABR": "04", "MAI": "05", "JUN": "06",
                        "JUL": "07", "AGO": "08", "SET": "09", "OUT": "10", "NOV": "11", "DEZ": "12"
                    }
                    mes = meses.get(mes_ext, "")

                match_siape_nome = re.search(r"(\d{7})\s+([A-Z\s\-ÇÁÉÍÓÚÂÊÔÃÕ]+)", linha_texto)
                if match_siape_nome:
                    siape = match_siape_nome.group(1)
                    nome_completo = match_siape_nome.group(2).strip()

                    nome_sem_cargo = re.sub(r"\s+(MÉDICO|CONTADOR|REDATOR|BIOMÉDICO|PRODUTOR|TRADUTOR|PSICÓLOGA|JORNALISTA|BIBLIOTECARIO|TRADUTOR|ESTATÍSTICO|ECONOMISTA|TÉC DE TECNOLOGIA|REVISOR DE TEXTOS|SECRETÁRIO|FARMACÊUTICO|BIBLIO|ENFERMEIRO|ADMINISTRADOR|ASSISTENTE|AUDITOR|FISIOTERAPEUTA|ANALISTA|PROGRAMADOR VISUAL|BIBLIOTECÁRIO\-DOCUMENTALISTA|PSICÓLOGO\-AREA|DIAGRAMADOR|FONOAUDIÓLOGO|NUTRICIONISTA|ZELADOR|AUXILIAR|ENGENHEIRO|TÉCNICO|COORDENADOR(?: [A-Z]+)*|BIBLIOTECÁRIO(?: [A-Z\-]+)*)\b.*", "", nome_completo).strip()       
                    if len(nome_sem_cargo.split()) >= 2:
                     registros.append({"SIAPE": siape, "Nome": nome_sem_cargo, "Dia": dia, "Mês": mes})

        return registros


if __name__ == "__main__":
    Tk().withdraw()
    caminhos_odt = filedialog.askopenfilenames(title="Selecione os arquivos .odt", filetypes=[("ODT files", "*.odt")])

    todos_registros = []
    for caminho in caminhos_odt:
        registros = extrair_dados_odt(caminho)
        todos_registros.extend(registros)

    if todos_registros:
        df_dados = pd.DataFrame(todos_registros)
        df_dados.to_excel("datas_extraidas.ods", index=False, engine="odf")
        print(f"✅ {len(df_dados)} registros extraídos e salvos em 'datas_extraidas.ods'")
    else:
        print("⚠️ Nenhuma data válida encontrada nos arquivos selecionados.")
