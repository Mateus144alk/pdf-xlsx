import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import calendar
import csv
from datetime import date
# Dados globais
dados_retroativo = pd.DataFrame()
dados_siape = pd.DataFrame()
df_valores_pgto = pd.DataFrame()
global valores_totais
#-------------------------------programação e logica etapa 1 dados GRCOSERRUB---------------------===
def mostrar_log_comparacao():
    pdf_antigo = entrada_pdf_antigo.get()
    pdf_atual = entrada_pdf_atual.get()

    if not pdf_antigo or not pdf_atual:
        messagebox.showerror("Erro", "Selecione ambos os arquivos PDF (antigo e atual).")
        return

    dados_antigo = extrair_dados_pdf(pdf_antigo)
    dados_atual = extrair_dados_pdf(pdf_atual)

    dict_antigo = {mat: (nome, valor) for mat, nome, valor in dados_antigo}
    dict_atual = {mat: (nome, valor) for mat, nome, valor in dados_atual}


def extrair_dados_pdf(path_pdf):
    import pdfplumber
    import re
    dados = []

    with pdfplumber.open(path_pdf) as pdf:
        for page in pdf.pages:
            texto = page.extract_text()
            if not texto:
                continue
            for linha in texto.splitlines():
                match = re.search(r"(.+?)\s+EST\d{2}\s+(\d{6,7})\s+R\s+0\s+([\d.,]+)\s+N", linha)
                if match:
                    nome = match.group(1).strip()
                    matricula = match.group(2)
                    valor = float(match.group(3).replace(".", "").replace(",", "."))
                    dados.append((matricula, nome, valor))
    return dados
#-------------------------------programação e logica etapa 2  DATA RETROATIVA DO PAGAMENTO---------------------===
def ler_planilha_retroativa(caminho):
    df = pd.read_excel(caminho)
    df.columns = df.columns.str.strip().str.upper()

    retroativos = {}

    for i, row in df.iterrows():
        try:
            siape = str(row["SIAPE"]).strip()
            nome = str(row["NOME"]).strip()

            dia_ini = int(row["DIA_INICIAL"])
            mes_ini = int(row["MÊS_INICIAL"])
            ano_ini = int(row["ANO_INICIAL"])
            dia_fim = int(row["DIA_FINAL"])
            mes_fim = int(row["MÊS_FINAL"])
            ano_fim = int(row["ANO_FINAL"])

            # Validações básicas
            if not (1 <= mes_ini <= 12 and 1 <= mes_fim <= 12):
                print(f"Linha {i+2}: mês inválido ({mes_ini} ou {mes_fim}) — ignorado.")
                continue
            if not (1 <= dia_ini <= 31 and 1 <= dia_fim <= 31):
                print(f"Linha {i+2}: dia inválido ({dia_ini} ou {dia_fim}) — ignorado.")
                continue

            data_ini = date(ano_ini, mes_ini, dia_ini)
            data_fim = date(ano_fim, mes_fim, dia_fim)

            if data_fim < data_ini:
                print(f"Linha {i+2}: data final anterior à inicial — ignorado.")
                continue

            # Calcula diferença básica
            anos = ano_fim - ano_ini
            meses = mes_fim - mes_ini + anos * 12
            dias = dia_fim - dia_ini

            # Corrige dias negativos
            if dias < 0:
                meses -= 1
                mes_corrigido = mes_fim - 1 if mes_fim > 1 else 12
                ano_corrigido = ano_fim if mes_fim > 1 else ano_fim - 1
                dias_no_mes_anterior = calendar.monthrange(ano_corrigido, mes_corrigido)[1]
                dias += dias_no_mes_anterior

            if (ano_ini, mes_ini) != (ano_fim, mes_fim) and dia_fim >= dia_ini:
                meses += 1
                dias = 0  # zera os dias, já contou o mês completo

            retroativos[siape] = (nome, dias, meses)

        except Exception as e:
            print(f" Erro na linha {i+2} ({row.to_dict()}): {e}")
            continue

    print(f" processada! {len(retroativos)} registros válidos.")
    return retroativos


def calcular_retroativos():
    pdf_antigo = entrada_pdf_antigo.get()
    pdf_atual = entrada_pdf_atual.get()
    planilha = entrada_planilha.get()

    if not pdf_antigo or not pdf_atual or not planilha:
        messagebox.showerror("Erro", "Selecione todos os arquivos: dois PDFs e uma planilha.")
        return

    dados_antigo = extrair_dados_pdf(pdf_antigo)
    dados_atual = extrair_dados_pdf(pdf_atual)
    retroativos = ler_planilha_retroativa(planilha)

    dict_antigo = {mat: (nome, valor) for mat, nome, valor in dados_antigo}
    dict_atual = {mat: (nome, valor) for mat, nome, valor in dados_atual}

    for siape in retroativos:
        nome_plan, dias, meses = retroativos[siape]
        nome_antigo, val_antigo = dict_antigo.get(siape, (nome_plan, 0.0))
        nome_atual, val_atual = dict_atual.get(siape, (nome_plan, 0.0))

        salario = val_atual - val_antigo
        retroativo = (salario * meses) + ((salario / 30) * dias)


#--------------------programação da logica etapa  3 DADOS EXTRATOR SIAPE:-----------------------------
import pandas as pd

def ler_dados_siape(caminho_planilha):

    try:
        df = pd.read_excel(caminho_planilha)

        dados = {}
        for _, row in df.iterrows():
            siape = str(row["SIAPE"]).strip()
            dv = str(row["DÍGITO VERIFICADOR MATRÍCULA"]).strip()
            origem = str(row["MATRÍCULA NA ORIGEM"]).strip()
            nome = str(row["NOME"]).strip()

            dados[siape] = {
                "DV": dv,
                "ORIGEM": origem
            }

            print(f"SIAPE: {siape} | DV: {dv} | ORIGEM: {origem} | NOME: {nome}")
        return dados
    except Exception as e:
        print(f" Erro ao ler planilha de dados SIAPE: {e}")
        return {}


#-----------------------------------------resultado para 1 e 2--------------------------------------------
def calcular_diferenca_bruta(dados_antigo, dados_atual):
    resultado = []
    dict_antigo = {mat: (nome, valor) for mat, nome, valor in dados_antigo}
    dict_atual = {mat: (nome, valor) for mat, nome, valor in dados_atual}

    for siape in dict_antigo:
        nome, valor_antigo = dict_antigo[siape]
        valor_atual = dict_atual.get(siape, (nome, 0.0))[1]
        diferenca = valor_atual - valor_antigo

        resultado.append({
            "SIAPE": siape,
            "NOME": nome,
            "VALOR ANTIGO": valor_antigo,
            "VALOR ATUAL": valor_atual,
            "DIFERENÇA": round(diferenca, 2)
        })

    return resultado


#-------------------ressultado para 1-----------------------
def calcular_valores_retroativos(dados_antigo, dados_atual, retroativos):
    resultado = []

    dict_antigo = {mat: (nome, valor) for mat, nome, valor in dados_antigo}
    dict_atual = {mat: (nome, valor) for mat, nome, valor in dados_atual}

    for siape, (nome, dias, meses) in retroativos.items():
        nome_antigo, valor_antigo = dict_antigo.get(siape, (nome, 0.0))
        nome_atual, valor_atual = dict_atual.get(siape, (nome, 0.0))

        diferenca = valor_atual - valor_antigo
        proporcional = diferenca * (meses + dias / 30)

        resultado.append({
            "SIAPE": siape,
            "NOME": nome,
            "ANTIGO": valor_antigo,
            "ATUAL": valor_atual,
            "MESES": meses,
            "DIAS": dias,
            "RETROATIVO": round(proporcional, 2)
        })

    return resultado
def exportar_resultado(resultado, nome_arquivo):
    df = pd.DataFrame(resultado)
    caminho = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Planilha Excel", "*.xlsx")],
        title=f"Salvar {nome_arquivo}"
    )
    if caminho:
        df.to_excel(caminho, index=False)
        messagebox.showinfo("Sucesso", f"{nome_arquivo} salvo com sucesso:\n{caminho}")
def gerar_diferenca_bruta():
    pdf_antigo = entrada_pdf_antigo.get()
    pdf_atual = entrada_pdf_atual.get()

    if not pdf_antigo or not pdf_atual:
        messagebox.showerror("Erro", "Selecione os dois arquivos PDF (antigo e atual) antes de gerar o resultado.")
        return

    dados_antigo = extrair_dados_pdf(pdf_antigo)
    dados_atual = extrair_dados_pdf(pdf_atual)

    resultado = calcular_diferenca_bruta(dados_antigo, dados_atual)
    exportar_resultado(resultado, "1º RESULTADO - Diferença Bruta")

#=---------------------resultado 2 -----------------------------------------------------

def gerar_valor_retroativo():
    if not entrada_pdf_antigo.get() or not entrada_pdf_atual.get() or not entrada_planilha.get():
        messagebox.showerror("Campos obrigatórios", "Por favor, selecione os dois PDFs (Antigo e Atual) e a Planilha de Retroativos.")
        return

    dados_antigo = extrair_dados_pdf(entrada_pdf_antigo.get())
    dados_atual = extrair_dados_pdf(entrada_pdf_atual.get())
    retroativos = ler_planilha_retroativa(entrada_planilha.get())
    resultado = calcular_valores_retroativos(dados_antigo, dados_atual, retroativos)
    exportar_resultado(resultado, "2º RESULTADO - Valor Retroativo")

    

#-------------------resultado 3-------------------------------------

def gerar_csv_carga_batch(dados_siape, resultado_diferenca, rubrica, sequencia, caminho_saida):
    try:
        with open(caminho_saida, mode='w', newline='', encoding='utf-8') as arquivo_csv:
            escritor = csv.writer(arquivo_csv, delimiter=';')  # <- ponto e vírgula como delimitador
            
            escritor.writerow([
                "MatSiape", "DVSiape", "Comando", "RendimentoDesconto", 
                "Rubrica", "Sequência", "Valor", "MatriculaOrigem"
            ])
            
            for item in resultado_diferenca:
                siape = item["SIAPE"]
                valor_final = item["ATUAL"]

                info = dados_siape.get(siape)
                if not info:
                    continue

                linha = [
                    siape,
                    info["DV"],
                    "4",  # Comando fixo
                    "1",  # Rendimento/Desconto fixo
                    rubrica,
                    sequencia,
                    f"{valor_final:.2f}".replace('.', ','),  # Valor final com vírgula
                    info["ORIGEM"]
                ]
                escritor.writerow(linha)

        return True, f" CSV gerado com sucesso em: {caminho_saida}"

    except Exception as e:
        return False, f"Erro ao gerar CSV: {e}"


    except Exception as e:
        return False, f"Erro ao gerar CSV: {e}"

# - - --------------------------Criar janela principal------------------------
janela = tk.Tk()
janela.title("Sistema de Pagamento Geral")
janela.geometry("800x600")

# ------------------- 1. VALORES ANTIGO / ATUAL (Preenchimento Manual) -------------------
frame_antigo = tk.LabelFrame(janela, text="1. Arquivo com VALORES ANTIGOS (PDF)")
frame_antigo.pack(fill="x", padx=10, pady=5)

entrada_pdf_antigo = tk.Entry(frame_antigo, width=60)
entrada_pdf_antigo.pack(side="left", padx=5)

def selecionar_pdf_antigo():
    caminho = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
    if caminho:
        entrada_pdf_antigo.delete(0, tk.END)
        entrada_pdf_antigo.insert(0, caminho)

tk.Button(frame_antigo, text="Selecionar PDF Anterior", command=selecionar_pdf_antigo).pack(side="left")

frame_atual = tk.LabelFrame(janela, text="1.1 PDF com VALORES ATUAIS")
frame_atual.pack(fill="x", padx=10, pady=5)

entrada_pdf_atual = tk.Entry(frame_atual, width=60)
entrada_pdf_atual.pack(side="left", padx=5)

def selecionar_pdf_atual():
    caminho = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if caminho:
        entrada_pdf_atual.delete(0, tk.END)
        entrada_pdf_atual.insert(0, caminho)

tk.Button(frame_atual, text="Selecionar PDF Atual", command=selecionar_pdf_atual).pack(side="left")



# ------------------- 2. PLANILHA DE DADOS RETROATIVOS -------------------
frame_planilha = tk.LabelFrame(janela, text="2. Planilha com Dados Retroativos")
frame_planilha.pack(fill="x", padx=10, pady=5)

entrada_planilha = tk.Entry(frame_planilha, width=60)
entrada_planilha.pack(side="left", padx=5)

def selecionar_planilha():
    caminho = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if caminho:
        entrada_planilha.delete(0, tk.END)
        entrada_planilha.insert(0, caminho)

tk.Button(frame_planilha, text="Selecionar Planilha", command=selecionar_planilha).pack(side="left")

# ------------------- 3. DADOS SIAPE (já extraídos) -------------------
frame_siape = tk.LabelFrame(janela, text="3. Dados SIAPE (para carga batch)")
frame_siape.pack(fill="x", padx=10, pady=10)

entrada_siape_path = tk.Entry(frame_siape, width=60)
entrada_siape_path.pack(side="left", padx=5)

def selecionar_planilha_siape():
    global dados_siape
    caminho = filedialog.askopenfilename(filetypes=[("Planilhas Excel", "*.xlsx *.xls")])
    if caminho:
        entrada_siape_path.delete(0, tk.END)
        entrada_siape_path.insert(0, caminho)
        try:
            dados_siape = pd.read_excel(caminho)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar planilha SIAPE:\n{e}")

tk.Button(frame_siape, text="Selecionar Planilha SIAPE", command=selecionar_planilha_siape).pack(side="left")

# ------------------- 4. PREENCHIMENTO MANUAL PARA BATCH -------------------
frame_batch = tk.LabelFrame(janela, text="3.1 Dados da Carga Batch (Preenchimento Manual)")
frame_batch.pack(fill="x", padx=10, pady=10)

tk.Label(frame_batch, text="Rubrica:").grid(row=0, column=0, padx=5, pady=5)
entrada_rubrica = tk.Entry(frame_batch, width=10)
entrada_rubrica.grid(row=0, column=1)

tk.Label(frame_batch, text="Sequência:").grid(row=0, column=2, padx=5)
entrada_seq = tk.Entry(frame_batch, width=10)
entrada_seq.grid(row=0, column=3)

# ------------------- BOTÃO GERAR ARQUIVOS -------------------

def gerar_resultado():
    # Passo 1: carregar e validar entradas
    pdf_antigo = entrada_pdf_antigo.get()
    pdf_atual = entrada_pdf_atual.get()
    caminho_planilha = entrada_planilha.get()
    siape_path = entrada_siape_path.get()

    if not (pdf_antigo and pdf_atual and caminho_planilha and siape_path):
        messagebox.showerror("Erro", "Preencha todos os campos e selecione os arquivos.")
        return

    # Passo 2: extrair dados
    dados_antigo = extrair_dados_pdf(pdf_antigo)
    dados_atual = extrair_dados_pdf(pdf_atual)
    retroativos = ler_planilha_retroativa(caminho_planilha)
    dados_siape_dict = ler_dados_siape(siape_path)

    # Passo 3: calcular valores finais para CSV (valor atual + retroativo)
    resultado = []
    dict_antigo = {mat: (nome, valor) for mat, nome, valor in dados_antigo}
    dict_atual = {mat: (nome, valor) for mat, nome, valor in dados_atual}

    for siape, (nome, dias, meses) in retroativos.items():
        nome_antigo, val_antigo = dict_antigo.get(siape, (nome, 0.0))
        nome_atual, val_atual = dict_atual.get(siape, (nome, 0.0))
        salario = val_atual - val_antigo
        retroativo = (salario * meses) + ((salario / 30) * dias)
        valor_final = val_atual + retroativo

        resultado.append({
            "SIAPE": siape,
            "NOME": nome,
            "ATUAL": valor_final
        })

    # Passo 4: gerar CSV
    rubrica = entrada_rubrica.get()
    sequencia = entrada_seq.get()
    caminho_csv = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])

    if caminho_csv:
            sucesso, mensagem = gerar_csv_carga_batch(dados_siape_dict, resultado, rubrica, sequencia, caminho_csv)
            if sucesso:
             messagebox.showinfo("Sucesso", mensagem)
            else:
             messagebox.showerror("Erro", mensagem)
tk.Button(
    janela,
    text="Gerar Arquivo de Pagamento (CSV) - 3º Resultado",
    bg="#4CAF50",  # verde mais forte
    fg="white",
    font=("Arial", 12, "bold"),
    command=gerar_resultado,
    padx=12,
    pady=8
).pack(pady=20)

btn_resultado1 = tk.Button(
    janela,
    text="1º Resultado - Diferença Bruta",
    bg="#2196F3",  # azul médio
    fg="white",
    font=("Arial", 11, "bold"),
    command=gerar_diferenca_bruta,
    padx=10,
    pady=5
)
btn_resultado1.pack(pady=10)

btn_resultado2 = tk.Button(
    janela,
    text="2º Resultado - Retroativo Proporcional",
    bg="#FF9800",  # laranja médio
    fg="white",
    font=("Arial", 11, "bold"),
    command=gerar_valor_retroativo,
    padx=10,
    pady=5
)
btn_resultado2.pack(pady=10)

janela.mainloop()
