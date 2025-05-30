import pdfplumber
import pandas as pd
import os
import re
import glob
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

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

    tipos_adicionais = sorted(set(r for _, _, r, _ in registros if "BASICO" not in r.upper()))

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
    df.rename(columns={
        col: re.sub(r"^\d{5} - ", "", col).replace(" - LEI 11.091/05 AT", "")
        for col in df.columns if re.match(r"^\d{5} - ", col)
    }, inplace=True)
    df.to_excel(caminho_saida, index=False)
    messagebox.showinfo("Sucesso", f"Planilha criada com sucesso em:\n{caminho_saida}")

def comparar_planilhas(caminho_abril, caminho_maio, caminho_saida, apenas_diferencas=False):
    df_abr = pd.read_excel(caminho_abril)
    df_mai = pd.read_excel(caminho_maio)

    colunas_excluir = ["Nome", "Matrícula"]
    colunas_valores = [col for col in df_abr.columns if col not in colunas_excluir]

    for col in colunas_valores:
        df_abr[col] = pd.to_numeric(df_abr[col], errors="coerce").fillna(0)
        df_mai[col] = pd.to_numeric(df_mai[col], errors="coerce").fillna(0)

    df_comparado = pd.merge(df_abr, df_mai, on="Matrícula", suffixes=(" (Abr)", " (Mai)"))
    df_comparado["Nome"] = df_comparado["Nome (Mai)"].combine_first(df_comparado["Nome (Abr)"])

    for col in colunas_valores:
        df_comparado[f"{col} (Diferença)"] = df_comparado[f"{col} (Mai)"] - df_comparado[f"{col} (Abr)"]

    colunas_ordenadas = ["Matrícula", "Nome"]
    for col in colunas_valores:
        colunas_ordenadas.extend([f"{col} (Abr)", f"{col} (Mai)", f"{col} (Diferença)"])

    df_final = df_comparado[colunas_ordenadas]

    if apenas_diferencas:
     diff_cols = [col for col in df_final.columns if col.endswith("(Diferença)")]
     df_final[diff_cols] = df_final[diff_cols].round(2)
     df_final = df_final[df_final[diff_cols].abs().sum(axis=1) > 0]



    df_final.to_excel(caminho_saida, index=False)
    messagebox.showinfo("Sucesso", f"Comparativo gerado com sucesso em:\n{caminho_saida}")
    # Aplica cor nas colunas de diferença
    wb = load_workbook(caminho_saida)
    ws = wb.active
    fill = PatternFill(start_color="ee0000", end_color="ee0000", fill_type="solid")  # vermelho claro

    for col in range(1, ws.max_column + 1):
        col_name = ws.cell(row=1, column=col).value
        if col_name and "(Diferença)" in col_name:
            for row in range(2, ws.max_row + 1):
                valor = ws.cell(row=row, column=col).value
                if isinstance(valor, (int, float)) and abs(valor) > 0.001:
                    ws.cell(row=row, column=col).fill = fill

    wb.save(caminho_saida)
    messagebox.showinfo("Sucesso", f"Comparativo gerado com sucesso em:\n{caminho_saida}")
# --- Interface ---
janela = tk.Tk()
janela.title("Consolidador e Comparador de Salários")
janela.geometry("700x350")

# Consolidação
def escolher_pasta_pdf():
    pasta = filedialog.askdirectory()
    if pasta:
        entrada_pasta.delete(0, tk.END)
        entrada_pasta.insert(0, pasta)

def salvar_excel_saida():
    caminho = filedialog.asksaveasfilename(defaultextension=".xlsx")
    if caminho:
        entrada_saida.delete(0, tk.END)
        entrada_saida.insert(0, caminho)

def acao_consolidar():
    pasta = entrada_pasta.get().strip()
    saida = entrada_saida.get().strip()
    if not pasta or not saida:
        messagebox.showwarning("Atenção", "Preencha todos os campos.")
        return
    arquivos = glob.glob(os.path.join(pasta, "*.pdf"))
    if not arquivos:
        messagebox.showwarning("Aviso", "Nenhum PDF encontrado na pasta.")
        return
    consolidar_dados(arquivos, saida)

# Comparar
def selecionar_planilha_abril():
    caminho = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
    if caminho:
        entrada_abril.delete(0, tk.END)
        entrada_abril.insert(0, caminho)

def selecionar_planilha_maio():
    caminho = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
    if caminho:
        entrada_maio.delete(0, tk.END)
        entrada_maio.insert(0, caminho)

def salvar_comparativo():
    caminho = filedialog.asksaveasfilename(defaultextension=".xlsx")
    if caminho:
        entrada_comparativo.delete(0, tk.END)
        entrada_comparativo.insert(0, caminho)

def acao_comparar():
    abr = entrada_abril.get()
    mai = entrada_maio.get()
    saida = entrada_comparativo.get()
    somente_diff = var_diferencas.get()
    if not abr or not mai or not saida:
        messagebox.showwarning("Atenção", "Preencha todos os campos da comparação.")
        return
    comparar_planilhas(abr, mai, saida, apenas_diferencas=somente_diff)

# Seção: Consolidador
frame1 = tk.LabelFrame(janela, text="Agrupar dados pdf em planilha")
frame1.pack(fill="x", padx=10, pady=5)

entrada_pasta = tk.Entry(frame1, width=70)
entrada_pasta.pack(side="left", padx=5, pady=5)
tk.Button(frame1, text="Selecionar Pasta", command=escolher_pasta_pdf).pack(side="left")

entrada_saida = tk.Entry(janela, width=70)
entrada_saida.pack(padx=15)
tk.Button(janela, text="Salvar como Excel", command=salvar_excel_saida).pack()
tk.Button(janela, text="📥 Gerar Planilha", bg="#4CAF50", fg="white", command=acao_consolidar).pack(pady=10)

# Seção: Comparar
frame2 = tk.LabelFrame(janela, text="Comparar Planilhas (anteriror vs. atual)")
frame2.pack(fill="x", padx=10, pady=5)

entrada_abril = tk.Entry(frame2, width=60)
entrada_abril.pack(side="left", padx=5)
tk.Button(frame2, text="Planilha Anterior", command=selecionar_planilha_abril).pack(side="left")

entrada_maio = tk.Entry(janela, width=60)
entrada_maio.pack(padx=15)
tk.Button(janela, text="Planilha Atual", command=selecionar_planilha_maio).pack()

entrada_comparativo = tk.Entry(janela, width=60)
entrada_comparativo.pack(padx=15)
tk.Button(janela, text="Salvar Comparativo", command=salvar_comparativo).pack()

var_diferencas = tk.BooleanVar()
tk.Checkbutton(janela, text="Exibir apenas diferenças", variable=var_diferencas).pack()
tk.Button(janela, text="📊 Comparar Meses", bg="#2196F3", fg="white", command=acao_comparar).pack(pady=10)

janela.mainloop()
