import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

df_batch = pd.DataFrame()  # Inicialmente vazio

def importar_csv():
    global df_batch
    caminho = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    if caminho:
        try:
            df_batch = pd.read_csv(caminho, dtype=str)  # Importa tudo como string
            messagebox.showinfo("Sucesso", "CSV importado com sucesso.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao importar CSV:\n{e}")

def gerar_arquivo_txt(df_batch, caminho_txt, config):
    try:
        with open(caminho_txt, "w", encoding="utf-8") as f:
            for _, row in df_batch.iterrows():
                mat_siape = str(row["MatSiape"]).zfill(7)
                dv_siape = str(row["DVSiape"])[-1]
                comando = str(row["Comando"]).zfill(1)
                rendimento = str(row["RendimentoDesconto"]).zfill(1)
                rubrica = str(row["Rubrica"]).zfill(5)
                sequencia = str(row["Sequencia"]).zfill(2)

                # Valor: substituir vírgula por ponto se necessário
                valor_str = str(row["Valor"]).replace(",", ".")
                valor_centavos = int(round(float(valor_str) * 100))
                valor_formatado = str(valor_centavos).zfill(9)

                matricula_origem = str(row["MatriculaOrigem"]).zfill(7)

                orgao = str(config.get("matriz_padrao", "00000")).zfill(5)
                mes_pgto = str(config.get("mes_pagto", "01")).zfill(2)
                ano_pgto = str(config.get("ano_pagto", "2025")).zfill(4)
                mes_rubrica = str(config.get("mes_rubrica", "01")).zfill(2)
                ano_rubrica = str(config.get("ano_rubrica", "2025")).zfill(4)
                nome_inst = str(config.get("nome_instituicao", "XXXXXX")).ljust(6)[:6]
                rubrica_header = str(config.get("rubrica_arquivo", "00000")).zfill(5)

                linha = (
                    orgao +
                    mes_pgto +
                    ano_pgto +
                    mes_rubrica +
                    ano_rubrica +
                    nome_inst +
                    rubrica_header +
                    comando +
                    rendimento +
                    mat_siape +
                    dv_siape +
                    rubrica +
                    sequencia +
                    ano_rubrica +
                    mes_rubrica +
                    valor_formatado +
                    matricula_origem
                )

                f.write(linha + "\n")

        print("✅ Arquivo TXT gerado com sucesso.")
    except KeyError as e:
        print(f"❌ Erro: Coluna não encontrada no DataFrame: {e}")
    except Exception as e:
        print(f"❌ Erro ao gerar o arquivo TXT: {e}")


def gerar_txt():
    if df_batch.empty:
        messagebox.showerror("Erro", "Nenhum CSV importado.")
        return

    config = {
        "matriz_padrao": entrada_orgao_siape.get(),
        "mes_pagto": entrada_mes_pagto.get(),
        "ano_pagto": entrada_ano_pagto.get(),
        "mes_rubrica": entrada_mes_rubrica.get(),
        "ano_rubrica": entrada_ano_rubrica.get(),
        "nome_instituicao": entrada_nome_inst.get()[:6].ljust(6),
        "rubrica_arquivo": entrada_rubrica_header.get().zfill(5)
    }

    caminho_txt = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("TXT", "*.txt")])
    if caminho_txt:
        try:
            gerar_arquivo_txt(df_batch, caminho_txt, config)
            messagebox.showinfo("Sucesso", f"Arquivo TXT gerado:\n{caminho_txt}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar TXT:\n{e}")

def alternar_campos_txt():
    visivel = var_gerar_txt.get()
    for widget in campos_txt:
        widget.grid() if visivel else widget.grid_remove()
    botao_gerar_txt.grid() if visivel else botao_gerar_txt.grid_remove()

# --- Interface ---
root = tk.Tk()
root.title("Gerador Movi-Financ")

frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

# Checkbox
var_gerar_txt = tk.BooleanVar()
check_txt = tk.Checkbutton(frame, text="Gerar arquivo TXT (Movi-Financ)", variable=var_gerar_txt, command=alternar_campos_txt)
check_txt.grid(row=0, column=0, columnspan=2, sticky="w")

# Entradas para cabeçalho
campos_txt = []

entrada_orgao_siape = tk.Entry(frame)
entrada_mes_pagto = tk.Entry(frame)
entrada_ano_pagto = tk.Entry(frame)
entrada_mes_rubrica = tk.Entry(frame)
entrada_ano_rubrica = tk.Entry(frame)
entrada_nome_inst = tk.Entry(frame)
entrada_rubrica_header = tk.Entry(frame)

labels_entradas = [
    ("Órgão SIAPE:", entrada_orgao_siape),
    ("Mês Pagamento:", entrada_mes_pagto),
    ("Ano Pagamento:", entrada_ano_pagto),
    ("Mês Rubrica:", entrada_mes_rubrica),
    ("Ano Rubrica:", entrada_ano_rubrica),
    ("Nome da Instituição:", entrada_nome_inst),
    ("Rubrica do Arquivo:", entrada_rubrica_header),
]

for i, (texto, entrada) in enumerate(labels_entradas, start=1):
    lbl = tk.Label(frame, text=texto)
    lbl.grid(row=i, column=0, sticky="e")
    entrada.grid(row=i, column=1, sticky="w")
    campos_txt.extend([lbl, entrada])

for widget in campos_txt:
    widget.grid_remove()

# Botões
botao_importar_csv = tk.Button(frame, text="Importar CSV", command=importar_csv)
botao_importar_csv.grid(row=10, column=0, pady=10)

botao_gerar_txt = tk.Button(frame, text="Gerar Arquivo TXT", command=gerar_txt)
botao_gerar_txt.grid(row=10, column=1, pady=10)
botao_gerar_txt.grid_remove()

root.mainloop()
