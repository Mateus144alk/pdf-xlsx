import tkinter as tk
from tkinter import filedialog
from tkinter.scrolledtext import ScrolledText
import pandas as pd

def ler_dados_siape_gui(caminho_planilha, log_widget):
    try:
        df = pd.read_excel(caminho_planilha)

        dados = {}
        log_widget.insert(tk.END, "\n=== DADOS LIDOS DA PLANILHA SIAPE ===\n")
        for _, row in df.iterrows():
            siape = str(row["SIAPE"]).strip()
            dv = str(row["DÍGITO VERIFICADOR MATRÍCULA"]).strip()
            origem = str(row["MATRÍCULA NA ORIGEM"]).strip()
            nome = str(row["NOME"]).strip()

            dados[siape] = {
                "DV": dv,
                "ORIGEM": origem
            }

            log_widget.insert(tk.END, f"SIAPE: {siape} | DV: {dv} | ORIGEM: {origem} | NOME: {nome}\n")

        log_widget.insert(tk.END, "=== FIM DA LEITURA ===\n\n")
        log_widget.see(tk.END)
        return dados

    except Exception as e:
        log_widget.insert(tk.END, f"❌ Erro ao ler planilha de dados SIAPE: {e}\n")
        return {}

# =============== INTERFACE ===============

def criar_interface():
    root = tk.Tk()
    root.title("Leitura Dados SIAPE")

    # Botão para carregar planilha
    def carregar_siape():
        caminho = filedialog.askopenfilename(filetypes=[("Planilhas Excel", "*.xlsx")])
        if caminho:
            ler_dados_siape_gui(caminho, log_text)

    btn_carregar = tk.Button(root, text="📂 Carregar dados SIAPE", command=carregar_siape)
    btn_carregar.pack(pady=10)

    # Log visual
    log_text = ScrolledText(root, width=100, height=20)
    log_text.pack(padx=10, pady=10)

    root.mainloop()

if __name__ == "__main__":
    criar_interface()
