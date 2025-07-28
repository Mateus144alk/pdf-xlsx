import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from openpyxl.utils import column_index_from_string

class AccurateSiapeComparator:
    def __init__(self, root):
        self.root = root
        self.root.title("Comparador de SIAPEs - Versão Precisão")
        self.root.geometry("750x500")
        self.root.configure(bg="#f0f0f0")
        
        # Variáveis
        self.xlsx_file = tk.StringVar()
        self.csv_file = tk.StringVar()
        self.xlsx_column = tk.StringVar(value="A")
        self.csv_column = tk.StringVar(value="1")
        self.has_header = tk.BooleanVar(value=True)
        
        # Interface
        self.create_widgets()
    
    def create_widgets(self):
        # Frame principal
        main_frame = tk.Frame(self.root, bg="#f0f0f0", padx=20, pady=20)
        main_frame.pack(expand=True, fill=tk.BOTH)
        
        # Estilos
        label_style = {'bg': '#f0f0f0', 'anchor': 'w', 'font': ('Arial', 10)}
        entry_style = {'width': 45, 'borderwidth': 2, 'relief': 'groove', 'font': ('Arial', 10)}
        button_style = {'bg': '#4CAF50', 'fg': 'white', 'borderwidth': 0, 'font': ('Arial', 10)}
        
        # Seção XLSX
        tk.Label(main_frame, text="Arquivo Excel (XLSX):", **label_style).pack(fill=tk.X)
        
        xlsx_frame = tk.Frame(main_frame, bg="#f0f0f0")
        xlsx_frame.pack(fill=tk.X, pady=5)
        tk.Entry(xlsx_frame, textvariable=self.xlsx_file, **entry_style).pack(side=tk.LEFT, padx=5)
        tk.Button(xlsx_frame, text="Procurar", command=self.browse_xlsx, **button_style).pack(side=tk.LEFT)
        
        tk.Label(main_frame, text="Coluna SIAPE (ex: A, B, C ou nome da coluna):", **label_style).pack(fill=tk.X)
        self.xlsx_col_entry = tk.Entry(main_frame, textvariable=self.xlsx_column, width=15, 
                                      borderwidth=2, relief='groove', font=('Arial', 10))
        self.xlsx_col_entry.pack(anchor=tk.W)
        
        # Seção CSV
        tk.Label(main_frame, text="\nArquivo CSV:", **label_style).pack(fill=tk.X)
        
        csv_frame = tk.Frame(main_frame, bg="#f0f0f0")
        csv_frame.pack(fill=tk.X, pady=5)
        tk.Entry(csv_frame, textvariable=self.csv_file, **entry_style).pack(side=tk.LEFT, padx=5)
        tk.Button(csv_frame, text="Procurar", command=self.browse_csv, **button_style).pack(side=tk.LEFT)
        
        tk.Label(main_frame, text="Coluna SIAPE (número: 1, 2, 3 ou nome da coluna):", **label_style).pack(fill=tk.X)
        self.csv_col_entry = tk.Entry(main_frame, textvariable=self.csv_column, width=15, 
                                    borderwidth=2, relief='groove', font=('Arial', 10))
        self.csv_col_entry.pack(anchor=tk.W)
        
        # Opções
        options_frame = tk.Frame(main_frame, bg="#f0f0f0")
        options_frame.pack(fill=tk.X, pady=10)
        tk.Checkbutton(options_frame, text="Arquivos têm cabeçalho", variable=self.has_header, 
                      bg="#f0f0f0", font=('Arial', 10)).pack(side=tk.LEFT)
        
        # Botão de comparação
        compare_btn = tk.Button(main_frame, text="COMPARAR SIAPEs (PRECISÃO)", command=self.compare_siapes,
                              bg="#2196F3", fg="white", font=('Arial', 10, 'bold'),
                              padx=20, pady=5)
        compare_btn.pack(pady=15)
        
        # Resultados
        result_frame = tk.LabelFrame(main_frame, text=" Resultados ", bg="#f0f0f0", 
                                   font=('Arial', 10, 'bold'), labelanchor='n')
        result_frame.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = tk.Scrollbar(result_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.result_text_widget = tk.Text(result_frame, height=12, wrap=tk.WORD, 
                                        yscrollcommand=scrollbar.set,
                                        font=('Consolas', 10), padx=10, pady=10)
        self.result_text_widget.pack(fill=tk.BOTH, expand=True)
        self.result_text_widget.insert(tk.END, "Selecione os arquivos e clique em COMPARAR SIAPEs")
        scrollbar.config(command=self.result_text_widget.yview)
        
        # Barra de status
        self.status_var = tk.StringVar(value="Pronto. Selecione os arquivos.")
        status_bar = tk.Label(self.root, textvariable=self.status_var, bd=1, 
                             relief=tk.SUNKEN, anchor=tk.W, bg="#e0e0e0", font=('Arial', 10))
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def browse_xlsx(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
        if filename:
            self.xlsx_file.set(filename)
            self.status_var.set(f"XLSX selecionado: {filename}")
    
    def browse_csv(self):
        filename = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("Text files", "*.txt"), ("All files", "*.*")])
        if filename:
            self.csv_file.set(filename)
            self.status_var.set(f"CSV selecionado: {filename}")
    
    def get_column_data(self, df, col_ref, has_header, file_type):
        """Obtém os dados de uma coluna usando referência de letra, número ou nome"""
        try:
            # Se for número (1-based)
            if str(col_ref).isdigit():
                col_idx = int(col_ref) - 1
                if col_idx < 0 or col_idx >= len(df.columns):
                    return None, f"Coluna {col_ref} inválida. O {file_type} tem {len(df.columns)} colunas."
                return df.iloc[1 if has_header else 0:, col_idx], None
            
            # Se for letra (A, B, C, ...)
            if len(str(col_ref)) <= 2 and str(col_ref).isalpha():
                col_idx = column_index_from_string(str(col_ref)) - 1
                if col_idx < 0 or col_idx >= len(df.columns):
                    return None, f"Coluna {col_ref} inválida. O {file_type} tem {len(df.columns)} colunas."
                return df.iloc[1 if has_header else 0:, col_idx], None
            
            # Se for nome de coluna (precisa ter header)
            if has_header:
                if str(col_ref) in df.columns:
                    return df.iloc[1:, df.columns.get_loc(str(col_ref))], None
                return None, f"Coluna '{col_ref}' não encontrada. Cabeçalhos disponíveis: {list(df.columns)}"
            
            return None, f"Para usar nome de coluna, o arquivo deve ter cabeçalho"
        
        except Exception as e:
            return None, f"Erro ao acessar coluna {col_ref}: {str(e)}"
    
    def clean_siape_data(self, series):
        """Limpa e padroniza os dados de SIAPE"""
        # Converter para string, remover espaços e caracteres especiais
        cleaned = series.astype(str).str.strip().str.replace(r'[^0-9]', '', regex=True)
        # Remover valores vazios e duplicados
        cleaned = cleaned[cleaned != ''].drop_duplicates().dropna()
        return cleaned
    
    def compare_siapes(self):
        xlsx_path = self.xlsx_file.get()
        csv_path = self.csv_file.get()
        
        if not xlsx_path or not csv_path:
            messagebox.showerror("Erro", "Por favor, selecione ambos os arquivos!")
            return
        
        try:
            self.status_var.set("Processando arquivos...")
            self.root.update()
            
            # Configurações
            xlsx_col = self.xlsx_column.get().strip()
            csv_col = self.csv_column.get().strip()
            has_header = self.has_header.get()
            
            # ===== Processar XLSX =====
            try:
                xlsx_df = pd.read_excel(xlsx_path, header=0 if has_header else None, dtype=str)
            except Exception as e:
                messagebox.showerror("Erro XLSX", f"Não foi possível ler o arquivo XLSX:\n{str(e)}")
                return
            
            siape_xlsx, error = self.get_column_data(xlsx_df, xlsx_col, has_header, "XLSX")
            if error:
                messagebox.showerror("Erro XLSX", error)
                return
            
            # ===== Processar CSV =====
            try:
                csv_df = pd.read_csv(csv_path, header=0 if has_header else None, dtype=str)
            except Exception as e:
                messagebox.showerror("Erro CSV", f"Não foi possível ler o arquivo CSV:\n{str(e)}")
                return
            
            siape_csv, error = self.get_column_data(csv_df, csv_col, has_header, "CSV")
            if error:
                messagebox.showerror("Erro CSV", error)
                return
            
            # ===== Processar SIAPEs =====
            # Limpar e padronizar os dados
            siape_xlsx_clean = self.clean_siape_data(siape_xlsx)
            siape_csv_clean = self.clean_siape_data(siape_csv)
            
            # Converter para conjuntos
            set_xlsx = set(siape_xlsx_clean)
            set_csv = set(siape_csv_clean)
            
            # Encontrar diferenças
            only_in_xlsx = sorted(set_xlsx - set_csv, key=lambda x: str(x))
            only_in_csv = sorted(set_csv - set_xlsx, key=lambda x: str(x))
            in_both = sorted(set_xlsx & set_csv, key=lambda x: str(x))
            
            # Mostrar resultados
            self.result_text_widget.delete(1.0, tk.END)
            
            # Estatísticas
            stats = (f"=== Estatísticas ===\n"
                    f"Total no XLSX: {len(set_xlsx)}\n"
                    f"Total no CSV: {len(set_csv)}\n"
                    f"Presentes em ambos: {len(in_both)}\n"
                    f"Presentes apenas no XLSX: {len(only_in_xlsx)}\n"
                    f"Presentes apenas no CSV: {len(only_in_csv)}\n\n")
            
            self.result_text_widget.insert(tk.END, stats)
            
            if in_both:
                self.result_text_widget.insert(tk.END, f"SIAPEs presentes em ambos arquivos ({len(in_both)}):\n")
                self.result_text_widget.insert(tk.END, "\n".join(in_both[:50]))  # Mostra no máximo 50
                if len(in_both) > 50:
                    self.result_text_widget.insert(tk.END, f"\n[...] e mais {len(in_both)-50} SIAPEs")
                self.result_text_widget.insert(tk.END, "\n\n")
            
            if only_in_xlsx:
                self.result_text_widget.insert(tk.END, f"SIAPEs apenas no XLSX ({len(only_in_xlsx)}):\n")
                self.result_text_widget.insert(tk.END, "\n".join(only_in_xlsx[:50]))  # Mostra no máximo 50
                if len(only_in_xlsx) > 50:
                    self.result_text_widget.insert(tk.END, f"\n[...] e mais {len(only_in_xlsx)-50} SIAPEs")
                self.result_text_widget.insert(tk.END, "\n\n")
            
            if only_in_csv:
                self.result_text_widget.insert(tk.END, f"SIAPEs apenas no CSV ({len(only_in_csv)}):\n")
                self.result_text_widget.insert(tk.END, "\n".join(only_in_csv[:50]))
                if len(only_in_csv) > 50:
                    self.result_text_widget.insert(tk.END, f"\n[...] e mais {len(only_in_csv)-50} SIAPEs")
            
            self.status_var.set("Comparação concluída com precisão!")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro inesperado:\n{str(e)}")
            self.status_var.set("Erro durante a comparação")
            self.result_text_widget.delete(1.0, tk.END)
            self.result_text_widget.insert(tk.END, f"ERRO: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = AccurateSiapeComparator(root)
    root.mainloop()