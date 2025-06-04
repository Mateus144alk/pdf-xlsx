import pandas as pd
import shutil

# Caminhos
arquivo_original = "C:/Ler arquivo/pdf-xlsx/3 - Pgto Retroativo PROGRESSÕES.ods"
arquivo_copia = "C:/Ler arquivo/pdf-xlsx/3 - Pgto Retroativo PROGRESSÕES - Atualizado.xlsx"
arquivo_datas = "C:/Ler arquivo/pdf-xlsx/datas_extraidas.ods"

# Converta o .ods para .xlsx antes (pode usar LibreOffice ou pyexcel-ods3 + openpyxl)

# Criar cópia da planilha original
shutil.copyfile(arquivo_original.replace(".ods", ".xlsx"), arquivo_copia)

# === 1. Ler planilha de datas ===
import ezodf
ezodf.config.set_table_expand_strategy('all')
ods_datas = ezodf.opendoc(arquivo_datas)
sheet_datas = ods_datas.sheets[0]

dia = int(sheet_datas['C2'].value)
mes = int(sheet_datas['D2'].value)
print(f"Dia: {dia} | Mês: {mes}")

siapes_datas = []
for row in range(1, sheet_datas.nrows()):
    siape = sheet_datas[row, 0].value
    if siape:
        siapes_datas.append(str(siape).strip())

print("SIAPEs extraídos:", siapes_datas)

# === 2. Ler a planilha de progressões com header correto ===
df_prog = pd.read_excel(arquivo_copia, engine="openpyxl", header=1)
df_prog.columns = df_prog.columns.str.strip().str.upper()

# Verificar se coluna "SIAPE" existe
if 'SIAPE' not in df_prog.columns:
    print("Coluna 'SIAPE' não encontrada.")
    exit()

# Preparar os valores para preenchimento
df_prog['SIAPE'] = df_prog['SIAPE'].astype(str).str.strip()

# === 3. Preencher colunas AH e AI com dia e mês ===
col_ah = 33  # AH (34ª coluna, zero-indexed)
col_ai = 34  # AI (35ª coluna)

for idx, row in df_prog.iterrows():
    if row['SIAPE'] in siapes_datas:
        df_prog.iat[idx, col_ah] = dia
        df_prog.iat[idx, col_ai] = mes

# === 4. Salvar resultado ===
df_prog.to_excel(arquivo_copia, index=False)
print(f"Arquivo atualizado salvo em: {arquivo_copia}")
