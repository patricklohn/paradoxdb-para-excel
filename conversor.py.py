import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import time
from pypxlib import Table
import datetime

# Dicionário de correções para caracteres corrompidos
CHAR_FIXES = {
    "Ã├O": "ÇÃO",
    "Ã£": "ã",
    "Ã¡": "á",
    "Ã©": "é",
    "Ãª": "ê",
    "Ã³": "ó",
    "Ã§": "ç",
    "Ãº": "ú",
    "Ã­": "í",
    "Ãµ": "õ",
    "Ã¤": "ä",
    "Ã¶": "ö",
}

def select_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Selecione o arquivo Paradox", filetypes=[("Paradox DB", "*.db")]
    )
    return file_path

def fix_text(value):
    """Corrige caracteres corrompidos em strings."""
    if isinstance(value, str):
        for wrong, correct in CHAR_FIXES.items():
            value = value.replace(wrong, correct)
    return value

def safe_get_value(row, col):
    """Obtém o valor da coluna e aplica correção de texto."""
    try:
        value = row[col]
        if isinstance(value, str):
            value = fix_text(value)
        if isinstance(value, datetime.date) and value.year < 1:
            return pd.NaT
        return value
    except Exception as e:
        print(f"Erro ao acessar valor na coluna {col}: {e}")
        return pd.NaT

def convert_paradox_to_excel(paradox_file):
    start_time = time.time()
    try:
        if not paradox_file:
            messagebox.showerror("Erro", "Nenhum arquivo selecionado.")
            return

        print(f"📂 Arquivo selecionado: {paradox_file}")

        print("📖 Lendo arquivo Paradox...")
        try:
            table = Table(paradox_file)
        except Exception as err:
            print(f"❌ Erro ao ler arquivo Paradox: {err}")
            messagebox.showerror("Erro", f"Erro ao ler arquivo Paradox: {err}")
            return

        columns = list(table.fields.keys())
        print(f"📊 Colunas encontradas: {columns}")

        data = []
        for row in table:
            data.append([safe_get_value(row, col) for col in columns])

        df = pd.DataFrame(data, columns=columns)

        print("📋 Prévia dos dados antes da conversão:")
        print(df.head())

        # Aplicar correção em todas as células do DataFrame
        df = df.applymap(fix_text)

        print("📋 Prévia dos dados após a correção:")
        print(df.head())

        excel_file = paradox_file.replace(".DB", ".xlsx")

        print("💾 Salvando dados no Excel...")
        df.to_excel(excel_file, index=False, engine="openpyxl")

        elapsed_time = time.time() - start_time
        messagebox.showinfo(
            "Sucesso",
            f"Conversão concluída em {elapsed_time:.2f} segundos!\nArquivo salvo como: {excel_file}",
        )
        print(f"✅ Arquivo salvo como: {excel_file}")

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")
        print(f"❌ Erro: {str(e)}")

paradox_file = select_file()
convert_paradox_to_excel(paradox_file)
