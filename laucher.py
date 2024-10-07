import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import asyncio
import subprocess
import os

# Função para chamar o script Node.js correspondente (F3.js ou F4.js)
def run_node_script(script, *args):
    try:
        print(f"Executando script: {script} com argumentos: {args}")  # Log para depuração
        result = subprocess.run(
            ['node', script] + list(args),
            capture_output=True,
            text=True,
            check=True,
            encoding='utf-8'  # Força o uso de UTF-8 para capturar a saída corretamente
        )
        print(f"Saída: {result.stdout}")  # Log para depuração
        return result.stdout
    except subprocess.CalledProcessError as e:
        print(f"Erro ao executar o script: {e.output}")  # Log para depuração
        return f"Erro: {e.output}"

# Função para iniciar o processamento
async def start_process(process, file_path, sequenciaNF):
    df = pd.read_excel(file_path)
    df.columns = df.columns.str.strip()

    script_file = 'F3.js' if process == 'F3' else 'F4.js'

    for index, row in df.iterrows():
        convenio = str(row["Convênio"])
        nrTitulo = str(row["Nr Retorno"])
        retorno = str(row.get("Tipo", ""))
        glosa = str(row.get("Glosa Total", ""))
        estabelecimento = str(row["Estabelecimento"])

        output = run_node_script(
            script_file,
            convenio, nrTitulo, sequenciaNF, estabelecimento, str(index + 1),
            retorno, glosa
        )
        print(output)

    messagebox.showinfo("Sucesso", "Processamento concluído com sucesso")

# Função para processar o arquivo
def process_file():
    process = process_var.get()
    sequenciaNF = sequenciaNR_entry.get()
    file_path = file_path_var.get()

    if not process or not sequenciaNF or not file_path:
        messagebox.showerror("Erro", "Todos os campos são obrigatórios!")
        return

    asyncio.run(start_process(process, file_path, sequenciaNF))

# Função para escolher o arquivo XLSX
def choose_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    file_path_var.set(file_path)

# Interface gráfica com Tkinter
root = tk.Tk()
root.title("Processador F3/F4")

tk.Label(root, text="Processo").grid(row=0, column=0, padx=10, pady=10)
process_var = tk.StringVar()
tk.Radiobutton(root, text="F3", variable=process_var, value="F3").grid(row=0, column=1)
tk.Radiobutton(root, text="F4", variable=process_var, value="F4").grid(row=0, column=2)

tk.Label(root, text="SequenciaNR").grid(row=1, column=0, padx=10, pady=10)
sequenciaNR_entry = tk.Entry(root)
sequenciaNR_entry.grid(row=1, column=1, columnspan=2)

tk.Label(root, text="Arquivo XLSX").grid(row=2, column=0, padx=10, pady=10)
file_path_var = tk.StringVar()
file_path_entry = tk.Entry(root, textvariable=file_path_var, state="readonly")
file_path_entry.grid(row=2, column=1, columnspan=2)
tk.Button(root, text="Escolher Arquivo", command=choose_file).grid(row=2, column=3, padx=10)

tk.Button(root, text="Iniciar Processamento", command=process_file).grid(row=3, column=1, columnspan=2, pady=20)

root.mainloop()