import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter

def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("All Files", "*.*")])
    if file_path:
        entry_file_path.delete(0, tk.END)
        entry_file_path.insert(0, file_path)

def convert_file():
    input_format = combo_input_format.get()
    output_format = combo_output_format.get()
    file_path = entry_file_path.get()

    if not file_path:
        messagebox.showerror("Erro", "Selecione um arquivo!")
        return

    if input_format == output_format:
        messagebox.showerror("Erro", "Os formatos de entrada e saída são iguais!")
        return

    try:
        if input_format == "XLSX" and output_format == "PDF":
            convert_xlsx_to_pdf(file_path)
        elif input_format == "PDF" and output_format == "XLSX":
            convert_pdf_to_xlsx(file_path)
        else:
            messagebox.showerror("Erro", "Conversão não suportada!")
            return
        messagebox.showinfo("Sucesso", "Conversão realizada com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao converter: {e}")

def convert_xlsx_to_pdf(file_path):
    df = pd.read_excel(file_path)
    pdf_path = file_path.replace(".xlsx", ".pdf")
    with open(pdf_path, "w") as pdf_file:
        pdf_file.write(df.to_string())

def convert_pdf_to_xlsx(file_path):
    reader = PdfReader(file_path)
    text = "\n".join([page.extract_text() for page in reader.pages])
    xlsx_path = file_path.replace(".pdf", ".xlsx")
    df = pd.DataFrame({"Conteúdo": [text]})
    df.to_excel(xlsx_path, index=False)

# Criar a janela principal
root = tk.Tk()
root.title("Conversor de Arquivos")
root.geometry("580x300")
root.configure(bg="#E8F9FD")

# Estilo
style = ttk.Style()
style.configure("TLabel", background="#E8F9FD", font=("Arial", 12))
style.configure("TButton", font=("Arial", 12))
style.configure("TCombobox", font=("Arial", 12))

# Widgets
frame = tk.Frame(root, bg="#E8F9FD", padx=20, pady=20)
frame.pack(expand=True)

lbl_input_format = ttk.Label(frame, text="Formato de Entrada:")
lbl_input_format.grid(row=0, column=0, pady=10, sticky="w")

combo_input_format = ttk.Combobox(frame, values=["XLSX", "PDF"], state="readonly")
combo_input_format.grid(row=0, column=1, pady=10)
combo_input_format.set("XLSX")

lbl_output_format = ttk.Label(frame, text="Formato de Saída:")
lbl_output_format.grid(row=1, column=0, pady=10, sticky="w")

combo_output_format = ttk.Combobox(frame, values=["XLSX", "PDF"], state="readonly")
combo_output_format.grid(row=1, column=1, pady=10)
combo_output_format.set("PDF")

lbl_file_path = ttk.Label(frame, text="Arquivo:")
lbl_file_path.grid(row=2, column=0, pady=10, sticky="w")

entry_file_path = ttk.Entry(frame, width=30, font=("Arial", 12))
entry_file_path.grid(row=2, column=1, pady=10)

btn_browse = ttk.Button(frame, text="Selecionar", command=select_file)
btn_browse.grid(row=2, column=2, padx=10)

btn_convert = ttk.Button(frame, text="Converter", command=convert_file)
btn_convert.grid(row=3, column=0, columnspan=3, pady=20)

# Rodar o aplicativo
root.mainloop()
