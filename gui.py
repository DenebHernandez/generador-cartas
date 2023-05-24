import os
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from app import Generador

root = tk.Tk()
root.title("Generador cartas")

generador = Generador()

def define_documents():
    word_file = word_file_entry.get()
    excel_file = excel_file_entry.get()
    output_dir = output_entry.get()
    generador.documents(excel_file=excel_file, word_template=word_file, outpur_dir=output_dir)

def render():
    export_word = to_word_var.get()
    export_pdf = to_pdf_var.get()
    define_documents()
    generador.render_template(to_word=export_word, to_pdf=export_pdf)

def browse_word_file():
    file_path = filedialog.askopenfilename(filetypes=[('Word files', '*.docx')])
    word_file_entry.delete(0, tk.END)
    word_file_entry.insert(0, file_path)

def browse_excel_file():
    file_path = filedialog.askopenfilename(filetypes=[('Excel files', '*.xlsx')])
    excel_file_entry.delete(0, tk.END)
    excel_file_entry.insert(0, file_path)

def browse_output():
    file_path = filedialog.askdirectory()
    output_entry.delete(0, tk.END)
    output_entry.insert(0, file_path)

def validation():
    define_documents()
    generador.validation_info()

def show_message():
    messagebox.showinfo("Procesando")

# def run_app():
#     word_file = word_file_entry.get()
#     excel_file = excel_file_entry.get()
#     export_word = include_address_var.get()
#     export_pdf = include_greeting_var.get()
#     os.system(f'python app.py {word_file} {excel_file}')

canvas= tk.Canvas(root, width=600, height=300)
canvas.grid(columnspan=3, rowspan=10)

#plantilla word
word_file_label = tk.Label(root, text="Plantilla Word:")
word_file_label.grid(column=0, row=0)

word_file_entry = tk.Entry(root, width=50)
word_file_entry.grid(column=1, row=0)

word_file_button = tk.Button(root, text="Buscar", command=browse_word_file)
word_file_button.grid(column=2, row=0)

#platilla excel
excel_file_label = tk.Label(root, text="Tabla Excel:")
excel_file_label.grid(column=0, row=1)

excel_file_entry = tk.Entry(root, width=50)
excel_file_entry.grid(column=1, row=1)

excel_file_button = tk.Button(root, text="Buscar", command=browse_excel_file)
excel_file_button.grid(column=2, row=1)

#destino
output_label = tk.Label(root, text="Destino:")
output_label.grid(column=0, row=2)

output_entry = tk.Entry(root, width=50)
output_entry.grid(column=1, row=2)

output_button = tk.Button(root, text="Buscar", command=browse_output)
output_button.grid(column=2, row=2)

#formatos exportación
to_pdf_var = tk.BooleanVar(value=True)
to_pdf_checkbutton = tk.Checkbutton(root, text=" Exportar a PDF", variable=to_pdf_var)
to_pdf_checkbutton.grid(column=1, row=5)

to_word_var = tk.BooleanVar(value=True)
to_word_checkbutton = tk.Checkbutton(root, text="Exportar a Word", variable=to_word_var)
to_word_checkbutton.grid(column=1, row=6)

#llamada inicio del proceso
run_button = tk.Button(root, text="Generar", command=render)
run_button.grid(column=1, row=9)

root.mainloop()