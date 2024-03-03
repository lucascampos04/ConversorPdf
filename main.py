import tkinter as tk
from tkinter import filedialog
from docx2pdf import convert
from tqdm import tqdm
from tkinter import ttk
import time

def converter_docx_para_pdf(caminho_docx):
    progresso_variavel.set(0)  
    progresso_barra["value"] = 0
    janela.update_idletasks()

    try:
        for _ in tqdm(range(100), desc="Convertendo", unit="%", dynamic_ncols=True):
            time.sleep(0.02)  
            progresso_variavel.set(progresso_variavel.get() + 1)
            progresso_barra["value"] = progresso_variavel.get()
            janela.update_idletasks()

        convert(caminho_docx)
        print(f"Conversão concluída. O arquivo PDF foi salvo em: {caminho_docx[:-4]}.pdf")
    except Exception as e:
        print(f"Erro na conversão: {e}")

def selecionar_arquivo():
    caminho_docx = filedialog.askopenfilename(filetypes=[("Arquivos DOCX", "*.docx")])
    if caminho_docx:
        converter_docx_para_pdf(caminho_docx)

def local_arquivos():
    caminho_docx = filedialog.askopenfilename(filetypes=[("Arquivos PDF", "*.pdf")])
    return caminho_docx

janela = tk.Tk()
janela.geometry("200x200")
janela.title("Conversor DOCX para PDF")

botao_selecionar = tk.Button(janela, text="Selecionar Arquivo DOCX", command=selecionar_arquivo)
botao_selecionar.pack(pady=20)

progresso_variavel = tk.DoubleVar()
progresso_barra = ttk.Progressbar(janela, variable=progresso_variavel, maximum=100)
progresso_barra.pack(pady=10)

button_open_local = tk.Button(janela, text="Local do arquivo", command=local_arquivos)
button_open_local.pack(pady=5)
janela.mainloop()
