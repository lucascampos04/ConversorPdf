import tkinter as tk
from tkinter import filedialog
from docx2pdf import convert

def converter_docx_para_pdf(caminho_docx):
    try:
        convert(caminho_docx)
        print(f"Conversão concluída. O arquivo PDF foi salvo em: {caminho_docx[:-4]}.pdf")
    except Exception as e:
        print(f"Erro na conversão: {e}")

def selecionar_arquivo():
    caminho_docx = filedialog.askopenfilename(filetypes=[("Arquivos DOCX", "*.docx")])
    if caminho_docx:
        converter_docx_para_pdf(caminho_docx)

janela = tk.Tk()
janela.title("Conversor DOCX para PDF")

botao_selecionar = tk.Button(janela, text="Selecionar Arquivo DOCX", command=selecionar_arquivo)
botao_selecionar.pack(pady=20)

janela.mainloop()
