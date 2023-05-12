import tkinter as tk
import os
from Insere_Alternativo_na_BOM import parametros
from PIL import Image, ImageTk
import datetime


def aux(path_analiseimpacto, responsavel, eco):
    janela3 = tk.Tk()
    janela3.title("Parametros")
    janela3.geometry("400x220")
    janela3.resizable(width=False, height=False)
    janela3.configure(bg="#EAEAEA")
    janela3.grid_columnconfigure(0, minsize=25)
    janela3.grid_columnconfigure(1, minsize=50)
    janela3.grid_columnconfigure(2, minsize=175)
    janela3.grid_rowconfigure(0, minsize=15)

    label_1 = tk.Label(janela3, text="Código Principal",
                       font=("calibri", 12), bg="#EAEAEA")
    label_2 = tk.Label(janela3, text="Código Alternativo",
                       font=("calibri", 12), bg="#EAEAEA")
    label_3 = tk.Label(janela3, text="Descreva a Alteração",
                       font=("calibri", 12), bg="#EAEAEA")
    label_4 = tk.Label(janela3, text="Engenheiro Responsável",
                       font=("calibri", 12), bg="#EAEAEA")
    label_5 = tk.Label(janela3, text="Previsão de Liberação",
                       font=("calibri", 12), bg="#EAEAEA")
    label_6 = tk.Label(janela3, text="")

    campo_1 = tk.Entry(janela3)
    campo_2 = tk.Entry(janela3)
    campo_3 = tk.Entry(janela3)
    campo_4 = tk.Entry(janela3)
    campo_5 = tk.Entry(janela3)

    campo_1.configure(bg="#DDEBF7", highlightbackground="#205A8C",
                      highlightthickness=1, borderwidth=0)
    campo_2.configure(bg="#DDEBF7", highlightbackground="#205A8C",
                      highlightthickness=1, borderwidth=0)
    campo_3.configure(bg="#DDEBF7", highlightbackground="#205A8C",
                      highlightthickness=1, borderwidth=0)
    campo_4.configure(bg="#DDEBF7", highlightbackground="#205A8C",
                      highlightthickness=1, borderwidth=0)
    campo_5.configure(bg="#DDEBF7", highlightbackground="#205A8C",
                      highlightthickness=1, borderwidth=0)
    
    botao_3 = tk.Button(text="Atualizar arquivos", bg="#205A8C", fg="white", activebackground="#205A8C", activeforeground="white", command=lambda: parametros(
        path_analiseimpacto, str(campo_1.get()), str(campo_2.get()), str(campo_3.get()), str(campo_4.get()), datetime.datetime.strptime(campo_5.get(), "%d/%m/%Y"), janela3, responsavel, eco))

    path = os.path.abspath(
        "C:/Users/tc.zcouto/OneDrive - Padtec/Área de Trabalho/Zaire/teste/Atualiza alternativos/featured_channel.png")
    imagem = Image.open(path)
    imagem = imagem.resize((30, 30), resample=Image.Resampling.LANCZOS)
    imagem = ImageTk.PhotoImage(imagem)
    label_imagem = tk.Label(janela3, image=imagem, bg="#EAEAEA")
    label_imagem.image = imagem

    label_imagem.grid(row=7, column=3, pady=10, sticky="NW")
    label_1.grid(row=1, column=1, sticky="W")
    label_2.grid(row=2, column=1, sticky="W")
    label_3.grid(row=3, column=1, sticky="W")
    label_4.grid(row=4, column=1, sticky="W")
    label_5.grid(row=5, column=1, sticky="W")
    label_6.grid(row=6, column=1, sticky="W")
    campo_1.grid(row=1, column=2)
    campo_2.grid(row=2, column=2)
    campo_3.grid(row=3, column=2)
    campo_4.grid(row=4, column=2)
    campo_5.grid(row=5, column=2)
    botao_3.grid(row=6, column=2, pady=(12,0))
