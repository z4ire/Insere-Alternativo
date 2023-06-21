import tkinter as tk
from os import listdir
from os.path import abspath
from PIL import Image, ImageTk
from pathlib import Path
from tkinter import messagebox
from sqlite3 import Row
from distutils.cmd import Command
from Insere_Alternativo_AUX import aux


def escolher_ECO(responsavel, janela):

    janela.destroy()
    janela2 = tk.Tk()
    janela2.title("ECO")
    janela2.geometry("345x160")
    janela2.resizable(width=False, height=False)
    janela2.configure(bg="#EAEAEA")
    janela2.grid_columnconfigure(0, minsize=25)
    janela2.grid_columnconfigure(1, minsize=50)
    janela2.grid_columnconfigure(2, minsize=175)
    janela2.grid_rowconfigure(0, minsize=20)
    janela2.grid_rowconfigure(4, minsize=0.5)

    label_4 = tk.Label(janela2, text="Ano da ECO", font=("Poppins", 10), bg="#EAEAEA")
    label_5 = tk.Label(janela2, text="Número da ECO", font=("Poppins", 10), bg="#EAEAEA")
    label_6 = tk.Label(janela2, text="", bg="#EAEAEA")
    label_7 = tk.Label(janela2, text="- Inserir apenas o número da ECO\n- O caminho da ECO deve conter \" - To do\" ", font=("Poppins", 7), bg="#EAEAEA")

    campo_3 = tk.Entry(janela2)
    campo_4 = tk.Entry(janela2)
    campo_3.configure(bg="#DDEBF7", highlightbackground="#205A8C", highlightthickness=1, borderwidth=0)
    campo_4.configure(bg="#DDEBF7", highlightbackground="#205A8C", highlightthickness=1, borderwidth=0)
    botao_2 = tk.Button(janela2, text="Avançar", font=("Poppins", 8),  bg="#205A8C", fg="white", activebackground="#205A8C", activeforeground="white", command=lambda: cria_path(str(campo_3.get()), str(campo_4.get()), responsavel, janela2))

    label_4.grid(row=1, column=1, sticky="W")
    label_5.grid(row=2, column=1, sticky="W")
    label_6.grid(row=4, column=1)
    label_7.grid(row=4, column=2)
    campo_3.grid(row=1, column=2)
    campo_4.grid(row=2, column=2)
    botao_2.grid(row=5, column=2, pady=(12, 0))

    # Adiciona a imagem
    path = abspath(
        r'\\terra\GER_PRODUTOS\0 GPd\10 - Trabalhos desenvolvidos\Inserir alternativos - Automação\Arquivos\featured_channel.png')
    imagem = Image.open(path)
    imagem = imagem.resize((30, 30), resample=Image.Resampling.LANCZOS)
    imagem = ImageTk.PhotoImage(imagem)
    label_imagem = tk.Label(janela2, image=imagem, bg="#EAEAEA")
    label_imagem.image = imagem
    label_imagem.grid(row=6, column=3, pady=(0, 20), sticky="NW")


def cria_path(ano, numero, responsavel, janela2):

    path_analiseimpacto = r'\\loki\PADTEC - Campinas\Tecnologia\Hardware\Transferencia_PRO\ECO\AnaliseImpacto'

    try:
        if int(numero) > 99:
            eco = 'ECO-' + ano + '-' + '0' + numero
            path_analiseimpacto = path_analiseimpacto + "\ECOs_ " + \
                ano + '\ECO-' + ano + '-' + '0' + numero + r' - To do' + r'\Backup'
        elif int(numero) < 10:
            eco = 'ECO-' + ano + '-' + '000' + numero
            path_analiseimpacto = path_analiseimpacto + "\ECOs_ " + \
                ano + '\ECO-' + ano + '-' + '000' + numero + r' - To do' + r'\Backup'
        else:
            eco = 'ECO-' + ano + '-' + '00' + numero
            path_analiseimpacto = path_analiseimpacto + "\ECOs_ " + \
                ano + '\ECO-' + ano + '-' + '00' + numero + r' - To do' + r'\Backup'
    except:
        messagebox.showerror(
            message="Ano e/ou número da BOM incorretos! ", title="ERRO")
    else:

        try:
            folder = listdir(path_analiseimpacto)

        except:
            messagebox.showerror(
                message="Ano e/ou número da BOM incorretos.", title="ERRO")
        else:
            janela2.destroy()
            aux(path_analiseimpacto, responsavel, eco)
