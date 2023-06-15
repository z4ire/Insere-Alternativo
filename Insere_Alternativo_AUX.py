import tkinter as tk
import os
import datetime
from tkinter import ttk
from Insere_Alternativo_na_BOM import parametros
from PIL import Image, ImageTk
from tkinter import messagebox


def aux(path_analiseimpacto, responsavel, eco):
    janela3 = tk.Tk()
    janela3.title("Parametros")
    janela3.geometry("428x250")
    janela3.resizable(width=False, height=False)
    janela3.configure(bg="#EAEAEA")
    janela3.grid_columnconfigure(0, minsize=25)
    janela3.grid_columnconfigure(1, minsize=50)
    janela3.grid_columnconfigure(2, minsize=175)
    janela3.grid_rowconfigure(0, minsize=15)

    label_0 = tk.Label(janela3, text="BOMs",
                       font=("Poppins", 10), bg="#EAEAEA")
    label_1 = tk.Label(janela3, text="Código principal",
                       font=("Poppins", 10), bg="#EAEAEA")
    label_2 = tk.Label(janela3, text="Código alternativo",
                       font=("Poppins", 10), bg="#EAEAEA")
    label_3 = tk.Label(janela3, text="Descreva a alteração",
                       font=("Poppins", 10), bg="#EAEAEA")
    label_4 = tk.Label(janela3, text="Engenheiro responsável",
                       font=("Poppins", 10), bg="#EAEAEA")
    label_5 = tk.Label(janela3, text="Previsão de liberação",
                       font=("Poppins", 10), bg="#EAEAEA")
    label_6 = tk.Label(janela3, text="Criar nova versão da BOM?",
                       font=("Poppins", 10), bg="#EAEAEA")
    label_7 = tk.Label(janela3, text="",
                       font=("Poppins", 10), bg="#EAEAEA")

    campo_1 = tk.Entry(janela3)
    campo_2 = tk.Entry(janela3)
    campo_3 = tk.Entry(janela3)
    campo_4 = tk.Entry(janela3)
    campo_5 = tk.Entry(janela3)
    campo_6 = tk.Entry(janela3)

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
    campo_6.configure(bg="#DDEBF7", highlightbackground="#205A8C",
                      highlightthickness=1, borderwidth=0)

    botao_3 = tk.Button(text="Atualizar arquivos", bg="#205A8C", fg="white",
                        activebackground="#205A8C", activeforeground="white", command=lambda: atualizar_arquivos())

    CheckVar1 = tk.IntVar()
    caixa_de_selecao = tk.Checkbutton(
        text="Atualizar todas as BOMs da pasta.", bg="#EAEAEA", variable=CheckVar1, font=("Poppins, 8"))

    options = ["Sim", "Não"]
    selected_option = tk.StringVar()
    selected_option.set("Sim")
    lista_suspensa = ttk.Combobox(
        janela3, textvariable=selected_option, values=options, width=17)

    path = os.path.abspath(
        r"\\terra\GER_PRODUTOS\0 GPd\10 - Trabalhos desenvolvidos\Inserir alternativos - Automação\Arquivos\featured_channel.png")
    imagem = Image.open(path)
    imagem = imagem.resize((30, 30), resample=Image.Resampling.LANCZOS)
    imagem = ImageTk.PhotoImage(imagem)
    label_imagem = tk.Label(janela3, image=imagem, bg="#EAEAEA")
    label_imagem.image = imagem

    label_imagem.grid(row=10, column=3, padx=(0, 20), sticky="NW")
    label_0.grid(row=1, column=1, sticky="W")
    label_1.grid(row=2, column=1, sticky="W")
    label_2.grid(row=3, column=1, sticky="W")
    label_3.grid(row=4, column=1, sticky="W")
    label_4.grid(row=5, column=1, sticky="W")
    label_5.grid(row=6, column=1, sticky="W")
    label_6.grid(row=7, column=1, sticky="W")
    label_7.grid(row=8, column=1, sticky="W")
    campo_1.grid(row=1, column=2)
    campo_2.grid(row=2, column=2)
    campo_3.grid(row=3, column=2)
    campo_4.grid(row=4, column=2)
    campo_5.grid(row=5, column=2)
    campo_6.grid(row=6, column=2)
    botao_3.grid(row=9, column=2, pady=(0, 0))
    lista_suspensa.grid(row=7, column=2)
    caixa_de_selecao.grid(row=8, column=1)

    def validar_data():
        try:
            datetime.datetime.strptime(campo_6.get(), "%d/%m/%Y")
            if CheckVar1.get() == 1:
                resposta = messagebox.askquestion("Confirmação", "Tem certeza que deseja atualizar todas as placas?")

                if resposta == "yes":
                    return True
            else:
                return True
                

        except ValueError:
            messagebox.showerror(
                message=f"O formato da data é invalido.", title="Erro")
            return False

    def atualizar_arquivos():
        if validar_data():
            parametros(path_analiseimpacto, str(campo_1.get()).strip().split(';'), str(campo_2.get()), str(campo_3.get()), str(campo_4.get()), str(
                campo_5.get()), datetime.datetime.strptime(campo_6.get(), "%d/%m/%Y"), str(selected_option.get()), CheckVar1.get(), janela3, responsavel, eco)
