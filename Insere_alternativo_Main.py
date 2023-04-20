import tkinter as tk
import os
from PIL import Image, ImageTk
from Valida_Usuario import validar_usuario


class LoginApp:
    def __init__(self, master):
        self.master = master
        master.title("Login")
        master.geometry("325x200")
        master.resizable(width=False, height=False)
        master.grid_columnconfigure(0, minsize=25)
        master.grid_columnconfigure(1, minsize=50)
        master.grid_columnconfigure(2, minsize=175)
        master.grid_rowconfigure(0, minsize=20)
        master.configure(bg="#EAEAEA")

        label_login = tk.Label(master, text="Usuário",font=("calibri", 12), bg="#EAEAEA")
        label_senha = tk.Label(master, text="Senha",font=("calibri", 12), bg="#EAEAEA")
        self.label_erro = tk.Label(master, text="", bg="#EAEAEA", fg="red")
        self.campo_login = tk.Entry(master)
        self.campo_login.configure(bg="#DDEBF7", highlightbackground="#205A8C", highlightthickness=1, borderwidth=0)
        self.campo_senha = tk.Entry(master, show="*")
        self.campo_senha.configure(bg="#DDEBF7", highlightbackground="#205A8C", highlightthickness=1, borderwidth=0)
        botao_acessar = tk.Button(text="Acessar", bg="#205A8C", fg="white", activebackground="#205A8C", activeforeground="white", command=self.validar_usuario)

        label_login.grid(row=1, column=1, sticky="W")
        label_senha.grid(row=2, column=1, sticky="w")
        self.label_erro.grid(row=3, column=2, sticky="NS")
        self.campo_login.grid(row=1, column=2)
        self.campo_senha.grid(row=2, column=2)
        botao_acessar.grid(row=6, column=2, sticky="n")

        # Adiciona a imagem
        path = os.path.abspath("C:/Users/tc.zcouto/OneDrive - Padtec/Área de Trabalho/Zaire/teste/Atualiza alternativos/logo-padtec.png")
        imagem = Image.open(path)
        imagem = imagem.resize((100, 50), resample=Image.Resampling.LANCZOS)
        imagem = ImageTk.PhotoImage(imagem)
        label_imagem = tk.Label(master, image=imagem, bg="#EAEAEA")
        label_imagem.image = imagem
        label_imagem.grid(row=0, column=2, pady=10)

    def validar_usuario(self):
        resultado = validar_usuario(
            self.campo_login.get(), self.campo_senha.get(), self.master)
        if resultado is not True:
            self.label_erro.configure(text="Usuário e/ou senha incorreto(s)")


janela = tk.Tk()
janela.wm_attributes("-topmost", 1)  # mantém a janela sempre no topo
# define vermelho como cor transparente
janela.wm_attributes("-transparentcolor", "red")
janela.config(background="red")
login_app = LoginApp(janela)
janela.mainloop()
