import tkinter as tk
from Escolhe_ECO import escolher_ECO


path_htm = r'\\loki\PADTEC - Campinas\Tecnologia\Hardware\Transferencia_PRO\Web\htm'
path_xls = r'\\loki\PADTEC - Campinas\Tecnologia\Hardware\Transferencia_PRO\Web\xls'
path_bom = r'\\loki\PADTEC - Campinas\Tecnologia\Hardware\Transferencia_PRO\BOM'
path_ep = r'\\loki\PADTEC - Campinas\Tecnologia\Hardware\Transferencia_PRO\Produto\EPs'


def validar_usuario(campo_1, campo_2, janela):
    login = campo_1
    senha = campo_2
    usuarios = {"zaire": ["Zaire Couto", "asdf"],
                "vinicius": ["Vinicius Lopes", "asdf"],
                "favrin": ["Guilherme Favrin", "asdf"],
                "tiago": ["Tiago Simões", "asdf"],
                "login": ["Usuario x", "senha"]}

    if login in usuarios and senha == usuarios[login][1]:
        responsavel = usuarios[login][0]
        escolher_ECO(responsavel, janela)
        return True
    else:
        return False
