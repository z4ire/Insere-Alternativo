from tkinter import messagebox
import sys


def escolher_versao(wb, BOM, excel):

    lista_de_versoes = []
    for j in range(1, wb.Sheets.Count):
        versao = wb.Sheets(j).Name
        if versao[:1] == 'V':
            try:
                lista_de_versoes.append(int(versao[1:]))
            except:
                messagebox.showerror(
                    message=f"O nome \"{versao}\" em {BOM} deve ser alterado.", title="ERRO")
                wb.Close(False)
                excel.Quit()
                sys.exit(0)

    try:
        versao_atual = 'V' + str(max(lista_de_versoes, default=0))
        versao_intermediaria = versao_atual + " (2)"
        nova_versao = 'V' + str(max(lista_de_versoes, default=0) + 1)
        n_nova_versao = str(max(lista_de_versoes, default=0) + 1)

    except:
        messagebox.showerror(message=f"Existe algum problema com algum dos nomes das planilhas no arquivo {BOM}.", title="ERRO")
        wb.Close(False)
        excel.Quit()
        sys.exit(0)

    return versao_atual, versao_intermediaria, nova_versao, n_nova_versao
