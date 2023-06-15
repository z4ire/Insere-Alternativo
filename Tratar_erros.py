from tkinter import messagebox

def erro(BOM, wb, erro):

    erros = {'Histórico de Modificação': f'Não foi encontrado a palavra \"Histórico de Modificação\" na BOM {BOM}',
             'Versao': f'Não foi encontrado a palavra \"Versão\" na coluna \"A\" da planilha \"Histórico de Modificação\" da BOM {BOM}',
             'Status': f'Não foi encontrado a palavra \"Status\" na planilha \"Histórico de Modificação\" da BOM {BOM}',
             'Versao da Lista':f'Não foi encontrado a palavra \"Versão da lista\" na coluna \"A\" da BOM {BOM}',
             'Liberacao':f'Não foi encontrado a palavra \"Liberação\" na coluna \"A\" da BOM {BOM}',
             'Modificada por': f'Não foi encontrado a palavra \"Modificada por\" na coluna \"A\" da BOM {BOM}',
             'Status_versao': f'Não foi encontrado a palavra \"Status\" na coluna \"A\" da BOM {BOM}',
             'Emitente': f'Não foi encontrado a palavra \"Emitente\" na coluna \"A\" da BOM {BOM}',
             'Codigo': f'Não foi encontrado a palavra \"Código\" na coluna \"A\" da BOM {BOM}',
             'Codigo_func2':f'Não foi encontrado a palavra \"Código\" na coluna \"A\" da BOM {BOM}'}

    wb.Close(False)
    messagebox.showerror(message=erros[erro], title="ERRO")
    raise ValueError(erros[erro])
