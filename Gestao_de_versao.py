import win32com.client as win32
import pywintypes
import datetime
from Tratar_erros import erro
from win32api import RGB
from tkinter import messagebox
from pathlib import Path


def criar_nova_versao(wb, versao_atual, versao_intermediaria, nova_versao, n_nova_versao, BOM, codigo_principal, responsavel):

    buscar = {"Versão da lista": "Versão da lista:",
              "Liberação": "Liberação:",
              "Modificada por": "Modificada por:",
              "Status": "Status:",
              "Emitente": "Emitente:",
              "Código": "Código"}

    worksheet_original = wb.Worksheets(versao_atual)
    worksheet_clonada = worksheet_original.Copy(None, worksheet_original)
    worksheet_clonada = wb.Worksheets(versao_intermediaria)
    worksheet_clonada.Name = nova_versao

    # Encontrar as células de interesse na coluna A
    versao_lista_celula = worksheet_clonada.Cells.Range(
        "A:A").Find(buscar["Versão da lista"])

    if versao_lista_celula is not None:
        versao_lista_celula.Offset(1, 2).Value = n_nova_versao
    else:
        erro(BOM, wb, 'Versao da Lista')
        # wb.Close(False)
        # messagebox.showerror( message=f"Não foi encontrado a palavra \"{buscar['Versão da lista']}\" na coluna \"A\" da versão \"{versao_atual}\" da BOM {BOM}", title="ERRO")

        # raise ValueError(f"Não foi encontrado a palavra \"{buscar['Versão da lista']}\" na coluna \"A\" da versão \"{versao_atual}\" da BOM {BOM}")

    liberacao_celula = worksheet_clonada.Cells.Range(
        "A:A").Find(buscar["Liberação"])

    if liberacao_celula is not None:
        liberacao_celula.Offset(1, 2).Value = pywintypes.Time(
            datetime.date.today())
    else:
        erro(BOM, wb, 'Liberacao')
        # wb.Close(False)
        # messagebox.showerror(
        #     message=f"Não foi encontrado a palavra \"{buscar['Liberação']}\" na coluna \"A\" da versão \"{versao_atual}\" da BOM {BOM}", title="ERRO")
        # raise ValueError(
        #     f"Não foi encontrado a palavra \"{buscar['Liberação']}\" na coluna \"A\" da versão \"{versao_atual}\" da BOM {BOM}")

    modificada_por_celula = worksheet_clonada.Cells.Range(
        "A:A").Find(buscar["Modificada por"])
    if modificada_por_celula is not None:
        modificada_por_celula.Offset(1, 2).Value = responsavel
    else:
        erro(BOM, wb, 'Modificada por')
        # wb.Close(False)
        # messagebox.showerror(
        #     message=f"Não foi encontrado a palavra \"{buscar['Modificada por']}\" na coluna \"A\" da versão \"{versao_atual}\" da BOM {BOM}", title="ERRO")
        # raise ValueError(
        #     f"Não foi encontrado a palavra \"{buscar['Modificada por']}\" na coluna \"A\" da versão \"{versao_atual}\" da BOM {BOM}")

    status_celula = worksheet_clonada.Cells.Range("A:A").Find(buscar["Status"])
    if status_celula is not None:
        status_celula.Offset(1, 2).Value = "Criado Nova Versão"
    else:
        erro(BOM, wb, 'Status_versao')
        # wb.Close(False)
        # messagebox.showerror(
        #     message=f"Não foi encontrado a palavra \"{buscar['Status']}\" na coluna \"A\" da versão \"{versao_atual}\" da BOM {BOM}", title="ERRO")
        # raise ValueError(
        #     f"Não foi encontrado a palavra \"{buscar['Status']}\" na coluna \"A\" da versão \"{versao_atual}\" da BOM {BOM}")

    emitente_celula = worksheet_clonada.Cells.Range(
        "A:A").Find(buscar["Emitente"])
    if emitente_celula is not None:
        emitente_celula.Offset(1, 2).Value = responsavel
    else:
        erro(BOM, wb, 'Emitente')
        # wb.Close(False)
        # messagebox.showerror(
        #     message=f"Não foi encontrado a palavra \"{buscar['Emitente']}\" na coluna \"A\" da versão \"{versao_atual}\" da BOM {BOM}", title="ERRO")
        # raise ValueError(
        #     f"Não foi encontrado a palavra \"{buscar['Emitente']}\" na coluna \"A\" da versão \"{versao_atual}\" da BOM {BOM}")

    celula_codigo = worksheet_clonada.Cells.Range("A:A").Find(buscar["Código"])

    # Verifique se a célula foi encontrada
    if celula_codigo:
        linha_inicial = celula_codigo.Row + 1
        linha_final = linha_inicial + 4999
        worksheet_clonada.Range(
            f"A{linha_inicial}:I{linha_final}").Interior.Color = win32.constants.xlColorIndexNone
    else:
        erro(BOM, wb, 'Codigo')
        # wb.Close(False)
        # messagebox.showerror(
        #     message=f"Não foi encontrado a palavra \"{buscar['Código']}\" na coluna \"A\" da versão \"{versao_atual}\" da BOM {BOM}", title="ERRO")
        # raise ValueError(
        #     f"Não foi encontrado a palavra \"{buscar['Código']}\" na coluna \"A\" da versão \"{versao_atual}\" da BOM {BOM}")

    for coluna in range(1, worksheet_clonada.UsedRange.Columns.Count + 1):
        if worksheet_clonada.Cells(celula_codigo.Row, coluna).Value is not None:
            ultima_coluna = coluna
    ultima_coluna = chr(ultima_coluna + 64)

    celula_codigo_principal = worksheet_clonada.Cells.Range(
        "A:A").Find(codigo_principal, LookIn=win32.constants.xlValues)

    if celula_codigo_principal is not None:
        merge_range = celula_codigo_principal.MergeArea

    else:
        print(f"A BOM {BOM} não contém o código {codigo_principal}\n")

    return worksheet_clonada, celula_codigo, ultima_coluna, celula_codigo_principal


def manter_versao_atual(wb, versao_atual, BOM, codigo_principal):

    worksheet_original = wb.Worksheets(versao_atual)

    buscar = {"Código": "Código"}

    celula_codigo = worksheet_original.Cells.Range(
        "A:A").Find(buscar["Código"])

    # Verifique se a célula foi encontrada
    if not celula_codigo:
        erro(BOM, wb, 'Codigo_func2')
        # wb.Close(False)
        # messagebox.showerror(message=f"Não foi encontrado a palavra \"{buscar['Código']}\" na coluna \"A\" da versão \"{versao_atual}\" da BOM {BOM}", title="ERRO")
        # raise ValueError(f"Não foi encontrado a palavra \"{buscar['Código']}\" na coluna \"A\" da versão \"{versao_atual}\" da BOM {BOM}")

    for coluna in range(1, worksheet_original.UsedRange.Columns.Count + 1):
        if worksheet_original.Cells(celula_codigo.Row, coluna).Value is not None:
            ultima_coluna = coluna

    ultima_coluna = chr(ultima_coluna + 64)

    celula_codigo_principal = worksheet_original.Cells.Range(
        "A:A").Find(codigo_principal, LookIn=win32.constants.xlValues)

    if celula_codigo_principal is not None:
        merge_range = celula_codigo_principal.MergeArea

    else:
        print(f"A BOM {BOM} não contém o código {codigo_principal}\n")

    return worksheet_original, celula_codigo, ultima_coluna, celula_codigo_principal
