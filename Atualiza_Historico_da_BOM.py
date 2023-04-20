import win32com.client as win32
import pywintypes
import datetime
from win32api import RGB
from pathlib import Path
from tkinter import messagebox


def atualizar_historico(wb, n_nova_versao, merge_range, eco, codigo_principal, codigo_alternativo, descricao, engenheiro, responsavel, descricao_comp, BOM):
    try:
        historico = wb.Worksheets("Histórico de Modificação")
    except:
        wb.Close(False)
        messagebox.showerror(message=f"Não foi encontrado uma planilha com o nome \"Histórico de Modificação\" na BOM {BOM}", title="ERRO")
        raise ValueError( f"Não foi encontrado uma planilha com o nome \"Histórico de Modificação\" na BOM {BOM}")


    celula_versao = historico.Cells.Range("A:A").Find("Versão")
    if celula_versao is None:
        wb.Close(False)
        messagebox.showerror(message=f"Não foi encontrado a palavra \"Versão\" na coluna \"A\" da planilha \"Histórico de Modificação\" da BOM {BOM}", title="ERRO")
        raise ValueError( f"Não foi encontrado a palavra \"Versão\" na coluna \"A\" da planilha \"Histórico de Modificação\" da BOM {BOM}")
    
    celula_status = historico.Range("A" + str(celula_versao.Row) + ":Z" + str(celula_versao.Row)).Find("Status", LookIn=win32.constants.xlValues, LookAt=win32.constants.xlPart, SearchDirection=win32.constants.xlNext, SearchOrder=win32.constants.xlByColumns)
    
    if celula_status is None:
        wb.Close(False)
        messagebox.showerror(message=f"Não foi encontrado a palavra \"Status\" na planilha \"Histórico de Modificação\" da BOM {BOM}", title="ERRO")
        raise ValueError( f"Não foi encontrado a palavra \"Status\" na planilha \"Histórico de Modificação\" da BOM {BOM}")
    
    strBusca = "EOL"
    celula_status_EOL = historico.Columns(celula_status.Column).Find(strBusca, LookIn=win32.constants.xlValues, LookAt=win32.constants.xlPart, SearchDirection=win32.constants.xlNext, SearchOrder=win32.constants.xlByColumns)

    if celula_status_EOL is not None:
        celula_status_EOL.Value = "Não Liberado"
        celula_status_EOL.FormatConditions.Delete()
        celula_status_EOL.Font.Bold = True
        celula_status_EOL.Font.Color = RGB(255, 255, 255)
        celula_status_EOL.Interior.Color = RGB(128, 0, 0)
    else:
        strBusca = "Liberado"
        celula_status_liberado = historico.Columns(celula_status.Column).Find(strBusca, LookIn=win32.constants.xlValues, LookAt=win32.constants.xlWhole,
                                                                              SearchDirection=win32.constants.xlNext, SearchOrder=win32.constants.xlByColumns)
        if celula_status_liberado is not None:
            celula_status_liberado.Value = "EOL"
            celula_status_liberado.FormatConditions.Delete()
            celula_status_liberado.Font.Bold = True
            celula_status_liberado.Font.Color = RGB(0, 0, 0)
            celula_status_liberado.Interior.Color = RGB(255, 255, 0)

    strBusca = "Liberado"
    celula_status_liberado = historico.Columns(celula_status.Column).Find(strBusca, LookIn=win32.constants.xlValues, LookAt=win32.constants.xlPart,
                                                                          SearchDirection=win32.constants.xlPrevious, SearchOrder=win32.constants.xlByColumns, After=celula_status)

    if celula_status_liberado is not None:
        celula_status_liberado.Value = "EOL"
        celula_status_liberado.FormatConditions.Delete()
        celula_status_liberado.Font.Bold = True
        celula_status_liberado.Font.Color = RGB(0, 0, 0)
        celula_status_liberado.Interior.Color = RGB(255, 255, 0)

    linha_atual = celula_status.Row + 1
    celula_atual = historico.Cells(linha_atual, celula_status.Column)

    while celula_atual.MergeCells:
        linha_atual = celula_atual.MergeArea.Row + celula_atual.MergeArea.Rows.Count
        celula_atual = historico.Cells(linha_atual, celula_status.Column)
    while celula_atual.Value is not None or celula_atual.MergeCells:
        linha_atual += 1
        celula_atual = historico.Cells(linha_atual, celula_status.Column)
        while celula_atual.MergeCells:
            linha_atual = celula_atual.MergeArea.Row + celula_atual.MergeArea.Rows.Count
            celula_atual = historico.Cells(linha_atual, celula_status.Column)

    ultima_linha = linha_atual

    # copiar a linha logo abaixo da linha da palavra "Versão"
    nova_linha = historico.Cells(int(celula_versao.Row)+1, 1).EntireRow
    # cria a nova versão copiando a V1
    historico.Rows(ultima_linha).Insert()
    nova_linha.Copy(historico.Cells(ultima_linha, 1))

    hiperlink = "\\\\loki\\PADTEC - Campinas\\Tecnologia\\Hardware\\Transferencia_PRO\\ECO\\AnaliseImpacto\\ECOs_ " + \
        eco[4:8] + "\\" + eco + "\\" + eco + ".xlsm"

    # Preenche a nova versão criada no histórico de modificações
    historico.Cells(ultima_linha, 1).Font.Strikethrough = False
    historico.Cells(ultima_linha, 1).Value = n_nova_versao
    historico.Cells(ultima_linha, 2).Font.Strikethrough = False
    historico.Cells(ultima_linha, 2).Value = eco
    historico.Cells(ultima_linha, 3).Font.Strikethrough = False
    historico.Cells(ultima_linha, 3).Hyperlinks.Add(Anchor=historico.Cells(ultima_linha, 3), Address=hiperlink)
    historico.Cells(ultima_linha, 3).Value = "Download"
    historico.Cells(ultima_linha, 4).Font.Strikethrough = False
    historico.Cells(ultima_linha, 4).Value = codigo_alternativo
    historico.Cells(ultima_linha, 4).Font.Color = RGB(0, 0, 255)
    historico.Cells(ultima_linha, 5).Font.Strikethrough = False
    historico.Cells(ultima_linha, 5).Value = descricao_comp
    historico.Cells(ultima_linha, 5).Font.Color = RGB(0, 0, 255)
    historico.Cells(ultima_linha, 6).Font.Strikethrough = False
    historico.Cells(ultima_linha, 6).Value = merge_range.Cells(1, 3).Value
    historico.Cells(ultima_linha, 6).Font.Color = RGB(0, 0, 255)
    historico.Cells(ultima_linha, 7).Font.Strikethrough = False
    historico.Cells(ultima_linha, 7).Value = merge_range.Cells(1, 4).Value
    historico.Cells(ultima_linha, 8).Font.Strikethrough = False
    historico.Cells(ultima_linha, 8).Value = descricao
    historico.Cells(ultima_linha, 9).Font.Strikethrough = False
    historico.Cells(ultima_linha, 9).Value = engenheiro
    historico.Cells(ultima_linha, 10).Font.Strikethrough = False
    historico.Cells(ultima_linha, 10).Value = pywintypes.Time(datetime.date.today())
    historico.Cells(ultima_linha, 11).Font.Strikethrough = False
    # Verificar com um IF se é SIM ou OK, dependendo do arquivo da
    historico.Cells(ultima_linha, 11).Value = "Ok"
    if (historico.Cells(str(celula_versao.Row), 12).Value == "Afeta Qualidade - Modificação Obrigatória"):
        historico.Cells(ultima_linha, 12).Font.Strikethrough = False
        historico.Cells(ultima_linha, 12).Value = "Não"

    historico.Cells(
        ultima_linha, celula_status.Column).Font.Strikethrough = False
    historico.Cells(ultima_linha, celula_status.Column).Value = "Liberado"
    historico.Cells(
        ultima_linha, celula_status.Column).FormatConditions.Delete()
    historico.Cells(ultima_linha, celula_status.Column).Font.Bold = True
    historico.Cells(
        ultima_linha, celula_status.Column).Font.Color = RGB(0, 0, 0)
    historico.Cells(ultima_linha, celula_status.Column).Interior.Color = RGB(
        51, 153, 102)
    quantidade = historico.Cells(ultima_linha, 6).Value
    designator = historico.Cells(ultima_linha, 7).Value
    wb.Save()
    wb.Close()
    return quantidade, designator
