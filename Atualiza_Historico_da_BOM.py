import win32com.client as win32
import pywintypes
from os.path import dirname
from os.path import join
from Tratar_erros import erro
from win32api import RGB
from win32com.client import constants


def atualizar_historico_V1(wb, path_analiseimpacto, n_nova_versao, merge_range, eco, codigo_alternativo, descricao, engenheiro, data, descricao_comp, BOM):
    try:
        historico = wb.Worksheets("Histórico de Modificação")
    except:
        erro(BOM, wb, 'Histórico de Modificação')

        # wb.Close(False)
        # messagebox.showerror(message=f"Não foi encontrado uma planilha com o nome \"Histórico de Modificação\" na BOM {BOM}", title="ERRO")
        # raise ValueError( f"Não foi encontrado uma planilha com o nome \"Histórico de Modificação\" na BOM {BOM}")

    celula_versao = historico.Cells.Range("A:A").Find("Versão")

    if celula_versao is None:
        erro(BOM, wb, 'Versao')

        # wb.Close(False)
        # messagebox.showerror(message=f"Não foi encontrado a palavra \"Versão\" na coluna \"A\" da planilha \"Histórico de Modificação\" da BOM {BOM}", title="ERRO")
        # raise ValueError(f"Não foi encontrado a palavra \"Versão\" na coluna \"A\" da planilha \"Histórico de Modificação\" da BOM {BOM}")

    celula_status = historico.Range("A" + str(celula_versao.Row) + ":Z" + str(celula_versao.Row)).Find("Status", LookIn=win32.constants.xlValues,
                                                                                                       LookAt=win32.constants.xlPart, SearchDirection=win32.constants.xlNext, SearchOrder=win32.constants.xlByColumns)

    if celula_status is None:
        erro(BOM, wb, 'Status')

        # wb.Close(False)
        # messagebox.showerror(message=f"Não foi encontrado a palavra \"Status\" na planilha \"Histórico de Modificação\" da BOM {BOM}", title="ERRO")
        # raise ValueError(f"Não foi encontrado a palavra \"Status\" na planilha \"Histórico de Modificação\" da BOM {BOM}")

    def formatacao_condicional(celula):
        celula.FormatConditions.Delete()
        # Cria uma lista de regras de formatação condicional
        regras_formatacao = [
            {"condicao": "Não Liberado", "cor_fundo": RGB(
                128, 0, 0), "cor_fonte": RGB(255, 255, 255)},
            {"condicao": "Liberado", "cor_fundo": RGB(
                51, 153, 102), "cor_fonte": RGB(0, 0, 0)},
            {"condicao": "EOL", "cor_fundo": RGB(
                255, 255, 0), "cor_fonte": RGB(0, 0, 0)}
        ]

        # Adiciona cada regra de formatação condicional e aplica as configurações
        for regra in regras_formatacao:
            nova_formatacao = celula.FormatConditions.Add(
                Type=win32.constants.xlCellValue, Operator=win32.constants.xlEqual, Formula1=regra["condicao"])
            nova_formatacao.Font.Color = regra["cor_fonte"]
            nova_formatacao.Font.Bold = True
            nova_formatacao.Interior.Color = regra["cor_fundo"]

    strBusca = "EOL"
    celula_status_EOL = historico.Columns(celula_status.Column).Find(strBusca, LookIn=win32.constants.xlValues,
                                                                     LookAt=win32.constants.xlPart, SearchDirection=win32.constants.xlNext, SearchOrder=win32.constants.xlByColumns)
    if celula_status_EOL is not None:
        formatacao_condicional(celula_status_EOL)
        # celula_status_EOL.Value = "Não Liberado"
        # celula_status_EOL.FormatConditions.Delete()
        # celula_status_EOL.Font.Bold = True
        # celula_status_EOL.Font.Color = RGB(255, 255, 255)
        # celula_status_EOL.Interior.Color = RGB(128, 0, 0)
    else:
        strBusca = "Liberado"
        celula_status_liberado = historico.Columns(celula_status.Column).Find(
            strBusca, LookIn=win32.constants.xlValues, LookAt=win32.constants.xlWhole, SearchDirection=win32.constants.xlNext, SearchOrder=win32.constants.xlByColumns)
        if celula_status_liberado is not None:
            formatacao_condicional(celula_status_liberado)
            celula_status_liberado.Value = "EOL"
            # celula_status_liberado.FormatConditions.Delete()
            # celula_status_liberado.Font.Bold = True
            # celula_status_liberado.Font.Color = RGB(0, 0, 0)
            # celula_status_liberado.Interior.Color = RGB(255, 255, 0)

    strBusca = "Liberado"
    celula_status_liberado = historico.Columns(celula_status.Column).Find(strBusca, LookIn=win32.constants.xlValues, LookAt=win32.constants.xlPart,
                                                                          SearchDirection=win32.constants.xlPrevious, SearchOrder=win32.constants.xlByColumns, After=celula_status)

    if celula_status_liberado is not None:
        formatacao_condicional(celula_status_liberado)
        celula_status_liberado.Value = "EOL"
        # celula_status_liberado.FormatConditions.Delete()
        # celula_status_liberado.Font.Bold = True
        # celula_status_liberado.Font.Color = RGB(0, 0, 0)
        # celula_status_liberado.Interior.Color = RGB(255, 255, 0)

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

    hiperlink = "\\\\urano\\Storage\\PADTEC - Campinas\\Tecnologia\\Hardware\\Transferencia_PRO\\ECO\\AnaliseImpacto\\ECOs_ " + \
        eco[4:8] + "\\" + eco + "\\" + eco + ".xlsm"

    # Preenche a nova versão criada no histórico de modificações
    historico.Cells(ultima_linha, 1).Font.Strikethrough = False
    historico.Cells(ultima_linha, 1).Value = n_nova_versao
    historico.Cells(ultima_linha, 2).Font.Strikethrough = False
    historico.Cells(ultima_linha, 2).Value = eco
    historico.Cells(ultima_linha, 3).Font.Strikethrough = False

    if historico.Cells(ultima_linha - 1, 3).Value == '-' or historico.Cells(ultima_linha - 1, 3).Value == ' - ':
        historico.Cells(
            ultima_linha, 3).Formula = f'=HIPERLINK("{hiperlink}", "Download")'

    else:
        celula_acima = historico.Cells(ultima_linha - 1, 3)
        # Verifica se a célula acima está mesclada
        if historico.Cells(ultima_linha - 1, 3).MergeCells:
            # Se estiver mesclada, seleciona a primeira célula da mesclagem
            celula_acima = celula_acima.MergeArea.Cells(1)

        # Copia a formatação da célula acima
        celula_acima.Copy()
        historico.Cells(ultima_linha, 3).PasteSpecial(constants.xlPasteAll)
        historico.Cells(ultima_linha, 3).Calculate
        historico.Cells(ultima_linha, 3).Borders.LineStyle = 1

    historico.Cells(ultima_linha, 4).Value = codigo_alternativo
    historico.Cells(ultima_linha, 4).WrapText = True
    historico.Cells(ultima_linha, 4).Font.Color = RGB(0, 0, 255)
    historico.Cells(ultima_linha, 4).Borders.LineStyle = 1
    historico.Cells(ultima_linha, 4).Font.Name = "Arial"
    historico.Cells(ultima_linha, 4).Font.Size = 8
    historico.Cells(ultima_linha, 4).HorizontalAlignment = -4108
    historico.Cells(ultima_linha, 4).VerticalAlignment = -4108
    historico.Cells(ultima_linha, 4).Font.Strikethrough = False

    historico.Cells(ultima_linha, 5).Value = descricao_comp
    historico.Cells(ultima_linha, 5).WrapText = True
    historico.Cells(ultima_linha, 5).Font.Color = RGB(0, 0, 255)
    historico.Cells(ultima_linha, 5).Borders.LineStyle = 1
    historico.Cells(ultima_linha, 5).Font.Name = "Arial"
    historico.Cells(ultima_linha, 5).Font.Size = 8
    historico.Cells(ultima_linha, 5).HorizontalAlignment = -4108
    historico.Cells(ultima_linha, 5).VerticalAlignment = -4108
    historico.Cells(ultima_linha, 5).Font.Strikethrough = False

    historico.Cells(ultima_linha, 6).Value = merge_range.Cells(1, 3).Value
    historico.Cells(ultima_linha, 6).WrapText = True
    historico.Cells(ultima_linha, 6).Font.Color = RGB(0, 0, 0)
    historico.Cells(ultima_linha, 6).Borders.LineStyle = 1
    historico.Cells(ultima_linha, 6).Font.Name = "Arial"
    historico.Cells(ultima_linha, 6).Font.Size = 8
    historico.Cells(ultima_linha, 6).HorizontalAlignment = -4108
    historico.Cells(ultima_linha, 6).VerticalAlignment = -4108
    historico.Cells(ultima_linha, 6).Font.Strikethrough = False

    historico.Cells(ultima_linha, 7).Value = merge_range.Cells(1, 4).Value
    historico.Cells(ultima_linha, 7).WrapText = True
    historico.Cells(ultima_linha, 7).Font.Color = RGB(0, 0, 0)
    historico.Cells(ultima_linha, 7).Borders.LineStyle = 1
    historico.Cells(ultima_linha, 7).Font.Name = "Arial"
    historico.Cells(ultima_linha, 7).Font.Size = 8
    historico.Cells(ultima_linha, 7).HorizontalAlignment = -4108
    historico.Cells(ultima_linha, 7).VerticalAlignment = -4108
    historico.Cells(ultima_linha, 7).Font.Strikethrough = False

    historico.Cells(ultima_linha, 8).Value = descricao
    historico.Cells(ultima_linha, 8).WrapText = True
    historico.Cells(ultima_linha, 8).Font.Color = RGB(0, 0, 0)
    historico.Cells(ultima_linha, 8).Borders.LineStyle = 1
    historico.Cells(ultima_linha, 8).Font.Name = "Arial"
    historico.Cells(ultima_linha, 8).Font.Size = 8
    historico.Cells(ultima_linha, 8).HorizontalAlignment = -4108
    historico.Cells(ultima_linha, 8).VerticalAlignment = -4108
    historico.Cells(ultima_linha, 8).Font.Strikethrough = False

    historico.Cells(ultima_linha, 9).Value = engenheiro
    historico.Cells(ultima_linha, 9).WrapText = True
    historico.Cells(ultima_linha, 9).Font.Color = RGB(0, 0, 0)
    historico.Cells(ultima_linha, 9).Borders.LineStyle = 1
    historico.Cells(ultima_linha, 9).Font.Name = "Arial"
    historico.Cells(ultima_linha, 9).Font.Size = 8
    historico.Cells(ultima_linha, 9).HorizontalAlignment = -4108
    historico.Cells(ultima_linha, 9).VerticalAlignment = -4108
    historico.Cells(ultima_linha, 9).Font.Strikethrough = False

    historico.Cells(ultima_linha, 10).Value = pywintypes.Time(data)
    historico.Cells(ultima_linha, 10).WrapText = True
    historico.Cells(ultima_linha, 10).Font.Color = RGB(0, 0, 0)
    historico.Cells(ultima_linha, 10).Borders.LineStyle = 1
    historico.Cells(ultima_linha, 10).Font.Name = "Arial"
    historico.Cells(ultima_linha, 10).Font.Size = 8
    historico.Cells(ultima_linha, 10).HorizontalAlignment = -4108
    historico.Cells(ultima_linha, 10).VerticalAlignment = -4108
    historico.Cells(ultima_linha, 10).Font.Strikethrough = False

    historico.Cells(ultima_linha, 11).Value = "Ok"
    historico.Cells(ultima_linha, 10).WrapText = True
    historico.Cells(ultima_linha, 11).Font.Color = RGB(0, 0, 0)
    historico.Cells(ultima_linha, 11).Borders.LineStyle = 1
    historico.Cells(ultima_linha, 11).Font.Name = "Arial"
    historico.Cells(ultima_linha, 11).Font.Size = 8
    historico.Cells(ultima_linha, 11).HorizontalAlignment = -4108
    historico.Cells(ultima_linha, 11).VerticalAlignment = -4108
    historico.Cells(ultima_linha, 11).Font.Strikethrough = False

    if (historico.Cells(str(celula_versao.Row), 12).Value == "Afeta Qualidade - Modificação Obrigatória"):

        historico.Cells(ultima_linha, 12).Value = "Não"
        historico.Cells(ultima_linha, 12).WrapText = True
        historico.Cells(ultima_linha, 12).Font.Color = RGB(0, 0, 0)
        historico.Cells(ultima_linha, 12).Borders.LineStyle = 1
        historico.Cells(ultima_linha, 12).Font.Name = "Arial"
        historico.Cells(ultima_linha, 12).Font.Size = 8
        historico.Cells(ultima_linha, 12).HorizontalAlignment = -4108
        historico.Cells(ultima_linha, 12).VerticalAlignment = -4108
        historico.Cells(ultima_linha, 12).Font.Strikethrough = False

    formatacao_condicional(historico.Cells(ultima_linha, celula_status.Column))
    historico.Cells(ultima_linha, celula_status.Column).Value = "Liberado"
    historico.Cells(ultima_linha, celula_status.Column).Borders.LineStyle = 1
    historico.Cells(ultima_linha, celula_status.Column).Font.Name = "Arial"
    historico.Cells(ultima_linha, celula_status.Column).Font.Size = 8
    historico.Cells(
        ultima_linha, celula_status.Column).HorizontalAlignment = -4108
    historico.Cells(
        ultima_linha, celula_status.Column).VerticalAlignment = -4108
    historico.Cells(
        ultima_linha, celula_status.Column).Font.Strikethrough = False

    quantidade = historico.Cells(ultima_linha, 6).Value
    designator = historico.Cells(ultima_linha, 7).Value

    path_analiseimpacto = join(dirname(path_analiseimpacto), "Atualizado")
    path_analiseimpacto = join(path_analiseimpacto, wb.Name)
    wb.SaveAs(path_analiseimpacto)
    wb.Close()

    return quantidade, designator


def atualizar_historico_teste(wb, path_analiseimpacto, merge_range, codigo_alternativo, descricao, engenheiro, descricao_comp, BOM):

    try:
        historico = wb.Worksheets("Histórico de Modificação")

    except:
        erro(BOM, wb, 'Histórico de Modificação')

    celula_versao = historico.Cells.Range("A:A").Find("Versão")

    if celula_versao is None:
        erro(BOM, wb, 'Versao')

    celula_status = historico.Range("A" + str(celula_versao.Row) + ":Z" + str(celula_versao.Row)).Find("Status", LookIn=win32.constants.xlValues,
                                                                                                       LookAt=win32.constants.xlPart, SearchDirection=win32.constants.xlNext, SearchOrder=win32.constants.xlByColumns)

    if celula_status is None:
        erro(BOM, wb, 'Status')

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

    quantidade = historico.Cells(ultima_linha - 1, 6).Value
    designator = historico.Cells(ultima_linha - 1, 7).Value

    historico.Range(historico.Cells(ultima_linha, 1), historico.Cells(ultima_linha - 1, 1)).Merge()
    historico.Cells(ultima_linha, 1).Borders.LineStyle = 1

    historico.Range(historico.Cells(ultima_linha, 2), historico.Cells(ultima_linha - 1, 2)).Merge()
    historico.Cells(ultima_linha, 2).Borders.LineStyle = 1

    historico.Range(historico.Cells(ultima_linha, 3), historico.Cells(ultima_linha - 1, 3)).Merge()
    historico.Cells(ultima_linha, 3).Borders.LineStyle = 1

    historico.Cells(ultima_linha, 4).Value = codigo_alternativo
    historico.Cells(ultima_linha, 4).WrapText = True
    historico.Cells(ultima_linha, 4).Font.Color = RGB(0, 0, 255)
    historico.Cells(ultima_linha, 4).Borders.LineStyle = 1
    historico.Cells(ultima_linha, 4).Font.Name = "Arial"
    historico.Cells(ultima_linha, 4).Font.Size = 8
    historico.Cells(ultima_linha, 4).HorizontalAlignment = -4108
    historico.Cells(ultima_linha, 4).VerticalAlignment = -4108
    historico.Cells(ultima_linha, 4).Font.Strikethrough = False

    historico.Cells(ultima_linha, 5).Value = descricao_comp
    historico.Cells(ultima_linha, 5).WrapText = True
    historico.Cells(ultima_linha, 5).Font.Color = RGB(0, 0, 255)
    historico.Cells(ultima_linha, 5).Borders.LineStyle = 1
    historico.Cells(ultima_linha, 5).Font.Name = "Arial"
    historico.Cells(ultima_linha, 5).Font.Size = 8
    historico.Cells(ultima_linha, 5).HorizontalAlignment = -4108
    historico.Cells(ultima_linha, 5).VerticalAlignment = -4108
    historico.Cells(ultima_linha, 5).Font.Strikethrough = False

    historico.Cells(ultima_linha, 6).Value = merge_range.Cells(1, 3).Value
    historico.Cells(ultima_linha, 6).WrapText = True
    historico.Cells(ultima_linha, 6).Font.Color = RGB(0, 0, 0)
    historico.Cells(ultima_linha, 6).Borders.LineStyle = 1
    historico.Cells(ultima_linha, 6).Font.Name = "Arial"
    historico.Cells(ultima_linha, 6).Font.Size = 8
    historico.Cells(ultima_linha, 6).HorizontalAlignment = -4108
    historico.Cells(ultima_linha, 6).VerticalAlignment = -4108
    historico.Cells(ultima_linha, 6).Font.Strikethrough = False

    historico.Cells(ultima_linha, 7).Value = merge_range.Cells(1, 4).Value
    historico.Cells(ultima_linha, 7).WrapText = True
    historico.Cells(ultima_linha, 7).Font.Color = RGB(0, 0, 0)
    historico.Cells(ultima_linha, 7).Borders.LineStyle = 1
    historico.Cells(ultima_linha, 7).Font.Name = "Arial"
    historico.Cells(ultima_linha, 7).Font.Size = 8
    historico.Cells(ultima_linha, 7).HorizontalAlignment = -4108
    historico.Cells(ultima_linha, 7).VerticalAlignment = -4108
    historico.Cells(ultima_linha, 7).Font.Strikethrough = False

    if descricao == "":
        historico.Range(historico.Cells(ultima_linha, 8), historico.Cells(ultima_linha - 1, 8)).Merge()
        historico.Cells(ultima_linha, 8).Borders.LineStyle = 1
    else:
        historico.Cells(ultima_linha, 8).Value = descricao
        historico.Cells(ultima_linha, 8).WrapText = True
        historico.Cells(ultima_linha, 8).Font.Color = RGB(0, 0, 0)
        historico.Cells(ultima_linha, 8).Borders.LineStyle = 1
        historico.Cells(ultima_linha, 8).Font.Name = "Arial"
        historico.Cells(ultima_linha, 8).Font.Size = 8
        historico.Cells(ultima_linha, 8).HorizontalAlignment = -4108
        historico.Cells(ultima_linha, 8).VerticalAlignment = -4108
        historico.Cells(ultima_linha, 8).Font.Strikethrough = False

    if engenheiro == "":
        historico.Range(historico.Cells(ultima_linha, 9), historico.Cells(ultima_linha - 1, 9)).Merge()
        historico.Cells(ultima_linha, 9).Borders.LineStyle = 1
    else:
        historico.Cells(ultima_linha, 9).Value = engenheiro
        historico.Cells(ultima_linha, 9).WrapText = True
        historico.Cells(ultima_linha, 9).Font.Color = RGB(0, 0, 0)
        historico.Cells(ultima_linha, 9).Borders.LineStyle = 1
        historico.Cells(ultima_linha, 9).Font.Name = "Arial"
        historico.Cells(ultima_linha, 9).Font.Size = 8
        historico.Cells(ultima_linha, 9).HorizontalAlignment = -4108
        historico.Cells(ultima_linha, 9).VerticalAlignment = -4108
        historico.Cells(ultima_linha, 9).Font.Strikethrough = False

    historico.Range(historico.Cells(ultima_linha, 10), historico.Cells(ultima_linha - 1, 10)).Merge()
    historico.Cells(ultima_linha - 1, 10).Borders.LineStyle = 1
    historico.Cells(ultima_linha, 10).Borders.LineStyle = 1

    historico.Range(historico.Cells(ultima_linha, 11), historico.Cells(ultima_linha - 1, 11)).Merge()
    historico.Cells(ultima_linha - 1, 11).Borders.LineStyle = 1
    historico.Cells(ultima_linha, 11).Borders.LineStyle = 1

    if (historico.Cells(str(celula_versao.Row), 12).Value == "Afeta Qualidade - Modificação Obrigatória"):
        historico.Range(historico.Cells(ultima_linha, 12), historico.Cells(ultima_linha - 1, 12)).Merge()
        historico.Cells(ultima_linha - 1, 12).Borders.LineStyle = 1
        historico.Cells(ultima_linha, 12).Borders.LineStyle = 1

    historico.Range(historico.Cells(ultima_linha, celula_status.Column), historico.Cells(ultima_linha - 1, celula_status.Column)).Merge()
    historico.Cells(ultima_linha, celula_status.Column).Borders.LineStyle = 1

    path_analiseimpacto = join(dirname(path_analiseimpacto), "Atualizado")
    path_analiseimpacto = join(path_analiseimpacto, wb.Name)
    wb.Save()
    wb.Close()

    return quantidade, designator
