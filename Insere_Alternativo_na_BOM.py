import glob
import os
import win32com.client as win32
import sys
import time
import pandas as pd
from os.path import dirname
from os.path import join
from win32api import RGB
from tkinter import messagebox
from pathlib import Path
from Escolhe_Versao_da_BOM import escolher_versao
from Gestao_de_versao import criar_nova_versao
from Gestao_de_versao import manter_versao_atual
from Atualiza_Historico_da_BOM import atualizar_historico_V1
from Atualiza_Historico_da_BOM import atualizar_historico_teste
from SQL import Part_Numbers, descricao_placa


def parametros(path_analiseimpacto, arquivos, codigo_principal, codigo_alternativo, descricao, engenheiro, data, check_versao, check_arquivos, janela, responsavel, eco):
    janela.destroy()

    start = time.time()

    if check_versao == "Não":
        path_analiseimpacto = join(dirname(path_analiseimpacto), "Atualizado")

    if check_arquivos == 1:
        xlsx_files = glob.glob(os.path.join(path_analiseimpacto, "*.xlsx"))

    else:
        xlsx_files = []
        for arquivo in glob.glob(os.path.join(path_analiseimpacto, '*.xlsx')):
            if os.path.splitext(os.path.basename(arquivo))[0] in arquivos:
                arquivo = os.path.normpath(arquivo)
                xlsx_files.append(os.path.normpath(arquivo))

    BOMs = ""
    BOMexc = ""
    dados = []
    fabricantes_e_PNs, descricao_comp = Part_Numbers(codigo_alternativo)

    inserir_alternativo(path_analiseimpacto, BOMs, BOMexc, codigo_principal, codigo_alternativo,
                        descricao, engenheiro, data, check_versao, responsavel, eco, xlsx_files, start, dados, fabricantes_e_PNs, descricao_comp)


def inserir_alternativo(path_analiseimpacto, BOMs, BOMexc, codigo_principal, codigo_alternativo, descricao, engenheiro, data, check_versao, responsavel, eco, xlsx_files, start, dados, fabricantes_e_PNs, descricao_comp):

    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.DisplayAlerts = False

    for xlsx in xlsx_files:
        if not os.path.basename(xlsx).startswith("6."):
            continue

        BOM = os.path.basename(xlsx)

        print(f"Atualizando {BOM[:12]}...")

        wb = excel.Workbooks.Open(xlsx)
        versao_atual, versao_intermediaria, nova_versao, n_nova_versao = escolher_versao(
            wb, BOM, excel)

        if check_versao == "Sim":

            worksheet_clonada, celula_codigo, ultima_coluna, celula_codigo_principal = criar_nova_versao(
                wb, versao_atual, versao_intermediaria, nova_versao, n_nova_versao, BOM, codigo_principal, responsavel)

        elif check_versao == "Não":
            worksheet_clonada, celula_codigo, ultima_coluna, celula_codigo_principal = manter_versao_atual(
                wb, versao_atual, BOM, codigo_principal)

        if celula_codigo_principal is None:
            BOMexc += BOM[:12] + ";\n"
            continue

        BOMs += BOM[:12] + ";\n"
        # Verifica se a célula onde a palavra foi encontrada está mesclada e exibe uma mensagem informando o range da mesclagem, se for o caso

        merge_range = celula_codigo_principal.MergeArea
        celula_alternativos = worksheet_clonada.Range("A" + str(celula_codigo.Row) + ":Z" + str(celula_codigo.Row)).Find(
            "Alternativo", LookIn=win32.constants.xlValues, LookAt=win32.constants.xlPart, SearchDirection=win32.constants.xlNext, SearchOrder=win32.constants.xlByColumns)
        if celula_alternativos is not None:
            celula_coluna_i = merge_range.Cells(1, celula_alternativos.Column)
            merge_range.Columns(
                "A:" + ultima_coluna).Interior.Color = RGB(255, 255, 204)
            if celula_coluna_i.Value is not None:
                celula_coluna_i.Value = celula_coluna_i.Value + \
                    chr(10) + fabricantes_e_PNs
            else:
                celula_coluna_i.Value = ""
                celula_coluna_i.Value = fabricantes_e_PNs

        else:
            celula_alternativos = worksheet_clonada.Range("A" + str(celula_codigo.Row) + ":Z" + str(celula_codigo.Row)).Find("Informação adicional", LookIn=win32.constants.xlValues, LookAt=win32.constants.xlPart,
                                                                                                                             SearchDirection=win32.constants.xlNext, SearchOrder=win32.constants.xlByColumns)
            celula_coluna_i = merge_range.Cells(1, celula_alternativos.Column)
            merge_range.Columns(
                "A:" + ultima_coluna).Interior.Color = RGB(255, 255, 204)

            if celula_coluna_i.Value is not None:
                if "Alternativo" in celula_coluna_i.Value:
                    celula_coluna_i.Value = celula_coluna_i.Value + \
                        chr(10) + fabricantes_e_PNs
                    start_pos = celula_coluna_i.Value.find("Alternativos:")
                    celula_coluna_i.GetCharacters(
                        start_pos + 14, -1).Font.Bold = False
                else:
                    celula_coluna_i.Value = celula_coluna_i.Value + \
                        chr(10) + chr(10) + "Alternativos:" + \
                        chr(10) + fabricantes_e_PNs
                    start_pos = celula_coluna_i.Value.find("Alternativos:")
                    celula_coluna_i.GetCharacters(
                        start_pos + 14, -1).Font.Bold = False

            else:
                celula_coluna_i.Value = "Alternativos:" + \
                    chr(10) + fabricantes_e_PNs
                celula_coluna_i.Font.Bold = False
                celula_coluna_i.GetCharacters(1, 13).Font.Bold = True

        if check_versao == "Sim":
            quantidade, designator = atualizar_historico_V1(
                wb, path_analiseimpacto, n_nova_versao, merge_range, eco, codigo_alternativo, descricao, engenheiro, data, descricao_comp, BOM)
        elif check_versao == "Não":
            quantidade, designator = atualizar_historico_teste(
                wb, path_analiseimpacto, merge_range, codigo_alternativo, descricao, engenheiro, descricao_comp, BOM)

        print(f"{BOM} atualizada.\n")

        desc_placa = descricao_placa(BOM[:12])

        dados.append([BOM[:12], desc_placa, versao_atual, nova_versao, codigo_alternativo,
                     descricao_comp, quantidade, designator, engenheiro])

    excel.Quit()

    path_analiseimpacto = join(dirname(path_analiseimpacto), "Atualizado")
    # path_analiseimpacto = join(path_analiseimpacto, wb.Name)

    df = pd.DataFrame(dados, columns=['BOM', 'Descrição da placa', 'Versão em EOL', 'Versão Liberada', 'Cóodigo Alternativo',
                      'Descrição do código alternativo', 'Quantidade', 'Designator', 'Engenheiro Responsável'])
    df.to_excel(os.path.join(path_analiseimpacto, "versoes.xlsx"), index=False)

    end = time.time()

    if BOMs == "":
        if BOMexc == "":
            messagebox.showwarning(
                message=f"Nenhuma BOM foi encontrada no local indicado.", title="BOMs")
        else:
            messagebox.showwarning(
                message=f" O código {codigo_principal} não foi encontrado em nenhuma BOM", title="BOMs")
    else:
        if BOMexc == "":
            messagebox.showinfo(
                message=f"TEMPO DE EXECUÇÃO: {int((end - start)//60)} minuto(s) e {int((end - start)%60)} segundos.\n\nAs seguintes BOMs foram atualizadas: \n\n{BOMs}", title="Finalizado")
        else:
            messagebox.showinfo(
                message=f"TEMPO DE EXECUÇÃO: {int((end - start)//60)} minuto(s) e {int((end - start)%60)} segundos.\n\nAs seguintes BOMs foram atualizadas: \n\n{BOMs} \n\nO código {codigo_principal} não foi encontrado na(s) BOM(s): \n\n{BOMexc}.", title="Finalizado")
    sys.exit(0)
