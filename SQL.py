import pandas as pd
import pyodbc
from tkinter import messagebox


def retornar_conexao_sql():
    server = "Kimera"
    database = "Repositorio_SAP"
    username = "leitura_sap"
    password = "Fx!74Xq0gBg@qit"

    string_conexao = 'Driver={SQL Server Native Client 11.0};Server=' + \
        server+';Database='+database+';UID='+username+';PWD='+password
    conexao = pyodbc.connect(string_conexao)
    return conexao.cursor()


def pns_sap(codigo):
    cursor = retornar_conexao_sql()
    df1 = pd.read_sql_query("SELECT U_ItemCode, U_NEO_DESCRI, U_NEO_PARTNUM, U_NEO_STPN, [desc-item] FROM CT_PF_OIDT "
                            "FULL JOIN CT_PF_IDT3 ON CT_PF_OIDT.Code = CT_PF_IDT3.Code "
                            "FULL JOIN OITM ON OITM.[cod-item] = CT_PF_OIDT.U_ItemCode "
                            f"WHERE U_NEO_STPN = 'ATIVO' AND U_ItemCode = '{codigo}' ", cursor.connection)

    cursor.close()

    # agrupar os valores por Fabricante e PN
    grouped = df1.groupby(['U_NEO_DESCRI', 'U_NEO_PARTNUM'])

    string_formatada = ' ('
    for i, (group, data) in enumerate(grouped):
        fabricante, pn = group
        if i != 0:
            string_formatada += ';'
        string_formatada += f' {fabricante} / {pn}'

    string_formatada += ')'

    # criar a string formatada
    # string_formatada = ''
    # for group, data in grouped:
    #     fabricante, pn = group
    #     string_formatada += f'\n({fabricante} / {pn})'

    try:
        # adicionar o código ao início da string
        string_formatada = f'{df1["U_ItemCode"][0]}' + string_formatada
        if df1['desc-item'].nunique() == 1:
            descricao_comp = df1.iloc[0]['desc-item']

        else:
            print(
                "Existem dua ou mais descrições diferentes para um mesmo código na Query")
    except:
        messagebox.showerror(
            message=f"Não foi encontrado no nosso banco de dados o código \"{codigo}\" ou Part Numbers associados a ele", title="ERRO")
        raise ValueError(
            f"Não foi encontrado no nosso banco de dados o código \"{codigo}\" ou Part Numbers associados a ele")

    # imprimir a string formatada
    return string_formatada, descricao_comp
