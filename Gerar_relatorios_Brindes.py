import openpyxl
import pandas as pd
import openpyxl
 
def Gerar_Relat_Normal():
    global convenio

    df        = pd.read_excel(r"\\10.0.0.239\automacao_faturamento\Brindes\Dados.xlsx", sheet_name='Dados')


    normal          = df[df["Tipo Brinde"] == "Normal"] 
    total_normal    = normal['Quantidade'].sum()
    diretoria       = df[df["Tipo Brinde"] == "Diretoria"]
    total_diretoria = diretoria['Quantidade'].sum()
    fora            = df[df["Tipo Brinde"] == "FORA"]
    total_fora      = fora['Quantidade'].sum()


    convenios = df["Convênio"]
    convenios = convenios.drop_duplicates()
    convenios = convenios.to_list()

    for convenio in convenios:
        dados             = df[df["Convênio"] == convenio]
        qtd_dados         = len(dados)
        book              = openpyxl.load_workbook(r"\\10.0.0.239\automacao_faturamento\Brindes\Template.xlsx")
        sheet             = book.active
        sheet['B7'].value = convenio
        sheet['B8'].value = qtd_dados
        sheet['F7'].value = total_normal
        sheet['F8'].value = total_diretoria
        sheet['F9'].value = total_fora
        writer            = pd.ExcelWriter(r"\\10.0.0.239\automacao_faturamento\Brindes\Relatorio_Brindes" + f"\\{convenio}.xlsx", engine='openpyxl')
        writer.book       = book
        writer.sheets     = dict((ws.title, ws) for ws in book.worksheets)
        dados.to_excel(writer, "Relatório_Brindes", startrow=13, startcol=0, header=False, index=False)
        writer.save()
        print()

Gerar_Relat_Normal()
        


