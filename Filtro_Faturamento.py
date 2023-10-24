import pandas as pd
from tkinter import filedialog
import os

def processar_planilha():
    global planilha
    planilha = filedialog.askopenfilename()
    df = pd.read_excel(planilha)
    arquivo = pd.ExcelFile(planilha)
    sheet_names = arquivo.sheet_names
    quantidade = len(sheet_names)
    count = 0
    df = pd.DataFrame()

    for i in range(count, quantidade):
        sheet = pd.read_excel(planilha, sheet_name=count , index_col=False, header=None)

        for i in range(18):
            sheet = sheet.drop(i)
        
        sheet = sheet.fillna(0)
        sheet = sheet.values.tolist()
        nova_lista = []
        INDICES = (0,2,3,4,6,9,10,11,14,15,17,20,21,22,23,25,28,29,31)

        for lista in sheet:
            lista_da_lista = []

            for i in range(len(lista)):
                item_na_lista = lista[i]

                if i in INDICES:
                    continue

                if i == 8 and item_na_lista != 'Matríc. Convênio':
                    item_na_lista = f'Nº - {item_na_lista}'

                lista_da_lista.append(item_na_lista)

            nova_lista.append(lista_da_lista)

        cabecalho = nova_lista.pop(0)
        sheet = pd.DataFrame(nova_lista)
        sheet.columns = cabecalho
        sheet = sheet[['Paciente', 'Controle', 'Matríc. Convênio', 'Nº Guia', 'Senha Aut.', 'Procedimento', 'Valor Cobrado']]
        sheet['Situação'] = ""
        sheet['Validação Carteira'] = ""
        sheet['Validação Proc.'] = ""
        sheet['Validação Senha.'] = ""
        sheet['Pesquisado no Portal'] = ""
        sheet['Matríc. Convênio'] = sheet['Matríc. Convênio'].astype(str)
        sheet = sheet.iloc[:-4]
        count = count + 1
        df_2 = pd.DataFrame(sheet)
        df = df.append(df_2, ignore_index=True)

    xlsx = planilha.replace('.xls', '.xlsx')
    df.to_excel(xlsx, index=False)
    
def remove():
    os.remove(planilha)
