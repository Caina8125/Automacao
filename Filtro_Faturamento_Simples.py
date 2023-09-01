import pandas as pd
from tkinter import filedialog
import os

def processar_planilha():
    global planilha
    planilha = filedialog.askopenfilename()
    arquivo = pd.ExcelFile(planilha)
    print(arquivo)
    sheet_names = arquivo.sheet_names
    print(sheet_names)
    quantidade = len(sheet_names)
    count = 0
    df = pd.DataFrame()

    for i in range(count, quantidade):
        sheet = pd.read_excel(planilha, header=18, sheet_name=count , index_col=False)
        sheet = sheet.fillna(0)
        sheet = sheet[['Paciente', 'Realizado em  Horário', 'Controle', 'AMHPTISS', 'Nº Guia', 'Valor Cobrado']]
        sheet['Situação'] = ""
        sheet['Pesquisado no Portal'] = ""
        sheet = sheet.iloc[:-3]
        print(sheet)
        count = count + 1
        df_2 = pd.DataFrame(sheet)
        df = df.append(df_2, ignore_index=True)
        print(df)

    xlsx = planilha.replace('.xls', '.xlsx')
    df.to_excel(xlsx, index=False)
    
def remove():
    os.remove(planilha)