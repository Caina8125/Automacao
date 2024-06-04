import pandas as pd
from tkinter import messagebox
from datetime import datetime
from tkinter import filedialog

class FiltroFaturas():
    def __init__(self, planilha1, planilha2) -> None:
        self.df_glosas = pd.read_excel(planilha1)
        self.df_glosas['Fatura'] = self.df_glosas['Fatura'].astype(str)
        self.df_faturas = pd.read_excel(planilha2)
        self.df_faturas['Nº Fatura'] = self.df_faturas['Nº Fatura'].astype(str)

    def filtrar(self):
        df_plan_nova = pd.DataFrame()

        for _, linha in self.df_faturas.iterrows():
            fatura = str(linha['Nº Fatura']).replace('.0', '')
            df_glosas_filtrado = self.df_glosas.loc[(self.df_glosas['Fatura'] == fatura)]

            if df_glosas_filtrado.empty == True:
                continue
            
            df_plan_nova = pd.concat([df_plan_nova, df_glosas_filtrado])

        if df_plan_nova.empty == True:
            messagebox.showinfo('Filtro Faturas', 'Não há dados para serem salvos!')
            return

        data_atual = datetime.now()
        data_e_hora_em_texto = data_atual.strftime('%d_%m_%Y_%H_%M_%S')
        df_plan_nova.to_excel(f"Output\\Relatorio_Filtrado_{data_e_hora_em_texto}.xlsx", sheet_name='Faturas_Glosadas', index=False)
        messagebox.showinfo('Filtro Matrícula', 'Planilha gerada!')

def filtrar_faturas():
    try:
        messagebox.showinfo('Filtro Faturas', 'Selecione a planilha de glosas.')
        planilha1 = filedialog.askopenfilename()
        messagebox.showinfo('Filtro Faturas', 'Selecione a planilha com as faturas.')
        planilha2 = filedialog.askopenfilename()
        filtro = FiltroFaturas(planilha1, planilha2)
        filtro.filtrar()
    except Exception as e:
        messagebox.showerror('Filtro Faturas', f'Ocorreu uma exceção nao tratada.\n{e.__class__.__name__}:\n{e}')