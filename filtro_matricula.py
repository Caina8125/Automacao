import pandas as pd
from tkinter import messagebox
from datetime import datetime
from filtro_matricula import FiltroMatricula
from tkinter import filedialog


class FiltroMatricula():
    def __init__(self, planilha1, planilha2) -> None:
        self.df_glosas = pd.read_excel(planilha1)
        self.df_glosas['Associado'] = self.df_glosas['Associado'].astype(str)
        self.df_matriculas = pd.read_excel(planilha2)
        self.df_matriculas['Matrícula'] = self.df_matriculas['Matrícula'].astype(str)

    def filtrar(self):
        df_plan_nova = pd.DataFrame()

        for index, linha in self.df_matriculas.iterrows():
            matricula = str(linha['Matrícula']).replace('.0', '')
            df_glosas_filtrado = self.df_glosas.loc[(self.df_glosas['Associado'] == matricula)]

            if df_glosas_filtrado.empty == True:
                continue
            
            df_plan_nova = pd.concat([df_plan_nova, df_glosas_filtrado])

        if df_plan_nova.empty == True:
            messagebox.showinfo('Filtro Matrícula', 'Não há dados para serem salvos!')
            return

        data_atual = datetime.now()
        data_e_hora_em_texto = data_atual.strftime('%d_%m_%Y_%H_%M_%S')
        df_plan_nova.to_excel(f"\\Output\\Relatorio_Filtrado_{data_e_hora_em_texto}.xlsx", sheet_name='Faturas_Glosadas', index=False)
        messagebox.showinfo('Filtro Matrícula', 'Planilha gerada!')

def filtrar_matricula():
    try:
        messagebox.showinfo('Filtro Matrícula', 'Selecione a planilha de glosas.')
        planilha1 = filedialog.askopenfilename()
        messagebox.showinfo('Filtro Matrícula', 'Selecione a planilha com as matrículas.')
        planilha2 = filedialog.askopenfilename()
        filtro = FiltroMatricula(planilha1, planilha2)
        filtro.filtrar()
    except Exception as e:
        messagebox.showerror('Filtro Matrícula', f'Ocorreu uma exceção nao tratada.\n{e.__class__.__name__}:\n{e}')