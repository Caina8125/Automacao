from pandas import DataFrame
from pandas import read_excel

class PlanilhaSerpro():
    
    def __init__(self, path_planilha: str) -> None:
        self.path_planilha: str = path_planilha
        self.df_planilha: DataFrame = read_excel(path_planilha)

    def to_be_named(self):
        lista_de_processos: list[str] = [f'{value}'.replace('.0', '') for value in list(set(self.df_planilha['Fatura'].values.tolist()))]