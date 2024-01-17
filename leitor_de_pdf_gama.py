import PyPDF2
import pandas as pd
from tkinter import *
from os import listdir
from tkinter import ttk
from tabula.io import read_pdf
from tkinter.messagebox import showinfo
from tkinter.filedialog import askdirectory

class PDFReader():
    def __init__(self, directory: str) -> None:
        self._directory: str = directory

    @property
    def directory(self) -> str:
        return self._directory
    
    @directory.setter
    def directory(self, directory: str) -> None:
        self._directory = directory

    def main(self) -> list[list]:
        LISTA_DE_ARQUIVOS: list = PDFReader.set_list(self.directory)
        ARQUIVO_CARTA_REMESSA: str | bool = PDFReader.get_carta_remessa(LISTA_DE_ARQUIVOS)
        if(ARQUIVO_CARTA_REMESSA):
            LISTA_DE_ARQUIVOS.remove(ARQUIVO_CARTA_REMESSA)
        else:
            showinfo('Automação', 'Carta remessa não encontrada!')
            return
        DF_TABELA_REMESSA: pd.DataFrame = PDFReader.get_df_remessa(ARQUIVO_CARTA_REMESSA)
        MATRIZ_DE_DADOS = []

        for arquivo in LISTA_DE_ARQUIVOS:
            arquivo_formatado = arquivo.replace(self.directory, '').replace('\\', '')

            result: dict = PDFReader(self.directory).ler_pdfs(arquivo)

            protocolo: str = result['pages_content_array'][0][0] 
            numero_fatura_pagina1: str = result['pages_content_array'][0][37].split('_')[0].replace('0000000000000','')

            numero_fatura_pagina2: str = PDFReader.get_numero_nota(result['pages_content_array'][1])

            if numero_fatura_pagina1 == numero_fatura_pagina2:
                numeros_remessa = PDFReader.number_in_df(DF_TABELA_REMESSA, numero_fatura_pagina1, protocolo)

                if numeros_remessa == False:
                    lista_de_dados: list[str] = [arquivo_formatado, numero_fatura_pagina1, protocolo, 'Não encontrado', 'Não encontrado', numero_fatura_pagina2, 'Fatura não econtrada na remessa']
                
                else:
                    processo_remessa, protocolo_remessa = numeros_remessa
                    if(processo_remessa and protocolo_remessa):
                        lista_de_dados: list[str] = [arquivo_formatado, numero_fatura_pagina1, protocolo, processo_remessa, protocolo_remessa, numero_fatura_pagina2, 'OK']

                    elif(processo_remessa and protocolo_remessa == False):
                        lista_de_dados: list[str] = [arquivo_formatado, numero_fatura_pagina1, protocolo, processo_remessa, 'Inválido', numero_fatura_pagina2, 'Número de protocolo diferente']

                    elif(processo_remessa == False and protocolo_remessa):
                        lista_de_dados: list[str] = [arquivo_formatado, numero_fatura_pagina1, protocolo, 'Inválido', protocolo_remessa, numero_fatura_pagina2, 'Número de processo diferente']
            else:
                lista_de_dados: list[str] = [arquivo_formatado, numero_fatura_pagina1, protocolo, '-', '-', numero_fatura_pagina2, 'Os números de fatura no arquivo são diferentes']

            MATRIZ_DE_DADOS.append(lista_de_dados)
        
        return MATRIZ_DE_DADOS

    @classmethod
    def set_list(cls, directory: str) -> list:
        lista_de_arquivos: list = [f'{directory}\\{arquivo}' for arquivo in listdir(directory) if arquivo.endswith('.pdf')]
        return lista_de_arquivos
    
    @classmethod
    def get_carta_remessa(cls, lista_de_arquivos: list) -> str:
        for arquivo in lista_de_arquivos:
            
            result: dict = cls.ler_pdfs(arquivo)
            
            if '(Carta)' in result['pages_content_array'][0]:
                return arquivo
            
        return False

    @classmethod                    
    def ler_pdfs(cls, arquivo: str) -> dict[list, int]:
        with open(arquivo, 'rb') as arquivo_pdf:
            leitor_pdf: PyPDF2.PdfReader = PyPDF2.PdfReader(arquivo_pdf)
            pages_array: list = []

            for pagina_numero in range(len(leitor_pdf.pages)):
                pagina: int = leitor_pdf.pages[pagina_numero]
                texto_sem_quebra: str = ' '.join(pagina.extract_text().split('\n'))
                pages_array.append(texto_sem_quebra.split(' '))

            arquivo_pdf.close()

        return {
            'pages_content_array': pages_array,
            'qtd': len(leitor_pdf.pages)
        }
    
    @classmethod
    def get_numero_nota(cls, array_pagina: list[str]) -> str:
        for string in array_pagina:
            if 'Informações' in string:
                return string.replace('Informações', '')

    @classmethod                               
    def get_df_remessa(cls, arquivo_carta_remessa: str) -> pd.DataFrame:
        teste = read_pdf(arquivo_carta_remessa, pages='1')
        df: pd.DataFrame = pd.DataFrame(teste[len(teste) - 1])
        if 'Nome completo' in df.columns[0]:
            matriz: list = df.values.tolist()

        else:
            matriz: list = [df.columns.to_list()]
            for lista in df.values.tolist():
                matriz.append(lista)

        df_carta_remessa:pd.DataFrame = pd.DataFrame(matriz)
        match len(df_carta_remessa.columns):
            case 5:
                df_carta_remessa.columns = ['Nº Fatura', 'Data de Emissão', 'Quantidade', 'Valor', 'Protocolo']
            case 6:
                df_carta_remessa.columns = ['Nº Fatura', 'Data de Emissão', 'Quantidade', 'Valor', 'Protocolo', 'Unnamed']
            case 7:
                df_carta_remessa.columns = ['Nº Fatura', 'Data de Emissão', 'Quantidade', 'Valor', 'Protocolo', 'Unnamed1', 'Unnamed2']
        
        return df_carta_remessa
    
    @classmethod
    def number_in_df(cls, df: pd.DataFrame, string_processo: str, string_protocolo) -> str | bool:
        for index, linha in df.iterrows():
            if f"{linha['Nº Fatura']}" == string_processo and f"{linha['Protocolo']}" == string_protocolo:
                return (f"{linha['Nº Fatura']}", f"{linha['Protocolo']}")
            
            elif f"{linha['Nº Fatura']}" == string_processo and f"{linha['Protocolo']}" != string_protocolo:
                return (f"{linha['Nº Fatura']}", False)
            
            elif f"{linha['Nº Fatura']}" != string_processo and f"{linha['Protocolo']}" == string_protocolo:
                return (False, f"{linha['Protocolo']}")
        return False
    
class TreeView():
    def __init__(self, dados: list[list], master: Tk=None):
        self.tree: ttk.Treeview = ttk.Treeview(master, selectmode='browse', column=('column1', 'column2', 'column3', 'column4', 'column5', 'column6', 'column7'), show='headings')

        self.tree.column('column1', width=200, minwidth=50, stretch=YES)
        self.tree.heading('#1', text='Arquivo')

        self.tree.column('column2', width=200, minwidth=50, stretch=YES)
        self.tree.heading('#2', text='Processo no PEG')

        self.tree.column('column3', width=200, minwidth=50, stretch=YES)
        self.tree.heading('#3', text='Protocolo no PEG')
        
        self.tree.column('column4', width=200, minwidth=50, stretch=YES)
        self.tree.heading('#4', text='Processo na Remessa')

        self.tree.column('column5', width=200, minwidth=50, stretch=YES)
        self.tree.heading('#5', text='Protocolo na Remessa')

        self.tree.column('column6', width=200, minwidth=50, stretch=YES)
        self.tree.heading('#6', text='N° na Nota Fiscal')

        self.tree.column('column7', width=200, minwidth=50, stretch=YES)
        self.tree.heading('#7', text='Validação')

        self.tree.grid(row=0, column=0)
        self.tree.pack(expand=True, fill=BOTH)
        for dado in dados:
            self.tree.insert('', END, values=dado)
    
def pdf_reader() -> None:
    caminho_do_pdf: str = askdirectory()
    pdf_reader: PDFReader = PDFReader(caminho_do_pdf)
    dados = pdf_reader.main()
    janela = Tk()
    janela.iconbitmap('Robo.ico')
    janela.title('Leitor de PDF GAMA')
    janela.eval('tk::PlaceWindow . center')
    tree_view = TreeView(dados, janela)
    janela.mainloop()