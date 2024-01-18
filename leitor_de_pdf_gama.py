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
        if type(directory) == str:
            self._directory = directory

    def main(self) -> list:
        LISTA_DE_ARQUIVOS: list = PDFReader.set_list(self.directory)
        ARQUIVO_CARTA_REMESSA: str | bool = PDFReader.get_carta_remessa(LISTA_DE_ARQUIVOS)
        if(ARQUIVO_CARTA_REMESSA):
            LISTA_DE_ARQUIVOS.remove(ARQUIVO_CARTA_REMESSA)
        else:
            showinfo('Automação', 'Carta remessa não encontrada!')
            return False
        DF_TABELA_REMESSA: pd.DataFrame = PDFReader.get_df_remessa(ARQUIVO_CARTA_REMESSA)
        MATRIZ_DE_DADOS: list = []

        for arquivo in LISTA_DE_ARQUIVOS:
            arquivo_formatado: str = arquivo.replace(self.directory, '').replace('\\', '')

            result: dict = PDFReader(self.directory).ler_pdfs(arquivo)

            protocolo: str = result['pages_content_array'][0][0] 
            numero_fatura_pagina1: str = result['pages_content_array'][0][37].split('_')[0].replace('0000000000000','')
            valor_peg = result['pages_content_array'][0][15]
            valor_nota = PDFReader.valor_is_equal(valor_peg, result['pages_content_array'][1])

            numero_fatura_pagina2: str = PDFReader.get_numero_nota(result['pages_content_array'][1])

            if numero_fatura_pagina1 == numero_fatura_pagina2:
                numeros_remessa: str | bool = PDFReader.number_in_df(DF_TABELA_REMESSA, numero_fatura_pagina1, protocolo)

                if numeros_remessa == False:
                    lista_de_dados: list = [arquivo_formatado, numero_fatura_pagina1, protocolo, 'Não encontrado', 'Não encontrado', numero_fatura_pagina2, 'Fatura não econtrada na remessa']
                
                else:
                    processo_remessa, protocolo_remessa = numeros_remessa
                    if(processo_remessa and protocolo_remessa):
                        lista_de_dados: list = [arquivo_formatado, numero_fatura_pagina1, protocolo, processo_remessa, protocolo_remessa, numero_fatura_pagina2, 'OK']

                    elif(processo_remessa and protocolo_remessa == False):
                        lista_de_dados: list = [arquivo_formatado, numero_fatura_pagina1, protocolo, processo_remessa, 'Inválido', numero_fatura_pagina2, 'Número de protocolo diferente']

                    elif(processo_remessa == False and protocolo_remessa):
                        lista_de_dados: list = [arquivo_formatado, numero_fatura_pagina1, protocolo, 'Inválido', protocolo_remessa, numero_fatura_pagina2, 'Número de processo diferente']
            else:
                lista_de_dados: list = [arquivo_formatado, numero_fatura_pagina1, protocolo, '-', '-', numero_fatura_pagina2, 'Os números de fatura no arquivo são diferentes']

            MATRIZ_DE_DADOS.append(lista_de_dados)
        
        return MATRIZ_DE_DADOS

    @classmethod
    def set_list(cls, directory: str) -> list:
        lista_de_arquivos: list = [f'{directory}\\{arquivo}'
                                    for arquivo in listdir(directory) 
                                    if arquivo.endswith('.pdf') or arquivo.endswith('.xls') or arquivo.endswith('.xlsx')]
        return lista_de_arquivos
    
    @classmethod
    def get_carta_remessa(cls, lista_de_arquivos: list) -> str:
        for arquivo in lista_de_arquivos:
            
            if '.xls' in arquivo.lower() or '.xlsx' in arquivo.lower():
                return arquivo

    @classmethod                    
    def ler_pdfs(cls, arquivo: str) -> dict:
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
    def get_numero_nota(cls, array_pagina: list) -> str:
        for string in array_pagina:
            if 'Informações' in string:
                return string.replace('Informações', '')

    @classmethod                               
    def get_df_remessa(cls, arquivo_carta_remessa: str) -> pd.DataFrame:
        df_carta_remessa: pd.DataFrame = pd.read_excel(arquivo_carta_remessa, header=23)
        df_carta_remessa = df_carta_remessa.iloc[:-6]
        
        return df_carta_remessa
    
    @classmethod
    def number_in_df(cls, df: pd.DataFrame, string_processo: str, string_protocolo) -> str | bool:
        for index, linha in df.iterrows():
            n_fatura: str = f"{linha['Nº Fatura']}".replace('.0', '')
            n_protocolo: str = f"{linha['Protocolo']}".replace('.0', '')
            if  n_fatura == string_processo and n_protocolo == string_protocolo:
                return (n_fatura, n_protocolo)
            
            elif f"{linha['Nº Fatura']}" == string_processo and f"{linha['Protocolo']}" != string_protocolo:
                return (n_fatura, False)
            
            elif f"{linha['Nº Fatura']}" != string_processo and f"{linha['Protocolo']}" == string_protocolo:
                return (False, n_protocolo)
        return False
    
class TreeView():
    def __init__(self, dados: list, master: Tk=None):
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
    dados: list | bool = pdf_reader.main()
    if dados:
        janela = Tk()
        janela.iconbitmap('Robo.ico')
        janela.title('Leitor de PDF GAMA')
        janela.eval('tk::PlaceWindow . center')
        tree_view = TreeView(dados, janela)
        janela.mainloop()