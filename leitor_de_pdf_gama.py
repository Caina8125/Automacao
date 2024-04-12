import PyPDF2
import pandas as pd
from tkinter import *
from os import listdir
from tkinter import ttk
from tkinter.filedialog import askdirectory
from tkinter.messagebox import showinfo, showerror

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
            protocolo: str = result['peg'][0] 
            numero_fatura_pagina1: str = result['peg'][37].split('_')[0].replace('0000000000000','')
            valor_peg: str = result['peg'][15]
            valor_nota = PDFReader.extrair_valor_monetario(result['nota_fiscal'])

            numero_fatura_pagina2: str = PDFReader.get_numero_nota(result['nota_fiscal'])

            if numero_fatura_pagina1 == numero_fatura_pagina2:
                numeros_remessa: str | bool = PDFReader.number_in_df(DF_TABELA_REMESSA, numero_fatura_pagina1, protocolo)

                if numeros_remessa == False:
                    lista_de_dados: list = [
                        arquivo_formatado, #Nome do arquivo
                        numero_fatura_pagina1, #Processo no PEG
                        protocolo, #Protocolo no PEG
                        valor_peg, #Valor no PEG
                        'Não encontrado', #Processo na remessa
                        'Não encontrado', #Protocolo na remessa 
                        'Não encontrado', #Valor na remessa
                        numero_fatura_pagina2, #Processo na nota fiscal
                        valor_nota, #Valor na nota fiscal
                        '-', # Validação Valor
                        'Fatura não econtrada na remessa' #Validação
                        ]
                
                else:
                    processo_remessa, protocolo_remessa, valor_remessa = numeros_remessa
                    is_same_price: bool = valor_nota == valor_peg

                    if is_same_price:

                        if valor_peg == valor_remessa:
                            validacao: str = 'Ok'
                        
                        else:
                            validacao: str = 'Valor divergente: (Arquivo - Remessa)'

                    else:
                        validacao: str = 'Valor divergente: (PEG - Nota Fiscal)'

                    if processo_remessa and protocolo_remessa:
                        lista_de_dados: list = [
                            arquivo_formatado, #Nome do arquivo
                            numero_fatura_pagina1, #Processo no PEG
                            protocolo, #Protocolo no PEG
                            valor_peg, #Valor no PEG
                            processo_remessa, #Processo na remessa
                            protocolo_remessa, #Protocolo na remessa 
                            valor_remessa, #Valor na remessa
                            numero_fatura_pagina2, #Processo na nota fiscal
                            valor_nota, #Valor na nota fiscal
                            validacao, #Validação Valor
                            'Ok'
                            ]

                    elif processo_remessa and protocolo_remessa == False:
                        lista_de_dados: list = [
                            arquivo_formatado, #Nome do arquivo
                            numero_fatura_pagina1, #Processo no PEG
                            protocolo, #Protocolo no PEG
                            valor_peg, #Valor no PEG
                            processo_remessa, #Processo na remessa
                            'Inválido', #Protocolo na remessa 
                            valor_remessa, #Valor na remessa
                            numero_fatura_pagina2, #Processo na nota fiscal
                            valor_nota, #Valor na nota fiscal
                            validacao, #Validação Valor
                            'Protocolo diferente' #Validação
                            ]

                    elif processo_remessa == False and protocolo_remessa:
                        lista_de_dados: list = [
                            arquivo_formatado, #Nome do arquivo
                            numero_fatura_pagina1, #Processo no PEG
                            protocolo, #Protocolo no PEG
                            valor_peg, #Valor no PEG
                            'Inválido', #Processo na remessa
                            protocolo_remessa, #Protocolo na remessa 
                            valor_remessa, #Valor na remessa
                            numero_fatura_pagina2, #Processo na nota fiscal
                            valor_nota, #Valor na nota fiscal
                            validacao, #Validação Valor
                            'Número de processo diferente' #Validação
                            ]
                        
            else:
                lista_de_dados: list = [
                    arquivo_formatado, #Nome do arquivo
                    numero_fatura_pagina1, #Processo no PEG
                    protocolo, #Protocolo no PEG
                    valor_peg, #Valor no PEG
                    '-', #Processo na remessa
                    '-', #Protocolo na remessa
                    '-', #Valor na remessa
                    numero_fatura_pagina2, #Processo na nota fiscal
                    valor_nota, #Valor na nota fiscal
                    '-', #Validação valor
                    'N° Fatura divergentes' #Validação
                    ]

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
                   
    def ler_pdfs(self, arquivo: str) -> dict:
        with open(arquivo, 'rb') as arquivo_pdf:
            leitor_pdf: PyPDF2.PdfReader = PyPDF2.PdfReader(arquivo_pdf)
            word_dict: dict = {}

            for pagina_numero in range(len(leitor_pdf.pages)):
                pagina = leitor_pdf.pages[pagina_numero]
                texto_sem_quebra: str = ' '.join(pagina.extract_text().split('\n'))
                if 'PROTOCOLO ENVIO LOTE DE CONTAS MÉDICAS' in texto_sem_quebra:
                    word_dict['peg'] = texto_sem_quebra.split(' ')
                elif 'Secretaria de Estado de Fazenda do Distrito Federal' in texto_sem_quebra:
                    word_dict['nota_fiscal'] = texto_sem_quebra.split(' ')

                else:
                    showinfo('Leitor de PDF Gama', f'Confira se há PEG e Nota Fiscal no arquivo {arquivo.replace(self._directory, "")}')
                    return None

            arquivo_pdf.close()

        return word_dict
    
    @classmethod
    def extrair_valor_monetario(cls, string_array: list) -> str | None:
        for index, string in enumerate(string_array):
            if (string == 'SERVICOS' or string == 'contas:') and string_array[index + 2].replace('.', '').replace(',', '').isdigit():
                return string_array[index + 2]
        
        return None
    
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
            valor_remessa: str = f"{linha['Valor']}"
            if  n_fatura == string_processo and n_protocolo == string_protocolo:
                return (n_fatura, n_protocolo, valor_remessa)
            
            elif f"{linha['Nº Fatura']}" == string_processo and f"{linha['Protocolo']}" != string_protocolo:
                return (n_fatura, False, valor_remessa)
            
            elif f"{linha['Nº Fatura']}" != string_processo and f"{linha['Protocolo']}" == string_protocolo:
                return (False, n_protocolo, valor_remessa)
            
        return False
    
class TreeView():
    def __init__(self, dados: list, master: Tk = None):
        self.style: ttk.Style = ttk.Style()
        self.style.configure("My.Treeview", rowheight=150)
        self.tree: ttk.Treeview = ttk.Treeview(master, style="My.Treeview", selectmode='browse', column=('column1', 'column2', 'column3', 'column4', 'column5', 'column6', 'column7', 'column8', 'column9', 'column10', 'column11'), show='headings')
        self.tree.pack()

        self.tree.column('column1', width=100, minwidth=50, stretch=YES,)
        self.tree.heading('#1', text='Arquivo')

        self.tree.column('column2', width=100, minwidth=50, stretch=YES)
        self.tree.heading('#2', text='Processo no PEG')

        self.tree.column('column3', width=100, minwidth=50, stretch=YES)
        self.tree.heading('#3', text='Protocolo no PEG')

        self.tree.column('column4', width=100, minwidth=50, stretch=YES)
        self.tree.heading('#4', text='Valor PEG')
        
        self.tree.column('column5', width=100, minwidth=50, stretch=YES)
        self.tree.heading('#5', text='Processo na Remessa')

        self.tree.column('column6', width=100, minwidth=50, stretch=YES)
        self.tree.heading('#6', text='Protocolo na Remessa')

        self.tree.column('column7', width=100, minwidth=50, stretch=YES)
        self.tree.heading('#7', text='Valor na Remessa')

        self.tree.column('column8', width=100, minwidth=50, stretch=YES)
        self.tree.heading('#8', text='N° na Nota Fiscal')

        self.tree.column('column9', width=100, minwidth=50, stretch=YES)
        self.tree.heading('#9', text='Valor Nota Fiscal')

        self.tree.column('column10', width=200, minwidth=50, stretch=YES)
        self.tree.heading('#10', text='Validação valor')

        self.tree.column('column11', width=200, minwidth=50, stretch=YES)
        self.tree.heading('#11', text='Validação')

        # self.tree.grid(row=0, column=0)
        self.tree.pack(expand=True, fill=BOTH)
        for dado in dados:
            self.tree.insert('', END, values=dado)       
    
def pdf_reader() -> None:
    try:
        caminho_do_pdf: str = askdirectory()
        if not caminho_do_pdf:
            showinfo('Leitor de PDF GAMA', 'Nenhuma pasta foi selecionado!')
            return
        pdf_reader: PDFReader = PDFReader(caminho_do_pdf)
        dados: list | bool = pdf_reader.main()
        if dados:
            window = Tk()
            window.iconbitmap('Robo.ico')
            window.title('Leitor de PDF GAMA')
            tree_view = TreeView(dados, window)
            window.mainloop()
    except Exception as err:
        showerror("Leitor de PDF GAMA", f"Ocorreu uma exceção não tratada\n{err}")