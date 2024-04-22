from datetime import date
import os
from tkinter import *
from tkinter import Tk, ttk
from tkinter.filedialog import askopenfilename
from tkinter.messagebox import showinfo
import PyPDF2
from pandas import read_excel

class ConferenciaProtocolos():
    lista_convenios_guias_digitalizadas = ['(433)', '(457)', '(381)', '(319)', '(225)', '(56)', '(23)']
    ano_atual = date.today().year

    def remessa_em_arquivo_unico(self, nome_arquivo, diretorio, cod_conv):
        diretorio_convenio = self.pegar_diretorio_convenio(cod_conv, diretorio)
        if diretorio_convenio == None:
            return None
        
        for root, _, files in os.walk(diretorio_convenio):
            arquivo = self.name_is_in_list(nome_arquivo, files)
            if arquivo and cod_conv in root:
                full_path = os.path.join(root, arquivo)
                with open(full_path, 'rb') as arquivo_pdf:
                    leitor_pdf: PyPDF2.PdfReader = PyPDF2.PdfReader(arquivo_pdf)

                    if not self.is_aceite(leitor_pdf.pages):
                        texto_completo = ''
                        for pagina_numero in range(len(leitor_pdf.pages)):
                            pagina = leitor_pdf.pages[pagina_numero]
                            if 'Encaminhamos guias de atendimentos, referentes aos serviços prestados' not in pagina.extract_text():
                                texto = pagina.extract_text().replace(' ', '')
                                texto_completo = texto_completo + ' ' + texto

                        return {'path': full_path, 'text': texto_completo}
        
        return None
    
    def remessa_is_diretorio(self, diretorio, remessa, cod_convenio):
        for root, _, files in os.walk(diretorio):
            if remessa in root.split('/')[-1] and cod_convenio in root:
                return [os.path.join(root, file) for file in files if file.endswith('.pdf')]
            
        return None
    
    def name_is_in_list(self, nome, str_list: list[str]):
        for name in str_list:
            if nome in name and name.endswith('.pdf'):
                return name
            
        return None
        
    def is_aceite(self, arquivo):
        for pagina_numero in range(len(arquivo)):
            pagina = arquivo[pagina_numero]
            texto = ' '.join(pagina.extract_text().split('\n'))
            if 'Relatório de procedimentos realizados (Carta)' not in texto:
                return False
        return True
    
    def n_remessa(self, path: str):
        return path.split('/')[-1].replace('.xlsx', '').replace('.xls', '')
    
    def cod_convenio(self, path):
        codigo_convenio = read_excel(path, engine='xlrd')['Unnamed: 2'][9].split(' - ')[-1].replace('(', '').replace(')', '')
        for k, _ in enumerate(codigo_convenio):
            if codigo_convenio[k] != "0":
                codigo_convenio = str(codigo_convenio[k:])
                break

        return f'({codigo_convenio})'
    
    def conferir_pasta(self, path_planilha):
        df_plan = read_excel(path_planilha, engine='xlrd')
        df_plan = read_excel(path_planilha, engine='xlrd', header=23).iloc[:-6].dropna(axis=1)
        directorio_finan_fat = r'\\10.0.0.239\financeiro - faturamento\Protocolos de Convênios - Aceite'
        diretorio_guiasscaneadas = r'\\10.0.0.239\guiasscaneadas'
        n_remessa = self.n_remessa(path_planilha)
        codigo_convenio = self.cod_convenio(path_planilha)
        lista_de_processos = df_plan['Nº Fatura'].astype(str).values.tolist()
        dados = []

        if codigo_convenio in self.lista_convenios_guias_digitalizadas:
            for processo in lista_de_processos:
                info_processo = self.procurar_arquivo_guias_digitalizadas(processo.replace('.0', ''), diretorio_guiasscaneadas, codigo_convenio)
                dados.append(info_processo)
            return dados

        info_arquivo_remessa = self.remessa_em_arquivo_unico(n_remessa, directorio_finan_fat, codigo_convenio)

        if info_arquivo_remessa:
            for processo in lista_de_processos:
                if processo.replace('.0', '') in info_arquivo_remessa['text']:
                    info_processo = [processo.replace('.0', ''), info_arquivo_remessa['path'], 'Encontrado']
                else:
                    info_processo = [processo.replace('.0', ''), '', 'Não encontrado']
                dados.append(info_processo)

            return dados

        info_pasta_remessa = self.remessa_is_diretorio(directorio_finan_fat, n_remessa, codigo_convenio)
        if info_pasta_remessa:
            for processo in lista_de_processos:
                info_processo = self.pegar_caminho_processo(processo.replace('.0', ''), info_pasta_remessa)
                dados.append(info_processo)
            return dados
                    
        else:
            return None

    def pegar_caminho_processo(self, nome, lista_de_arquivos):
        for arquivo in lista_de_arquivos:
            if nome in arquivo:
                return [nome, arquivo, 'Encontrado']
            
        return [nome, '', 'Não encontrado']
    
    def procurar_arquivo_guias_digitalizadas(self, processo, diretorio, cod_convenio):
        convenio = self.pegar_nome_convenio(cod_convenio)
        lista_anos = [self.ano_atual, self.ano_atual - 1]
        for ano in lista_anos:
            for root, dirs, files in os.walk(f"{diretorio}\\{ano}\\{convenio}"):
            # Verifica se o arquivo está na lista de arquivos
                if f'{processo}.pdf' in files:
                    path = os.path.join(root, f'{processo}.pdf')
                    return [processo, path, 'Encontrado']
        return [processo, '', 'Não encontrado']
    
    def pegar_nome_convenio(self, cod_convenio):
        match cod_convenio:
            case '(433)':
                return 'GDF'
            case '(457)':
                return 'PMDF'
            case '(381)':
                return 'PMDF'
            case '(319)':
                return 'TJDFT'
            case '(225)':
                return 'GEAP'
            case '(56)':
                return 'AFFEGO'
            case '(23)':
                return 'ASETE'
            
    def pegar_diretorio_convenio(self, cod_convenio, diretorio):
        lista_de_pastas = os.listdir(diretorio)
        for pasta in lista_de_pastas:
            if cod_convenio in pasta:
                return os.path.join(diretorio, pasta)
            
class TreeView():
    def __init__(self, dados: list, master: Tk = None):
        self.style: ttk.Style = ttk.Style()
        self.style.configure("My.Treeview", rowheight=50)
        self.style.layout("My.Treeview", [('Treeitem.padding', {'sticky': 'nswe', 'children': [('Treeitem.indicator', {'side': 'left', 'sticky': ''}), ('Treeitem.image', {'side': 'left', 'sticky': ''}), ('Treeitem.text', {'side': 'left', 'sticky': ''})]})])
        self.tree: ttk.Treeview = ttk.Treeview(master, style="My.Treeview", selectmode='browse', column=('column1', 'column2', 'column3'), show='headings')
        self.tree.pack()

        self.tree.column('column1', width=50, minwidth=0, stretch=YES)
        self.tree.heading('#1', text='Fatura')

        self.tree.column('column2', width=0, minwidth=250, stretch=YES)
        self.tree.heading('#2', text='Caminho')

        self.tree.column('column3', width=0, minwidth=100, stretch=YES)
        self.tree.heading('#3', text='Situação')

        # self.tree.grid(row=0, column=0)
        self.tree.pack(expand=True, fill=BOTH)
        for dado in dados:
            self.tree.insert('', END, values=dado)   

def executar_conferencia_arquivos():
    conf_protocolos = ConferenciaProtocolos()
    planilha = 'inicio'
    while True:

        planilha = askopenfilename()
        if not planilha:
            break

        infos = conf_protocolos.conferir_pasta(planilha)

        if infos:
            window = Tk()
            window.iconbitmap('Robo.ico')
            window.title('Automação')
            # window.state('zoomed')
            window.geometry("600x400")
            tree_view = TreeView(infos, window)
            window.mainloop()

        else:
            showinfo('', "Nenhum registro encontrado na pasta do convênio!")

if __name__ == '__main__':
    executar_conferencia_arquivos()