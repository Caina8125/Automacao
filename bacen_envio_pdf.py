from datetime import datetime
from time import sleep
from tkinter.messagebox import showerror, showinfo
import pandas as pd
from selenium import webdriver
from seleniumwire import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from tkinter.filedialog import askdirectory, askopenfilenames
import pandas as pd
from seleniumwire import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import os
import zipfile
from bacen_protocolo import BuscarProtocolo
from page_element import PageElement

class BacenMapa(PageElement):
    usuario_input = (By.XPATH, '/html/body/table/tbody/tr/td/div/div/div/div/div/div/div/div/div/div/div[2]/table/tbody/tr[2]/td/form/table/tbody/tr[1]/td[2]/input')
    senha_input = (By.XPATH, '/html/body/table/tbody/tr/td/div/div/div/div/div/div/div/div/div/div/div[2]/table/tbody/tr[2]/td/form/table/tbody/tr[2]/td[2]/input')
    login_button = (By.XPATH, '//*[@id="rp1_edt"]/table/tbody/tr[2]/td/form/table/tbody/tr[3]/td/div/input')
    faturamento = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/table/tbody/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr[2]/td/nobr/a')
    aguardando_fisico = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr/td[1]/div[2]/div[4]/a')
    input_pesquisar = (By.XPATH, '//*[@id="tsk_toolbar"]/div/div/div/div/div/div/div/div[2]/div/table/tbody/tr/td/form/div[1]/input[1]')
    lupa_pesquisa = (By.XPATH, '//*[@id="tsk_toolbar"]/div/div/div/div/div/div/div/div[2]/div/table/tbody/tr/td/form/div[1]/input[2]')
    lupa_ver_fatura = (By.XPATH, '//*[@id="FormMain"]/table/tbody/tr[1]/td/div/div/div/div/div/div/div/div/div[2]/table/tbody/tr[2]/td[1]/a/img')
    tabela_guia_com_anexo = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr[1]/td/div/div/div/div/div/div/div/div/div[2]/table')
    tbody_guia_com_anexo = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr[1]/td/div/div/div/div/div/div/div/div/div[2]/table/tbody')
    tbody_guia_sem_anexo = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[4]/td/table/tbody/tr[1]/td/div/div/div/div/div/div/div/div/div[2]/table/tbody')
    botao_novo = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td/div/div/div/div/div/div/div/div/div[1]/div[2]/a')
    procurar_arquivo = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[1]/td/div/div[3]/div/div/div/div/div/div/form/table/tbody/tr[2]/td[2]/div/div[2]/a/img')
    input_file = (By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/form/div/div/div/div/div/div/div/div/div[2]/table/tbody/tr/td/table/tbody/tr[2]/td[2]/input')
    botao_enviar = (By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/form/div/div/div/div/div/div/div/div/div[2]/table/tbody/tr/td/table/tbody/tr[3]/td/input')
    botao_salvar = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[1]/td/div/div[2]/div/div/div/div[3]/a')
    detalhes = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td/div/div/div/div/div/div/div/div/div[2]/table/tbody/tr[3]/td/div/div[2]/a')
    lupa_conta_fisica = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr/td/form/table/tbody/tr[1]/td/div/div/div/div/div/div/div/div/div[2]/table/tbody/tr[2]/td[1]/a[1]')
    processar_anexo = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr/td/div[1]/div/div/div/div/div/div/div/div[3]/div/table/tbody/tr/td/div/nobr/a')
    a_numero_protocolo = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr/td/div[1]/div/div/div/div/div/div/div/div[1]/a[2]')
    tbody_conta_fisica_anexada = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td/div/div/div/div/div/div/div/div/div[2]/table/tbody')

    def __init__(self, driver, url, usuario, senha, dir_planilha, dir_processos, buscar_protocolo):
        super().__init__(driver, url)
        self.usuario = usuario
        self.senha = senha
        self.planilhas = [planilha for planilha in dir_planilha if planilha.endswith('.xlsx')]
        self.lista_de_pastas = [{'path': f"{dir_processos}/{pasta}", 'numero_processo': pasta}
                                for pasta in os.listdir(dir_processos) if pasta.isdigit()]
        self.buscar_protocolo: BuscarProtocolo = buscar_protocolo

    def login(self):
        self.driver.implicitly_wait(30)
        self.driver.find_element(*self.usuario_input).send_keys(self.usuario)
        sleep(2)
        self.driver.find_element(*self.senha_input).send_keys(self.senha)
        sleep(2)
        self.driver.find_element(*self.login_button).click()
        sleep(2)

    def exe_caminho(self):
        self.driver.find_element(*self.faturamento).click()
        sleep(2)
        self.driver.find_element(*self.aguardando_fisico).click()
        sleep(2)

    def pesquisar_protocolo(self, protocolo):
        self.driver.find_element(*self.input_pesquisar).send_keys(protocolo)
        sleep(2)
        self.driver.find_element(*self.lupa_pesquisa).click()
        sleep(2)

    def zipar_arquivos(self, pasta, nome_arquivo_zip) -> str:
        lista_de_arquivos = [f"{pasta}/{arquivo}" for arquivo in os.listdir(pasta) if arquivo.endswith('.pdf')]
        with zipfile.ZipFile(f"{pasta}/{nome_arquivo_zip}", "w", zipfile.ZIP_DEFLATED) as zipf:
            for arquivo in lista_de_arquivos:
                if arquivo.endswith('.pdf') and "PEG" in arquivo and "GUIAPRESTADOR" in arquivo:
                    zipf.write(arquivo, os.path.relpath(arquivo, pasta))
        return f"{pasta}/{nome_arquivo_zip}"

    def anexar_guias(self, arquivo, numero_processo, numero_protocolo):
        tbody_conta_fisica_anexada = self.driver.find_element(*self.tbody_conta_fisica_anexada).text

        if "Nenhum registro cadastrado." not in tbody_conta_fisica_anexada:
            informacoes = [numero_processo, numero_protocolo, "Já há um envio neste processo", '', '']
            return informacoes
        
        self.driver.find_element(*self.botao_novo).click()
        sleep(2)
        self.driver.find_element(*self.procurar_arquivo).click()
        sleep(2)
        self.driver.switch_to.window(self.driver.window_handles[-1])
        sleep(1)
        self.driver.find_element(*self.input_file).send_keys(arquivo)
        sleep(1)
        self.driver.find_element(*self.botao_enviar).click()
        sleep(2)
        self.driver.switch_to.window(self.driver.window_handles[0])
        sleep(1)
        self.driver.find_element(*self.botao_salvar).click()
        sleep(1)
        tbody_guia_sem_anexo = self.driver.find_element(*self.tbody_guia_sem_anexo).text
        count = 0

        while "Nenhum registro cadastrado." not in tbody_guia_sem_anexo and count!= 3:
            self.driver.find_element(*self.detalhes).click()
            sleep(2)
            self.driver.find_element(*self.lupa_conta_fisica).click()
            sleep(1.5)
            self.driver.find_element(*self.processar_anexo).click()
            sleep(1.5)
            self.driver.find_element(*self.a_numero_protocolo).click()
            sleep(2)
            tbody_guia_sem_anexo = self.driver.find_element(*self.tbody_guia_sem_anexo).text
            count +=1

        tbody_guia_com_anexo = self.driver.find_element(*self.tbody_guia_com_anexo).text

        if "Nenhum registro cadastrado." not in tbody_guia_com_anexo:
            tabela = self.driver.find_element(*self.tabela_guia_com_anexo)
            tabela_html = tabela.get_attribute('outerHTML')
            df_tabela = pd.read_html(tabela_html, header=0)[0]
            guia_prestador = df_tabela['Guia Prestador'][0]

            if guia_prestador == 'NaN':
                erro_matricula = "Sim"
            
            else:
                erro_matricula = "Não"
        
        tbody_guia_sem_anexo = self.driver.find_element(*self.tbody_guia_sem_anexo).text

        if "Nenhum registro cadastrado." in tbody_guia_sem_anexo:
            erro_guia_sem_anexo = "Não"
        
        else:
            erro_guia_sem_anexo = "Sim"
        
        informacoes = [numero_processo, numero_protocolo, "Enviado", erro_matricula, erro_guia_sem_anexo]

        return informacoes
        
    def encontrar_protocolo(self, numero_processo):
        if not self.planilhas:
            return None

        for planilha in self.planilhas:
            df_planilha = pd.read_excel(planilha)
            try:
                for _, linha in df_planilha.iterrows():
                    if str(numero_processo) in str(linha['N° Fatura']).replace(".0", ''):
                        protocolo = str(linha['N° Protocolo']).replace(".0", '')
                        if protocolo.isdigit():
                            return protocolo
            except:
                continue
                    
    def renomear_arquivo(self, pasta, arquivo, protocolo, amhptiss):
        os.rename(arquivo, f"{pasta}\\PEG{protocolo}_GUIAPRESTADOR{amhptiss}.pdf")
                    
    def confere_nome_arquivos(self, path_pasta, protocolo):
        lista_de_arquivos = [f"{path_pasta}/{arquivo}" for arquivo in os.listdir(path_pasta) if arquivo.endswith('.pdf')]
        for arquivo in lista_de_arquivos:
            if "PEG" not in arquivo or "GUIAPRESTADOR" not in arquivo:
                n_amhptiss = arquivo.replace(f'{path_pasta}/', '').replace('_Guia.pdf', '').replace(".pdf", '')
                if not n_amhptiss.isdigit():
                    raise Exception('Nome de arquivo invalido.')
                self.renomear_arquivo(path_pasta, arquivo, protocolo, n_amhptiss)
    
    def enviar_pdf_bacen(self):
        self.open()
        self.login()
        lista_de_dados = []
        self.exe_caminho()
        self.driver.execute_script("window.open('');") 
        self.driver.switch_to.window(self.driver.window_handles[-1])
        self.buscar_protocolo.open()
        self.buscar_protocolo.login_layout_novo()
        self.buscar_protocolo.caminho()
        self.driver.switch_to.window(self.driver.window_handles[0])
        
        for pasta in self.lista_de_pastas:
            numero_processo = pasta['numero_processo']
            path_pasta_processo = pasta['path']
            protocolo = self.encontrar_protocolo(numero_processo)

            if protocolo == None:
                self.driver.switch_to.window(self.driver.window_handles[-1])
                protocolo = self.buscar_protocolo.buscar_protocolo(numero_processo)

                if protocolo == 'Nenhum registro encontrado.':
                    informacoes = [numero_processo, protocolo, "Não Enviado, protocolo não encontrado", '', '']
                    lista_de_dados.append(informacoes)
                    continue

                self.driver.switch_to.window(self.driver.window_handles[0])
                
            self.confere_nome_arquivos(path_pasta_processo, protocolo)

            arquivo_zip = self.zipar_arquivos(path_pasta_processo, f'{numero_processo}.zip')
            sz = (os.path.getsize(arquivo_zip) / 1024) / 1024

            if sz >= 25.00:
                informacoes = [numero_processo, protocolo, "Não Enviado, arquivo .zip maior que 25MB", '', '']
                lista_de_dados.append(informacoes)
                continue

            self.pesquisar_protocolo(protocolo)

            if 'Nenhum registro foi encontrado.' in self.driver.find_element(*self.body).text:
                informacoes = [numero_processo, protocolo, "Não Enviado, protocolo não encontrado em Aguardando Físico", '', '']
                lista_de_dados.append(informacoes)
                self.exe_caminho()
                continue

            self.driver.find_element(*self.lupa_ver_fatura).click()
            sleep(2)
            informacoes = self.anexar_guias(arquivo_zip, numero_processo, protocolo)
            lista_de_dados.append(informacoes)
            self.exe_caminho()

        cabecalho = ["N° Fatura", "N° Protocolo", "Enviado", 'Erro Matricula', 'Guia sem anexo']
        df = pd.DataFrame(lista_de_dados, columns=cabecalho)
        data_e_hora_atuais = datetime.now()
        data_e_hora_em_texto = data_e_hora_atuais.strftime('%d_%m_%Y_%H_%M')
        segundo = data_e_hora_atuais.second
        df.to_excel(f'Bacen\\Envio_pdf_{data_e_hora_em_texto}_{segundo}.xlsx', index=False)
#----------------------------------------------------------------------------------------------------------------------------------------------


def enviar_pdf_bacen(user, password):
    try:
        showinfo('', 'Selecione uma pasta com as subpastas dos processos.')
        diretorio = askdirectory()
        showinfo('', 'Selecione planilhas com os protocolos. Caso não tenha, fechar a tela de seleção de arquivos.')
        planilhas = askopenfilenames()

        url = 'https://www3.bcb.gov.br/pasbcmapa/login.aspx'

        options = {
                    'proxy' : {
                        'http': f'http://{user}:{password}@10.0.0.230:3128',
                        'https': f'http://{user}:{password}@10.0.0.230:3128'
                    }
                }

        chrome_options = Options()
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument('--ignore-certificate-errors')
        chrome_options.add_argument('--ignore-ssl-errors')

        try:
            servico = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=servico, options=chrome_options, seleniumwire_options=options)
        except:
            driver = webdriver.Chrome(options=chrome_options, seleniumwire_options=options)

        url_bc_saude = 'https://www3.bcb.gov.br/portalbcsaude/Login'
        buscar_protocolo = BuscarProtocolo('00735860000173', 'Amhp2024!!', driver, url_bc_saude)

        envio_bacen = BacenMapa(driver, url, '00735860000173', 'Amhp2024!!', planilhas, diretorio, buscar_protocolo)
        envio_bacen.enviar_pdf_bacen()
        showinfo('', 'Envio realizado com sucesso!')

    except Exception as e:
        showerror('', f'Ocorreu uma exceção não tratada\n{e.__class__.__name__}:\n{e}')