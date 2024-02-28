from time import sleep
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
from page_element import PageElement

class LoginLayoutAntigo(PageElement):
    usuario_input = (By.XPATH, '/html/body/table/tbody/tr/td/div/div/div/div/div/div/div/div/div/div/div[2]/table/tbody/tr[2]/td/form/table/tbody/tr[1]/td[2]/input')
    senha_input = (By.XPATH, '/html/body/table/tbody/tr/td/div/div/div/div/div/div/div/div/div/div/div[2]/table/tbody/tr[2]/td/form/table/tbody/tr[2]/td[2]/input')
    login_button = (By.XPATH, '//*[@id="rp1_edt"]/table/tbody/tr[2]/td/form/table/tbody/tr[3]/td/div/input')

    def login(self, usuario, senha):
        self.driver.implicitly_wait(30)
        self.driver.find_element(*self.usuario_input).send_keys(usuario)
        sleep(2)
        self.driver.find_element(*self.senha_input).send_keys(senha)
        sleep(2)
        self.driver.find_element(*self.login_button).click()
        sleep(2)

class EnvioPDF(PageElement):
    body = (By.XPATH, '/html/body')
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
        self.driver.find_element(*self.lupa_ver_fatura).click()
        sleep(2)

    def renomear_arquivo(self, pasta, arquivo, protocolo, amhptiss):
        os.rename(arquivo, f"{pasta[0]}\\PEG{protocolo}_GUIAPRESTADOR{amhptiss}.pdf")
    
    def anexar_guias(self, arquivo, numero_processo, numero_protocolo):
        tbody_guia_com_anexo = self.driver.find_element(*self.tbody_guia_com_anexo).text

        if "Nenhum registro cadastrado." in tbody_guia_com_anexo:
            sleep(1)
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
            self.driver.find_element(*self.detalhes).click()
            sleep(2)
            self.driver.find_element(*self.lupa_conta_fisica).click()
            sleep(1.5)
            self.driver.find_element(*self.processar_anexo).click()
            sleep(1.5)
            self.driver.find_element(*self.a_numero_protocolo).click()
            sleep(2)
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
    
    def zipar_arquivos(self, pasta, nome_arquivo_zip, lista_de_arquivos):
        with zipfile.ZipFile(f"{pasta[0]}/{nome_arquivo_zip}", "w", zipfile.ZIP_DEFLATED) as zipf:
            for arquivo in lista_de_arquivos:
                if arquivo.endswith('.pdf') and "PEG" in arquivo and "GUIAPRESTADOR" in arquivo:
                    zipf.write(arquivo, os.path.relpath(arquivo, pasta[0]))

    def ler_planilhas(self, lista_de_planilhas, numero_processo):
        for planilha in lista_de_planilhas:
            df_planilha = pd.read_excel(planilha)
            for index, linha in df_planilha.iterrows():
                if numero_processo in f"{linha['N° Fatura']}".replace(".0", ''):
                    protocolo = f"{linha['N° Protocolo']}".replace(".0", '')
                    if protocolo.isdigit():
                        return protocolo

def enviar_bacen():
    diretorio = askdirectory()
    planilhas = askopenfilenames()
    planilhas = [planilha for planilha in planilhas if planilha.endswith('.xlsx')]
    lista_de_pastas = [[f"{diretorio}/{pasta}", pasta] for pasta in os.listdir(diretorio) if pasta.isdigit()]

    url = 'https://www3.bcb.gov.br/pasbcmapa/login.aspx'

    options = {
                'proxy' : {
                    'http': 'http://lucas.paz:RDRsoda90901@@10.0.0.230:3128',
                    'https': 'http://lucas.paz:RDRsoda90901@@10.0.0.230:3128'
                }
            }

    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument('--ignore-certificate-errors')
    chrome_options.add_argument('--ignore-ssl-errors')

    try:
        servico = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=servico, seleniumwire_options= options, options = chrome_options)
    except:
        driver = webdriver.Chrome(seleniumwire_options= options, options = chrome_options)

    envio_bacen = EnvioPDF(driver=driver, url=url)
    envio_bacen.open()
    LoginLayoutAntigo(driver=driver, url=url).login(usuario = "00735860000173", senha = "Amhpdf!2023)")
    lista_de_dados = []

    for pasta in lista_de_pastas:
        numero_processo = pasta[1]
        protocolo = envio_bacen.ler_planilhas(planilhas, numero_processo) #acrescentar alguma lógica aqui para pegar esse número de peg
        lista_de_arquivos = [f"{pasta[0]}/{arquivo}" for arquivo in os.listdir(pasta[0]) if arquivo.endswith('.pdf')]

        for arquivo in lista_de_arquivos:
            if "PEG" not in arquivo or "GUIAPRESTADOR" not in arquivo:
                n_amhptiss = arquivo.replace(f'{pasta[0]}/', '').replace('_Guia.pdf', '').replace(".pdf", '')
                envio_bacen.renomear_arquivo(pasta, arquivo, protocolo, n_amhptiss)
                lista_de_arquivos = [f"{pasta[0]}/{arquivo}" for arquivo in os.listdir(pasta[0]) if arquivo.endswith('.pdf')]
                
        nome_arquivo_zip = f'{numero_processo}.zip'
        envio_bacen.zipar_arquivos(pasta, nome_arquivo_zip, lista_de_arquivos)

        sz = (os.path.getsize(f"{pasta[0]}\\{nome_arquivo_zip}") / 1024) / 1024

        arquivo_zipado = f"{pasta[0]}\\{nome_arquivo_zip}"

        if sz >= 25.00:
            informacoes = [numero_processo, protocolo, "Não Enviado, arquivo .zip maior que 25MB"]
            continue

        envio_bacen.exe_caminho()
        envio_bacen.pesquisar_protocolo(protocolo)
        informacoes = envio_bacen.anexar_guias(arquivo_zipado, numero_processo, protocolo)
        lista_de_dados.append(informacoes)

    cabecalho = ["Número Processo", "Número Protocolo", "Envio", "Erro na Matricula", "Guias Não Anexadas"]
    df = pd.DataFrame(lista_de_dados, columns=cabecalho)