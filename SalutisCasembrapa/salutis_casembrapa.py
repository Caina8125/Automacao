from abc import ABC, abstractmethod
from tkinter import filedialog
import pandas as pd
from selenium.webdriver import Chrome
from seleniumwire import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
import time
from os import listdir
import json
# from page_element import PageElement

class PageElement(ABC):
    body: tuple = (By.XPATH, '/html/body')
    a: tuple = (By.TAG_NAME, 'a')
    p: tuple = (By.TAG_NAME, 'p')
    h1: tuple = (By.TAG_NAME, 'h1')
    h2: tuple = (By.TAG_NAME, 'h2')
    h3: tuple = (By.TAG_NAME, 'h3')
    h4: tuple = (By.TAG_NAME, 'h4')
    h5: tuple = (By.TAG_NAME, 'h5')
    h6: tuple = (By.TAG_NAME, 'h6')

    def __init__(self, driver: Chrome, url: str = '') -> None:
        self.driver: Chrome = driver
        self._url:str = url

    def open(self) -> None:
        self.driver.get(self._url)

    @property
    @abstractmethod
    def url(self):...

class SalutisCasembrapa(PageElement):
    usuario_input: tuple = (By.XPATH, '//*[@id="username"]')
    senha_input: tuple = (By.XPATH, '//*[@id="password"]')
    entrar: tuple = (By.XPATH, '//*[@id="submit-login"]')
    salutis: tuple = (By.XPATH, '//*[@id="menuButtons"]/td[1]')
    websaude: tuple = (By.XPATH, '//*[@id="divTreeNavegation"]/div[7]/span[2]')
    credenciados: tuple = (By.XPATH, '//*[@id="divTreeNavegation"]/div[8]/span[2]')
    lotes: tuple = (By.XPATH, '//*[@id="divTreeNavegation"]/div[11]/span[2]')
    lotes_de_credenciados: tuple = (By.XPATH, '/html/body/div[8]/div[2]/div[20]/span[2]')
    fechar_lotes_de_credenciados: tuple = (By.XPATH, '//*[@id="tabs"]/td[1]/table/tbody/tr/td[4]/span')
    numero_lote_pesquisa: tuple = (By.XPATH, '//*[@id="grdPesquisa"]/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[5]/td[2]/table/tbody/tr/td[1]/input')
    buscar_lotes: tuple = (By.XPATH, '//*[@id="buttonsContainer_1"]/td[1]/span[2]')
    numero_lote_operadora: tuple = (By.CSS_SELECTOR, '#form-view-label_gridLote_NUMERO > table > tbody > tr > td:nth-child(1) > input')
    processamento_de_guias: tuple = (By.XPATH, '//*[@id="divTreeNavegation"]/div[6]/span[2]')
    recurso_de_glosa: tuple = (By.XPATH, '//*[@id="divTreeNavegation"]/div[9]/span[2]')
    input_guia = (By.XPATH, '//*[@id="pesquisaParametro"]/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input')
    input_processo = (By.XPATH, '//*[@id="pesquisaParametro"]/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr[1]/td[2]/table/tbody/tr/td[1]/input')
    input_lote = (By.XPATH, '//*[@id="pesquisaParametro"]/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr[7]/td[2]/table/tbody/tr/td[1]/input')
    buscar = (By.ID, 'buttonsContainer_2')
    mudar_visao_1 = (By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[17]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[1]/div')
    mudar_visao_2 = (By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[17]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[1]/div')
    mudar_visao_3 = (By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[17]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[7]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[1]/div')
    mudar_visao_4 = (By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[17]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[7]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[1]/div')
    lupa_localizar_guia = (By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[17]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[2]/td[1]/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div')
    lupa_localizar_servico = (By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[17]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[2]/td[1]/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div')
    input_localizar_guia = (By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[17]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[3]/td[1]/table/tbody/tr[1]/td[3]/input')
    input_localizar_servico = (By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[17]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[3]/td[1]/table/tbody/tr[1]/td[3]/input')
    radio_todos_os_campos_guia = (By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[17]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[3]/td[1]/table/tbody/tr[2]/td[3]/table/tbody/tr[1]/td[1]/input')
    radio_todos_os_campos_servicos = (By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[17]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[3]/td[1]/table/tbody/tr[2]/td[3]/table/tbody/tr[1]/td[1]/input')



    def __init__(self, driver: Chrome, url: str, usuario: str, senha: str, diretorio: str) -> None:
        super().__init__(driver=driver, url=url)
        self.usuario: str = usuario
        self.senha: str = senha
        self.diretorio: str = diretorio
        self.lista_de_planilhas = [
            f'{diretorio}\\{arquivo}' 
            for arquivo in listdir(diretorio)
            if arquivo.endswith('.xlsx')
            ]

    @property
    def url(self):
        return self._url

    def login(self):
        self.open()
        self.driver.implicitly_wait(30)
        self.driver.find_element(*self.usuario_input).send_keys(self.usuario)
        time.sleep(2)
        self.driver.find_element(*self.senha_input).send_keys(self.senha)
        time.sleep(2)
        self.driver.find_element(*self.entrar).click()
        time.sleep(2)

    def abrir_opcoes_menu(self):
        self.driver.find_element(*self.salutis).click()
        time.sleep(2)
        self.driver.find_element(*self.websaude).click()
        time.sleep(2)
        self.driver.find_element(*self.credenciados).click()
        time.sleep(2)
        self.driver.find_element(*self.lotes).click()
        time.sleep(2)
        self.driver.find_element(*self.processamento_de_guias).click()
        time.sleep(2)

    def pegar_numero_lote(self, numero_fatura: str):
        self.driver.implicitly_wait(30)
        self.driver.find_element(*self.lotes_de_credenciados).click()
        time.sleep(2)
        self.driver.switch_to.frame('inlineFrameTabId1')
        self.driver.find_element(*self.numero_lote_pesquisa).clear()
        time.sleep(2)
        self.driver.find_element(*self.numero_lote_pesquisa).send_keys(numero_fatura)
        time.sleep(2)
        self.driver.switch_to.default_content()
        
        for i in range(1, 5):
            try:
                self.driver.find_element(*self.buscar_lotes).click()
                time.sleep(2)
                texto_no_botao = self.driver.find_element(*self.buscar_lotes).text
                if texto_no_botao == 'Pesquisar Lotes':
                    break
            except:
                continue

        time.sleep(2)
        self.driver.switch_to.frame('inlineFrameTabId1')
        selector = self.driver.find_element(*self.numero_lote_operadora)
        lote_operadora = selector.get_attribute('value')
        self.driver.switch_to.default_content()
        self.driver.find_element(*self.fechar_lotes_de_credenciados).click()
        return lote_operadora
    
    def busca_fatura(self, numero_lote):
        self.driver.switch_to.frame('inlineFrameTabId2')
        time.sleep(2)
        self.driver.find_element(*self.input_guia).clear()
        time.sleep(1)
        self.driver.find_element(*self.input_processo).clear()
        time.sleep(1)
        self.driver.find_element(*self.input_lote).clear()
        time.sleep(1)
        self.driver.find_element(*self.input_lote).send_keys(numero_lote)
        self.driver.switch_to.default_content()

        for i in range(1, 3):
            try:
                self.driver.find_element(*self.buscar).click()
                time.sleep(4)
                texto_no_botao = self.driver.find_element(*self.buscar).text
                if texto_no_botao == 'Nova Busca':
                    break
            except:
                continue

    def abrir_divs(self):
        self.driver.switch_to.frame('inlineFrameTabId2')
        time.sleep(1)
        self.driver.find_element(*self.mudar_visao_1).click()
        time.sleep(1)
        self.driver.find_element(*self.mudar_visao_2).click()
        time.sleep(1)
        self.driver.find_element(*self.mudar_visao_3).click()
        time.sleep(1)
        self.driver.find_element(*self.mudar_visao_4).click()
        time.sleep(2)
        self.driver.find_element(*self.lupa_localizar_guia).click()
        time.sleep(2)
        self.driver.find_element(*self.lupa_localizar_servico).click()
        time.sleep(2)
        self.driver.find_element(*self.radio_todos_os_campos_guia).click()
        time.sleep(2)
        self.driver.find_element(*self.radio_todos_os_campos_servicos).click()
        time.sleep(2)
    
    def executa_recurso(self):
        self.login()
        self.abrir_opcoes_menu()

        for planilha in self.lista_de_planilhas:
            df = pd.read_excel(planilha)
            numero_fatura = str(df['Fatura'][0]).replace('.0', '')
            lote_operadora = self.pegar_numero_lote(numero_fatura)
            time.sleep(2)
            print(lote_operadora)

            # Entra em Recurso de Glosa
            self.driver.find_element(*self.salutis).click()
            time.sleep(2)
            self.driver.find_element(*self.recurso_de_glosa).click()
            time.sleep(2)
            self.busca_fatura(lote_operadora)
            time.sleep(2)
            self.abrir_divs()

            for index, linha in df.iterrows():
                ...

    
def teste(user, password):
    diretorio = filedialog.askdirectory()
    url = 'http://170.84.17.131:22101/sistema'

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
    chrome_options.add_argument('--kiosk-printing')
    servico = Service(ChromeDriverManager().install())

    driver = webdriver.Chrome(service=servico, seleniumwire_options= options, options = chrome_options)

    usuario = "00735860000173"
    senha = "0073586@"

    login_page = SalutisCasembrapa(driver, url, usuario, senha, diretorio)
    login_page.executa_recurso()

teste('lucas.paz', 'WheySesc2024*')