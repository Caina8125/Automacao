from tkinter import filedialog
import pandas as pd
from selenium.webdriver import Chrome
from seleniumwire import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import time
import json
# from page_element import PageElement

class PageElement():
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
        self.url:str = url

    def open(self) -> None:
        self.driver.get(self.url)

class SalutisCasembrapa(PageElement):
    usuario_input = (By.XPATH, '//*[@id="username"]')
    senha_input = (By.XPATH, '//*[@id="password"]')
    entrar = (By.XPATH, '//*[@id="submit-login"]')
    salutis = (By.XPATH, '//*[@id="menuButtons"]/td[1]')
    websaude = (By.XPATH, '//*[@id="divTreeNavegation"]/div[7]/span[2]')
    credenciados = (By.XPATH, '//*[@id="divTreeNavegation"]/div[8]/span[2]')
    lotes = (By.XPATH, '//*[@id="divTreeNavegation"]/div[11]/span[2]')
    lotes_de_credenciados = (By.XPATH, '//*[@id="divTreeNavegation"]/div[16]/span[2]')
    numero_lote_pesquisa = (By.XPATH, '//*[@id="grdPesquisa"]/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[5]/td[2]/table/tbody/tr/td[1]/input')
    buscar_lotes = (By.XPATH, '//*[@id="buttonsContainer_1"]/td[1]/span[2]')
    numero_lote_operadora = (By.CSS_SELECTOR, '#form-view-label_gridLote_NUMERO > table > tbody > tr > td:nth-child(1) > input')
    processamento_de_guias = (By.XPATH, '//*[@id="divTreeNavegation"]/div[6]/span[2]')
    recurso_de_glosa = (By.XPATH, '//*[@id="divTreeNavegation"]/div[9]/span[2]')

    def __init__(self, driver: Chrome, url: str, usuario: str, senha: str) -> None:
        super().__init__(driver=driver, url=url)
        self.usuario: str = usuario
        self.senha: str = senha

    def login(self):
        self.open()
        self.driver.implicitly_wait(30)
        self.driver.find_element(*self.usuario_input).send_keys(self.usuario)
        time.sleep(2)
        self.driver.find_element(*self.senha_input).send_keys(self.senha)
        time.sleep(2)
        self.driver.find_element(*self.entrar).click()
        time.sleep(2)

    def buscar_numero_lote(self, numero_fatura: str):
        self.driver.implicitly_wait(30)
        self.driver.find_element(*self.salutis).click()
        time.sleep(2)
        self.driver.find_element(*self.websaude).click()
        time.sleep(2)
        self.driver.find_element(*self.credenciados).click()
        time.sleep(2)
        self.driver.find_element(*self.lotes).click()
        time.sleep(2)
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
        
        return selector.get_attribute('value')
    
    def executa_recurso(self, planilha):
        self.login()
        df = pd.read_excel(planilha)
        for index, linha in df.iterrows():
            numero_fatura = str(linha['Fatura']).replace('.0', '')
            lote_operadora = self.buscar_numero_lote(numero_fatura)
            print(lote_operadora)
    
def teste(user, password):
    planilha = filedialog.askopenfilename()
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

    login_page = SalutisCasembrapa(driver, url, usuario, senha)
    login_page.executa_recurso(planilha)

teste('lucas.paz', 'WheySesc2024*')