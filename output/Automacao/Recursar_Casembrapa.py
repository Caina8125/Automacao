from tkinter import filedialog
import pandas as pd
from selenium import webdriver
from seleniumwire import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from abc import ABC
import time
import datetime
import os
import tkinter
import json
import shutil

class PageElement(ABC):
    def __init__(self, driver, url=''):
        self.driver = driver
        self.url = url
    def open(self):
        self.driver.get(self.url)

class Login(PageElement):
    usuario = (By.XPATH, '//*[@id="username"]')
    senha = (By.XPATH, '//*[@id="password"]')
    entrar = (By.XPATH, '//*[@id="submit-login"]')

    def exe_login(self, usuario, senha):
        self.driver.implicitly_wait(30)
        self.driver.find_element(*self.usuario).send_keys(usuario)
        time.sleep(2)
        self.driver.find_element(*self.senha).send_keys(senha)
        time.sleep(2)
        self.driver.find_element(*self.entrar).click()
        time.sleep(2)

class Caminho(PageElement):
    salutis = (By.XPATH, '//*[@id="menuButtons"]/td[1]')
    websaude = (By.XPATH, '//*[@id="divTreeNavegation"]/div[7]/span[2]')
    credenciados = (By.XPATH, '//*[@id="divTreeNavegation"]/div[8]/span[2]')
    lotes = (By.XPATH, '//*[@id="divTreeNavegation"]/div[11]/span[2]')
    lotes_de_credenciados = (By.XPATH, '//*[@id="divTreeNavegation"]/div[18]/span[2]')
    numero_lote_pesquisa = (By.XPATH, '//*[@id="grdPesquisa"]/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[5]/td[2]/table/tbody/tr/td[1]/input')
    buscar_lotes = (By.XPATH, '//*[@id="buttonsContainer_1"]/td[1]/span[2]')
    numero_lote_operadora = (By.CSS_SELECTOR, '#form-view-label_gridLote_NUMERO > table > tbody > tr > td:nth-child(1) > input')
    processamento_de_guias = (By.XPATH, '//*[@id="divTreeNavegation"]/div[6]/span[2]')
    recurso_de_glosa = (By.XPATH, '//*[@id="divTreeNavegation"]/div[9]/span[2]')

    def buscar_numero_lote(self, numero_fatura):
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
        numero_operadora = selector.get_attribute('value')
        return numero_operadora
    
    def exe_caminho(self):
        self.driver.implicitly_wait(30)
        self.driver.find_element(*self.salutis).click()
        time.sleep(2)
        self.driver.find_element(*self.processamento_de_guias).click()
        time.sleep(2) 
        self.driver.find_element(*self.recurso_de_glosa).click()

class Recursar(PageElement):
    input_guia = (By.XPATH, '//*[@id="pesquisaParametro"]/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input')
    input_processo = (By.XPATH, '//*[@id="pesquisaParametro"]/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr[1]/td[2]/table/tbody/tr/td[1]/input')
    input_lote = (By.XPATH, '//*[@id="pesquisaParametro"]/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr[7]/td[2]/table/tbody/tr/td[1]/input')
    buscar = (By.XPATH, '//*[@id="buttonsContainer_5"]/td/span[2]')
    mudar_visao_1 = '//*[@id="changeViewButton"]'
    mudar_visao_2 = '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[17]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[1]/div'
    mudar_visao_3 = '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[17]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[7]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[1]/div'
    mudar_visao_4 = (By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[17]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[7]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[1]/div')
    localizar = (By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[17]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[2]/td[1]/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div')
    pesquisar_guia = (By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[17]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[3]/td[1]/table/tbody/tr[1]/td[3]/input')
    todos_os_campos = (By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[17]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[3]/td[1]/table/tbody/tr[2]/td[3]/table/tbody/tr[1]/td[1]/input')
    proxima = (By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[17]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[3]/td[1]/table/tbody/tr[1]/td[4]')


    def recursar(self, planilha):
        df = pd.read_excel(planilha)
        for index, linha in df.iterrows():
            numero_fatura = str(linha['Fatura']).replace('.0', '')
            lote_operadora = Caminho(driver, url).buscar_numero_lote(numero_fatura)
            self.driver.switch_to.frame('inlineFrameTabId5')
            time.sleep(1)
            self.driver.find_element(*self.input_guia).clear()
            time.sleep(1)
            self.driver.find_element(*self.input_processo).clear()
            time.sleep(1)
            self.driver.find_element(*self.input_lote).clear()
            time.sleep(1)
            self.driver.find_element(*self.input_lote).send_keys(lote_operadora)
            self.driver.switch_to.default_content()

            for i in range(1, 3):
                try:
                    self.driver.find_element(*self.buscar).click()
                    time.sleep(2)
                    texto_no_botao = self.driver.find_element(*self.buscar).text
                    if texto_no_botao == 'Nova Busca':
                        break
                except:
                    continue
            
            self.driver.switch_to.frame('inlineFrameTabId5')
            time.sleep(1)
            self.driver.find_element(*self.mudar_visao_1).click()
            time.sleep(1)
            self.driver.find_element(*self.mudar_visao_2).click()
            time.sleep(1)
            self.driver.find_element(*self.mudar_visao_3).click()
            time.sleep(1)
            self.driver.find_element(*self.mudar_visao_4).click()
            time.sleep(1)
            valor_glosa_planilha = linha['Valor Glosa'].replace('R$ ', '')
            self.driver.find_element(*self.pesquisar_guia).send_keys("")
            time.sleep(1)
            self.driver.find_element(*self.localizar).click()
            time.sleep(1)
            self.driver.find_element(*self.todos_os_campos).click()
            time.sleep(1)
            self.driver.find_element(*self.proxima).click()
            time.sleep(1)
            valor_glosa_planilha = linha['Valor Glosa'].replace('R$ ', '')
            codigo_procedimento_planilha = str(linha['Procedimento']).replace('.0', '')

            css_valor_portal = self.driver.find_element(By.CSS_SELECTOR, '#gridLote_gridguia_griditensdaguia > tbody > tr:nth-child(1) > td:nth-child(1) > table > tbody > tr:nth-child(4) > td > table > tbody > tr.grid-record.formView.odd > td:nth-child(2) > table > tbody > tr > td:nth-child(1) > input')
            valor_glosa_portal = css_valor_portal.get_attribute('value')

            css_procedimento = self.driver.find_element(By.CSS_SELECTOR, '#gridLote_gridguia_griditensdaguia > tbody > tr:nth-child(1) > td:nth-child(1) > table > tbody > tr:nth-child(1) > td > table > tbody > tr:nth-child(5) > td:nth-child(2) > table > tbody > tr > td:nth-child(1) > input')
            procedimento_portal = css_procedimento.get_attribute('value')

            checagem_codigo_1 = (codigo_procedimento_planilha[0] == '1' and len(codigo_procedimento_planilha) == 9)
            checagem_codigo_2 = codigo_procedimento_planilha[0] == "0" or codigo_procedimento_planilha[0] == "6" or codigo_procedimento_planilha[0] == "7" or codigo_procedimento_planilha[0] == "8"
            checagem_codigo_3 = codigo_procedimento_planilha[0] == "9" and codigo_procedimento_planilha[1] != "8"

            checagem_matmed = checagem_codigo_1 or checagem_codigo_2 or checagem_codigo_3

            

def iniciar():
    global url
    planilha = filedialog.askopenfilename()
    url = 'http://170.84.17.131:22101/sistema'

    settings = {
       "recentDestinations": [{
            "id": "Save as PDF",
            "origin": "local",
            "account": "",
        }],
        "selectedDestinationId": "Save as PDF",
        "version": 2
    }

    options = {
        'proxy' : {
            'http': 'http://lucas.paz:Gsw2022&@10.0.0.230:3128',
            'https': 'http://lucas.paz:Gsw2022&@10.0.0.230:3128'
        }
    }

    chrome_options = Options()
    chrome_options.add_experimental_option('prefs', {
        "printing.print_to_pdf": True,
        "download.default_directory": r"C:\Users\lucas.paz\Documents\Financeiro",
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,
        "printing.print_preview_sticky_settings.appState": json.dumps(settings),
        "savefile.default_directory": r"C:\Users\lucas.paz\Documents\Financeiro\Renomear"
})
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument('--ignore-certificate-errors')
    chrome_options.add_argument('--ignore-ssl-errors')
    chrome_options.add_argument('--kiosk-printing')
    servico = Service(ChromeDriverManager().install())

    global driver
    driver = webdriver.Chrome(service=servico, seleniumwire_options= options, options = chrome_options)

    global usuario, senha
    usuario = "00735860000173"
    senha = "0073586@"

    global login_page
    login_page = Login(driver, url)
    login_page.open()
    login_page.exe_login(usuario, senha)
    Recursar(driver, url).recursar(planilha)
    # BaixarDemonstrativo(driver, url).baixar_demonstrativo(planilha)

iniciar()
