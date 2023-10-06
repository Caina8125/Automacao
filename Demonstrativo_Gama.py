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

class PageElement(ABC):
    def __init__(self, driver, url=''):
        self.driver = driver
        self.url = url
    def open(self):
        self.driver.get(self.url)

class Login(PageElement):
    usuario = (By.XPATH, '//*[@id="username"]')
    senha = (By.XPATH, '//*[@id="password"]')
    entrar = (By.XPATH, '//*[@id="submitPrestador"]')

    def exe_login(self, usuario, senha):
        self.driver.find_element(*self.usuario).send_keys(usuario)
        self.driver.find_element(*self.senha).send_keys(senha)
        self.driver.find_element(*self.entrar).click()

class Caminho(PageElement):
    gama = (By.XPATH, '//*[@id="TB_ajaxContent"]/div/div[1]/div/div/a/img')
    extrato = (By.XPATH, '//*[@id="header"]/div[2]/div/div[2]/ul/li[4]/a')
    demonstrativo_tiss = (By.XPATH, '//*[@id="header"]/div[2]/div/div[2]/ul/li[4]/ul/li[2]/a')
    periodo_30_dias = (By.XPATH, '//*[@id="cmbPeriodo"]/option[2]')
    consultar = (By.XPATH, '//*[@id="btnConsultarExtratoPeriodo"]')
    detalhar_demonstrativo = (By.XPATH, '/html/body/div[2]/div/div/div[2]/div/table/tbody/tr[1]/td[4]/form/a')
    botao_aceitar_mensagem = (By.XPATH, '//*[@id="TB_ajaxContent"]/div/div/div[2]/div/a')

    def fechar_comunicado(self):

        while True:
            try:
                self.driver.implicitly_wait(5)
                self.driver.find_element(*self.botao_aceitar_mensagem).click()
                time.sleep(1.5)

            except:
                break


    def exe_caminho(self):
        self.driver.implicitly_wait(30)
        self.driver.find_element(*self.gama).click()
        self.fechar_comunicado()
        time.sleep(2)
        self.driver.implicitly_wait(30)
        self.driver.find_element(*self.extrato).click()
        time.sleep(2)
        self.driver.find_element(*self.demonstrativo_tiss).click()
        self.driver.implicitly_wait(30)
        time.sleep(2)
        self.driver.find_element(*self.periodo_30_dias).click()
        time.sleep(2)
        self.driver.find_element(*self.consultar).click()
        self.driver.implicitly_wait(30)
        time.sleep(2)
        self.driver.find_element(*self.detalhar_demonstrativo).click()
        time.sleep(4)

class BaixarDemonstrativos(PageElement):
    tabela = (By.XPATH, '/html/body/div[2]/div[1]/div/div[2]/div[1]/div[7]/table')
    baixar_xml = (By.XPATH, '//*[@id="demonstrativoLotePagamento"]/div[2]/a')
    fechar = (By.XPATH, '//*[@id="demonstrativoLotePagamento"]/a')

    def baixar_demonstrativos(self):
        self.driver.implicitly_wait(30)
        tabela = self.driver.find_element(*self.tabela)
        codigo_html_tabela = tabela.get_attribute('outerHTML')
        data_frame_tabela = pd.read_html(codigo_html_tabela)[0]
        quantidade_de_demonstrativos = len(data_frame_tabela) + 1
        time.sleep(2)

        for i in range(1, quantidade_de_demonstrativos):
            valor_glosa = self.driver.find_element(By.XPATH, f'/html/body/div[2]/div[1]/div/div[2]/div[1]/div[7]/table/tbody/tr[{i}]/td[7]').text
            time.sleep(2)

            if valor_glosa == 'R$ 0,00':
                numero_do_lote_clicado = False

                while numero_do_lote_clicado == False:
                    try:
                        self.driver.find_element(By.XPATH, f'/html/body/div[2]/div[1]/div/div[2]/div[1]/div[7]/table/tbody/tr[{i}]/td[1]/a').click()
                        numero_do_lote_clicado = True
                    except:
                        self.driver.execute_script('scrollBy(0,50)')

                time.sleep(3)
                self.driver.implicitly_wait(60)
                time.sleep(1.5)
                self.driver.find_element(*self.baixar_xml).click()
                time.sleep(2)
                self.driver.find_element(*self.fechar).click()
                time.sleep(2)

            else:
                continue

        tkinter.messagebox.showinfo( 'Demonstrativos Gama' , f"Downloads concluídos!" )
        self.driver.quit()

#--------------------------------------------------------------------------------------------------------------------

def demonstrativo_gama():
    url = 'https://wwwt.connectmed.com.br/conectividade/prestador/home.htm'

    options = {
        'proxy' : {
            'http': 'http://lucas.paz:Gsw2022&@10.0.0.230:3128',
            'https': 'http://lucas.paz:Gsw2022&@10.0.0.230:3128'
        }
    }

    chrome_options = Options()
    chrome_options.add_experimental_option('prefs', {
        "download.default_directory": r"\\10.0.0.239\automacao_financeiro\GAMA",
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,
        'safebrowsing.enabled': 'false'
})
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument('--ignore-certificate-errors')
    chrome_options.add_argument('--ignore-ssl-errors')
    servico = Service(ChromeDriverManager().install())

    driver = webdriver.Chrome(service=servico, seleniumwire_options= options, options = chrome_options)
    
    usuario = "amhpfinanceiro"
    senha = "H!4Pdgf4"

    login_page = Login(driver, url)
    login_page.open()
    login_page.exe_login(usuario, senha)
    Caminho(driver, url).exe_caminho()
    BaixarDemonstrativos(driver, url).baixar_demonstrativos()