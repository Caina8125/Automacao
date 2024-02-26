from tkinter import filedialog
import tkinter.messagebox
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium import webdriver
from openpyxl import load_workbook
from abc import ABC
import pandas as pd
import time
import os
from selenium.webdriver.chrome.options import Options
from seleniumwire import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service

class PageElement(ABC):
    def __init__(self, driver, url=''):
        self.driver = driver
        self.url = url
    def open(self):
        self.driver.maximize_window()
        self.driver.get(self.url)

class Login(PageElement):
    usuario = (By.XPATH, '//*[@id="UserName"]')
    senha = (By.XPATH, '//*[@id="Password"]')
    acessar = (By.XPATH, '//*[@id="LoginButton"]')

    def exe_login(self, usuario, senha):
        self.driver.find_element(*self.usuario).send_keys(usuario)
        self.driver.find_element(*self.senha).send_keys(senha)
        self.driver.find_element(*self.acessar).click()

class Xml(PageElement):
    envio_xml = (By.XPATH, '//*[@id="sidebar_envioXML"]/span[1]')
    enviar_xml = (By.XPATH, '//*[@id="ctl00_SidebarMenu"]/li[10]/ul/li[1]/a')
    anexar_xml = (By.XPATH, '//*[@id="ctl00_Body"]/input')
    gravar = (By.XPATH, '//*[@id="ctl00_Main_WIDGETID_636040089093201024_toolbar"]/a[2]')

    def caminho(self):
        self.driver.find_element(*self.envio_xml).click()
        time.sleep(2)
        self.driver.find_element(*self.enviar_xml).click()
        time.sleep(2)
    
    def arquivos(self):
        xml = os.listdir(pasta)
        return xml

    def envio(self):
        self.caminho()
        arquivos = self.arquivos()
        for xml in arquivos:
            caminho_arquivo = pasta + '/' + xml
            self.driver.find_element(*self.anexar_xml).send_keys(caminho_arquivo)
            time.sleep(1)
            self.driver.find_element(*self.gravar).click()
            time.sleep(1)
            print(xml)
            
#---------------------------------------------------------------------------------------------------------------------------------
def Enviar_caixa(user, password):
    global pasta
    pasta = filedialog.askdirectory()

    global url
    url = 'https://saude.caixa.gov.br/PORTALPRD/'

    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument('--ignore-certificate-errors')
    chrome_options.add_argument('--ignore-ssl-errors')

    options = {
    'proxy': {
            'http': f'http://{user}:{password}@10.0.0.230:3128',
                'https': f'http://{user}:{password}@10.0.0.230:3128'
        }
    }

    try:
        servico = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=servico, seleniumwire_options=options, options=chrome_options)
    except:
        driver = webdriver.Chrome(seleniumwire_options=options, options=chrome_options)

    try:
        login_page = Login(driver, url)
        login_page.open()

        login_page.exe_login(
            usuario = "00735860000173",
            senha = "Saude@2024!"
        )

        Xml(driver, url).envio()

        tkinter.messagebox.showinfo( 'Automação Saúde Caixa Recurso de Glosa' , 'Recursos do Saúde Caixa Concluídos 😎✌' )
    except Exception as e:
        tkinter.messagebox.showerror( 'Erro Automação' , f'Ocorreu uma excessão não tratada \n {e.__class__.__name__}: {e}' )
        driver.quit()