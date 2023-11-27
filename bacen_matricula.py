from abc import ABC
from time import sleep
import pandas as pd
from selenium.webdriver.common.by import By

class PageElement(ABC):
    def __init__(self, driver, url)-> None:
        self.driver = driver
        self.url = url
        
    def open(self)-> None:
        self.driver.get(self.url)

class LoginLayoutNovo(PageElement):
    usuario_input = (By.XPATH, '//*[@id="UserName"]')
    senha_input = (By.XPATH, '//*[@id="Password"]')
    login_button = (By.XPATH, '//*[@id="LoginButton"]')
    
    def login_layout_novo(self, usuario, senha):
        self.driver.find_element(*self.usuario_input).send_keys(usuario)
        sleep(2)
        self.driver.find_element(*self.senha_input).send_keys(senha)
        sleep(2)
        self.driver.find_element(*self.login_button).click()
        sleep(2)

class BuscarProtocolo(PageElement):
    envio_de_arquivo_xml = (By.XPATH, '//*[@id="sidebar_envioXML"]/span[1]')
    consultar_arquivo = (By.XPATH, '//*[@id="ctl00_SidebarMenu"]/li[10]/ul/li[2]/a/span[1]/span')
    input_numero_lote = (By.XPATH, '/html/body/form/div[3]/div[3]/div[3]/div/div/div[1]/div/div/div[2]/div/div[2]/span/div/div/div/div[3]/div/div/div[2]/div[1]/span/input')
    a_numero_protocolo = (By.XPATH, '/html/body/form/div[3]/div[3]/div[3]/div/div/div[1]/div/div/div[2]/div/div[2]/div[4]/table/tbody/tr/td[4]/a')

    def buscar_protocolo(self, planilha):
        df = pd.read_excel(planilha)
        numero_fatura = df.loc[13, 'Unnamed: 13']
        self.driver.find_element(*self.envio_de_arquivo_xml).click()
        sleep(2)
        self.driver.find_element(*self.consultar_arquivo).click()
        sleep(2)
        self.driver.find_element(*self.input_numero_lote).send_keys(numero_fatura)
        sleep(2)
        numero_protocolo = self.driver.find_element(*self.a_numero_protocolo).text
        return numero_protocolo