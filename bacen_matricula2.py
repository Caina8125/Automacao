from abc import ABC
from time import sleep
import pandas as pd
from seleniumwire import webdriver
from selenium.webdriver.common.by import By
from tkinter.filedialog import askopenfilename
import bacen_matricula
import requests
from bs4 import BeautifulSoup

class PageElement(ABC):
    def __init__(self, driver, url)-> None:
        self.driver = driver
        self.url = url
        
    def open(self)-> None:
        self.driver.get(self.url)
        sleep(2)

class LoginLayoutAntigo(PageElement):
    usuario_input = (By.XPATH, '//*[@id="rp1_edt"]/table/tbody/tr[2]/td/form/table/tbody/tr[1]/td[2]/input')
    senha_input = (By.XPATH, '//*[@id="rp1_edt"]/table/tbody/tr[2]/td/form/table/tbody/tr[2]/td[2]/input')
    login_button = (By.XPATH, '//*[@id="rp1_edt"]/table/tbody/tr[2]/td/form/table/tbody/tr[3]/td/div/input')

    def login(self, usuario, senha):
        self.driver.find_element(*self.usuario_input).send_keys()
        sleep(2)
        self.driver.find_element(*self.senha_input).send_keys()
        sleep(2)
        self.driver.find_element(*self.login_button).click()
        sleep(2)

class ConferirFatura(PageElement):
    ...

url = 'https://www3.bcb.gov.br/portalbcsaude/saude/a/portal/prestador/tiss/ConsultarArquivoTiss.aspx?i=PORTAL_SUBMENU_XML_CONSULTARARQUIVO&m=MENU_ENVIO_XML_AGRUPADO'

proxy = {
    'http': '10.0.0.230:3128',
    'https': 'lucas.paz:RDRsoda90901@@10.0.0.230:3128'
    }
data = {
    'ctl00$Main$WIDGETID_636039305340582320$FilterControl$GERAL$3__NUMEROLOTE': '2195645'
}

response = requests.post(url=url, proxies=proxy, data=data)
html = response.text
soup = BeautifulSoup(html)
print(soup.get_text('\n'))