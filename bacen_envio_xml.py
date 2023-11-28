from abc import ABC
from time import sleep
import pandas as pd
from selenium import webdriver
from seleniumwire import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from tkinter.filedialog import askdirectory
from bacen_protocolo import BuscarProtocolo, PageElement
from abc import ABC
import pandas as pd
from seleniumwire import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from tkinter import messagebox
import os

class EnviarXML(PageElement):
    opcao_envio_de_xml = (By.XPATH, '/html/body/form/div[3]/div[3]/div[1]/div/ul/li[10]/a')
    opcao_enviar_arquivo = (By.XPATH, '/html/body/form/div[3]/div[3]/div[1]/div/ul/li[10]/ul/li[1]/a')
    input_file = (By.XPATH, "/html/body/input")
    salvar_novo = (By.XPATH, '/html/body/form/div[3]/div[3]/div[3]/div/div/div/div/div/div[2]/div/div[2]/div[1]/a[2]')

    def caminho(self):
        self.driver.find_element(*self.opcao_envio_de_xml).click()
        sleep(2)
        self.driver.find_element(*self.opcao_enviar_arquivo).click()
        sleep(2)
    
    def enviar_arquivo(self, lista_de_arquivos):
        for arquivo in lista_de_arquivos:
            self.driver.find_element(*self.input_file).send_keys(arquivo)
            sleep(1)
            self.driver.find_element(*self.salvar_novo).click()
            sleep(1.5)
        sleep(10)
        return ...

    def confere_envio(self):
        ...

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------

# def fazer_envio():
messagebox.showwarning("Automação Bacen", "Selecione uma pasta!")
pasta = askdirectory()
lista_de_arquivos = [f"{pasta}\\{arquivo}" for arquivo in os.listdir(pasta) if arquivo.endswith(".xml")]
url = 'https://www3.bcb.gov.br/portalbcsaude/Login'
global driver

chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument('--ignore-certificate-errors')
chrome_options.add_argument('--ignore-ssl-errors')

options = {
'proxy': {
        'http': 'http://lucas.paz:RDRsoda90901@@10.0.0.230:3128',
        'https': 'http://lucas.paz:RDRsoda90901@@10.0.0.230:3128'
    }
}
try:
    servico = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=servico, seleniumwire_options=options, options=chrome_options)
except:
    driver = webdriver.Chrome(seleniumwire_options=options, options=chrome_options)

try:
    acessar_portal = BuscarProtocolo(driver, url)
    acessar_portal.open()

    acessar_portal.login_layout_novo(
        usuario = "00735860000173",
        senha = "Amhpdf!2023"
    )
    enviar_xml = EnviarXML(driver, url='')
    enviar_xml.caminho()
    lista_de_processos = enviar_xml.enviar_arquivo(lista_de_arquivos)
    driver.quit()
    messagebox.showinfo( 'Automação Bacen' , 'Todas as pesquisas foram concluídas.' )
except Exception as e:
    messagebox.showerror( 'Erro Automação' , f'Ocorreu uma excessão não tratada \n {e.__class__.__name__}: {e}' )
    driver.quit()