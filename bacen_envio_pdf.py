from abc import ABC
from time import sleep
import pandas as pd
from selenium import webdriver
from seleniumwire import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from tkinter.filedialog import askdirectory
from bacen_protocolo import BuscarProtocolo
from abc import ABC
import pandas as pd
from seleniumwire import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from tkinter import messagebox
import os
import zipfile

class PageElement(ABC):
    def __init__(self, driver=None, url=None)-> None:
        self.driver = driver
        self.url = url
        
    def open(self)-> None:
        self.driver.get(self.url)
        sleep(2)

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
    tbody_guia_com_anexo = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr[1]/td/div/div/div/div/div/div/div/div/div[2]/table/tbody')
    botao_novo = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td/div/div/div/div/div/div/div/div/div[1]/div[2]/a')
    procurar_arquivo = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[1]/td/div/div[3]/div/div/div/div/div/div/form/table/tbody/tr[2]/td[2]/div/div[2]/a/img')
    input_file = (By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/form/div/div/div/div/div/div/div/div/div[2]/table/tbody/tr/td/table/tbody/tr[2]/td[2]/input')
    botao_enviar = (By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/form/div/div/div/div/div/div/div/div/div[2]/table/tbody/tr/td/table/tbody/tr[3]/td/input')
    botao_salvar = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[1]/td/div/div[2]/div/div/div/div[3]/a')
    lupa_conta_fisica = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td/div/div/div/div/div/div/div/div/div[2]/table/tbody/tr[2]/td[1]/a[1]')
    processar_anexo = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr/td/div[1]/div/div/div/div/div/div/div/div[3]/div/table/tbody/tr/td/div/nobr/a')

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

    def renomear_arquivo(self, pasta, arquivo, protocolo, amhptiss):
        os.rename(arquivo, f"{pasta}\\PEG{protocolo}_GUIAPRESTADOR{amhptiss}.pdf")
    
    def anexar_guias(self, arquivo):
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
    
    def zipar_arquivos(self, pasta, nome_arquivo_zip, lista_de_arquivos):
        with zipfile.ZipFile(f"{pasta[0]}/{nome_arquivo_zip}", "w", zipfile.ZIP_DEFLATED) as zipf:
            for arquivo in lista_de_arquivos:
                if arquivo.endswith('.pdf') and "PEG" in arquivo and "GUIAPRESTADOR" in arquivo:
                    zipf.write(arquivo, os.path.relpath(arquivo, pasta[0]))

diretorio = askdirectory()
lista_de_pastas = [[f"{diretorio}/{pasta}", pasta] for pasta in os.listdir(diretorio) if pasta.isdigit()]
teste = EnvioPDF()

for pasta in lista_de_pastas:
    numero_processo = pasta[1]
    protocolo = ... #acrescentar alguma lógica aqui para pegar esse número de peg
    lista_de_arquivos = [f"{pasta[0]}/{arquivo}" for arquivo in os.listdir(pasta[0]) if arquivo.endswith('.pdf')]

    for arquivo in lista_de_arquivos:
        if "_Guia" in arquivo:
            n_amhptiss = arquivo.replace(f'{pasta}/', '').replace('_Guia.pdf', '')
            teste.renomear_arquivo(pasta, arquivo, protocolo, n_amhptiss)
            lista_de_arquivos = [f"{pasta}/{arquivo}" for arquivo in os.listdir(pasta) if arquivo.endswith('.pdf')]
            
    nome_arquivo_zip = f'{numero_processo}.zip'
    teste.zipar_arquivos(pasta, nome_arquivo_zip, lista_de_arquivos)

    sz = (os.path.getsize(f"{pasta[0]}\\{nome_arquivo_zip}") / 1024) / 1024

    if sz >= 25.00:
        continue

    print('')