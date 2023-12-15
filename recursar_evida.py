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
import Pidgin
import tkinter
import os

class PageElement(ABC):
    def __init__(self, driver, url=''):
        self.driver = driver
        self.url = url
    def open(self):
        self.driver.get(self.url)

class Login(PageElement):
    prestador_pf = (By.XPATH, '//*[@id="tipoAcesso"]/option[6]')
    usuario = (By.XPATH, '//*[@id="login-entry"]')
    senha = (By.XPATH, '//*[@id="password-entry"]')
    entrar = (By.XPATH, '//*[@id="BtnEntrar"]')

    def exe_login(self, usuario, senha):
        self.driver.find_element(*self.prestador_pf).click()
        self.driver.find_element(*self.usuario).send_keys(usuario)
        self.driver.find_element(*self.senha).send_keys(senha)
        self.driver.find_element(*self.entrar).click()
        time.sleep(5)

class Caminho(PageElement):
    faturas = (By.XPATH, '//*[@id="menuPrincipal"]/div/div[11]/a')
    relatorio_de_faturas = (By.XPATH, '/html/body/header/div[4]/div/div/div/div[11]/div[1]/div[2]/div/div[2]/div/div/div/div[1]/a')

    def exe_caminho(self):
        try:
            self.driver.implicitly_wait(10)
            time.sleep(2)
            self.driver.find_element(*self.faturas)
        except:
            self.driver.refresh()
            time.sleep(2)
            login_page.exe_login(usuario, senha)

        self.driver.implicitly_wait(30)
        time.sleep(3)
        self.driver.find_element(*self.faturas).click()
        time.sleep(2)
        self.driver.find_element(*self.relatorio_de_faturas).click()
        time.sleep(2)

class Recurso(PageElement):
    codigo = (By.XPATH, '/html/body/main/div/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/input-text[1]/div/div/input')
    pesquisar = (By.XPATH, '//*[@id="filtro"]/div[2]/div[2]/button')
    recurso_de_glosa = (By.XPATH, '/html/body/main/div/div[1]/div[4]/div/div/div[1]/div/div[2]/a[5]/i')

    def fazer_recurso(self, pasta):
        lista_de_planilhas = [f"{pasta}//{arquivo}" for arquivo in pasta if arquivo.endswith(".xlsx")]

        for planilha in lista_de_planilhas:
            df = pd.read_excel(planilha)
            protocolo = f"{df['Protocolo Glosa'][0]}".replace(".0", "")
            self.driver.find_element(*self.codigo).send_keys(protocolo)
            time.sleep(2)
            self.driver.find_element(*self.pesquisar).click()
            time.sleep(2)
            self.driver.find_element(*self.recurso_de_glosa).click()
            time.sleep(2)

            for index, linha in df.iterrows():
                numero_guia = f'{linha["Nro. Guia"]}'.replace('.0', '')
                codigo_procedimento = f'{linha["Procedimento"]}'.replace('.0', '')

                

def demonstrativo_evida():
    try:
        url = 'https://novowebplanevida.facilinformatica.com.br/GuiasTISS/Logon'
        planilha = filedialog.askopenfilename()

        options = {
            'proxy' : {
                'http': 'http://lucas.paz:RDRsoda90901@@10.0.0.230:3128',
                'https': 'http://lucas.paz:RDRsoda90901@@10.0.0.230:3128'
            }
        }

        chrome_options = Options()
        chrome_options.add_argument("--start-maximized")
        try:
            servico = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=servico, seleniumwire_options= options, options = chrome_options)
        except:
            driver = webdriver.Chrome(seleniumwire_options= options, options = chrome_options)
        
        global usuario, senha, login_page, caminho
        usuario = "00735860000173"
        senha = "00735860000173"

        login_page = Login(driver, url)
        login_page.open()
        login_page.exe_login(usuario, senha)
    
    except:
        ...