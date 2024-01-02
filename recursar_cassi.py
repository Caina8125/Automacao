import tkinter.messagebox
from tkinter import filedialog
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
from abc import ABC
import pandas as pd
import time
import os
from selenium.webdriver.chrome.options import Options
from seleniumwire import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import Pidgin

class PageElement(ABC):
    def __init__(self, driver, url=''):
        self.driver = driver
        self.url = url
    def open(self):
        self.driver.get(self.url)

class Login(PageElement):
    usuario = (By.XPATH, '//*[@id="cpfOuCnpj"]')
    senha = (By.XPATH, '//*[@id="senha"]')
    acessar = (By.XPATH, '//*[@id="loginGeral"]')

    def exe_login(self, usuario, senha):
        self.driver.implicitly_wait(30)
        time.sleep(1.5)
        self.driver.find_element(*self.usuario).send_keys(usuario)
        time.sleep(1.5)
        self.driver.find_element(*self.senha).send_keys(senha)
        time.sleep(1.5)
        self.driver.find_element(*self.acessar).click()
        time.sleep(1.5)

class caminho(PageElement):
    finalizar = (By.XPATH, '//*[@id="step-0"]/nav/button')
    demonstrativo_tiss = (By.XPATH, '/html/body/div[1]/aside/section/div/div/div[1]/div[1]/ul/li[11]/a')
    demonstrativo_de_analises = (By.XPATH, '/html/body/div[1]/aside/section/div/div/div[1]/div[1]/ul/li[11]/ul/li[1]/a')

    def exe_caminho(self):
        self.driver.implicitly_wait(30)
        self.driver.find_element(*self.finalizar).click()
        time.sleep(4)
        self.driver.get('https://servicosonline.cassi.com.br/Prestador/RecursoRevisaoPagamento/TISS/DemonstrativoAnaliseContas/Index')

class Recursar(PageElement):
    protocolo_input = (By.XPATH, '//*[@id="ProtocoloPagamento"]')
    consultar = (By.XPATH, '/html/body/div[1]/div[5]/section/div/form/fieldset/div[4]/div/button')
    voltar = (By.XPATH, '//*[@id="btnVoltar"]')
    xpath_corpo_da_pagina = (By.XPATH, '/html/body')
    demonstrativo_de_analise = (By.XPATH, '/html/body/div[1]/div[5]/section/div/fieldset/div/table/tbody/tr/td[3]/form/input[3]')
    recursar = (By.XPATH, '/html/body/div[1]/div[5]/section/div/div[1]/div[1]/div/div[2]/button[2]')
    guia_input = (By.XPATH, '/html/body/div[1]/div[5]/section/div/div[7]/div[2]/div/div/fieldset/div[1]/div/input')
    pesquisar = (By.XPATH, '/html/body/div[1]/div[5]/section/div/div[7]/div[2]/div/div/fieldset/div[1]/div/div/a')
    guia_click = (By.XPATH, '/html/body/div[1]/div[5]/section/div/div[7]/div[2]/div/div/fieldset/div[2]/table/tbody/tr/td[1]')
    valor_recursando_input = (By.XPATH, '/html/body/div[1]/div[5]/section/div/div[3]/div/div/div[2]/div[1]/div[4]/div[2]/div[3]/div/div/input')
    mostrar_textarea_just = (By.XPATH, '/html/body/div[1]/div[5]/section/div/div[3]/div/div/div[2]/div[1]/div[6]/div[2]/div/div/div[1]/h4/a')
    textarea_justificativa = (By.XPATH, '/html/body/div[1]/div[5]/section/div/div[3]/div/div/div[2]/div[1]/div[5]/div/div/div/div/textarea')
    fechar_justificativa = (By.XPATH, '/html/body/div[1]/div[5]/section/div/div[3]/div/div/div[2]/div[1]/div[5]/div/div/div/div/div/button')
    salvar_recurso = (By.XPATH, '/html/body/div[1]/div[5]/section/div/div[3]/div/div/div[2]/div[1]/div[8]/div/button[1]')
    fechar_recurso = (By.XPATH, '/html/body/div[1]/div[5]/section/div/div[3]/div/div/div[2]/div[1]/div[8]/div/button[2]')
    ok = (By.XPATH, '/html/body/div[1]/div[5]/section/div/div[6]/div/div/div[3]/button')
    solicitar = (By.XPATH, '/html/body/div[1]/div[5]/section/nav/div/div[2]/ul[1]/li[3]/a')
    editar_recurso = (By.XPATH, '/html/body/div[1]/div[5]/section/div/div[10]/div/table/tbody/tr/td[10]/a[1]')

    def recurso(self, pasta):
        lista_de_planilhas = [f"{pasta}/{arquivo}" for arquivo in os.listdir(pasta) if arquivo.endswith('.xlsx')]

        for planilha in lista_de_planilhas:
            df = pd.read_excel(planilha)
            protocolo = f"{df['Protocolo Aceite'][0]}"
            self.driver.find_element(*self.protocolo_input).send_keys(protocolo)
            time.sleep(2)
            self.driver.find_element(*self.consultar).click()
            time.sleep(2)
            corpo_pagina = self.driver.find_element(*self.xpath_corpo_da_pagina).text
            time.sleep(1)

            if not "Não foram encontrados resultados para a pesquisa" in corpo_pagina:
                self.driver.find_element(*self.demonstrativo_de_analise).click()
                time.sleep(2)
                self.driver.find_element(*self.recursar).click()
                time.sleep(2)
                corpo_pagina = self.driver.find_element(*self.xpath_corpo_da_pagina).text
                if 'Ação não permitida. Já existe um recurso/revisão em digitação para este protocolo.' in corpo_pagina:
                    self.driver.find_element(*self.ok).click()
                    time.sleep(2)
                    self.driver.find_element(*self.solicitar).click()
                    time.sleep(2)
                    self.driver.find_element(*self.editar_recurso).click()
                    time.sleep(2)

                time.sleep(2)
                for index, linha in df.iterrows():
                    numero_controle = f"{linha['Controle Inicial']}".replace('.0', '')
                    procedimento = f"{linha['Procedimento']}".replace('.0', '')
                    valor_glosado = f"{linha['Valor Glosa']}".replace('-', '').replace('.', ',')
                    valor_recursar = f"{linha['Valor Recursado']}"
                    justificativa = f"{linha['Recurso Glosa']}"
                    self.driver.find_element(*self.guia_input).clear()
                    time.sleep(1)
                    self.driver.find_element(*self.guia_input).send_keys(numero_controle)
                    time.sleep(2)
                    self.driver.find_element(*self.pesquisar).click()
                    time.sleep(2)
                    self.driver.find_element(*self.guia_click).click()
                    time.sleep(2)
                    count = 1
                    while count != 0:
                        try:
                            codigo_portal = self.driver.find_element(By.XPATH, f'/html/body/div[1]/div[5]/section/div/div[7]/div[3]/div[1]/div[2]/fieldset/div/table/tbody/tr[{count}]/td[1]').text
                            valor_glosa_portal = self.driver.find_element(By.XPATH, f'/html/body/div[1]/div[5]/section/div/div[7]/div[3]/div[1]/div[2]/fieldset/div/table/tbody/tr[{count}]/td[5]').text
                            span_class = self.driver.find_element(By.XPATH, f'/html/body/div[1]/div[5]/section/div/div[7]/div[3]/div[1]/div[2]/fieldset/div/table/tbody/tr[{count}]/td[7]/span').get_attribute("class")
                            if procedimento == codigo_portal and valor_glosado == valor_glosa_portal and 'parcial' in span_class:
                                self.driver.find_element(By.XPATH, f'/html/body/div[1]/div[5]/section/div/div[7]/div[3]/div[1]/div[2]/fieldset/div/table/tbody/tr[{count}]/td[7]/a[2]').click()
                                time.sleep(2)
                                self.driver.find_element(*self.valor_recursando_input).clear()
                                time.sleep(2)
                                self.driver.find_element(*self.valor_recursando_input).send_keys(valor_recursar)
                                time.sleep(2)
                                self.driver.find_element(*self.mostrar_textarea_just).click()
                                time.sleep(2)
                                self.driver.find_element(*self.textarea_justificativa).send_keys(justificativa)
                                time.sleep(2)
                                self.driver.find_element(*self.fechar_justificativa).click()
                                time.sleep(2)
                                self.driver.find_element(*self.salvar_recurso).click()
                                time.sleep(2)
                                self.driver.find_element(*self.fechar_recurso).click()
                                break
                            else:
                                count += 1
                        except:
                            break

                self.driver.get('https://servicosonline.cassi.com.br/Prestador/RecursoRevisaoPagamento/TISS/DemonstrativoAnaliseContas/Index')

            else:
                tkinter.messagebox.showinfo( 'Demonstrativo Cassi' , 'Não foram encontrados resultados para a pesquisa' )

        tkinter.messagebox.showinfo( 'Recurso CASSI' , f"Recurso concluído!" )
        self.driver.quit()

#---------------------------------------------------------------------------------------------------------------
def recursar_cassi():
    try:
        pasta = filedialog.askdirectory()
        global url
        url = 'https://servicosonline.cassi.com.br/GASC/v2/Usuario/Login/Prestador'
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
            driver = webdriver.Chrome(service=servico, options=chrome_options, seleniumwire_options=options)
        except:
            driver = webdriver.Chrome(options=chrome_options, seleniumwire_options=options)

        login_page = Login(driver, url)
        login_page.open()

        login_page.exe_login(
            usuario = "00735860000173",
            senha = "amhpdf123"
        )
        caminho(driver, url).exe_caminho()
        Recursar(driver, url).recurso(pasta=pasta)
    
    except Exception as err:
        tkinter.messagebox.showerror("Automação", f"Ocorreu uma exceção não tratada. \n {err.__class__.__name__} - {err}")
        Pidgin.financeiroDemo(f"Ocorreu uma exceção não tratada. \n {err.__class__.__name__} - {err}")