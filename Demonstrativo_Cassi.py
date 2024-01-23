import tkinter.messagebox
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
from datetime import datetime
import Pidgin
from xml.dom import minidom

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

class BaixarDemonstrativo(PageElement):
    data_inicial = (By.XPATH, '//*[@id="DataInicial"]')
    data_final = (By.XPATH, '//*[@id="DataFinal"]')
    consultar = (By.XPATH, '/html/body/div[1]/div[5]/section/div/form/fieldset/div[4]/div/button')
    xpath_tabela = (By.XPATH, '/html/body/div[1]/div[5]/section/div/fieldset/div/table')
    download_xml = (By.XPATH, '//*[@id="formExportar"]/button[2]')
    download_pdf = (By.XPATH, '//*[@id="formExportar"]/button[1]')
    voltar = (By.XPATH, '//*[@id="btnVoltar"]')
    xpath_corpo_da_pagina = (By.XPATH, '/html/body')

    def baixar_demontrativo(self, data_inicial, data_final):

        self.driver.find_element(*self.data_inicial).send_keys(data_inicial)
        time.sleep(2)
        self.driver.find_element(*self.data_final).send_keys(data_final)
        self.driver.find_element(*self.data_inicial).send_keys(Keys.ESCAPE)
        time.sleep(2)
        self.driver.find_element(*self.consultar).click()
        time.sleep(2)
        corpo_pagina = self.driver.find_element(*self.xpath_corpo_da_pagina).text

        if "Não foram encontrados resultados para a pesquisa" in corpo_pagina:
            tkinter.messagebox.showinfo( 'Demonstrativo Cassi' , 'Não foram encontrados resultados para a pesquisa' )

        else:
            tabela = self.driver.find_element(*self.xpath_tabela).get_attribute('outerHTML')
            df_tabela = pd.read_html(tabela)[0]
            quantidade_demonstrativos = len(df_tabela) + 1

            for index, linha in df_tabela.iterrows():
                numero_do_protocolo_df = str(linha['Número do Protocolo']).replace(".0","")

                for i in range(1, quantidade_demonstrativos):
                    numero_protocolo_portal = str(self.driver.find_element(By.XPATH, f'/html/body/div[1]/div[5]/section/div/fieldset/div/table/tbody/tr[{i}]/td[1]').text).replace(".0", "")

                    if numero_do_protocolo_df != numero_protocolo_portal:
                        continue
                    
                    time.sleep(2)
                    self.driver.find_element(By.XPATH, f'/html/body/div[1]/div[5]/section/div/fieldset/div/table/tbody/tr[{i}]/td[3]/form/input[3]').click()
                    time.sleep(2)
                    self.driver.find_element(*self.download_xml).click()
                    time.sleep(2)
                    count = 0

                    while count < 20:
                        try:
                            with open(f"\\\\10.0.0.239\\automacao_financeiro\\CASSI\\{numero_do_protocolo_df}.xml", "r", encoding="utf-8") as f:
                                valor_total_glosa = 0.00
                                xml = minidom.parse(f)
                                glosa = xml.getElementsByTagName("ans:valorGlosaGeral")

                                for tag in glosa:
                                    valor_total_glosa = tag.firstChild.data
                            break

                        except:
                            count += 1

                            if count == 18:
                                time.sleep(10)

                            else:
                                time.sleep(2)

                    f.close()    
                    
                    if valor_total_glosa != "0":
                        self.driver.find_element(*self.download_pdf).click()
                        data_atual = datetime.now().strftime('%d_%m_%Y')
                        endereco = r"\\10.0.0.239\automacao_financeiro\CASSI"
                        novo_nome = f"{endereco}\\{numero_do_protocolo_df}.pdf"
                        contador = 0

                        while contador < 20:
                            try:
                                os.rename(f"{endereco}\\Demonstrativo de Analise de Conta {data_atual}.pdf", novo_nome)
                                break
                            
                            except:
                                print("Download ainda não foi feito")
                                contador += 1

                                if contador == 18:
                                    time.sleep(10)
                                
                                else:
                                    time.sleep(2)

                    time.sleep(2)
                    self.driver.find_element(*self.voltar).click()
                    time.sleep(2)
                    self.driver.refresh()
                    break

            tkinter.messagebox.showinfo( 'Demonstrativos CASSI' , f"Downloads concluídos!" )
            self.driver.quit()

#---------------------------------------------------------------------------------------------------------------
def demonstrativo_cassi(data_inicial, data_final, user, password):
    try:

        global url
        url = 'https://servicosonline.cassi.com.br/GASC/v2/Usuario/Login/Prestador'

        settings = {
        "recentDestinations": [{
                "id": "Save as PDF",
                "origin": "local",
                "account": "",
            }],
            "selectedDestinationId": "Save as PDF",
            "version": 2
        }

        chrome_options = Options()
        chrome_options.add_experimental_option('prefs', {
            "download.default_directory": r"\\10.0.0.239\automacao_financeiro\CASSI",
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True,
            'safebrowsing.enabled': 'false'
    })
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument('--ignore-certificate-errors')
        chrome_options.add_argument('--ignore-ssl-errors')
        chrome_options.add_argument('--kiosk-printing')

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

        login_page = Login(driver, url)
        login_page.open()

        login_page.exe_login(
            usuario = "00735860000173",
            senha = "amhpdf123"
        )
        caminho(driver, url).exe_caminho()
        BaixarDemonstrativo(driver, url).baixar_demontrativo(data_inicial, data_final)
    
    except Exception as err:
        tkinter.messagebox.showerror("Automação", f"Ocorreu uma exceção não tratada. \n {err.__class__.__name__} - {err}")
        Pidgin.financeiroDemo(f"Ocorreu uma exceção não tratada. \n {err.__class__.__name__} - {err}")
