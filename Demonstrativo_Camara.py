import os
import time
import shutil
import tkinter
import pandas as pd
from abc import ABC
from datetime import datetime
from selenium import webdriver
from tkinter import filedialog
from seleniumwire import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


class PageElement(ABC):
    def __init__(self, driver, url=''):
        self.driver = driver
        self.url = url
    def open(self):
        self.driver.get(self.url)


class Login(PageElement):
    email = (By.XPATH, '//*[@id="Email"]')
    senha = (By.XPATH, '//*[@id="Senha"]')
    logar = (By.XPATH, '//*[@id="btnLogin"]')

    def exe_login(self, email, senha):
        # caminho().Alert()
        self.driver.find_element(*self.email).send_keys(email)
        self.driver.find_element(*self.senha).send_keys(senha)
        self.driver.find_element(*self.logar).click()

class caminho(PageElement):
    demonstrativo = (By.XPATH, '//*[@id="sidebar-menu"]/li[24]/a/span[1]')
    analise_conta = (By.XPATH, '//*[@id="sidebar-menu"]/li[24]/ul/li[3]/a/span')
    selecionar_convenio = (By.XPATH, '//*[@id="s2id_OperadorasCredenciadas_HandleOperadoraSelected"]/a/span[2]/b')
    opcao_camara = (By.XPATH, '/html/body/div[14]/ul/li[2]/div')
    inserir_protocolo = (By.XPATH, '//*[@id="Protocolo"]')
    baixar_demonstrativo = (By.XPATH, '//*[@id="btn-Baixar_Demonstrativo"]')
    fechar_botao = (By.XPATH, '//*[@id="bcInformativosModal"]/div/div/div[3]/button[2]')
    fechar_alerta = (By.XPATH, '/html/body/bc-modal-evolution/div/div/div/div[3]/button[3]')

    def exe_caminho(self):
        time.sleep(1)
        self.driver.find_element(*self.demonstrativo).click()
        time.sleep(1)
        self.driver.find_element(*self.analise_conta).click()
        time.sleep(2)
        caminho(driver, url).Alert()
        self.driver.find_element(*self.selecionar_convenio).click()
        time.sleep(1)
        self.driver.find_element(*self.opcao_camara).click()
        time.sleep(1)

    def buscar_demonstrativo(self):
        df = pd.read_excel(planilha, header=5)
        df = df.iloc[:-1]
        df = df.dropna()

        for index, linha in df.iterrows():

            global fatura
            try:
                protocolo =  f"{linha['Nº do Protocolo']}".replace(".0","")
            except:
                protocolo =  f"{linha['Nº do Protocolo']}"
            try:
                fatura =  f"{linha['Nº Fatura']}".replace(".0","")
            except:
                fatura =  f"{linha['Nº Fatura']}"

            self.driver.find_element(*self.inserir_protocolo).send_keys(protocolo)
            time.sleep(1)
            self.driver.find_element(*self.baixar_demonstrativo).click()
            time.sleep(4)
            caminho(driver, url).movePath()
            time.sleep(2)
            self.driver.find_element(*self.inserir_protocolo).clear()
            time.sleep(1)





    def movePath(self):
        for i in range(10):
            pasta = r"\\10.0.0.239\automacao_financeiro\CAMARA\Renomear"
            nomes_arquivos = os.listdir(pasta)
            time.sleep(2)
            for nome in nomes_arquivos:
                nomepdf = os.path.join(pasta, nome)

            renomear = r"\\10.0.0.239\automacao_financeiro\CAMARA\Renomear" +f"\\{fatura}"  +  ".pdf"
            arqDest = r"\\10.0.0.239\automacao_financeiro\CAMARA" + f"\\{fatura}"  +  ".pdf"

            try:
                os.rename(nomepdf,renomear)
                shutil.move(renomear,arqDest)
                time.sleep(2)
                print("Arquivo renomeado e guardado com sucesso")
                break


            except Exception as e:
                print(e)
                print("Download ainda não foi feito/Arquivo não renomeado")
                time.sleep(2)



    def Alert(self):
        try:
            modal = WebDriverWait(self.driver, 3.0).until(EC.presence_of_element_located((By.XPATH, '//*[@id="bcInformativosModal"]/div/div')))
            self.driver.find_element(*self.fechar_botao).click()

            while True:
                try:
                    proximo_botao = WebDriverWait(self.driver, 0.2).until(EC.presence_of_element_located((By.XPATH, '//*[@id="bcInformativosModal"]/div/div/div[3]/button[2]')))
                    proximo_botao.click()
                except:
                    break

            try:
                fechar_botao = WebDriverWait(self.driver, 0.2).until(EC.presence_of_element_located((By.XPATH, '/html/body/bc-modal-evolution/div/div/div/div[3]/button[3]')))
                fechar_botao.click()
            except:
                print("Não foi possível encontrar o botão de fechar.")
                pass

        except:
            print("Não teve Modal")
            pass

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
def demonstrativo_camara():
    global planilha

    try:
        global pasta
        global driver
        global url

        edge_options = Options()

        edge_options.add_experimental_option('prefs', { "download.default_directory": r"\\10.0.0.239\automacao_financeiro\CAMARA\Renomear",
                                                "download.prompt_for_download": False,
                                                "download.directory_upgrade": True,
                                                "plugins.always_open_pdf_externally": True
                                                })
        
        edge_options.add_argument("--start-maximized")
        edge_options.add_argument('--ignore-certificate-errors')
        edge_options.add_argument('--ignore-ssl-errors')
        edge_options.add_argument('--kiosk-printing')

        url = 'https://portalconectasaude.com.br/Account/Login'
        
        planilha = filedialog.askopenfilename()

        proxy = {
        'proxy': {
                'http': 'http://lucas.paz:Gsw2022&@10.0.0.230:3128',
                'https': 'http://lucas.paz:Gsw2022&@10.0.0.230:3128'
            }
        }

        driver = webdriver.Edge(seleniumwire_options=proxy, options=edge_options)

    except:
        tkinter.messagebox.showerror( 'Erro Automação' , 'Ocorreu um erro inesperado' )

    try:
        login_page = Login(driver, url)
        login_page.open()
        login_page.exe_login(
            email="negociacao.gerencia@amhp.com.br",
            senha="Amhp@0073"
        )

        print('Pegar Alerta Acionado!')
        caminho(driver, url).Alert()

        
        caminho(driver, url).exe_caminho()

        caminho(driver, url).buscar_demonstrativo()

    except:
        tkinter.messagebox.showerror( 'Erro Automação' , 'Ocorreu um erro enquanto o Robô trabalhava, provavelmente o portal da Benner caiu 😢' )
    driver.quit()