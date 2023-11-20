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
import tkinter
import Pidgin
import tkinter.messagebox

class PageElement(ABC):
    def __init__(self, driver, url=''):
        self.driver = driver
        self.url = url
    def open(self):
        self.driver.get(self.url)

class Login(PageElement):
    medico = (By.XPATH, '//*[@id="tipoAcesso"]/option[9]')
    usuario = (By.XPATH, '//*[@id="login-entry"]')
    senha = (By.XPATH, '//*[@id="password-entry"]')
    entrar = (By.XPATH, '//*[@id="BtnEntrar"]')

    def exe_login(self, usuario, senha):
        self.driver.implicitly_wait(30)
        self.driver.find_element(*self.medico).click()
        self.driver.find_element(*self.usuario).send_keys(usuario)
        self.driver.find_element(*self.senha).send_keys(senha)
        self.driver.find_element(*self.entrar).click()
        time.sleep(4)

class Caminho(PageElement):
    faturas = (By.XPATH, '//*[@id="menuPrincipal"]/div/div[7]/a')
    relatorio_de_faturas = (By.XPATH, '/html/body/header/div[4]/div/div/div/div[7]/div[1]/div[2]/div/div[2]/div/div/div/div[1]/a')

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
        self.driver.find_element(*self.faturas).click()
        time.sleep(2)
        self.driver.find_element(*self.relatorio_de_faturas).click()
        time.sleep(2)

class BaixarDemonstrativo(PageElement):
    lote = (By.XPATH, '//*[@id="txtLote"]')
    pesquisar = (By.XPATH, '//*[@id="filtro"]/div[2]/div[2]/button')
    ver_xml = (By.XPATH, '//*[@id="div-Servicos"]/div[1]/div[4]/div/div/div[1]/div/div[2]/a[2]')
    radio_button = (By.XPATH, '//*[@id="divEscolhaProtocoloXml"]/div/input')
    exportar_todos = (By.XPATH, '//*[@id="escolha-protocolo-modal"]/div/div/div[3]/button[2]')
    salvar = (By.XPATH, '//*[@id="btn-salxar-xml-servico"]')
    fechar = (By.XPATH, '//*[@id="operation-modal"]/div/div/div[3]/button[2]')
    botao_ok = (By.XPATH, '//*[@id="button-0"]')


    def baixar_demonstrativo(self, planilha):
        df = pd.read_excel(planilha, header=5)
        df = df.iloc[:-1]
        df = df.dropna()
        df['Concluído'] = ''
        count = 0
        quantidade_de_faturas = len(df)

        for tentativa in range(1, 6):
            print(f'Tentativa {tentativa}')

            try:
                for index, linha in df.iterrows():

                    if df['Concluído'][index] == "Sim":
                        continue
                    
                    numero_fatura = f"{linha['Nº Fatura']}".replace(".0", "")
                    print(numero_fatura)
                    self.driver.implicitly_wait(30)
                    time.sleep(1.5)
                    self.driver.find_element(*self.lote).send_keys(numero_fatura)
                    time.sleep(1.5)
                    self.driver.find_element(*self.pesquisar).click()
                    time.sleep(1.5)
                    try:
                        self.driver.implicitly_wait(10)
                        self.driver.find_element(*self.ver_xml).click()
                    except:
                        self.driver.find_element(*self.botao_ok).click()
                        time.sleep(1.5)
                        self.driver.find_element(*self.lote).clear()
                        df.loc[index, 'Concluído'] = 'Sim'
                        continue
                    
                    self.driver.implicitly_wait(30)
                    self.driver.find_element(*self.radio_button).click()
                    time.sleep(1.5)

                    time.sleep(1.5)
                    self.driver.find_element(*self.exportar_todos).click()
                    time.sleep(1.5)
                    self.driver.find_element(*self.salvar).click()
                    time.sleep(4)

                    count += 1
                    print(f"Download da fatura {numero_fatura} concluído com sucesso")

                    df.loc[index, 'Concluído'] = 'Sim'

                    print('---------------------------------------------------------------')

                    self.driver.find_element(*self.fechar).click()
                    time.sleep(1)
                    self.driver.find_element(*self.lote).clear()
        
                tkinter.messagebox.showinfo( 'Demonstrativos BRB' , f"Downloads concluídos: {count} de {quantidade_de_faturas}." )
                self.driver.quit()
                break

            except Exception as err:
                print(err)
                self.driver.get(self.url)
                login_page.exe_login(usuario, senha)
                caminho.exe_caminho()


#--------------------------------------------------------------------------------------------------------------------

def demonstrativo_brb():
    try:
        url = 'https://portal.saudebrb.com.br/GuiasTISS/Logon'
        planilha = filedialog.askopenfilename()

        options = {
            'proxy' : {
                'http': 'http://lucas.paz:RDRsoda90901@@10.0.0.230:3128',
                'https': 'http://lucas.paz:RDRsoda90901@@10.0.0.230:3128'
            }
        }

        chrome_options = Options()
        chrome_options.add_experimental_option('prefs', {
            "download.default_directory": r"\\10.0.0.239\automacao_financeiro\BRB",
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True
    })
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument('--ignore-certificate-errors')
        chrome_options.add_argument('--ignore-ssl-errors')
        chrome_options.add_argument('--kiosk-printing')
        try:
            servico = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=servico, seleniumwire_options= options, options = chrome_options)
        except:
            driver = webdriver.Chrome(seleniumwire_options= options, options = chrome_options)

        global usuario, senha, login_page, caminho

        usuario = "00735860000173"
        senha = "amhpdf0073"

        login_page = Login(driver, url)
        login_page.open()
        login_page.exe_login(usuario, senha)

        caminho = Caminho(driver, url)
        caminho.exe_caminho()
        BaixarDemonstrativo(driver, url).baixar_demonstrativo(planilha)

    except FileNotFoundError as err:
        tkinter.messagebox.showerror('Automação', f'Nenhuma planilha foi selecionada!')
    
    except Exception as err:
        tkinter.messagebox.showerror("Automação", f"Ocorreu uma exceção não tratada. \n {err.__class__.__name__} - {err}")
        Pidgin.financeiroDemo(f"Ocorreu uma exceção não tratada. \n {err.__class__.__name__} - {err}")
    driver.quit()