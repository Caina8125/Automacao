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
import json
import os

class PageElement(ABC):
    def __init__(self, driver, url=''):
        self.driver = driver
        self.url = url
    def open(self):
        self.driver.get(self.url)

class Login(PageElement):
    prestador_pj = (By.XPATH, '//*[@id="tipoAcesso"]/option[7]')
    usuario = (By.XPATH, '//*[@id="login-entry"]')
    senha = (By.XPATH, '//*[@id="password-entry"]')
    entrar = (By.XPATH, '//*[@id="BtnEntrar"]')

    def exe_login(self, usuario, senha):
        self.driver.implicitly_wait(30)
        self.driver.find_element(*self.prestador_pj).click()
        self.driver.find_element(*self.usuario).send_keys(usuario)
        self.driver.find_element(*self.senha).send_keys(senha)
        self.driver.find_element(*self.entrar).click()
        time.sleep(4)

class Caminho(PageElement):
    faturas = (By.XPATH, '//*[@id="menuPrincipal"]/div/div[12]/a')
    relatorio_de_faturas = (By.XPATH, '/html/body/header/div[4]/div/div/div/div[12]/div[1]/div[2]/div/div[2]/div/div/div/div[1]/a')
    somente_com_glosa = (By.XPATH, '//*[@id="chkSomenteComGlosa"]')

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
        self.driver.find_element(*self.somente_com_glosa).click()
        time.sleep(2)

class BaixarDemonstrativo(PageElement):
    codigo = (By.XPATH, '/html/body/main/div/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/input-text[1]/div/div/input')
    pesquisar = (By.XPATH, '//*[@id="filtro"]/div[2]/div[2]/button')
    ver_xml = (By.XPATH, '//*[@id="div-Servicos"]/div[1]/div[4]/div/div/div[1]/div/div[2]/a[2]')
    radio_button = (By.XPATH, '//*[@id="divEscolhaProtocoloXml"]/div/input')
    exportar_todos = (By.XPATH, '//*[@id="escolha-protocolo-modal"]/div/div/div[3]/button[2]')
    salvar = (By.XPATH, '//*[@id="btn-salxar-xml-servico"]')
    fechar = (By.XPATH, '//*[@id="operation-modal"]/div/div/div[3]/button[2]')
    botao_ok = (By.XPATH, '//*[@id="button-0"]')
    detalhes_da_fatura = (By.XPATH, '/html/body/main/div/div[1]/div[4]/div/div/div[1]/div/div[2]/a[1]/i')
    relatorio_de_servico = (By.XPATH, '/html/body/main/div/div[1]/div[4]/div/div/div[3]/div[2]/div[1]/div[2]/input[4]')

    def baixar_demonstrativo(self, planilha):
        df = pd.read_excel(planilha, header=5)
        df = df.iloc[:-1]
        df = df.dropna()
        df['Concluído'] = ''
        count = 0
        quantidade_de_faturas = len(df)
        erro_portal = False

        for tentativa in range(1, 6):
            print(f'Tentativa {tentativa}')

            try:
                for index, linha in df.iterrows():

                    if df['Concluído'][index] == "Sim":
                        continue
                    
                    numero_fatura = f"{linha['Nº Fatura']}".replace(".0", "")
                    numero_protocolo = f"{linha['Nº do Protocolo']}".replace(".0", "")
                    print(numero_fatura)
                    self.driver.implicitly_wait(30)
                    time.sleep(1.5)
                    self.driver.find_element(*self.codigo).send_keys(numero_protocolo)
                    time.sleep(1.5)
                    self.driver.find_element(*self.pesquisar).click()
                    time.sleep(1.5)

                    try:
                        self.driver.implicitly_wait(10)
                        self.driver.find_element(*self.ver_xml).click()

                    except:
                        self.driver.find_element(*self.botao_ok).click()
                        time.sleep(1.5)
                        self.driver.find_element(*self.codigo).clear()
                        df.loc[index, 'Concluído'] = 'Sim'
                        continue

                    self.driver.implicitly_wait(30)
                    time.sleep(1.5)
                    # self.driver.find_element(*self.radio_button).click()
                    # time.sleep(1.5)
                    self.driver.find_element(*self.exportar_todos).click()
                    time.sleep(1.5)
                    self.driver.implicitly_wait(180)
                    time.sleep(1.5)
                    self.driver.find_element(*self.salvar).click()
                    time.sleep(4)
                    self.driver.implicitly_wait(30)
                    print(f"Download do xml da fatura {numero_fatura} concluído com sucesso")
                    time.sleep(3)
                    self.driver.find_element(*self.fechar).click()
                    time.sleep(2)
                    self.driver.find_element(*self.detalhes_da_fatura).click()
                    time.sleep(1.5)
                    self.driver.find_element(*self.relatorio_de_servico).click()

                    novo_nome = r"\\10.0.0.239\automacao_financeiro\FASCAL" + f"\\{numero_fatura}.pdf"
                    lista_faturas_com_erro = []
                    download_feito = False
                    endereco = r"\\10.0.0.239\automacao_financeiro\FASCAL"

                    for i in range(10):

                        try:
                            os.rename(f"{endereco}\\RelatorioServicos_{numero_protocolo}.pdf", novo_nome)
                            download_feito = True
                            break

                        except:
                            print("Download ainda não foi feito")

                            if i == 9:
                                erro_portal = True
                                self.driver.quit()

                            time.sleep(2)

                    if download_feito == True:
                        count += 1
                        print(f"Download da fatura {numero_fatura} concluído com sucesso")

                    else:
                        print(f"Download da fatura {numero_fatura} não foi feito ou não foi possível renomear.")
                        lista_faturas_com_erro.append(numero_fatura)

                    df.loc[index, 'Concluído'] = 'Sim'

                    print('---------------------------------------------------------------')

                    time.sleep(1)
                    self.driver.find_element(*self.codigo).clear()
        
                tkinter.messagebox.showinfo( 'Demonstrativos Fascal' , f"Downloads concluídos: {count} de {quantidade_de_faturas}." )
                self.driver.quit()
                break

            except Exception as err:
                if erro_portal == True:
                    tkinter.messagebox.showerror("Automação", "Ocorreu algum erro ao fazer download no site.")

                print(err)
                self.driver.get(self.url)
                login_page.exe_login(usuario, senha)
                caminho.exe_caminho()



#--------------------------------------------------------------------------------------------------------------------

def demonstrativo_fascal():
    try:
        url = 'https://novowebplanfascal.facilinformatica.com.br/GuiasTISS/Logon'
        planilha = filedialog.askopenfilename()

        options = {
            'proxy' : {
                'http': 'http://lucas.paz:RDRsoda90901@@10.0.0.230:3128',
                'https': 'http://lucas.paz:RDRsoda90901@@10.0.0.230:3128'
            }
        }

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
            "printing.print_to_pdf": True,
            "download.default_directory": r"\\10.0.0.239\automacao_financeiro\FASCAL",
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True,
            "printing.print_preview_sticky_settings.appState": json.dumps(settings),
            "savefile.default_directory": r"\\10.0.0.239\automacao_financeiro\FASCAÇ"
    })
        chrome_options.add_argument("--start-maximized")

        try:
            driver = webdriver.Chrome(service=servico, seleniumwire_options= options, options = chrome_options)
            servico = Service(ChromeDriverManager().install())

        except:
            driver = webdriver.Chrome(seleniumwire_options= options, options = chrome_options)

        global usuario, senha, login_page, caminho

        usuario = "AMHPDF-ADM"
        senha = "00735860000173"

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