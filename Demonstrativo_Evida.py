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
    iframe = (By.XPATH, '/html/body/div[6]/iframe[4]')
    fechar_chat = (By.XPATH, '/html/body/div/div/div/div[2]')

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
        id = self.driver.find_element(*self.iframe).get_attribute("id")
        self.driver.switch_to.frame(id)
        self.driver.find_element(*self.fechar_chat).click()
        time.sleep(1)
        self.driver.switch_to.default_content()

class BaixarDemonstrativo(PageElement):
    lote = (By.XPATH, '//*[@id="txtLote"]')
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

        for tentativa in range(1, 6):
            print(f'Tentativa {tentativa}')

            try:
                for index, linha in df.iterrows():
                    numero_fatura = str(linha['Nº Fatura']).replace('.0', '')

                    if df['Concluído'][index] == "Sim":
                        continue

                    print(numero_fatura)
                    self.driver.implicitly_wait(30)
                    time.sleep(1.5)
                    self.driver.find_element(*self.lote).send_keys(numero_fatura)
                    time.sleep(1.5)
                    self.driver.find_element(*self.pesquisar).click()
                    time.sleep(1.5)
                    ver_xml_clicado = False

                    for k in range(5):
                        try:
                            self.driver.implicitly_wait(10)
                            self.driver.find_element(*self.ver_xml).click()
                            ver_xml_clicado = True
                            break
                        except:
                            time.sleep(2)
                                
                    if ver_xml_clicado == False:
                        self.driver.find_element(*self.botao_ok).click()
                        time.sleep(1.5)
                        self.driver.find_element(*self.lote).clear()
                        df.loc[index, 'Concluído'] = 'Sim'
                        continue
                    
                    self.driver.implicitly_wait(30)
                    time.sleep(1.5)
                    self.driver.find_element(*self.radio_button).click()
                    time.sleep(1.5)
                    self.driver.find_element(*self.exportar_todos).click()
                    time.sleep(1.5)
                    self.driver.implicitly_wait(30)
                    time.sleep(1.5)
                    self.driver.find_element(*self.salvar).click()
                    time.sleep(4)
                    self.driver.find_element(*self.fechar).click()
                    time.sleep(1.5)
                    self.driver.find_element(*self.detalhes_da_fatura).click()
                    time.sleep(1.5)
                    codigo = self.driver.find_element(By.XPATH, '//*[@id="div-Servicos"]/div[1]/div[4]/div/div/div[1]/div/div[1]/div[1]/div[2]/span').text

                    for j in range(0, 10):
                        try:
                            self.driver.find_element(*self.relatorio_de_servico).click()
                            break
                        except:
                            time.sleep(2)
                            continue

                    novo_nome = r"\\10.0.0.239\automacao_financeiro\E-VIDA" + f"\\{numero_fatura}.pdf"
                    lista_faturas_com_erro = []
                    download_feito = False
                    endereco = r"\\10.0.0.239\automacao_financeiro\E-VIDA"

                    for i in range(10):

                        try:
                            os.rename(f"{endereco}\\RelatorioServicos_{codigo}.pdf", novo_nome)
                            download_feito = True
                            break

                        except:
                            print("Download ainda não foi feito")

                            if i == 9:
                                erro_portal = True
                                self.driver.quit()

                            time.sleep(2)
                    count += 1
                    print(f"Download do XML da fatura {numero_fatura} concluído com sucesso")

                    df.loc[index, 'Concluído'] = 'Sim'

                    print('---------------------------------------------------------------')
                    self.driver.find_element(*self.lote).clear()
                    time.sleep(2)
                
                tkinter.messagebox.showinfo( 'Demonstrativos E-Vida' , f"Downloads concluídos: {count} de {quantidade_de_faturas}." )

                self.driver.quit()

                break
            
            except Exception as err:
                print(err)
                self.driver.get(self.url)
                login_page.exe_login(usuario, senha)
                caminho.exe_caminho()
                
#--------------------------------------------------------------------------------------------------------------------

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
        chrome_options.add_experimental_option('prefs', {
            "download.default_directory": r"\\10.0.0.239\automacao_financeiro\E-VIDA",
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True,
            "safebrowsing.enabled": 'false',
            "safebrowsing.disable_download_protection,": True,
            "safebrowsing_for_trusted_sources_enabled": False,
    })
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

        caminho = Caminho(driver, url)
        caminho.exe_caminho()
        BaixarDemonstrativo(driver, url).baixar_demonstrativo(planilha)

    except FileNotFoundError as err:
        tkinter.messagebox.showerror('Automação', f'Nenhuma planilha foi selecionada!')
    
    except Exception as err:
        tkinter.messagebox.showerror("Automação", f"Ocorreu uma exceção não tratada. \n {err.__class__.__name__} - {err}")
        Pidgin.financeiroDemo(f"Ocorreu uma exceção não tratada. \n {err.__class__.__name__} - {err}")
    driver.quit()