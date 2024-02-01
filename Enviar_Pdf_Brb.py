import pandas as pd
import pyautogui
import time
import os
from abc import ABC
from tkinter import filedialog
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import tkinter

class PageElement(ABC):
    def __init__(self, driver, url=''):
        self.driver = driver
        self.url = url
    def open(self):
        self.driver.get(self.url)

class Login(PageElement):
    prestador_pj = (By.XPATH, '//*[@id="tipoAcesso"]/option[9]')
    usuario = (By.XPATH, '//*[@id="login-entry"]')
    senha = (By.XPATH, '//*[@id="password-entry"]')
    entrar = (By.XPATH, '//*[@id="BtnEntrar"]')

    def logar(self, usuario, senha):
        self.driver.implicitly_wait(30)
        self.driver.find_element(*self.prestador_pj).click()
        time.sleep(2)
        self.driver.find_element(*self.usuario).send_keys(usuario)
        time.sleep(2)
        self.driver.find_element(*self.senha).send_keys(senha)
        time.sleep(2)
        self.driver.find_element(*self.entrar).click()
        time.sleep(5)

class Caminho(PageElement):
    envio_xml = (By.XPATH, '//*[@id="menuPrincipal"]/div/div[4]/a')

    def exe_caminho(self):
        try:
            WebDriverWait(self.driver, 15).until(EC.presence_of_element_located((self.envio_xml))).click()
            time.sleep(2)
            
        except:
            self.driver.refresh()
            time.sleep(5)
            login_page.logar(usuario = '00735860000173', senha = 'amhpdf0073')
            WebDriverWait(self.driver, 30).until(EC.presence_of_element_located((self.envio_xml))).click()
    
class EnviarPdf(PageElement):
    pesquisar = (By.XPATH, '//*[@id="filtro"]/div[2]/div[2]/button[1]')
    botao_enviar_xml = (By.XPATH, '/html/body/main/div/div[5]/div/div/div[3]/button[1]')
    arquivo = (By.XPATH, '/html/body/main/div/div[4]/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/table/tbody/tr/td[1]/span')
    botao_sim = (By.XPATH, '//*[@id="button-0"]')
    lista_planilhas = []
    resultado_upload = (By.XPATH, '/html/body/main/div/div[5]/div/div/div[2]/div[5]/div/form/div')
    cem_itens = (By.XPATH, '//*[@id="div-lote-xml"]/div[2]/div/div[2]/div[11]/div[1]/div[2]/div/div/select/option[7]')
    caminho_arquivo = (By.XPATH, '/html/body/main/div/div[4]/div/div/div[2]/div/div/div/div[2]/div[1]/input-file/div/div/div[1]/div/div/input')
    proxima_pagina = (By.XPATH, '//*[@id="div-lote-xml"]/div[2]/div/div[2]/div[101]/div[1]/div[1]/div/nav/ul/li[8]/a')
    adicionar_arquivo = (By.XPATH, '/html/body/main/div/div[4]/div/div/div[2]/div/div/div/div[2]/div[1]/input-file/div/div/div[1]/div/div/div[1]/div[2]/button')
    lixeira = (By.XPATH, '/html/body/main/div/div[4]/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/table/thead/tr/th[1]/i')
    fechar = (By.XPATH, '/html/body/main/div/div[4]/div/div/div[3]/button')
    primeira_pagina = (By.XPATH, '//*[@id="div-lote-xml"]/div[2]/div/div[2]/div[101]/div[1]/div[1]/div/nav/ul/li[1]/a')
    corpo_envio_arquivo = (By.XPATH, '//*[@id="modal-arquivos"]/div/div')

    def arquivos(self):
        nomesarquivos = os.listdir(pasta)
        self.lista_planilhas = [f"{pasta}/{nome}" for nome in nomesarquivos]

    def enviar_pdf(self):
        self.arquivos()
        self.driver.find_element(*self.pesquisar).click()
        WebDriverWait(self.driver, 30).until(EC.presence_of_element_located((self.cem_itens))).click()
        WebDriverWait(self.driver, 30).until(EC.presence_of_element_located((By.XPATH, f'/html/body/main/div/div[2]/div/div[2]/div[1]/div[1]/div/div[5]/button[1]'))).send_keys(Keys.CONTROL + Keys.HOME)

        for planilha in self.lista_planilhas:
            if ".xls" in planilha and "lock" not in planilha:
                faturas_df = pd.read_excel(planilha, header=23)
                print(faturas_df)
                faturas_df = faturas_df.iloc[:-6]

            else:
                continue

            for index, linha in faturas_df.iterrows():
                processo_planilha = f'{linha["Nº Fatura"]}'.replace('.0', '')
                caminho_planilha = linha['Observações']
                guia_achada = False

                # if linha['Enviado'] == "Sim":
                #     continue

                for page in range(0, 5):
                    time.sleep(1)
                    print(f'Pagina {page + 1}')

                    for i in range(1,101):
                        self.driver.implicitly_wait(30)
                        slot = WebDriverWait(self.driver, 30).until(EC.presence_of_element_located((By.XPATH, f'/html/body/main/div/div[2]/div/div[2]/div[{i}]'))).text
                        
                        if processo_planilha in slot and "Arquivo com inconsistências" not in slot:
                            print(processo_planilha)
                            botao_anexo_encontrado = False
                            time.sleep(1)

                            while botao_anexo_encontrado == False:
                                try:
                                    anexo = self.driver.find_element(By.XPATH, f'/html/body/main/div/div[2]/div/div[2]/div[{i}]/div/div/div[5]/button[5]').click()
                                    time.sleep(1)
                                    break
                                except:
                                    self.driver.execute_script('scrollBy(0,100)')

                            arquivo_enviado = False
                            
                            try:
                                self.driver.implicitly_wait(5)
                                self.driver.find_element(*self.arquivo)
                                arquivo_enviado = True

                            except:
                                pass

                            if arquivo_enviado == True:
                                time.sleep(1)
                                self.driver.save_screenshot(f"{pasta}//{linha['Nº Fatura']}_enviado_anteriormente.png")
                                guia_achada = True
                                self.driver.find_element(*self.fechar).click()
                                time.sleep(1)
                                break

                            self.driver.implicitly_wait(30)
                            self.driver.find_element(*self.caminho_arquivo).send_keys(caminho_planilha)
                            time.sleep(0.5)
                            self.driver.find_element(*self.adicionar_arquivo).click()
                            time.sleep(0.5)
                            botao_lixeira_encontrado = False

                            while botao_lixeira_encontrado == False:
                                try:
                                    self.driver.find_element(*self.lixeira).click()
                                    botao_lixeira_encontrado = True
                                except:
                                    pass

                            self.driver.find_element(*self.lixeira).click()
                            self.driver.save_screenshot(f"{pasta}//{linha['Nº Fatura']}.png")
                            self.driver.find_element(*self.fechar).click()
                            guia_achada = True
                            break

                        elif processo_planilha in slot and "Arquivo com inconsistências" in slot:
                            print("Arquivo com inconsistências.")
                
                    if guia_achada == True:
                        WebDriverWait(self.driver, 30).until(EC.presence_of_element_located((By.XPATH, f'/html/body/main/div/div[2]/div/div[2]/div[{i}]/div[1]/div/div[5]/button[1]'))).send_keys(Keys.CONTROL + Keys.END)
                        time.sleep(1)
                        self.driver.find_element(*self.primeira_pagina).click()
                        WebDriverWait(self.driver, 30).until(EC.presence_of_element_located((By.XPATH, f'/html/body/main/div/div[2]/div/div[2]/div[{i}]/div[1]/div/div[5]/button[1]'))).send_keys(Keys.CONTROL + Keys.HOME)
                        user = False

                        while user == False:
                            try:
                                usuario = self.driver.find_element(By.XPATH, '//*[@id="menu_78B1E34CFC8E414D8EB4F83B534E4FB4"]').click()
                                user = True
                            except:
                                pass

                        break

                    elif guia_achada == False:
                        WebDriverWait(self.driver, 30).until(EC.presence_of_element_located((By.XPATH, f'/html/body/main/div/div[2]/div/div[2]/div[{i}]/div[1]/div/div[5]/button[1]'))).send_keys(Keys.CONTROL + Keys.END)
                        time.sleep(1)
                        self.driver.find_element(*self.proxima_pagina).click()
                        user = False

                        while user == False:
                            try:
                                usuario = self.driver.find_element(By.XPATH, '//*[@id="menu_78B1E34CFC8E414D8EB4F83B534E4FB4"]').click()
                                user = True
                            except:
                                pass
                        WebDriverWait(self.driver, 30).until(EC.presence_of_element_located((By.XPATH, f'/html/body/main/div/div[2]/div/div[2]/div[1]/div[1]/div/div[5]/button[1]'))).send_keys(Keys.CONTROL + Keys.HOME)

                #Se não achar o processo nas 5 páginas, irá cair nesse bloco de código.        
                WebDriverWait(self.driver, 30).until(EC.presence_of_element_located((By.XPATH, f'/html/body/main/div/div[2]/div/div[2]/div[{i}]/div[1]/div/div[5]/button[1]'))).send_keys(Keys.CONTROL + Keys.END)
                time.sleep(1)
                self.driver.find_element(*self.primeira_pagina).click()
                WebDriverWait(self.driver, 30).until(EC.presence_of_element_located((By.XPATH, f'/html/body/main/div/div[2]/div/div[2]/div[{i}]/div[1]/div/div[5]/button[1]'))).send_keys(Keys.CONTROL + Keys.HOME)
                user = False

                while user == False:
                    try:
                        usuario = self.driver.find_element(By.XPATH, '//*[@id="menu_78B1E34CFC8E414D8EB4F83B534E4FB4"]').click()
                        user = True
                    except:
                        pass

#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
def enviar_pdf(user, password):
    try:
        global pasta, login_page

        pasta = filedialog.askdirectory()

        url = 'https://portal.saudebrb.com.br/GuiasTISS/Logon'

        options = {
        'proxy': {
                'http': f'http://{user}:{password}@10.0.0.230:3128',
                'https': f'https://{user}:{password}@10.0.0.230:3128'
            }
        }

        chrome_options = Options()

        chrome_options.add_argument("--start-maximized")
        # chrome_options.add_argument('--ignore-certificate-errors')
        # chrome_options.add_argument('--ignore-ssl-errors')

        driver = webdriver.Chrome(options=chrome_options)

        login_page = Login(driver, url)
        login_page.open()
        driver.maximize_window()
        time.sleep(4)
        pyautogui.write(user)
        pyautogui.press("TAB")
        time.sleep(1)
        pyautogui.write(password)
        pyautogui.press("enter")
        time.sleep(4)

        login_page.logar(
            usuario = '00735860000173',
            senha = 'AMHP7860@'
            )

        Caminho(driver,url).exe_caminho()

        envio_xml = EnviarPdf(driver, url)
        envio_xml.enviar_pdf()
        tkinter.messagebox.showinfo( 'Automação' , f"Envios concluídos!" )
    except Exception as e:
        tkinter.messagebox.showinfo( 'Automação' , f"Ocorreu uma exceção não tratada.\n{e.__class__.__name__} - {e}" )