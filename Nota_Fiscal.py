import time
import pyautogui
import time
import Pidgin
from abc import ABC
import pandas as pd
from selenium import webdriver
from tkinter import filedialog
from openpyxl import load_workbook
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

class PageElement(ABC):
    def __init__(self, driver, url='') -> None:
        self.driver = driver
        self.url = url
    def open(self):
        self.driver.get(self.url)


class Login(PageElement):
    logarCertificado = (By.XPATH, '//*[@id="btnAcionaCertificado"]')
    def exe_login(self):
        self.driver.find_element(*self.logarCertificado).click()

class Caminho(PageElement):
    declararServico = (By.XPATH, '//*[@id="Menu1_MenuPrincipal"]/ul/li[3]/div/span[3]')
    incluir = (By.XPATH, '//*[@id="Menu1_MenuPrincipal"]/ul/li[3]/ul/li[1]/div')
    fecharModal = (By.XPATH, '//*[@id="base-modal"]/div/div/div[1]/button')

    def exe_caminho(self):
        time.sleep(2)
        self.driver.find_element(*self.declararServico).click()
        time.sleep(2)
        self.driver.find_element(*self.incluir).click()
        time.sleep(3)
        self.driver.switch_to.frame('iframe')
        time.sleep(2)
        self.driver.find_element(*self.fecharModal).click()
        time.sleep(2)


class Nf(PageElement):
    inserirCNPJ = (By.XPATH, '//*[@id="dgContratados__ctl2_txtCPF_CNPJ"]')
    inserirNumDoc = (By.XPATH, '//*[@id="dgContratados__ctl2_txtNum_Doc"]')
    botaoGravar = (By.XPATH, '//*[@id="btnGravar"]')
    campoVlDoc = (By.XPATH, '//*[@id="dgContratados__ctl2_txtValor_Doc"]')
    fecharModalErro = (By.XPATH, '//*[@id="base-modal"]/div/div/div[1]/button')
    botaoCancelar = (By.XPATH, '//*[@id="Button4"]')
    # //*[@id="TxtErro"]

    def inserirDadosNf(self):
        faturas_df = pd.read_excel(planilha)
        for index, linha in faturas_df.iterrows():
            cnpj = str(f"{linha['CNPJCPF']}")
            nf = str(f"{linha['NFENUMERO']}")

            while not len(cnpj) == 14:
                cnpj = "0" + cnpj
            
            time.sleep(1)
            self.driver.find_element(*self.inserirCNPJ).send_keys(cnpj)
            self.driver.find_element(*self.inserirNumDoc).click()
            time.sleep(2)
            self.driver.find_element(*self.inserirNumDoc).send_keys(nf)
            self.driver.find_element(*self.campoVlDoc).click()
            time.sleep(2)
            self.driver.find_element(*self.botaoGravar).click()
            time.sleep(2)
            try:
                self.driver.switch_to.frame('iframeModal')
                erro = self.driver.find_element(By.ID, "TxtErro").text
                Pidgin.notaFiscal(f"Teste Robô => Erro ao grava NF: {erro}   Número NF: {linha['NFENUMERO']}")
                self.driver.switch_to.default_content()
                self.driver.switch_to.frame('iframe')
                self.driver.find_element(*self.fecharModalErro).click()
                self.driver.find_element(*self.botaoCancelar).click()
            except:
                pass

#---------------------------------------------------------------------------------------------------------------------------------------------

def subirNF():
    try:
        global planilha
        planilha = filedialog.askopenfilename()


        url = 'https://www.amhp.com.br/'
        refresh = "https://df.issnetonline.com.br/online/Login/Login.aspx#"
        driver = webdriver.Chrome()

        login_page = Login(driver, url)
        driver.maximize_window()
        driver.get(url)

        pyautogui.write('lucas.timoteo')
        pyautogui.press("TAB")
        pyautogui.write('Caina8125')
        pyautogui.press("enter")

        driver.get(refresh)
        login_page.exe_login()
        time.sleep(2)
        pyautogui.press("enter")
        time.sleep(2)

        Caminho(driver,url).exe_caminho()
        time.sleep(1)
        Nf(driver,url).inserirDadosNf()
    except:
        driver.quit()


