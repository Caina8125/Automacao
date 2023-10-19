import time
import pyautogui
import time
import Pidgin
from abc import ABC
import pandas as pd
from datetime import datetime
from selenium import webdriver
from tkinter import filedialog
from openpyxl import load_workbook
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
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
    botaoMenu = (By.XPATH, '//*[@id="menu-toggle"]')
    

    def exe_caminho(self):
        
        time.sleep(2)
        
        try:
            self.driver.find_element(*self.declararServico).click()
            time.sleep(2)
            self.driver.find_element(*self.incluir).click()
            time.sleep(3)
        except:
            self.driver.find_element(*self.botaoMenu).click()
            time.sleep(1)
            try:
                self.driver.find_element(*self.declararServico).click()
                time.sleep(2)
                self.driver.find_element(*self.incluir).click()
            except:
                time.sleep(2)
                self.driver.find_element(*self.incluir).click()

        self.driver.switch_to.frame('iframe')
        time.sleep(2)
        self.driver.find_element(*self.fecharModal).click()
        time.sleep(2)
        self.driver.switch_to.default_content()



class Nf(PageElement):
    inserirCNPJ     = (By.XPATH, '//*[@id="dgContratados__ctl2_txtCPF_CNPJ"]')
    inserirNumDoc   = (By.XPATH, '//*[@id="dgContratados__ctl2_txtNum_Doc"]')
    botaoGravar     = (By.XPATH, '//*[@id="btnGravar"]')
    campoVlDoc      = (By.XPATH, '//*[@id="dgContratados__ctl2_txtValor_Doc"]')
    fecharModalErro = (By.XPATH, '//*[@id="base-modal"]/div/div/div[1]/button')
    botaoCancelar   = (By.XPATH, '//*[@id="Button4"]')
    selectMes       = (By.XPATH, '//*[@id="ddlMes"]')
    janeiro         = (By.XPATH, '//*[@id="ddlMes"]/option[2]')
    fevereiro       = (By.XPATH, '//*[@id="ddlMes"]/option[3]')
    marco           = (By.XPATH, '//*[@id="ddlMes"]/option[4]')
    abril           = (By.XPATH, '//*[@id="ddlMes"]/option[5]')
    maio            = (By.XPATH, '//*[@id="ddlMes"]/option[6]')
    junho           = (By.XPATH, '//*[@id="ddlMes"]/option[7]')
    julho           = (By.XPATH, '//*[@id="ddlMes"]/option[8]')
    agosto          = (By.XPATH, '//*[@id="ddlMes"]/option[9]')
    setembro        = (By.XPATH, '//*[@id="ddlMes"]/option[10]')
    outubro         = (By.XPATH, '//*[@id="ddlMes"]/option[11]')
    novembro        = (By.XPATH, '//*[@id="ddlMes"]/option[12]')
    dezembro        = (By.XPATH, '//*[@id="ddlMes"]/option[13]')
    confirmarMes    = (By.XPATH, '//*[@id="btnAlterarCompetencia"]')

    def alterarCopetencia(self):
        driver.get("https://df.issnetonline.com.br/online/Default/alterar_competencia.aspx")
        time.sleep(1)
        self.driver.find_element(*self.selectMes).click()
        time.sleep(1)
        self.verificarMesCompetente(mesPorExtenso)
        time.sleep(1)
        self.driver.find_element(*self.confirmarMes).click()
        time.sleep(1)
        Caminho(driver,url).exe_caminho()


    def verificarMesCompetente(self,mesPlanilha):
        match mesPlanilha:
            case "Janeiro":
                self.driver.find_element(*self.janeiro).click()
            case "Fevereiro":
                self.driver.find_element(*self.fevereiro).click()
            case "Março":
                self.driver.find_element(*self.marco).click()
            case "Abril":
                self.driver.find_element(*self.abril).click()
            case "Maio":
                self.driver.find_element(*self.maio).click()
            case "Junho":
                self.driver.find_element(*self.junho).click()
            case "Julho":
                self.driver.find_element(*self.julho).click()
            case "Agosto":
                self.driver.find_element(*self.agosto).click()
            case "Setembro":
                self.driver.find_element(*self.setembro).click()
            case "Outubro":
                self.driver.find_element(*self.outubro).click()
            case "Novembro":
                self.driver.find_element(*self.novembro).click()
            case "Dezembro":
                self.driver.find_element(*self.dezembro).click()



    def inserirDadosNf(self):
        global mesPorExtenso
        faturas_df = pd.read_excel(planilha)
        for index, linha in faturas_df.iterrows():
            #Tratando Data Copetencia
            dataCompetencia = f"{linha['NFECOMPETENCIA']}"
            dataCompetencia = dataCompetencia.replace(" 00:00:00","")

            # Verificando se Data copetencia está vazio, se estiver pular para próxima
            if not '-' in dataCompetencia:
                continue

            # Converter dataCompetencia para DateTime
            convertData = datetime.strptime(dataCompetencia, '%Y-%m-%d').date()
            convertMes = convertData.month
            mesPorExtenso = self.mesCopetencia(convertMes)

            # Pegar o mês vigente do portal
            try:
                self.driver.switch_to.default_content()
                campoMesPortal = driver.find_element(By.XPATH, '/html/body/form/nav/div/div[2]/ul[2]/li[1]/a/span[1]').text
            except:
                campoMesPortal = driver.find_element(By.XPATH, '/html/body/form/nav/div/div[2]/ul[2]/li[1]/a/span[1]').text

            # Verificar se o mês vigente do portal é o mesmo da NF,
            # Se não for, altera no portal com a função alterarCopetencia()
            if mesPorExtenso in campoMesPortal:
                print("Mês vigente da NF")
            else:
                self.alterarCopetencia()

            # Alterar cnpj e nf para string
            cnpj = str(f"{linha['CNPJCPF']}")
            nf = str(f"{linha['NFENUMERO']}")

            # Inserir zero a esquerda quando o CNPJ não vir com 14 caracteres
            while not len(cnpj) == 14:
                cnpj = "0" + cnpj
            
            time.sleep(1)
            self.driver.switch_to.frame('iframe')
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
                Pidgin.notaFiscal(f"Erro ao grava NF: {erro}   Número NF: {linha['NFENUMERO']}")
                self.driver.switch_to.default_content()
                self.driver.switch_to.frame('iframe')
                self.driver.find_element(*self.fecharModalErro).click()
                self.driver.find_element(*self.botaoCancelar).click()
            except:
                pass

    def mesCopetencia(self,mes):
        match mes:
            case 1:
                return "Janeiro"
            case 2:
                return "Fevereiro"
            case 3:
                return "Março"
            case 4:
                return "Abril"
            case 5:
                return "Maio"
            case 6:
                return "Junho"
            case 7:
                return "Julho"
            case 8:
                return "Agosto"
            case 9:
                return "Setembro"
            case 10:
                return "Outubro"
            case 11:
                return "Novembro"
            case 12:
                return "Dezembro"

#---------------------------------------------------------------------------------------------------------------------------------------------

def subirNF():
    # try:
    global planilha
    global driver
    global url
    planilha = filedialog.askopenfilename()


    url = 'https://www2.geap.com.br/AuditoriaDigital/login'
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
    time.sleep(1)
    login_page.exe_login()
    time.sleep(2)
    pyautogui.press("enter")
    time.sleep(2)

    Caminho(driver,url).exe_caminho()
    time.sleep(1)
    Nf(driver,url).inserirDadosNf()
# except:
    #     pass

    # driver.quit()