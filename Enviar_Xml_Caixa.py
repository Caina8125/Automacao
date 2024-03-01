from datetime import datetime
from tkinter import filedialog
import tkinter.messagebox
from xml.dom import minidom
from selenium.webdriver.common.by import By
from selenium import webdriver
import pandas as pd
import time
import os
from selenium.webdriver.chrome.options import Options
from seleniumwire import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from page_element import PageElement

class Login(PageElement):
    usuario = (By.XPATH, '//*[@id="UserName"]')
    senha = (By.XPATH, '//*[@id="Password"]')
    acessar = (By.XPATH, '//*[@id="LoginButton"]')

    def exe_login(self, usuario, senha):
        self.driver.find_element(*self.usuario).send_keys(usuario)
        self.driver.find_element(*self.senha).send_keys(senha)
        self.driver.find_element(*self.acessar).click()

class Xml(PageElement):
    body = (By.XPATH, '/html/body')
    envio_xml = (By.XPATH, '//*[@id="sidebar_envioXML"]/span[1]')
    enviar_xml = (By.XPATH, '//*[@id="ctl00_SidebarMenu"]/li[10]/ul/li[1]/a')
    anexar_xml = (By.XPATH, '//*[@id="ctl00_Body"]/input')
    gravar = (By.XPATH, '//*[@id="ctl00_Main_WIDGETID_636040089093201024_toolbar"]/a[2]')
    numero_lote_input = (By.XPATH, '/html/body/form/div[3]/div[3]/div[3]/div/div/div[1]/div/div/div[2]/div/div[2]/span/div/div/div/div[3]/div/div/div[2]/div[1]/span/input')
    filtrar = (By.XPATH, '/html/body/form/div[3]/div[3]/div[3]/div/div/div[1]/div/div/div[2]/div/div[2]/span/div/div/div/div[2]/a[2]/i')
    protocolo = (By.XPATH, '/html/body/form/div[3]/div[3]/div[3]/div/div/div[1]/div/div/div[2]/div/div[2]/div[4]/table/tbody/tr/td[4]/a')

    def caminho(self):
        self.driver.find_element(*self.envio_xml).click()
        time.sleep(2)
        self.driver.find_element(*self.enviar_xml).click()
        time.sleep(2)
    
    def arquivos(self):
        xml = os.listdir(pasta)
        return xml
    
    def altera_xml_enviado_pasta(self,caminho,xml):
        destino = r"C:\Convenios Medicos\Server\XML"
        os.rename(caminho, destino + '\\' + xml)

    def get_all_zeroes(self, string) -> str:
        zeroes = ''

        for char in string:

            if char != '0':
                break

            zeroes += char

        return zeroes


    def envio(self) -> list:
        self.caminho()
        arquivos = self.arquivos()
        lista_dados = []

        for xml in arquivos:
            caminho_arquivo = pasta + '/' + xml
            sequencial = xml.split("_")[0].replace(self.get_all_zeroes(xml.split("_")[0]), '')
            self.driver.find_element(*self.anexar_xml).send_keys(caminho_arquivo)
            content = self.driver.find_element(*self.body).text

            while "Aguarde" in content:
                time.sleep(2)
                content = self.driver.find_element(*self.body).text

            time.sleep(1)
            self.driver.find_element(*self.gravar).click()
            time.sleep(1)

            with open(caminho_arquivo, "r",) as f:
                xml_document = minidom.parse(f)
                tag_processo = xml_document.getElementsByTagName("ans:numeroLote")

                for tag in tag_processo:
                    processo = tag.firstChild.data

            lista_dados.append(
                {
                    'processo': processo,
                    'sequencial': sequencial
                }
            )

            self.altera_xml_enviado_pasta(caminho_arquivo, xml)

        return lista_dados
    
    def buscar_protocolo(self, numero_fatura) -> str:
        self.driver.find_element(*self.numero_lote_input).send_keys(numero_fatura)
        time.sleep(2)
        self.driver.find_element(*self.filtrar).click()
        time.sleep(2)
        body = self.driver.find_element(*self.body).text

        if "Nenhum registro encontrado." in body:
            self.driver.find_element(*self.numero_lote_input).clear()
            return 'Nenhum registro encontrado.'

        else:
            protocolo = self.driver.find_element(*self.protocolo).text
            time.sleep(2)
            self.driver.find_element(*self.numero_lote_input).clear()
            return protocolo

    def confere_envio(self, lista_dados):
        self.driver.get("https://saude.caixa.gov.br/PORTALPRD/saude/a/portal/prestador/tiss/ConsultarArquivoTiss.aspx?i=PORTAL_SUBMENU_XML_CONSULTARARQUIVO&m=MENU_ENVIO_XML_AGRUPADO")
        listas_para_df = []   

        for data in lista_dados:
            protocolo = self.buscar_protocolo(data['processo'])
            lista = [data['processo'], data['sequencial'], protocolo]
            listas_para_df.append(lista)

        cabecalho = ["Fatura", "Sequencial", "Protocolo"]
        df = pd.DataFrame(listas_para_df, columns=cabecalho)
        data_e_hora_atuais = datetime.now()
        data_e_hora_em_texto = data_e_hora_atuais.strftime('%d_%m_%Y_%H_%M')
        segundo = data_e_hora_atuais.second
        df.to_excel(f'Output\\Envio_xml_{data_e_hora_em_texto}_{segundo}.xlsx', index=False)
#---------------------------------------------------------------------------------------------------------------------------------
def Enviar_caixa(user, password):
    global pasta
    pasta = filedialog.askdirectory()

    global url
    url = 'https://saude.caixa.gov.br/PORTALPRD/'

    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument('--ignore-certificate-errors')
    chrome_options.add_argument('--ignore-ssl-errors')

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

    try:
        login_page = Login(driver, url)
        login_page.open()

        login_page.exe_login(
            usuario = "00735860000173",
            senha = "!Saude2024"
        )

        lista_dados = Xml(driver, url).envio()
        Xml(driver, url).confere_envio(lista_dados)

        tkinter.messagebox.showinfo( 'Automação Saúde Caixa' , 'Arquivos XML enviados!' )
    except Exception as e:
        tkinter.messagebox.showerror( 'Erro Automação' , f'Ocorreu uma excessão não tratada \n {e.__class__.__name__}: {e}' )
        driver.quit()