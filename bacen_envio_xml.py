from abc import ABC
import time
import pandas as pd
from selenium import webdriver
from seleniumwire import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from tkinter.filedialog import askdirectory
from bacen_protocolo import BuscarProtocolo, PageElement
from abc import ABC
import pandas as pd
from seleniumwire import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from tkinter import messagebox
import os
from datetime import datetime

class EnviarXML(PageElement):
    opcao_envio_de_xml = (By.XPATH, '/html/body/form/div[3]/div[3]/div[1]/div/ul/li[10]/a')
    opcao_enviar_arquivo = (By.XPATH, '/html/body/form/div[3]/div[3]/div[1]/div/ul/li[10]/ul/li[1]/a')
    input_file = (By.XPATH, "/html/body/input")
    salvar_novo = (By.XPATH, '/html/body/form/div[3]/div[3]/div[3]/div/div/div/div/div/div[2]/div/div[2]/div[1]/a[2]')
    body = (By.XPATH, '/html/body')

    def caminho(self):
        self.driver.find_element(*self.opcao_envio_de_xml).click()
        time.sleep(2)
        self.driver.find_element(*self.opcao_enviar_arquivo).click()
        time.sleep(2)
    
    def enviar_arquivo(self, lista_de_arquivos):
        lista_processos = []
        self.driver.implicitly_wait(30)
        try:
            for arquivo in lista_de_arquivos:
                string_vet = arquivo.split("_")
                numero_processo = string_vet[0].replace("0000000000000", '').replace(pasta, '').replace('\\', '')
                time.sleep(2)
                self.driver.find_element(*self.input_file).send_keys(arquivo)
                content = self.driver.find_element(*self.body).text

                while "Aguarde" in content:
                    time.sleep(2)
                    content = self.driver.find_element(*self.body).text
                    
                time.sleep(2)
                self.driver.find_element(*self.salvar_novo).click()
                lista_processos.append(numero_processo)
                time.sleep(3)
            time.sleep(10)
            return lista_processos
        
        except:
            return lista_processos

    def confere_envio(self, lista_de_processos):
        self.driver.get("https://www3.bcb.gov.br/portalbcsaude/saude/a/portal/prestador/tiss/ConsultarArquivoTiss.aspx?i=PORTAL_SUBMENU_XML_CONSULTARARQUIVO&m=MENU_ENVIO_XML_AGRUPADO")
        listas_para_df = []
        for processo in lista_de_processos:
            protocolo = acessar_portal.buscar_protocolo(processo)
            lista = [processo, protocolo]
            listas_para_df.append(lista)
        cabecalho = ["N° Fatura", "N° Protocolo"]
        df = pd.DataFrame(listas_para_df, columns=cabecalho)
        data_e_hora_atuais = datetime.now()
        data_e_hora_em_texto = data_e_hora_atuais.strftime('%d_%m_%Y_%H_%M')
        segundo = data_e_hora_atuais.second
        df.to_excel(f'Bacen\\Envio_xml_{data_e_hora_em_texto}_{segundo}.xlsx', index=False)
#--------------------------------------------------------------------------------------------------------------------------------------------------------------------

def fazer_envio_xml(user, password):
    messagebox.showwarning("Automação Bacen", "Selecione uma pasta!")
    global pasta
    pasta = askdirectory()
    lista_de_arquivos = [f"{pasta}\\{arquivo}" for arquivo in os.listdir(pasta) if arquivo.endswith(".xml") or arquivo.endswith(".zip")]
    url = 'https://www3.bcb.gov.br/portalbcsaude/Login'
    global driver

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
        global acessar_portal
        acessar_portal = BuscarProtocolo(driver, url)
        acessar_portal.open()

        acessar_portal.login_layout_novo(
            usuario = "00735860000173",
            senha = "Amhpdf!2023"
        )
        enviar_xml = EnviarXML(driver, url='')
        enviar_xml.caminho()
        lista_de_processos = enviar_xml.enviar_arquivo(lista_de_arquivos)
        enviar_xml.confere_envio(lista_de_processos)
        driver.quit()
        messagebox.showinfo( 'Automação Bacen' , 'Todas os envios foram concluídos.' )
    except Exception as e:
        messagebox.showerror( 'Erro Automação' , f'Ocorreu uma excessão não tratada \n {e.__class__.__name__}: {e}' )
        driver.quit()