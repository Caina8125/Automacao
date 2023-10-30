from selenium import webdriver
import time
import pandas as pd
import tkinter
from tkinter import filedialog
from openpyxl import load_workbook
import pyautogui
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import Pidgin

class ConsultaCnpj():
    digita_cnpj = (By.XPATH, '//*[@id="Cnpj"]')
    button_consulta = (By.XPATH, '//*[@id="consultarForm"]/button')
    situacao = (By.XPATH, '//*[@id="conteudo"]/div[2]/div[2]/span[1]')

    def __init__(self, driver, url) -> None:
        self.driver = driver
        self.url = url

    def open(self):
        self.driver.get(self.url)
    
    def consulta(self, planilha):
        self.driver.get(url)
        self.driver.implicitly_wait(30)
        time.sleep(15)
        df = pd.read_excel(planilha)
        df.columns = ['Matrícula', 'CNPJ', 'Nome']
        df['Resposta'] = ''

        for index, linha in df.iterrows():
            cnpj = linha['CNPJ']
            self.driver.find_element(*self.digita_cnpj).send_keys(cnpj)
            time.sleep(1)
            self.driver.find_element(*self.button_consulta).click()
            time.sleep(2)
            situacao = self.driver.find_element(*self.situacao).text
            dados = {"Resposta" : [situacao]}
            df_resposta = pd.DataFrame(dados)
            book = load_workbook(planilha)
            writer = pd.ExcelWriter(planilha, engine='openpyxl')
            writer.book = book
            writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
            df_resposta.to_excel(writer, 'Planilha1', startrow=index, startcol=4, header=False, index=False)
            writer.save()

try:
    login_usuario = 'lucas.paz'
    senha_usuario = 'Gsw2022&'
    url = 'https://consopt.www8.receita.fazenda.gov.br/consultaoptantes'
    planilha = filedialog.askopenfilename()

    options = {
        'proxy' : {
            'http': 'http://lucas.paz:Gsw2022&@10.0.0.230:3128',
            'https': 'http://lucas.paz:Gsw2022&@10.0.0.230:3128'
        }
    }

    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    servico = Service(ChromeDriverManager().install())

    driver = webdriver.Chrome(service=servico, options = chrome_options)

    fazer_consulta = ConsultaCnpj(driver, url)
    fazer_consulta.open()
    time.sleep(4)
    pyautogui.write(login_usuario)
    pyautogui.press("TAB")
    time.sleep(1)
    pyautogui.write(senha_usuario)
    pyautogui.press("enter")
    time.sleep(4)
    fazer_consulta.consulta(planilha)

except FileNotFoundError as err:
    tkinter.messagebox.showerror('Automação', f'Nenhuma planilha foi selecionada!')

except Exception as err:
    tkinter.messagebox.showerror("Automação", f"Ocorreu uma exceção não tratada. \n {err.__class__.__name__} - {err}")
    Pidgin.main(f"Ocorreu uma exceção não tratada. \n {err.__class__.__name__} - {err}")