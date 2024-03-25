from abc import ABC
from datetime import date
from os import listdir, rename
from tkinter.filedialog import askdirectory
from openpyxl import load_workbook
from pandas import DataFrame, ExcelWriter, concat, read_excel, read_html
import pyautogui
from selenium.webdriver.remote.webelement import WebElement
from selenium import webdriver
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import time
from google_sheets import GoogleSheets
import requests
from tkinter.messagebox import showerror, showinfo
# from page_element import PageElement

class PageElement(ABC):
    body: tuple = (By.XPATH, '/html/body')
    a: tuple = (By.TAG_NAME, 'a')
    p: tuple = (By.TAG_NAME, 'p')
    h1: tuple = (By.TAG_NAME, 'h1')
    h2: tuple = (By.TAG_NAME, 'h2')
    h3: tuple = (By.TAG_NAME, 'h3')
    h4: tuple = (By.TAG_NAME, 'h4')
    h5: tuple = (By.TAG_NAME, 'h5')
    h6: tuple = (By.TAG_NAME, 'h6')
    table: tuple = (By.TAG_NAME, 'table')

    def __init__(self, driver: WebDriver, url: str) -> None:
        self.driver: WebDriver = driver
        self._url:str = url

    def open(self) -> None:
        self.driver.implicitly_wait(30)
        self.driver.get(self._url)

    def get_attribute_value(self, element: tuple, atributo: str) -> str | None:
        return self.driver.find_element(*element).get_attribute(atributo)

    def confirma_valor_inserido(self, element: tuple, valor: str) -> None:
        """Este método verifica se um input recebeu os valores que foram enviados.
           Caso não tenha recebido, tenta enviar novamente até 10x."""
        try:
            self.driver.find_element(*element).clear()
            valor_inserido: str = self.driver.find_element(*element).get_attribute('value')
            count: int = 0

            while valor_inserido == '':
                self.driver.find_element(*element).send_keys(valor)
                time.sleep(0.5)
                valor_inserido: str = self.driver.find_element(*element).get_attribute('value')
                count += 1

                if count == 10:
                    raise Exception("Element not interactable")

        except Exception as e:
            raise Exception(e)
        
    def get_element_visible(self, element: tuple=None,  web_element: WebElement=None) -> bool:
        """Este método observa se irá ocorrer ElementClickInterceptedException. Caso ocorra
        irá dar um scroll até 10x na página conforme o comando passado até achar o click do elemento"""
        for i in range(10):
            try:
                if element != None:
                    self.driver.find_element(*element).click()
                    return True
                
                elif web_element != None:
                    web_element.click()
                    return True
            
            except:
                if i == 10:
                    return False
                
                self.driver.execute_script('scrollBy(0,100)')
                continue

    def get_click(self, element: tuple, valor: str) -> None:
        for i in range(10):
            self.driver.find_element(*element).click()
            time.sleep(3)
            content: str = self.driver.find_element(*self.body).text

            if valor in content:
                break
            
            else:
                if i == 10:
                    raise Exception('Element not interactable')
                
                time.sleep(2)
                continue

class Trt(PageElement):
    input_usuario = (By.ID, 'j_idt85:usuario')
    input_senha = (By.ID, 'j_idt85:senha')
    btn_entrar = (By.ID, 'j_idt85:j_idt91')
    div_lotes_pendentes = (By.ID, 'j_idt85:pendenteanalise_content')
    btn_voltar = (By.ID, 'frmRecursoGlosaCabecalho:j_idt89')
    google_sheets = GoogleSheets('1hmPxV74gKEPf1x8b-xwuxk6YjU3qIuF5iJdmI8wHR6E')
    data_hoje = date.today()

    def __init__(self, driver: WebDriver, url: str, usuario: str, senha: str) -> None:
        super().__init__(driver, url)
        self.usuario = usuario
        self.senha = senha

    def login(self):
        self.driver.find_element(*self.input_usuario).send_keys(self.usuario)
        time.sleep(1.5)
        self.driver.find_element(*self.input_senha).send_keys(self.senha)
        time.sleep(1.5)
        self.driver.find_element(*self.btn_entrar).click()
        time.sleep(1.5)

    def pegar_lotes_pendentes(self):
        self.open()
        self.login()

        creds = self.google_sheets.authenticate()
        columns_planilha = self.google_sheets.sheet_columns(creds)
        df_valores_planilha = DataFrame(self.google_sheets.sheet_values(creds))
        df_valores_planilha.columns = columns_planilha
        conteudo_div_lotes = self.driver.find_element(*self.div_lotes_pendentes).text
        lista_de_lotes = [text.split(' ')[1] for text in conteudo_div_lotes.split('\n') if 'Lote' in text]
        
        for lote in lista_de_lotes:
            for index, linha in df_valores_planilha.iterrows():
                if lote == linha['Protocolo']:
                    self.google_sheets.send_values(creds, f'C{index + 2}', [['Sim']])
                    self.google_sheets.send_values(creds, f'F{index + 2}:H{index + 2}', [['Luquinhas', self.data_hoje.strftime('%d/%m/%Y'), '30']])
            self.driver.find_element(By.PARTIAL_LINK_TEXT, lote).click()
            time.sleep(1.5)
            self.driver.find_element(*self.btn_voltar).click()
            time.sleep(1.5)

def buscar_lotes_trt():
    URL = 'https://faturamentoeletronico.trt10.jus.br/faturamentoeletronico/jsf/login.jsf'
    USUARIO = 'recurso.amhpdf'
    SENHA = 'amhprecurso2023'

    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")

    try:
        servico = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=servico, options = chrome_options)
    except:
        driver = webdriver.Chrome(options = chrome_options)

    trt = Trt(driver, URL, USUARIO, SENHA)
    trt.pegar_lotes_pendentes()

if __name__ == '__main__':
    buscar_lotes_trt()