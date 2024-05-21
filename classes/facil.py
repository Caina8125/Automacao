from tkinter import filedialog
import pandas as pd
from selenium import webdriver
from seleniumwire import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import time
from openpyxl import load_workbook
import tkinter.messagebox
import tkinter
import os
from page_element import PageElement

class Facil(PageElement):
    prestador_pj = (By.XPATH, '//*[@id="tipoAcesso"]/option[7]')
    usuario = (By.XPATH, '//*[@id="login-entry"]')
    senha = (By.XPATH, '//*[@id="password-entry"]')
    entrar = (By.XPATH, '//*[@id="BtnEntrar"]')
    faturas = (By.XPATH, '//*[@id="menuPrincipal"]/div/div[12]/a')
    relatorio_de_faturas = (By.XPATH, '/html/body/header/div[4]/div/div/div/div[12]/div[1]/div[2]/div/div[2]/div/div/div/div[1]/a')
    body = (By.XPATH, '/html/body')
    codigo = (By.XPATH, '/html/body/main/div/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/input-text[1]/div/div/input')
    pesquisar = (By.XPATH, '//*[@id="filtro"]/div[2]/div[2]/button')
    recurso_de_glosa = (By.XPATH, '/html/body/main/div/div[1]/div[4]/div/div/div[1]/div/div[2]/a[5]/i')
    table = (By.ID, 'recursoGlosaTabelaServicos')
    text_area_justificativa = (By.ID, 'txtJustificativa')
    close_warning = (By.XPATH, '/html/body/ul/li/div/div/span/h4/i')
    recurso_de_glosa_menu = (By.XPATH, '//*[@id="menuPrincipal"]/div/div[10]/a')
    fatura_input = (By.XPATH, '/html/body/main/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/input-number/div/div/input')
    pesquisar_recurso = (By.XPATH, '/html/body/main/div[1]/div[1]/div[2]/div[2]/button[1]')
    recurso_de_glosa_2 = (By.XPATH, '/html/body/main/div[1]/div[2]/div/table/tbody/tr/td[11]/i')
    alerta = (By.XPATH, '/html/body/ul/li/div/div[2]/button[2]')
    guia_op = (By.XPATH, '/html/body/main/div[1]/div[1]/div[2]/div[1]/div[2]/input-text-search/div/div/div/input') 
    buscar = (By.XPATH, '/html/body/main/div[1]/div[1]/div[2]/div[1]/div[2]/input-text-search/div/div/div/span/span')
    
    def __init__(self, driver, url: str, usuario: str, senha: str, diretorio: str) -> None:
        super().__init__(driver=driver, url=url)
        self.usuario: str = usuario
        self.senha: str = senha
        self.diretorio: str = diretorio
        self.lista_de_planilhas: list[str] = [
            f'{diretorio}\\{arquivo}' 
            for arquivo in os.listdir(diretorio)
            if arquivo.endswith('.xlsx')
            ]

    def exe_login(self):
        self.driver.find_element(*self.prestador_pj).click()
        self.driver.find_element(*self.usuario).send_keys(self.usuario)
        self.driver.find_element(*self.senha).send_keys(self.senha)
        self.driver.find_element(*self.entrar).click()
        time.sleep(5)

    def exe_caminho(self):
        try:
            self.driver.implicitly_wait(10)
            time.sleep(2)
            self.driver.find_element(*self.faturas)
        except:
            self.driver.refresh()
            time.sleep(2)
            self.exe_login()

        self.driver.implicitly_wait(30)
        time.sleep(3)
        self.driver.find_element(*self.faturas).click()
        time.sleep(2)
        self.driver.find_element(*self.relatorio_de_faturas).click()
        time.sleep(2)

    def executar_recurso(self):
        for planilha in self.lista_de_planilhas:

            if "Enviado" in planilha or "Sem_Pagamento" in planilha:
                continue

            df = pd.read_excel(planilha)
            protocolo = f"{df['Protocolo Glosa'][0]}".replace(".0", "")
            self.driver.find_element(*self.codigo).send_keys(protocolo)
            time.sleep(2)
            self.driver.find_element(*self.pesquisar).click()
            time.sleep(2)
            self.driver.find_element(*self.recurso_de_glosa).click()
            time.sleep(2)
            content = self.driver.find_element(*self.body).text
            recurso_iniciado = False

            if 'Não existe informação de pagamento para a fatura recursada.' in content:
                planilha_anterior = planilha
                sem_extensao = planilha.replace('.xlsx', '')
                novo_nome = sem_extensao + '_Sem_Pagamento.xlsx'
                try:
                    time.sleep(2)
                    os.rename(planilha_anterior, novo_nome)
                    continue
                except PermissionError as err:
                    print(err)
                    continue

            if 'A fatura não possui itens para gerar o lote de recurso de glosa ou já existem lotes gerados para a mesma.' in content:
                recurso_iniciado = True
                self.driver.find_element(*self.close_warning).click()
                time.sleep(2)
                self.driver.find_element(*self.recurso_de_glosa_menu).click()
                time.sleep(2)
                self.driver.find_element(*self.fatura_input).send_keys(protocolo)
                time.sleep(2)
                self.driver.find_element(*self.pesquisar_recurso).click()
                time.sleep(2)
                self.driver.find_element(*self.recurso_de_glosa_2).click()

                tamanho_table = len(pd.read_html(self.driver.find_element(*self.table))[0])

                for i in range(1, tamanho_table + 1):
                    ...

    def abrir_guias(self):
        i_elements = [i_element for i_element in self.driver.find_elements(By.TAG_NAME, 'i') if i_element.get_attribute('target') != '']

        for i_element in i_elements:
            i_element.click()