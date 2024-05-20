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
    #-----------------------------------------------------------------------------------------------------------------------------
    table = (By.ID, 'recursoGlosaTabelaServicos')
    text_area_justificativa = (By.ID, 'txtJustificativa')
    #----------------------------------------------------------------------------------------------------------------------------------
    button_ok = ()
    salvar_parcialmente = ()
    i_close = ()
    proxima_pagina = ()
    label_registros = ()
    primeira_pagina = ()
    fechar = ()
    ul = ()
    #-----------------------------------------------------------------------------------------------------------------------------------------------
    close_warning = (By.XPATH, '/html/body/ul/li/div/div/span/h4/i')
    recurso_de_glosa_menu = (By.XPATH, '//*[@id="menuPrincipal"]/div/div[10]/a')
    fatura_input = (By.XPATH, '/html/body/main/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/input-number/div/div/input')
    pesquisar_recurso = (By.XPATH, '/html/body/main/div[1]/div[1]/div[2]/div[2]/button[1]')
    recurso_de_glosa_2 = (By.XPATH, '/html/body/main/div[1]/div[2]/div/table/tbody/tr/td[11]/i')
    alerta = (By.XPATH, '/html/body/ul/li/div/div[2]/button[2]')
    guia_op = (By.XPATH, '/html/body/main/div[1]/div[1]/div[2]/div[1]/div[2]/input-text-search/div/div/div/input') 
    buscar = (By.XPATH, '/html/body/main/div[1]/div[1]/div[2]/div[1]/div[2]/input-text-search/div/div/div/span/span')

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

    def inicializar_atributos(self, recurso_iniciado):
        if recurso_iniciado == False:
            self.button_ok = (By.XPATH, '/html/body/main/div/div[3]/div/div/div[3]/button[1]')
            self.salvar_parcialmente = (By.XPATH, '/html/body/main/div/div[2]/div/div/div[2]/div/div[1]/div/div/div[3]/div/div[2]/button')
            self.i_close = (By.XPATH, '/html/body/ul/li/div/div/span/h4/i')
            # self.proxima_pagina = (By.XPATH, '/html/body/main/div/div[2]/div/div/div[2]/div/div[1]/div/div/div[3]/div/div[1]/div/nav/ul/li[6]/a/span')
            self.label_registros = (By.XPATH, '/html/body/main/div/div[2]/div/div/div[2]/div/div[1]/div/div/div[3]/div/div[1]/label')
            self.primeira_pagina = (By.XPATH, '/html/body/main/div/div[2]/div/div/div[2]/div/div[1]/div/div/div[3]/div/div[1]/div/nav/ul/li[1]/a/span')
            self.fechar = (By.XPATH, '/html/body/main/div/div[2]/div/div/div[3]/button')
            self.ul = (By.XPATH, '/html/body/main/div/div[2]/div/div/div[2]/div/div[1]/div/div/div[3]/div/div[1]/div/nav/ul')
        else:
            self.button_ok = (By.XPATH, '/html/body/main/div[1]/div[4]/div/div/div[3]/button[1]')
            self.salvar_parcialmente = (By.XPATH, '/html/body/main/div[1]/div[3]/div/div/div[2]/div/div[1]/div/div/div[3]/div/div[2]/button')
            self.i_close = (By.XPATH, '/html/body/ul/li/div/div/span/h4/i')
            # self.proxima_pagina = (By.XPATH, '/html/body/main/div[1]/div[3]/div/div/div[2]/div/div[1]/div/div/div[3]/div/div[1]/div/nav/ul/li[6]/a/span')
            self.label_registros = (By.XPATH, '/html/body/main/div[1]/div[3]/div/div/div[2]/div/div[1]/div/div/div[3]/div/div[1]/label')
            self.primeira_pagina = (By.XPATH, '/html/body/main/div[1]/div[3]/div/div/div[2]/div/div[1]/div/div/div[3]/div/div[1]/div/nav/ul/li[1]/a/span')
            self.fechar = (By.XPATH, '/html/body/main/div[1]/div[3]/div/div/div[3]/button')
            self.ul = (By.XPATH, '/html/body/main/div[1]/div[3]/div/div/div[2]/div/div[1]/div/div/div[3]/div/div[1]/div/nav/ul')