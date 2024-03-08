from abc import ABC
from datetime import timedelta
from datetime import datetime
from tkinter import filedialog
import pandas as pd
from selenium.webdriver import Chrome
from seleniumwire import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from openpyxl import load_workbook
from webdriver_manager.chrome import ChromeDriverManager
from typing import Any
import time
from os import listdir, rename
from page_element import PageElement

class SalutisCasembrapa(PageElement):
    usuario_input: tuple = (By.XPATH, '//*[@id="username"]')
    senha_input: tuple = (By.XPATH, '//*[@id="password"]')
    entrar: tuple = (By.XPATH, '//*[@id="submit-login"]')
    salutis: tuple = (By.XPATH, '//*[@id="menuButtons"]/td[1]')
    websaude: tuple = (By.XPATH, '//*[@id="divTreeNavegation"]/div[7]/span[2]')
    credenciados: tuple = (By.XPATH, '//*[@id="divTreeNavegation"]/div[8]/span[2]')
    lotes: tuple = (By.XPATH, '//*[@id="divTreeNavegation"]/div[11]/span[2]')
    lotes_de_credenciados: tuple = (By.XPATH, '/html/body/div[8]/div[2]/div[20]/span[2]')
    fechar_recurso_de_glosa: tuple = (By.XPATH, '/html/body/div[3]/div/table/tbody/tr/td[1]/table/tbody/tr/td[4]/span/div')
    fechar_lotes_de_credenciados: tuple = (By.XPATH, '//*[@id="tabs"]/td[1]/table/tbody/tr/td[4]/span')
    numero_lote_pesquisa: tuple = (By.XPATH, '//*[@id="grdPesquisa"]/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[5]/td[2]/table/tbody/tr/td[1]/input')
    numero_lote_na_operadora: tuple = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input')
    buscar_lotes: tuple = (By.XPATH, '/html/body/div[4]/div/div[1]/table/tbody/tr/td[2]/table/tbody/tr/td[1]/span[2]')
    numero_lote_operadora: tuple = (By.CSS_SELECTOR, '#form-view-label_gridLote_NUMERO > table > tbody > tr > td:nth-child(1) > input')
    processamento_de_guias: tuple = (By.XPATH, '//*[@id="divTreeNavegation"]/div[6]/span[2]')
    recurso_de_glosa: tuple = (By.XPATH, '//*[@id="divTreeNavegation"]/div[9]/span[2]')
    input_guia = (By.XPATH, '//*[@id="pesquisaParametro"]/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input')
    input_processo = (By.XPATH, '//*[@id="pesquisaParametro"]/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr[1]/td[2]/table/tbody/tr/td[1]/input')
    input_lote = (By.XPATH, '//*[@id="pesquisaParametro"]/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr[7]/td[2]/table/tbody/tr/td[1]/input')
    buscar = (By.XPATH, '/html/body/div[4]/div/div[1]/table/tbody/tr/td[2]/table/tbody/tr')
    mudar_visao_1 = (By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[17]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[1]/div')
    mudar_visao_2 = (By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[17]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[1]/div')
    mudar_visao_3 = (By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[17]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[7]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[1]/div')
    mudar_visao_4 = (By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[17]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[7]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[1]/div')
    lupa_localizar_guia = (By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[17]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[2]/td[1]/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div')
    lupa_localizar_servico = (By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[17]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[2]/td[1]/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div')
    input_localizar_guia = (By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[17]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[3]/td[1]/table/tbody/tr[1]/td[3]/input')
    input_localizar_servico = (By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[17]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[3]/td[1]/table/tbody/tr[1]/td[3]/input')
    radio_todos_os_campos_guia = (By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[17]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[3]/td[1]/table/tbody/tr[2]/td[3]/table/tbody/tr[1]/td[1]/input')
    radio_todos_os_campos_servicos = (By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[17]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[3]/td[1]/table/tbody/tr[2]/td[3]/table/tbody/tr[1]/td[1]/input')
    data_inicio_lote = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[10]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input')
    data_fim_lote = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[10]/td/table/tbody/tr[3]/td[4]/table/tbody/tr/td[1]/input')
    data_inicio_operadora = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[12]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input')
    data_fim_operadora = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[12]/td/table/tbody/tr[3]/td[4]/table/tbody/tr/td[1]/input')
    bold_proxima_da_guia = (By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[17]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[3]/td[1]/table/tbody/tr[1]/td[4]/b')
    bold_proxima_do_servico = (By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[17]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[3]/td[1]/table/tbody/tr[1]/td[4]/b')
    input_proc_portal = (By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[17]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[1]/td/table/tbody/tr[5]/td[2]/table/tbody/tr/td[1]/input')
    input_valor_portal = (By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[17]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[4]/td/table/tbody/tr[3]/td[4]/table/tbody/tr/td[1]/input')
    inserir = (By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[17]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[7]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[2]/div')
    ok_alerta_inserir_recurso = (By.XPATH, '/html/body/div[10]/div[2]/table/tbody/tr[2]/td/div/button')
    input_numero_processo = (By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[17]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[7]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[1]/td/table/tbody/tr[1]/td[2]/table/tbody/tr/td[1]/input')
    input_motivo = (By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[17]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[7]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[4]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input')
    text_area_recurso = (By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[17]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[7]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[4]/td/table/tbody/tr[5]/td[2]/table/tbody/tr/td[1]/textarea')
    input_valor_recursado = (By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[17]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[7]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[4]/td/table/tbody/tr[9]/td[2]/table/tbody/tr/td[1]/input')
    confirmar_recurso_button = (By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[17]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[7]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]/div')
    gravar_recurso_button = (By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[17]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[7]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[2]/td[1]/table/tbody/tr/td[2]/div/table/tbody/tr/td[1]/span[2]/p')
    sim = (By.XPATH, '/html/body/div[9]/table/tbody/tr/td[2]/div/div[2]')
    span_quantidade_recurso = (By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[17]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[7]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[9]/span[2]')
    proximo_recurso = (By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[17]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[7]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[2]/td[1]/table/tbody/tr/td[1]/table/tbody/tr/td[4]/div')
    frame = (By.XPATH, '/html/body/div[4]/div/div[2]/iframe')
    input_valor_glosado = (By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[17]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[7]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[4]/td/table/tbody/tr[7]/td[2]/table/tbody/tr/td[1]/input')
    botao_ok_erro = (By.XPATH, '/html/body/div[10]/div[2]/table/tbody/tr[2]/td/div/button')
    guia_pesquisada = (By.XPATH, '/html/body/table/tbody/tr/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[17]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[1]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input')
    ok_alerta_lote = (By.XPATH, '/html/body/div[10]/div[2]/table/tbody/tr[2]/td/div/button')
    ultimo_lote = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[2]/td[1]/table/tbody/tr/td[1]/table/tbody/tr/td[5]/div')

    def __init__(self, driver: Chrome, url: str, usuario: str, senha: str, diretorio: str) -> None:
        super().__init__(driver=driver, url=url)
        self.usuario: str = usuario
        self.senha: str = senha
        self.diretorio: str = diretorio
        self.lista_de_planilhas: list[str] = [
            f'{diretorio}\\{arquivo}' 
            for arquivo in listdir(diretorio)
            if arquivo.endswith('.xlsx')
            ]

    def login(self):
        self.open()
        self.driver.implicitly_wait(30)
        self.driver.find_element(*self.usuario_input).send_keys(self.usuario)
        time.sleep(1)
        self.driver.find_element(*self.senha_input).send_keys(self.senha)
        time.sleep(1)
        self.driver.find_element(*self.entrar).click()
        time.sleep(1)

    def abrir_opcoes_menu(self):
        self.get_click(self.salutis, "Atalho")
        time.sleep(1)
        self.driver.find_element(*self.websaude).click()
        time.sleep(1)
        self.driver.find_element(*self.credenciados).click()
        time.sleep(1)
        self.driver.find_element(*self.lotes).click()
        time.sleep(1)
        self.driver.find_element(*self.processamento_de_guias).click()
        time.sleep(1)
        self.driver.find_element(*self.salutis).click()

    def pegar_numero_lote(self, numero_fatura: str):
        data_atual: datetime = datetime.now()
        data_seis_meses_atras: str = (data_atual - timedelta(days=180)).strftime('%d/%m/%Y')
        self.driver.implicitly_wait(30)
        self.driver.find_element(*self.lotes_de_credenciados).click()
        time.sleep(2)
        self.driver.switch_to.frame(self.get_attribute_value(self.frame, 'id'))
        self.confirma_valor_inserido(self.data_inicio_lote, data_seis_meses_atras)
        time.sleep(1)
        self.confirma_valor_inserido(self.data_fim_lote, data_atual.strftime('%d/%m/%Y'))
        time.sleep(1)
        self.confirma_valor_inserido(self.data_inicio_operadora, data_seis_meses_atras)
        time.sleep(1)
        self.confirma_valor_inserido(self.data_fim_operadora, data_atual.strftime('%d/%m/%Y'))
        time.sleep(1)
        self.driver.find_element(*self.numero_lote_pesquisa).clear()
        time.sleep(2)
        self.driver.find_element(*self.numero_lote_na_operadora).clear()
        time.sleep(1)
        self.confirma_valor_inserido(self.numero_lote_pesquisa, numero_fatura)
        time.sleep(1)
        self.driver.switch_to.default_content()
        
        for i in range(1, 5):
            try:
                self.driver.implicitly_wait(5)
                self.driver.find_element(*self.buscar_lotes).click()
                time.sleep(2)
                texto_no_botao = self.driver.find_element(*self.buscar_lotes).text

                if texto_no_botao == 'Pesquisar Lotes' or 'Lote cancelado' in self.driver.find_element(*self.body).text:
                    break

            except:
                if 'Lote cancelado' in self.driver.find_element(*self.body).text:
                    break

                continue

        time.sleep(2)

        self.driver.implicitly_wait(30)
        if 'Lote cancelado' in self.driver.find_element(*self.body).text:
            self.driver.switch_to.default_content()
            self.driver.find_element(*self.ok_alerta_lote).click()
            time.sleep(2)
            self.driver.switch_to.frame(self.get_attribute_value(self.frame, 'id'))
            self.driver.find_element(*self.ultimo_lote).click()
            self.driver.switch_to.default_content()
            time.sleep(2)


        self.driver.switch_to.frame(self.get_attribute_value(self.frame, 'id'))
        selector = self.driver.find_element(*self.numero_lote_operadora)
        lote_operadora = selector.get_attribute('value')
        self.driver.switch_to.default_content()
        self.driver.find_element(*self.fechar_lotes_de_credenciados).click()
        return lote_operadora
    
    def numero_lote_is_searchable(self, numero_lote: str) -> bool:
        self.driver.switch_to.frame(self.get_attribute_value(self.frame, 'id'))
        time.sleep(2)
        self.driver.find_element(*self.input_guia).clear()
        time.sleep(1)
        self.driver.find_element(*self.input_processo).clear()
        time.sleep(1)
        self.confirma_valor_inserido(self.input_lote, numero_lote)
        self.driver.switch_to.default_content()

        for i in range(1, 3):
            try:
                self.driver.find_element(*self.buscar).click()
                time.sleep(4)

                if "Erro" in self.driver.find_element(*self.body).text:
                    self.driver.switch_to.default_content()
                    self.driver.find_element(*self.botao_ok_erro).click()
                    time.sleep(2)
                    self.driver.find_element(*self.fechar_recurso_de_glosa).click()
                    return False

                texto_no_botao = self.driver.find_element(*self.buscar).text
                if texto_no_botao == 'Nova Busca':
                    return True
            except:
                continue

    def abrir_divs(self):
        self.driver.switch_to.frame(self.get_attribute_value(self.frame, 'id'))
        time.sleep(1)
        self.driver.find_element(*self.mudar_visao_1).click()
        time.sleep(1)
        self.driver.find_element(*self.mudar_visao_2).click()
        time.sleep(1)
        self.driver.find_element(*self.mudar_visao_3).click()
        time.sleep(1)
        self.driver.find_element(*self.mudar_visao_4).click()
        time.sleep(1)
        self.driver.find_element(*self.lupa_localizar_guia).click()
        time.sleep(1)
        self.driver.find_element(*self.lupa_localizar_servico).click()
        time.sleep(1)
        self.driver.find_element(*self.radio_todos_os_campos_guia).click()
        time.sleep(1)
        self.driver.find_element(*self.radio_todos_os_campos_servicos).click()
        time.sleep(1)

    def procedimento_is_valid(self, procedimento: str, valor: str):
        self.driver.find_element(*self.input_proc_portal).click()
        self.driver.find_element(*self.input_valor_portal).click()
        proced_portal: str = self.driver.find_element(*self.input_proc_portal).get_attribute('value')
        valor_portal: str = self.driver.find_element(*self.input_valor_portal).get_attribute('value')
        return procedimento in proced_portal and valor_portal in valor
    
    def salvar_valor_planilha(self, path_planilha: str, valor: Any, coluna: int, linha: int):
        if isinstance(valor, list) or isinstance(valor, tuple):
            dados = {"Recursado no Portal" : valor}
        else:
            dados = {"Recursado no Portal" : [valor]}

        df_dados = pd.DataFrame(dados)
        book = load_workbook(path_planilha)
        writer = pd.ExcelWriter(path_planilha, engine='openpyxl')
        writer.book = book
        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
        df_dados.to_excel(writer, 'Recurso', startrow=linha, startcol=coluna, header=False, index=False)
        writer.save()
        writer.close()

    def acrescenta_motivos_glosa(self, motivo_glosa: str, quantidade_recurso: int) -> list[str]:
        lista_motivos = []

        for i in range(0, quantidade_recurso):
            lista_motivos.append(motivo_glosa)

        return lista_motivos

    def lancar_recurso(self, path_planilha: str, linha: int, valor_recurso: str, codigo_mot_glosa: str, justificativa: str):
        self.driver.find_element(*self.inserir).click()
        time.sleep(2)
        self.driver.switch_to.default_content()
        self.driver.find_element(*self.ok_alerta_inserir_recurso).click()
        self.driver.switch_to.frame(self.get_attribute_value(self.frame, 'id'))
        numero_processo = self.driver.find_element(*self.input_numero_processo).get_attribute('value')
        valor_glosa_portal = self.driver.find_element(*self.input_valor_glosado).get_attribute('value')

        if valor_recurso > valor_glosa_portal:
            self.salvar_valor_planilha(
                path_planilha=path_planilha,
                valor="Valor de recurso maior que o valor glosado no portal",
                coluna=25,
                linha=linha+1
            )
            return
        
        time.sleep(2)
        self.salvar_valor_planilha(
            path_planilha=path_planilha,
            valor=numero_processo,
            coluna=3,
            linha=linha + 1
        )
        
        quantidade_recurso: int = int(self.driver.find_element(*self.span_quantidade_recurso).text)

        lista_de_codigos_mot_glosa: list[str] = codigo_mot_glosa.split(', ')

        if len(lista_de_codigos_mot_glosa) != quantidade_recurso and len(lista_de_codigos_mot_glosa) == 1:
            lista_de_codigos_mot_glosa: list[str] = self.acrescenta_motivos_glosa(lista_de_codigos_mot_glosa[0], quantidade_recurso)

        for codigo in lista_de_codigos_mot_glosa:
            readonly = self.driver.find_element(*self.input_motivo).get_attribute('readonly')

            if readonly == 'true':

                if quantidade_recurso > 1:
                    self.driver.find_element(*self.proximo_recurso).click()
                    time.sleep(1)

                continue

            self.driver.find_element(*self.input_motivo).click()
            self.driver.find_element(*self.input_motivo).send_keys(codigo)
            time.sleep(1)
            self.driver.find_element(*self.text_area_recurso).click()
            time.sleep(4)

            if self.driver.find_element(*self.input_motivo).get_attribute('value') == '':
                self.driver.switch_to.default_content()
                self.driver.find_element(*self.botao_ok_erro).click()
                time.sleep(2)
                self.driver.switch_to.frame(self.get_attribute_value(self.frame, 'id'))
                self.salvar_valor_planilha(
                    path_planilha=path_planilha,
                    valor='Código de motivo de glosa inválido',
                    coluna=25,
                    linha=linha+1
                )
                return

            self.confirma_valor_inserido(self.text_area_recurso, justificativa)
            time.sleep(1)
            self.confirma_valor_inserido(self.input_valor_recursado, valor_recurso)
            time.sleep(1)
            self.driver.find_element(*self.confirmar_recurso_button).click()
            time.sleep(1)
            
            if quantidade_recurso > 1:
                self.driver.find_element(*self.proximo_recurso).click()
                time.sleep(1)

        self.driver.find_element(*self.gravar_recurso_button).click()
        self.driver.find_element(*self.gravar_recurso_button).click()
        time.sleep(1)
        self.driver.switch_to.default_content()
        self.driver.find_element(*self.sim).click()
        time.sleep(1)
        self.driver.switch_to.frame(self.get_attribute_value(self.frame, 'id'))


        self.salvar_valor_planilha(
            path_planilha=path_planilha,
            valor='Sim',
            coluna=25,
            linha=linha+1
        )

    def renomear_planilha(self, path_planilha: str, msg: str):
        novo_nome: str = path_planilha.replace('.xlsx', '') + f'_{msg}.xlsx'
        rename(path_planilha, novo_nome)

    def content_has_value(self, element: tuple, valor: str) -> bool:
        content = self.driver.find_element(*element).text
        return valor in content

    def executa_recurso(self):
        self.login()
        self.abrir_opcoes_menu()

        for planilha in self.lista_de_planilhas:
            if "Atalho" not in self.driver.find_element(*self.body).text:
                self.get_click(self.salutis, "Atalho")

            if 'Enviado' in planilha or 'Nao_Enviado' in planilha:
                continue

            df = pd.read_excel(planilha)
            numero_fatura = str(df['Fatura Inicial'][0]).replace('.0', '')
            lote_operadora = str(df['Lote'][0]).replace('.0', '')

            if lote_operadora == '' or lote_operadora == 'nan':
                lote_operadora = self.pegar_numero_lote(numero_fatura)

                if lote_operadora == '':
                    self.renomear_planilha(planilha, "Nao_Enviado")
                    continue

            time.sleep(2)

            df['Lote'] = lote_operadora
            self.salvar_valor_planilha(
                path_planilha=planilha,
                valor=df['Lote'].values.tolist(),
                coluna=2,
                linha=1
            )

            # Entra em Recurso de Glosa
            self.get_click(self.salutis, "Atalho")   
            time.sleep(2)
            self.driver.find_element(*self.recurso_de_glosa).click()
            time.sleep(2)
            if not self.numero_lote_is_searchable(lote_operadora):
                self.renomear_planilha(planilha, "Nao_Enviado")
                continue

            # if self.content_has_value(self.body, 'Erro'):
            #     self.renomear_planilha(planilha, 'Nao_Enviado')
            #     continue

            time.sleep(2)
            self.abrir_divs()

            for index, linha in df.iterrows():

                if f'{linha["Recursado no Portal"]}' == 'Sim' or f'{linha["Recursado no Portal"]}' == 'Procedimento não encontrado':
                    continue

                numero_guia = f'{linha["Nro. Guia"]}'.replace('.0', '')
                codigo_procedimento = f'{linha["Procedimento"]}'.replace('.0', '')

                if isinstance(linha['Valor Glosa'], float) or isinstance(linha['Valor Glosa'], int):
                    valor_glosa = "{:.2f}".format(linha["Valor Glosa"])

                else:
                    valor_glosa = "{:.2f}".format(float(linha["Valor Glosa"].replace('.', '').replace(',', '.')))

                if isinstance(linha['Valor Recursado'], float) or isinstance(linha['Valor Recursado'], int):
                    valor_recurso = "{:.2f}".format(linha["Valor Recursado"])
                    
                else: 
                    valor_recurso = "{:.2f}".format(float(linha["Valor Recursado"].replace('.', '').replace(',', '.')))
                
                codigo_motivo_glosa = f'{linha["Motivo Glosa"]}'
                justificativa = f'{linha["Recurso Glosa"]}'.replace('\t', ' ')

                guia_pesquisada = self.driver.find_element(*self.guia_pesquisada).get_attribute('value')
                
                if guia_pesquisada != numero_guia:
                    self.confirma_valor_inserido(self.input_localizar_guia, numero_guia)
                    self.driver.find_element(*self.bold_proxima_da_guia).click()

                    time.sleep(3)
                    self.driver.switch_to.default_content()
                    
                    if self.content_has_value(self.body, 'Valor não encontrado'):
                        self.salvar_valor_planilha(
                            path_planilha=planilha,
                            valor='Guia não encontrada',
                            coluna=25,
                            linha=index + 1
                        )
                        self.driver.find_element(*self.ok_alerta_inserir_recurso).click()
                        time.sleep(1)
                        self.driver.switch_to.frame(self.get_attribute_value(self.frame, 'id'))
                        continue

                    self.driver.switch_to.frame(self.get_attribute_value(self.frame, 'id'))

                self.confirma_valor_inserido(self.input_localizar_servico, codigo_procedimento)
                self.driver.find_element(*self.bold_proxima_do_servico).click()
                self.driver.find_element(*self.bold_proxima_do_servico).click()
                time.sleep(5)
                self.driver.switch_to.default_content()

                if self.content_has_value(self.body, 'Valor não encontrado'):
                    self.salvar_valor_planilha(
                        path_planilha=planilha,
                        valor='Procedimento não encontrado',
                        coluna=25,
                        linha=index + 1
                    )
                    self.driver.find_element(*self.ok_alerta_inserir_recurso).click()
                    time.sleep(1)
                    self.driver.switch_to.frame(self.get_attribute_value(self.frame, 'id'))
                    continue

                self.driver.switch_to.frame(self.get_attribute_value(self.frame, 'id'))

                if self.procedimento_is_valid(codigo_procedimento, valor_glosa):

                    self.lancar_recurso(
                        path_planilha=planilha,
                        linha=index,
                        valor_recurso=valor_recurso,
                        codigo_mot_glosa=codigo_motivo_glosa,
                        justificativa=justificativa
                    )

            self.renomear_planilha(planilha, 'Enviado')
            self.driver.switch_to.default_content()
            self.driver.find_element(*self.fechar_recurso_de_glosa).click()
            time.sleep(1)

def recursar_casembrapa(user, password):
    diretorio = filedialog.askdirectory()
    url = 'http://170.84.17.131:22101/sistema'

    options = {
        'proxy' : {
            'http': f'http://{user}:{password}@10.0.0.230:3128',
            'https': f'http://{user}:{password}@10.0.0.230:3128'
        }
    }

    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument('--ignore-certificate-errors')
    chrome_options.add_argument('--ignore-ssl-errors')
    servico = Service(ChromeDriverManager().install())

    driver = webdriver.Chrome(service=servico, seleniumwire_options= options, options = chrome_options)

    usuario = "00735860000173"
    senha = "0073586@"

    salutis_casembrapa = SalutisCasembrapa(driver, url, usuario, senha, diretorio)
    salutis_casembrapa.executa_recurso()