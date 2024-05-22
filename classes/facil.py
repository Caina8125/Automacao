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
# from page_element import PageElement

from abc import ABC
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.remote.webelement import WebElement
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.webdriver.common.by import By
from time import sleep

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

    def __init__(self, driver: WebDriver, url: str = '') -> None:
        self.driver: WebDriver = driver
        self._url:str = url

    def open(self) -> None:
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
                sleep(0.5)
                valor_inserido: str = self.driver.find_element(*element).get_attribute('value')
                count += 1

                if count == 10:
                    raise Exception("Element not interactable")

        except Exception as e:
            raise Exception(e)
        
    def get_element_visible(self, element: tuple | None = None,  web_element: WebElement | None = None) -> bool:
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
            
            except ElementClickInterceptedException as e:
                if i == 10:
                    return False
                
                self.driver.execute_script('scrollBy(0,100)')
                continue

    def get_click(self, element: tuple, valor: str) -> None:
        for i in range(10):
            self.driver.find_element(*element).click()
            sleep(3)
            content: str = self.driver.find_element(*self.body).text

            if valor in content:
                break
            
            else:
                if i == 10:
                    raise Exception('Element not interactable')
                
                sleep(2)
                continue

    def click(self, element, tempo):
        self.driver.find_element(*element).click()
        sleep(tempo)

    def send_keys(self, element, texto, tempo):
        self.driver.find_element(*element).send_keys(texto)
        sleep(tempo)

    def clear(self, element, tempo):
        self.driver.find_element(*element).clear()
        sleep(tempo)

    def elemento_existe(self, element: tuple):
        try: 
            self.driver.find_element(*element)
            return True
        except:
            return False

class Facil(PageElement):
    prestador_pj = (By.XPATH, '//*[@id="tipoAcesso"]/option[7]')
    usuario_input = (By.XPATH, '//*[@id="login-entry"]')
    senha_input = (By.XPATH, '//*[@id="password-entry"]')
    entrar = (By.XPATH, '//*[@id="BtnEntrar"]')
    faturas = (By.XPATH, '//*[@id="menuPrincipal"]/div/div[12]/a')
    relatorio_de_faturas = (By.XPATH, '/html/body/header/div[4]/div/div/div/div[12]/div[1]/div[2]/div/div[2]/div/div/div/div[1]/a')
    body = (By.XPATH, '/html/body')
    codigo = (By.XPATH, '/html/body/main/div/div[1]/div[2]/div/div/div[2]/div[1]/div[1]/input-text[1]/div/div/input')
    pesquisar = (By.XPATH, '//*[@id="filtro"]/div[2]/div[2]/button')
    recurso_de_glosa = (By.XPATH, '/html/body/main/div/div[1]/div[4]/div/div/div[1]/div/div[2]/a[5]/i')
    table = (By.ID, 'recursoGlosaTabelaServicos')
    text_area_justificativa = (By.ID, 'txtJustificativa')
    button_ok = ()
    salvar_parcialmente = ()
    i_close = ()
    proxima_pagina = ()
    label_registros = ()
    primeira_pagina = ()
    fechar = ()
    ul = ()
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
        self.driver.find_element(*self.usuario_input).send_keys(self.usuario)
        self.driver.find_element(*self.senha_input).send_keys(self.senha)
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
        self.open()
        self.exe_login()
        self.exe_caminho()
        
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

            self.inicializar_atributos(recurso_iniciado)
            self.abrir_guias()

            qtd_registros = int(self.driver.find_element(*self.label_registros).text.split(' ')[9])
            numero_paginas = int(qtd_registros / 5)

            for _ in range(0, numero_paginas + 1):
                tamanho_table = len(pd.read_html(self.driver.find_element(*self.table).get_attribute('outerHTML'))[0])
                qtd_itens = int(self.driver.find_element(*self.label_registros).text.split(' ')[4])
                self.lançar_recursos(tamanho_table+1, recurso_iniciado, df, planilha)
                
                if qtd_itens != qtd_registros:
                    self.driver.find_element(*self.proxima_pagina).click()

            self.driver.find_element(*self.fechar).click()
            sleep(2)
            self.driver.get('https://novowebplanfascal.facilinformatica.com.br/GuiasTISS/Relatorios/ViewRelatorioServicos')    

    def abrir_guias(self):
        i_elements = [i_element for i_element in self.driver.find_elements(By.TAG_NAME, 'i') if i_element.get_attribute('class') == 'fa fa-plus fa-fw open-close-servicos']

        for i_element in i_elements:

            if i_element.get_attribute('style') == 'display: none;':
                continue
            else:
                i_element.click()

    def img_in_element(self, element):
        try:
            element.find_element(By.TAG_NAME, 'img')
            return True
        except:
            return False

    def inicializar_atributos(self, recurso_iniciado):
        if recurso_iniciado == False:
            self.button_ok = (By.XPATH, '/html/body/main/div/div[3]/div/div/div[3]/button[1]')
            self.salvar_parcialmente = (By.XPATH, '/html/body/main/div/div[2]/div/div/div[2]/div/div[1]/div/div/div[3]/div/div[2]/button')
            self.i_close = (By.XPATH, '/html/body/ul/li/div/div/span/h4/i')
            self.proxima_pagina = (By.XPATH, '/html/body/main/div/div[2]/div/div/div[2]/div/div[1]/div/div/div[3]/div/div[1]/div/nav/ul/li[6]/a/span')
            self.label_registros = (By.XPATH, '/html/body/main/div/div[2]/div/div/div[2]/div/div[1]/div/div/div[3]/div/div[1]/label')
            self.primeira_pagina = (By.XPATH, '/html/body/main/div/div[2]/div/div/div[2]/div/div[1]/div/div/div[3]/div/div[1]/div/nav/ul/li[1]/a/span')
            self.fechar = (By.XPATH, '/html/body/main/div/div[2]/div/div/div[3]/button')
            self.ul = (By.XPATH, '/html/body/main/div/div[2]/div/div/div[2]/div/div[1]/div/div/div[3]/div/div[1]/div/nav/ul')
        else:
            self.button_ok = (By.XPATH, '/html/body/main/div[1]/div[4]/div/div/div[3]/button[1]')
            self.salvar_parcialmente = (By.XPATH, '/html/body/main/div[1]/div[3]/div/div/div[2]/div/div[1]/div/div/div[3]/div/div[2]/button')
            self.i_close = (By.XPATH, '/html/body/ul/li/div/div/span/h4/i')
            self.proxima_pagina = (By.XPATH, '/html/body/main/div[1]/div[3]/div/div/div[2]/div/div[1]/div/div/div[3]/div/div[1]/div/nav/ul/li[8]/a')
            self.label_registros = (By.XPATH, '/html/body/main/div[1]/div[3]/div/div/div[2]/div/div[1]/div/div/div[3]/div/div[1]/label')
            self.primeira_pagina = (By.XPATH, '/html/body/main/div[1]/div[3]/div/div/div[2]/div/div[1]/div/div/div[3]/div/div[1]/div/nav/ul/li[1]/a/span')
            self.fechar = (By.XPATH, '/html/body/main/div[1]/div[3]/div/div/div[3]/button')
            self.ul = (By.XPATH, '/html/body/main/div[1]/div[3]/div/div/div[2]/div/div[1]/div/div/div[3]/div/div[1]/div/nav/ul')

    def get_values(self, i, recurso_iniciado):
        if recurso_iniciado == False:
            for j in range(10):
                try:
                    nro_guia_portal = self.driver.find_element(By.XPATH, f'/html/body/main/div/div[2]/div/div/div[2]/div/div[1]/div/div/div[3]/div/table/tbody/tr[{i}]/td[2]/span').text
                    codigo_proc_portal = self.driver.find_element(By.XPATH, f'/html/body/main/div/div[2]/div/div/div[2]/div/div[1]/div/div/div[3]/div/table/tbody/tr[{i}]/td[4]/span').text.replace('.', '').replace('-', '')
                    valor_glosa_portal = self.driver.find_element(By.XPATH, f'/html/body/main/div/div[2]/div/div/div[2]/div/div[1]/div/div/div[3]/div/table/tbody/tr[{i}]/td[6]/span').text
                    valor_recursado_portal = self.driver.find_element(By.XPATH, f'/html/body/main/div/div[2]/div/div/div[2]/div/div[1]/div/div/div[3]/div/table/tbody/tr[{i}]/td[8]/span').text
                    nome_paciente_portal = self.driver.find_element(By.XPATH, f'/html/body/main/div/div[2]/div/div/div[2]/div/div[1]/div/div/div[3]/div/table/tbody/tr[{i}]/td[3]/span').text.replace('-', '')
                    break
                except:
                    print("Variáveis não encontradas.")
                    time.sleep(2)
                    continue

        else:
            for j in range(10):
                try:
                    nro_guia_portal = self.driver.find_element(By.XPATH, f'/html/body/main/div[1]/div[3]/div/div/div[2]/div/div[1]/div/div/div[3]/div/table/tbody/tr[{i}]/td[2]/span').text
                    codigo_proc_portal = self.driver.find_element(By.XPATH, f'/html/body/main/div[1]/div[3]/div/div/div[2]/div/div[1]/div/div/div[3]/div/table/tbody/tr[{i}]/td[4]/span').text.replace('.', '').replace('-', '')
                    valor_glosa_portal = self.driver.find_element(By.XPATH, f'/html/body/main/div[1]/div[3]/div/div/div[2]/div/div[1]/div/div/div[3]/div/table/tbody/tr[{i}]/td[6]/span').text
                    valor_recursado_portal = self.driver.find_element(By.XPATH, f'/html/body/main/div[1]/div[3]/div/div/div[2]/div/div[1]/div/div/div[3]/div/table/tbody/tr[{i}]/td[8]/span').text
                    nome_paciente_portal = self.driver.find_element(By.XPATH, f'/html/body/main/div[1]/div[3]/div/div/div[2]/div/div[1]/div/div/div[3]/div/table/tbody/tr[{i}]/td[3]/span').text.replace('-', '')
                    break
                except:
                    print("Variáveis não encontradas.")
                    time.sleep(2)
        return (nro_guia_portal, codigo_proc_portal, valor_glosa_portal, valor_recursado_portal, nome_paciente_portal)
    
    def lançar_recursos(self, tamanho_table, recurso_iniciado, df, planilha):
        for i in range(1, tamanho_table):
            nro_guia_portal, codigo_proc_portal, valor_glosa_portal, valor_recursado_portal, nome_paciente_portal = self.get_values(i, recurso_iniciado)

            for index, linha in df.iterrows():
                if f"{linha['Recursado no Portal']}" == "Sim" or f"{linha['Recursado no Portal']}" == "Não":
                    continue
                nome_paciente = f'{linha["Paciente"]}'
                numero_guia = f'{linha["Nro. Guia"]}'.replace('.0', '')
                codigo_procedimento = f'{linha["Procedimento"]}'.replace('.0', '')
                valor_glosa = f'{linha["Valor Glosa"]}'
                valor_recurso = f'{linha["Valor Recursado"]}'
                justificativa = f'{linha["Recurso Glosa"]}'.replace('\t', ' ')
                tabela_convenio = f'{linha["Tabela"]}'
                matricula = f'{linha["Matrícula"]}'

                validacao_paciente = nome_paciente in nome_paciente_portal
                validacao_matricula = matricula in nome_paciente_portal
                validacao_numero_guia = numero_guia in nro_guia_portal
                validacao_valor_glosa = valor_glosa in valor_glosa_portal
                validacao_valor_recursado = valor_recursado_portal == "R$0,00"
                validacao_codigo = codigo_procedimento in codigo_proc_portal
                validacao_codigo_taxa = "TAXAS" in tabela_convenio and "Taxa" in codigo_proc_portal

                validacao_normal = (validacao_numero_guia or validacao_paciente or validacao_matricula) and validacao_codigo and validacao_valor_glosa and validacao_valor_recursado
                validacao_taxa = (validacao_numero_guia or validacao_paciente or validacao_matricula) and validacao_codigo_taxa and validacao_valor_glosa and validacao_valor_recursado

                if validacao_normal or validacao_taxa:
                    input_valor_recursado, preencher_justificativa = self.xpath_preencher_valores(i, recurso_iniciado)
                    self.driver.find_element(*input_valor_recursado).send_keys(valor_recurso)
                    time.sleep(2)
                    self.driver.find_element(*preencher_justificativa).click()
                    time.sleep(2)

                    for _ in range(0, 10):
                        try:
                            self.driver.find_element(*self.text_area_justificativa).send_keys(justificativa)
                            break
                        except:
                            time.sleep(2)
                            continue

                    time.sleep(2)
                    self.driver.find_element(*self.button_ok).click()
                    time.sleep(2)
                    self.driver.find_element(*self.salvar_parcialmente).click()
                    time.sleep(2)
                    self.driver.find_element(*self.i_close).click()
                    dados = {"Recursado no Portal" : ['Sim']}
                    df_dados = pd.DataFrame(dados)
                    book = load_workbook(planilha)
                    writer = pd.ExcelWriter(planilha, engine='openpyxl')
                    writer.book = book
                    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
                    df_dados.to_excel(writer, 'Recurso', startrow=index + 1, startcol=21, header=False, index=False)
                    writer.save()
                    break
    
    def xpath_preencher_valores(self, i, recurso_iniciado):
        if recurso_iniciado == False:
            input_valor_recursado = (By.XPATH, f'/html/body/main/div/div[2]/div/div/div[2]/div/div[1]/div/div/div[3]/div/table/tbody/tr[{i}]/td[7]/input')
            preencher_justificativa = (By.XPATH, f'/html/body/main/div/div[2]/div/div/div[2]/div/div[1]/div/div/div[3]/div/table/tbody/tr[{i}]/td[9]/a/i')
        else:
            input_valor_recursado = (By.XPATH, f'/html/body/main/div[1]/div[3]/div/div/div[2]/div/div[1]/div/div/div[3]/div/table/tbody/tr[{i}]/td[7]/input')
            preencher_justificativa = (By.XPATH, f'/html/body/main/div[1]/div[3]/div/div/div[2]/div/div[1]/div/div/div[3]/div/table/tbody/tr[{i}]/td[9]/a/i')
        return (input_valor_recursado, preencher_justificativa)

def recursar_fascal(user, password):
    try:
        url = 'https://novowebplanfascal.facilinformatica.com.br/GuiasTISS/Logon'
        pasta = filedialog.askdirectory()

        options = {
            'proxy' : {
                'http': f'http://{user}:{password}@10.0.0.230:3128',
                'https': f'http://{user}:{password}@10.0.0.230:3128'
            },
            'verify_ssl': False
        }

        chrome_options = Options()
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument('--ignore-certificate-errors')
        try:
            servico = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=servico, seleniumwire_options= options, options = chrome_options)
        except Exception as err:
            print(err)
            driver = webdriver.Chrome(seleniumwire_options= options, options = chrome_options)
        usuario = "AMHPDF-ADM"
        senha = "00735860000173"
        Facil(driver, url, usuario, senha, pasta).executar_recurso()
    
    except Exception as e:
        tkinter.messagebox.showerror( 'Erro Automação' , f'Ocorreu uma excessão não tratada \n {e.__class__.__name__}: {e}' )
        driver.quit()

if __name__ == '__main__':
    recursar_fascal('lucas.paz', 'WheySesc2024*')