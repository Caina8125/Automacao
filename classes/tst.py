from abc import ABC
from os import listdir
from time import sleep
from tkinter import filedialog
from pandas import read_excel, read_html
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium import webdriver
from seleniumwire import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By

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
        
    def get_element_visible(self, element: tuple | None = None,  web_element = None) -> bool:
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
            sleep(3)
            content: str = self.driver.find_element(*self.body).text

            if valor in content:
                break
            
            else:
                if i == 10:
                    raise Exception('Element not interactable')
                
                sleep(2)
                continue

class Tst(PageElement):
    input_usuario = (By.XPATH, '/html/body/div[2]/div[2]/form/table/tbody/tr[2]/td[2]/input')
    input_senha = (By.XPATH, '/html/body/div[2]/div[2]/form/table/tbody/tr[3]/td[2]/input')
    btn_entrar = (By.XPATH, '/html/body/div[2]/div[2]/form/table/tbody/tr[4]/td[2]/input')
    faturamento = (By.XPATH, '/html/body/div[1]/div[2]/ul/li[2]/a')
    manter_demonstrativo_de_pag = (By.XPATH, '/html/body/div[1]/div[2]/ul/li[2]/ul/li[3]/a')
    input_numero_lote = (By.ID, 'numeroLote')
    btn_consultar = (By.ID, 'botaoConsultar')
    visualizar_demonst_pagamento = (By.XPATH, '/html/body/div[2]/div[2]/div/div[2]/form/table[2]/tbody/tr/td[10]/a[1]/img')
    clique_aqui = (By.LINK_TEXT, 'clique aqui')
    select_tipo_guia = (By.ID, 'codigoTipoGuia')
    adicionar_guias = (By.XPATH, '/html/body/div[2]/div[2]/div/form/div[1]/fieldset/table/tbody/tr/td/fieldset[2]/table/tbody/tr/td/input')
    table_guias = (By.ID, 'guia')
    btn_adicionar = (By.XPATH, '/html/body/div[2]/div[2]/div/form/div[2]/div[11]/div/button[2]/span')

    def __init__(self, driver: WebDriver, usuario: str, senha: str, diretorio: str, url: str = '') -> None:
        super().__init__(driver, url)
        self.usuario = usuario
        self.senha = senha
        self.lista_de_arquivos = [
            {
                'numero_processo': arquivo.replace('.xlsx', ''),
                'path': f'{diretorio}\\{arquivo}',
                }
             for arquivo in listdir(diretorio)
             if arquivo.endswith('.xlsx')]

    def login(self):
        self.driver.find_element(*self.input_usuario).send_keys(self.usuario)
        sleep(1.5)
        self.driver.find_element(*self.input_senha).send_keys(self.senha)
        sleep(1.5)
        self.driver.find_element(*self.btn_entrar).click()
        sleep(1.5)

    def caminho(self):
        self.driver.find_element(*self.faturamento).click()
        sleep(1.5)
        self.driver.find_element(*self.manter_demonstrativo_de_pag).click()
        sleep(1.5)

    def to_be_named(self):
        self.open()
        self.login()
        self.caminho()

        for arquivo in self.lista_de_arquivos:
            numero_processo = arquivo['numero_processo']
            path_planilha = arquivo['path']
            df_processo = read_excel(path_planilha)
            lista_de_guias = set(df_processo['Controle Inicial'].astype(str).values.tolist())
            tipo_guia = self.pegar_xpath_tipo_guia(int(df_processo['Tipo Guia'][0]))
            self.driver.find_element(*self.input_numero_lote).send_keys(numero_processo)
            sleep(1.5)
            self.driver.find_element(*self.btn_consultar).click()
            sleep(1.5)
            self.driver.find_element(*self.visualizar_demonst_pagamento).click()
            sleep(1.5)
            self.driver.find_element(*self.clique_aqui).click()
            sleep(1.5)
            self.driver.find_element(*self.select_tipo_guia).click()
            sleep(1.5)
            self.driver.find_element(By.PARTIAL_LINK_TEXT, tipo_guia).click()
            sleep(1.5)
            self.driver.find_element(*self.adicionar_guias).click()
            sleep(1.5)

            while 'Aguarde, consultando...' in self.driver.find_element(*self.body).text:
                sleep(2)
            
            df_tabela_guia = read_html(self.driver.find_element(*self.table_guias).get_attribute('outerHTML'), header=0)[0]

            for ind, l in df_tabela_guia.iterrows():
                n_guia_prestador = f"{l['N° da Guia do Prestador']}".replace('.0', '')

                if n_guia_prestador in lista_de_guias:
                    self.driver.find_element(By.XPATH, f'/html/body/div[2]/div[2]/div/form/div[2]/div[2]/fieldset/div/div/table/tbody/tr[1]/td[1]/input[{ind+2}]').click()
                    sleep(1)

            self.driver.find_element(*self.btn_adicionar).click()
            sleep(2)
            count = 1
            self.driver.find_element(*self.input_numero_lote).send_keys(f'{numero_processo}-{count}')

            while 'Aguarde, processando...' in self.driver.find_element(*self.body).text:
                sleep(2)
    
    def pegar_xpath_tipo_guia(self, codigo_tipo_guia):
        match codigo_tipo_guia:
            case 1:
                return '/html/body/div[2]/div[2]/div/form/div[1]/fieldset/table/tbody/tr/td/fieldset[1]/select/option[2]'
            case 2:
                return '/html/body/div[2]/div[2]/div/form/div[1]/fieldset/table/tbody/tr/td/fieldset[1]/select/option[3]'
            case 3:
                return '/html/body/div[2]/div[2]/div/form/div[1]/fieldset/table/tbody/tr/td/fieldset[1]/select/option[4]'
            case 4:
                return '/html/body/div[2]/div[2]/div/form/div[1]/fieldset/table/tbody/tr/td/fieldset[1]/select/option[5]'
            case _:
                raise Exception('Tipo de guia inválido!')

if __name__ == '__main__':
    user = 'lucas.paz'
    password = 'WheySesc2024*'
    url = ' http://aplicacao7.tst.jus.br/tstsaude'
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
    except:
        driver = webdriver.Chrome(seleniumwire_options= options, options = chrome_options)
    usuario = "U000971"
    senha = "X11254S"
    tst = Tst(driver, usuario, senha, pasta, url)
    tst.to_be_named()
    driver.quit()