from abc import ABC
from datetime import date
from os import listdir
from tkinter.filedialog import askdirectory
from pandas import DataFrame, concat, read_excel, read_html
from copy import deepcopy
from selenium.webdriver.remote.webelement import WebElement
from selenium import webdriver
from selenium.webdriver.chrome.webdriver import WebDriver
from seleniumwire import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import time
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
                time.sleep(0.5)
                valor_inserido: str = self.driver.find_element(*element).get_attribute('value')
                count += 1

                if count == 10:
                    raise Exception("Element not interactable")

        except Exception as e:
            raise Exception(e)
        
    def get_element_visible(self, element: tuple) -> bool:
        """Este método observa se irá ocorrer ElementClickInterceptedException. Caso ocorra
        irá dar um scroll até 10x na página conforme o comando passado até achar o click do elemento"""
        for i in range(10):
            try:
                self.driver.find_element(*element).click()
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

class ConnectMed(PageElement):
    data_atual = date.today()
    input_usuario = (By.ID, 'username')
    input_senha = (By.ID, 'password')
    button_entrar = (By.ID, 'submitPrestador')
    extrato = (By.LINK_TEXT, 'Extrato')
    visualizar = (By.LINK_TEXT, 'Visualizar')
    abrir_filtro_extrato = (By.ID, 'abrir-fechar')
    opt_60_dias = (By.XPATH, '//*[@id="cmbPeriodo"]/option[3]')
    btn_consultar = (By.ID, 'btnConsultarExtratoPeriodo')
    button_detalhar_extrato = (By.ID, 'linkDetalharExtrato')
    option_glosados = (By.XPATH, '//*[@id="dadosLotesRecursoAberto_statusRecurso"]/option[2]')
    input_conta_prestador = (By.ID, 'dadosLotesRecursoAberto_numeroContaPrestador')
    input_buscar = (By.ID, 'dadosLotesRecursoAberto_btnBuscar')
    table_guias = (By.ID, 'dadosLotesRecursoAberto-contas_grid')
    td_guia = (By.XPATH, '/html/body/div[2]/div[1]/div[2]/div[2]/div[2]/div[2]/fieldset/div[1]/div/div[1]/div[1]/div/div/div[3]/div[3]/div/table/tbody/tr[2]/td[4]')
    table_proc = (By.ID, 'dadosLotesRecursoAberto-procedimentos_grid')
    text_area_resposta = (By.ID, 'dadosLotesRecursoAberto_mensagem')
    input_file = (By.XPATH, '/html/body/div[2]/div[1]/div[2]/div[2]/div[2]/div[3]/div[1]/fieldset/div[1]/div/div[2]/fieldset/div/div/div/div/form/div/div[4]/input')
    # tabela_de_extratos = (By.CLASS_NAME, 'span-22 tab-administracao size-11')

    def __init__(self, driver: WebDriver, url: str, usuario: str, senha: str, diretorio: str='') -> None:
        super().__init__(driver, url)
        self.diretorio: str = diretorio
        self.usuario: str = usuario
        self.senha: str = senha
        self.dados_planilhas: list[dict] = [
            {
                'caminho': f'{diretorio}\\{arquivo}',
                'lote': arquivo.replace('.xlsx', '')
            } 
            for arquivo in listdir(diretorio)
            if arquivo.endswith('.xlsx')
            ]

    def login(self) -> None:
        self.driver.find_element(*self.input_usuario).send_keys(self.usuario)
        time.sleep(2)
        self.driver.find_element(*self.input_senha).send_keys(self.senha)
        time.sleep(2)
        self.driver.find_element(*self.button_entrar).click()

    def acessar_extrato(self):
        self.driver.find_element(*self.extrato).click()
        time.sleep(2)
        self.driver.find_element(*self.visualizar).click()
        time.sleep(2)

    def get_extrato_df(self) -> DataFrame:
        df_extrato = read_html(self.driver.find_element(*self.table).get_attribute('outerHTML'))[0]
        return df_extrato
    
    def get_lote_no_extrato(self) -> list:
        tables = self.driver.find_elements(*self.table)
        lista_df: list[DataFrame] = []

        for table in tables:
            table_df = read_html(table.get_attribute('outerHTML'))[0]
            lista_df.append(table_df)
        
        return concat(lista_df)['Capa de Lote'].astype(str).values.tolist()
    
    def get_planilhas_dos_protocolos(self) -> list[str]:
        protocolos = self.get_lote_no_extrato()
        return [dado for dado in self.dados_planilhas if dado['lote'] in protocolos]
    
    def filtrar_guia(self, numero_guia: str):
        self.driver.find_element(*self.option_glosados).click()
        time.sleep(2)
        self.driver.find_element(*self.input_conta_prestador).send_keys(numero_guia)
        time.sleep(2)
        self.driver.find_element(*self.input_buscar).click()

    def converter_numero_para_string(self, valor: int | float | str) -> str:
        if isinstance(valor, float) or isinstance(valor, int):
            return "{:.2f}".format(valor).replace('.', ',')

        else:
            return "{:.2f}".format(float(valor.replace('.', '').replace(',', '.')))

    def is_recurso(self, procedimento: str, valor_glosado: str, valor_recurso: str) -> bool:
        table_procedimentos_outer: str = self.driver.find_element(*self.table_proc).get_attribute('outerHTML')
        xaover_essa_table_meno = read_html(table_procedimentos_outer)[0]
        table_procedimentos_lenth: int = len(read_html(table_procedimentos_outer)[0])

        for i in range(2, table_procedimentos_lenth + 1):
            tr_proc = (By.XPATH, 
            f'/html/body/div[2]/div[1]/div[2]/div[2]/div[2]/div[2]/fieldset/div[1]/div/div[2]/fieldset/div/div/div/div/div/div[3]/div[3]/div/table/tbody/tr[{i}]/td[21]/div')

            procedimento_portal: WebElement = self.driver.find_element(*tr_proc)

            valor_glosado_portal: str = self.driver.find_element(By.XPATH,
            f'/html/body/div[2]/div[1]/div[2]/div[2]/div[2]/div[2]/fieldset/div[1]/div/div[2]/fieldset/div/div/div/div/div/div[3]/div[3]/div/table/tbody/tr[{i}]/td[39]').text

            checkbox_procedimento: tuple = (By.XPATH,
            f'/html/body/div[2]/div[1]/div[2]/div[2]/div[2]/div[2]/fieldset/div[1]/div/div[2]/fieldset/div/div/div/div/div/div[3]/div[3]/div/table/tbody/tr[{i}]/td[1]/input')

            valor_a_recursar_proc: WebElement = self.driver.find_element(By.XPATH,
            f'/html/body/div[2]/div[1]/div[2]/div[2]/div[2]/div[2]/fieldset/div[1]/div/div[2]/fieldset/div/div/div/div/div/div[3]/div[3]/div/table/tbody/tr[{i}]/td[41]')

            valor_a_recursar_editabel: tuple = (By.XPATH,
            f'/html/body/div[2]/div[1]/div[2]/div[2]/div[2]/div[2]/fieldset/div[1]/div/div[2]/fieldset/div/div/div/div/div/div[3]/div[3]/div/table/tbody/tr[{i}]/td[41]/input')


            if procedimento_portal.text == procedimento and valor_glosado_portal.replace('R$ ', '') == valor_glosado and valor_a_recursar_proc.text == 'R$0.00':
                self.get_element_visible(tr_proc)
                time.sleep(2)
                self.driver.find_element(*checkbox_procedimento).click()
                time.sleep(1)
                valor_a_recursar_proc.click()
                time.sleep(1)
                valor_a_recursar_proc.find_element(*valor_a_recursar_editabel).clear()
                valor_a_recursar_proc.find_element(*valor_a_recursar_editabel).send_keys(valor_recurso)
                time.sleep(1)
                return True
        
        return False

    def to_be_named(self):
        self.driver.implicitly_wait(30)
        self.open()
        self.login()
        self.acessar_extrato()
        self.driver.find_element(*self.abrir_filtro_extrato).click()
        time.sleep(2)
        self.driver.find_element(*self.opt_60_dias).click()
        time.sleep(2)
        self.driver.find_element(*self.btn_consultar).click()
        time.sleep(2)
        df_extrato = self.get_extrato_df()

        for index, linha in df_extrato.iterrows():
            mes_extrato: int = int(f"{linha['Extrato']}".split('/')[1])

            if not mes_extrato == self.data_atual.month - 1:
                continue

            lupa_extrato = (By.XPATH, f'/html/body/div[2]/div/div/div[2]/div[1]/table/tbody/tr[{index+1}]/td[5]/form/a')
            self.driver.find_element(*lupa_extrato).click()
            time.sleep(2)
            self.driver.find_element(*self.button_detalhar_extrato).click()
            time.sleep(2)
            lista_de_dados_no_extrato = self.get_planilhas_dos_protocolos()

            for dado in lista_de_dados_no_extrato:
                caminho = dado['caminho']
                lote = dado['lote']

                id = f'formularioBuscarLote_{lote}'

                self.driver.find_element(By.ID, id).click()
                time.sleep(2)

                df_planilha = read_excel(caminho)

                for i, l in df_planilha.iterrows():
                    numero_guia = f"{l['Nro. Guia']}".replace('.0', '')
                    codigo_procedimento = f'{l["Procedimento"]}'.replace('.0', '')
                    justificativa = f'{l["Recurso Glosa"]}'.replace('\t', ' ')
                    valor_glosa = self.converter_numero_para_string(l['Valor Glosa']).replace('-', '')
                    valor_recurso = self.converter_numero_para_string(l['Valor Recursado'])
                    tabela_guias = self.driver.find_element(*self.table_guias)

                    if numero_guia not in tabela_guias.text:
                        continue

                    self.filtrar_guia(numero_guia)
                    self.driver.find_element(*self.td_guia).click()

                    if not self.is_recurso(codigo_procedimento, valor_glosa, valor_recurso):
                        ... #TODO lançar na planilha que o procedimento não passou pela validação
                    
                    self.driver.find_element(*self.text_area_resposta).send_keys(justificativa)
                    time.sleep(1)
                    self.driver.find_element(*self.input_file).send_keys() #ver como fazer para enviar o arquivo da guia
                    time.sleep(1)
                    ...

def recursar(user, password):
    diretorio = askdirectory()

    url = 'https://wwwt.connectmed.com.br/conectividade/prestador/home.htm'

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
    try:
        servico = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=servico, seleniumwire_options= options, options = chrome_options)
    except:
        driver = webdriver.Chrome(seleniumwire_options= options, options = chrome_options)

    usuario = "glosaamhp"
    senha = "D8nEUqb!"

    connect_med = ConnectMed(driver, url, usuario, senha, diretorio)
    connect_med.to_be_named()

if __name__ == '__main__':
    recursar('lucas.paz', 'WheySesc2024*')