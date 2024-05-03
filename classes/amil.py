from abc import ABC
from os import listdir
from tkinter.filedialog import askdirectory
from pandas import DataFrame, read_excel, read_html
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from seleniumwire import webdriver
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

    def click(self, element, tempo):
        self.driver.find_element(*element).click()
        sleep(tempo)

    def send_keys(self, element, texto, tempo):
        self.driver.find_element(*element).send_keys(texto)
        sleep(tempo)

    def clear(self, element, tempo):
        self.driver.find_element(*element).clear()
        sleep(tempo)

class Amil(PageElement):
    dict_guias_list = [{}]
    input_usuario = By.ID, 'login-usuario'
    input_senha = By.ID, 'login-senha'
    btn_entrar = By.XPATH, '/html/body/as-main-app/as-login-container/div[1]/div/as-login/div[2]/form/fieldset/button'
    acessar_sis_amil = By.XPATH, '/html/body/as-main-app/as-home/as-base-layout/section/div/as-navbar/nav/div[2]/form/button'
    menu = By.ID, 'mostraMenu'
    input_menu = By.ID, 'txtProcuraMenu'
    solicitacao_de_recurso = By.LINK_TEXT, 'Solicitação de Recurso de Glosas'
    option_amil = By.XPATH, '/html/body/div/div/form/div[3]/div/div/select/option[2]'
    segundo_mes = By.XPATH, '/html/body/div/div/form/div[4]/table/tbody/tr[3]/td[1]/input'
    table_segundo_mes = By.XPATH, '/html/body/div/div/form/div[4]/table/tbody/tr[4]/td/div/table'
    btn_avancar = By.ID, 'btnavancar'
    table_guias = By.XPATH, '/html/body/div/div/form/div[3]/div[4]/table'
    td_n_guia_prestador = By.XPATH, '/html/body/form/table/tbody/tr[9]/td/table/tbody/tr[2]/td[1]/font'
    fechar_dicas = By.ID, 'finalizar-walktour'
    procurar_texto_menu = By.ID, 'lnkProcuraMenu'

    def __init__(self, usuario: str, senha: str, diretorio: str, driver: WebDriver, url: str = '') -> None:
        super().__init__(driver, url)
        self.usuario = usuario
        self.senha = senha
        self.diretorio = diretorio
        self.lista_de_arquivos = [f"{diretorio}\\{arquivo}" for arquivo in listdir(diretorio) if arquivo.endswith('.xlsx')]
        self.driver.implicitly_wait(30)
    
    def login(self):
        self.click(self.fechar_dicas, 1)
        self.send_keys(self.input_usuario, self.usuario, 1.5)
        self.send_keys(self.input_senha, self.senha, 1.5)
        self.click(self.btn_entrar, 1.5)

    def caminho(self):
        self.get_element_visible(self.acessar_sis_amil)
        sleep(2)
        self.driver.switch_to.window(self.driver.window_handles[-1])
        self.click(self.menu, 2)
        self.driver.switch_to.frame('menu')
        self.send_keys(self.input_menu, 'Solicitação de Recurso de Glosas', 1)
        self.click(self.procurar_texto_menu, 1)
        self.click(self.solicitacao_de_recurso, 2)
        self.driver.switch_to.default_content()
        self.driver.switch_to.frame('principal')

    def exec_recurso(self):
        self.open()
        self.login()
        self.caminho()

        for arquivo in self.lista_de_arquivos:
            df_fatura = read_excel(arquivo)
            lista_de_guias = list(set(df_fatura['Controle'].values.tolist()))
            numero_processo = arquivo.replace(self.diretorio+'\\', '').replace('.xlsx', '')
            self.dict_guias_list = self.dict_guias(lista_de_guias, df_fatura)
            self.click(self.option_amil, 2)
            self.click(self.segundo_mes, 2)

            qtd = self.tamanho_tabela(self.table_segundo_mes) + 1

            xpath_input_lote = self.encontrar_xpath_processo(numero_processo, qtd)

            if xpath_input_lote == '':
                #TODO acrescentar não enviado no nome
                continue
            
            self.click((By.XPATH, xpath_input_lote), 1)
            self.driver.switch_to.default_content()
            self.driver.switch_to.frame('toolbar')
            self.click(self.btn_avancar, 2)

            self.driver.switch_to.default_content()
            self.driver.switch_to.frame('principal')
            tamanho_tabela_guias = len(
                [
                    element
                    for element in self.driver.find_elements(By.TAG_NAME, 'input')
                    if element.get_attribute('type') == 'checkbox' and element.get_attribute('name') == 'contas'
                ]
            )
            quantidade_tr = self.calcular_quant_tr(tamanho_tabela_guias)

            self.selecionar_guias_e_procedimentos(quantidade_tr)
    
    def encontrar_xpath_processo(self, numero_processo, qtd):
        for i in range(1, qtd):
            td_lote = By.XPATH, f'/html/body/div/div/form/div[4]/table/tbody/tr[4]/td/div/table/tbody/tr[{i}]/td[1]'
            n_lote = self.driver.find_element(*td_lote).text.replace(' ', '')

            if n_lote != numero_processo:
                continue

            return f'/html/body/div/div/form/div[4]/table/tbody/tr[4]/td/div/table/tbody/tr[{i}]/td[1]/input'
        
        return ''

    def tamanho_tabela(self, tabela):
        return len(read_html(self.driver.find_element(*tabela).get_attribute('outerHTML'))[0])
    
    def dict_guias(self, lista, df_fatura: DataFrame):
        dict_guias_list = []
        for num in lista:
            df: DataFrame = df_fatura.loc[(df_fatura["Controle"] == int(num))]
            numero_amhptiss = f"{df['Amhptiss'].values.tolist()[0]}"
            numero_nro_guia = f"{df['Nro. Guia'].values.tolist()[0]}"
            numero_guia_prestador = f"{df['Guia Prestador'].values.tolist()[0]}"
            numero_controle_inicial = f"{df['Controle Inicial'].values.tolist()[0]}"

            dict_cols = {j: i for i, j in enumerate(df.columns)}
            lista_procedimento = []

            for row in df.values:
                col = dict_cols["Procedimento"]
                lista_procedimento.append(f'{row[col]}'.replace('.0', ''))

            dict_guias_list.append(
                {num: {
                    'numeros_da_guia': [numero_amhptiss, numero_nro_guia, numero_guia_prestador, numero_controle_inicial],
                    'procedimentos': lista_procedimento
                }}
            )
        return dict_guias_list
            
    def teste2(self, qtd):
        lista_numeros_clicar = []
        for i in range(1, qtd):
            td_guia = By.XPATH, f'/html/body/div/div/form/div[3]/div[4]/table/tbody/tr[{i}]/td[2]/a/u'
            numero_na_table = self.driver.find_element(*td_guia).text

            if self.num_in_dict(numero_na_table):
                lista_numeros_clicar.append(numero_na_table)

            else:
                self.click(td_guia, 1.5)
                self.driver.switch_to.window(self.driver.window_handles[-1])
                self.driver.switch_to.frame('principal2')
                num_guia_portal = self.driver.find_element(*self.td_n_guia_prestador).text

                if self.num_in_dict(num_guia_portal):
                    lista_numeros_clicar.append()

    def num_in_dict(self, num):
        for dictionary in self.dict_guias_list:

            for _, value in dictionary.items():

                if num in value['numeros_da_guia']:
                    return value['procedimentos']
                
        return []
    
    def calcular_quant_tr(self, len_tabela):
        count = 1
        for _ in range(1, len_tabela):
            count += 3
        return count + 1
    
    def selecionar_guias_e_procedimentos(self, quant_tr):
        for i in range(1, quant_tr, 3):
            u_element_guia = self.driver.find_element(By.XPATH, f'/html/body/div/div/form/div[3]/div[4]/table/tbody/tr[{i}]/td[2]/a/u')
            num_guia_portal = u_element_guia.text
            procedimentos_guia = self.num_in_dict(num_guia_portal)

            if len(procedimentos_guia) <= 0:
                u_element_guia.click()
                self.driver.switch_to.window(self.driver.window_handles[-1])
                self.driver.switch_to.frame('principal2')
                num_guia_portal = self.driver.find_element(*self.td_n_guia_prestador).text
                procedimentos_guia = self.num_in_dict(num_guia_portal)
                self.driver.close()
                self.driver.switch_to.window(self.driver.window_handles[-1])
                self.driver.switch_to.frame('principal')
                
                if len(procedimentos_guia) <= 0:
                    continue

            checkbox_guia = (By.XPATH, f'/html/body/div/div/form/div[3]/div[4]/table/tbody/tr[{i}]/td[1]/input')
            self.click(checkbox_guia, 2)
            #TODO tratar o erro da requisição

            table_procedimentos = By.XPATH, f'/html/body/div/div/form/div[3]/div[4]/table/tbody/tr[{i+1}]/td/div/table'
            quantidade_de_procedimentos = self.tamanho_tabela(table_procedimentos) + 1

            self.selecionar_procedimentos(quantidade_de_procedimentos, i+1, procedimentos_guia)

    def selecionar_procedimentos(self, qtd_proc, num_tr, lista_procedimentos):
        for i in range(1, qtd_proc):
            procedimento_guia_portal = self.driver.find_element(By.XPATH, f'/html/body/div/div/form/div[3]/div[4]/table/tbody/tr[{num_tr}]/td/div/table/tbody/tr[{i}]/td[4]').text
            checkbox_procedimento = By.XPATH, f'/html/body/div/div/form/div[3]/div[4]/table/tbody/tr[{num_tr}]/td/div/table/tbody/tr[{i}]/td[1]/input'

            if procedimento_guia_portal not in lista_procedimentos:
                self.click(checkbox_procedimento, 1)

user = 'lucas.paz'
password = 'WheySesc2024*'
url = 'https://credenciado.amil.com.br/login'
diretorio = askdirectory()

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

amil = Amil('10019642', 'Amhpdf2024', diretorio, driver, url)
amil.exec_recurso()