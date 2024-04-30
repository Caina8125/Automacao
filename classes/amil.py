from abc import ABC
from os import listdir
from tkinter.filedialog import askdirectory
from pandas import DataFrame, read_excel, read_html
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.remote.webelement import WebElement
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
    utlimo_mes = By.XPATH, '/html/body/div/div/form/div[4]/table/tbody/tr[1]/td[1]/input'
    table_primeiro_mes = By.XPATH, '/html/body/div/div/form/div[4]/table/tbody/tr[2]/td/div/table'
    btn_avancar = By.ID, 'btnavancar'
    table_guias = By.XPATH, '/html/body/div/div/form/div[3]/div[4]/table'
    td_n_guia_prestador = By.XPATH, '/html/body/form/table/tbody/tr[9]/td/table/tbody/tr[2]/td[1]/font'

    def __init__(self, usuario: str, senha: str, diretorio: str, driver: WebDriver, url: str = '') -> None:
        super().__init__(driver, url)
        self.usuario = usuario
        self.senha = senha
        self.diretorio = diretorio
        self.lista_de_arquivos = [f"{diretorio}\\{arquivo}" for arquivo in listdir(diretorio) if arquivo.endswith('.xlsx')]
    
    def login(self):
        self.send_keys(self.input_usuario, self.usuario, 1.5)
        self.send_keys(self.input_senha, self.senha, 1.5)
        self.click(self.btn_entrar, 1.5)

    def caminho(self):
        self.click(self.acessar_sis_amil, 2)
        self.driver.switch_to.window(-1)
        self.click(self.menu, 2)
        self.send_keys(self.input_menu, 'Solicitação de Recurso de Glosas', 1)
        self.driver.switch_to.frame('principal')
        self.click(self.solicitacao_de_recurso, 2)

        for arquivo in self.lista_de_arquivos:
            df_fatura = read_excel(arquivo)
            lista_de_guias = list(set(df_fatura['Controle'].values.tolist()))
            numero_processo = arquivo.replace(self.diretorio, '').replace('.xlsx', '')
            self.click(self.option_amil, 2)
            self.click(self.utlimo_mes, 2)

            qtd = self.tamanho_tabela(self.table_primeiro_mes) + 1

            xpath_input_lote = self.encontrar_xpath_processo(numero_processo, qtd)

            if xpath_input_lote == '':
                #TODO acrescentar não enviado no nome
                continue
            
            self.click((By.XPATH, xpath_input_lote), 1)
            self.click(self.btn_avancar, 2)

            tamanho_tabela_guias = self.tamanho_tabela(self.table_guias)

    
    def encontrar_xpath_processo(self, numero_processo, qtd):
        for i in range(1, qtd):
            input_radio_lote = f'/html/body/div/div/form/div[4]/table/tbody/tr[2]/td/div/table/tbody/tr[{i}]/td[1]'
            n_lote = self.driver.find_element(*input_radio_lote).text

            if n_lote != numero_processo:
                continue

            return input_radio_lote
        
        return ''

    def tamanho_tabela(self, tabela):
        return len(read_html(self.driver.find_element(*tabela).get_attribute('outerHTML')[0]))
    
    def teste(self, lista, df_fatura: DataFrame):
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
                lista_procedimento.append(row[col])

            dict_guias_list.append(
                {num: {
                    'numeros': [numero_amhptiss, numero_nro_guia, numero_guia_prestador, numero_controle_inicial],
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
                self.driver.switch_to.window(-1)
                self.driver.switch_to.frame('principal2')
                num_guia_portal = self.driver.find_element(*self.td_n_guia_prestador).text

                if self.num_in_dict(num_guia_portal):
                    lista_numeros_clicar.append()

    def num_in_dict(self, num):
        for dictionary in self.dict_guias_list:

            for _, value in dictionary.items():

                if num in value:
                    return True
                
        return False
    
diretorio = askdirectory()
amil = Amil('', '', diretorio, '', '')
df = read_excel(r"C:\Users\lucas.paz\Documents\Nova pasta\PEG_68523.xlsx")
lista = list(set(df['Controle'].values.tolist()))
amil.teste(lista, df)