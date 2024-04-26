from abc import ABC
from os import listdir, rename
from tkinter import filedialog
from pandas import read_html
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.remote.webelement import WebElement
# from page_element import PageElement
from selenium.webdriver.common.by import By
from time import sleep
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from seleniumwire import webdriver

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

class Geap(PageElement):
    acessar_portal = By.XPATH, '/html/body/div[3]/div[3]/div[1]/form/div[1]/div[1]/div/a'
    usuario = By.XPATH, '/html/body/div[1]/div[2]/div/div/div[1]/div/div/div/div/div[2]/div/div[1]/div/label[1]/div/div[1]/div/input'
    input_senha = By.XPATH, '/html/body/div[1]/div[2]/div/div/div[1]/div/div/div/div/div[2]/div/div[1]/div/label[2]/div/div[1]/div[1]/input'
    entrar = By.XPATH, '/html/body/div[1]/div[2]/div/div/div[1]/div/div/div/div/div[2]/div/div[2]/button/div[2]/div/div'
    fechar = By.XPATH, '/html/body/div[4]/div[2]/div/div[3]/button'
    portal_tiss = (By.XPATH, '/html/body/div[1]/div/div[2]/div/div[1]/div[1]/div[1]/div/div')
    alerta = By.XPATH,' /html/body/div[2]/div/center/a'
    baixar_arquivo = By.XPATH, '/html/body/main/div/div[1]/div[2]/div/article/div[6]/div[4]/div[3]/div/div[2]/ul/li[2]/a'
    input_lote_prestador = By.XPATH, '/html/body/main/div/div/div[2]/div[2]/article/form/div/div[6]/input'
    btn_baixar = By.XPATH, '/html/body/main/div/div/div[2]/div[2]/article/form/div/a'
    detalhe = By.XPATH, '/html/body/main/div/div/div/table/tbody/tr[2]/td[9]/a/i'
    div_prestador = By.XPATH, '/html/body/main/div/div[1]/div[2]'
    input_file = By.XPATH, '/html/body/form/table/tbody/tr[7]/td/input'
    table_arquivo = By.XPATH, '/html/body/form/table/tbody/tr[9]/td/table'
    btn_salvar = By.XPATH, '/html/body/form/table/tbody/tr[11]/td/table/tbody/tr/td[1]/input'
    quantidade_de_nao_enviados = 0

    def __init__(self, driver: WebDriver, cpf: str, senha: str, url: str = '') -> None:
        super().__init__(driver, url)
        self.cpf = cpf
        self.senha = senha

    def exe_login(self):
        sleep(4)
        try:
            self.driver.find_element(*self.fechar).click()
        except:
            pass
        sleep(2)
        try:
            self.driver.implicitly_wait(15)
            self.driver.find_element(*self.acessar_portal).click()
            self.driver.implicitly_wait(4)
            self.driver.find_element(*self.usuario).send_keys(self.cpf)
            self.driver.find_element(*self.input_senha).click()
            sleep(2)
            self.driver.find_element(*self.input_senha).send_keys(self.senha)
            sleep(2)
            self.driver.find_element(*self.entrar).click()
            self.driver.find_element(*self.entrar).click()
            sleep(2)
            self.driver.find_element(*self.portal_tiss)
            sleep(2)
            self.driver.find_element(*self.portal_tiss).click()
            sleep(2)

        except Exception as e:
            self.driver.implicitly_wait(180)
            self.driver.find_element(*self.portal_tiss)
            sleep(2)
            self.driver.find_element(*self.portal_tiss).click()
            sleep(2)
            self.driver.implicitly_wait(15)


    def exe_caminho(self):
        self.driver.switch_to.window(self.driver.window_handles[-1])
        self.driver.implicitly_wait(4)
        try:
            self.driver.find_element(*self.alerta).click()
        except:
            print('Alerta não apareceu')
        self.driver.implicitly_wait(15)
        sleep(2)
        self.driver.find_element(*self.baixar_arquivo).click()
        sleep(2)
        self.driver.switch_to.window(self.driver.window_handles[-1])
    
    def exec_a_bagaca_toda(self):
        path_anexos_geap = r"\\10.0.0.239\automacao_integralis\ANEXOS GEAP FECHAMENTO"
        path_enviados = r"\\10.0.0.239\automacao_integralis\ANEXOS GEAP FECHAMENTO\Enviados"
        lista_de_diretorios = [pasta for pasta in listdir(path_anexos_geap) if pasta.isdigit()]
        self.open()
        self.exe_login()
        self.exe_caminho()
        
        for diretorio in lista_de_diretorios:
            path_diretorio = path_anexos_geap + '\\' + diretorio #Caminho do diretório do processo
            path_diretorio_enviados = path_enviados + '\\' + diretorio #Caminho que o diretório será jogado após finalizá-lo
            self.driver.find_element(*self.input_lote_prestador).send_keys(diretorio)
            sleep(1)
            self.driver.find_element(*self.btn_baixar).click()
            sleep(2)

            if 'Sua pesquisa não retornou nenhum registro!' in self.driver.find_element(*self.body).text:
                self.renomear_arquivo(path_diretorio, path_diretorio_enviados + "_nao_encontrado")
                self.driver.back()
                continue
                
            self.driver.switch_to.window(self.driver.window_handles[-1])
            self.driver.find_element(*self.detalhe).click()
            sleep(2)
            self.driver.switch_to.window(self.driver.window_handles[-1])
            self.driver.maximize_window()
            df_table_guias = read_html(self.driver.find_element(*self.table).get_attribute('outerHTML'), header=0)[0]
            dados_processo = self.get_dados_processo(path_diretorio)

            for dado in dados_processo:
                guia = dado['guia'].split('-')[0]
                path_arquivo = dado['path_arquivo']
                path_arquivo_enviado = path_diretorio + '\\' + dado['guia']

                if not self.guia_na_table(df_table_guias, guia):
                    self.renomear_arquivo(path_arquivo, path_arquivo_enviado + '_nao_encontrado.pdf')
                    self.quantidade_de_nao_enviados += 1
                    continue

                self.driver.switch_to.window(self.driver.window_handles[-1])
                table_content = self.driver.find_element(*self.table_arquivo).text

                if '.pdf' in table_content:
                    self.quantidade_de_nao_enviados += 1
                    self.renomear_arquivo(path_arquivo, path_arquivo_enviado + '_enviado_anteriormente.pdf')
                    continue

                self.driver.find_element(*self.input_file).send_keys(path_arquivo)
                sleep(1)
                self.driver.find_element(*self.btn_salvar).click()
                sleep(2)
                self.driver.close()
                self.driver.switch_to.window(self.driver.window_handles[-1])

            if self.quantidade_de_nao_enviados > 0:
                self.renomear_arquivo(path_diretorio, path_diretorio_enviados + '_parcialmente_enviado')
            else:
                self.renomear_arquivo(path_diretorio, path_diretorio_enviados)
                
            self.driver.close()
            self.driver.switch_to.window(self.driver.window_handles[-1])
            self.driver.back()
            self.driver.back()

    def get_dados_processo(self, path: str) -> dict:
        """Este método retorna um dict com as informações Nº Guia e o caminho do arquivo"""
        return [
            {
                'guia': arquivo.replace('.pdf', ''),
                'path_arquivo': path + '\\' + arquivo
            }
            for arquivo in listdir(path)
        ]
    
    def guia_na_table(self, df_table_guias, guia):
        for i, linha in df_table_guias.iterrows():
            if f"{linha['Cód Guia Prestador']}" == guia:
                self.driver.find_element(By.XPATH, f'/html/body/main/div/div/div/table/tbody/tr[{i+2}]/td[7]/a').click()
                return True
        return False
    
    @staticmethod
    def renomear_arquivo(path, path_novo) -> None:
        rename(path, path_novo)

if __name__ == '__main__':
    user = 'lucas.paz'
    password = 'WheySesc2024*'
    url = 'https://www2.geap.com.br/auth/prestadorVue.asp'

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

    geap = Geap(driver, '05613381160', 'Sarah608@', url)
    geap.exec_a_bagaca_toda()