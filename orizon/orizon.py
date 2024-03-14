from abc import ABC
from datetime import date, timedelta
from os import listdir, rename
from tkinter.filedialog import askdirectory
from pandas import DataFrame, concat, read_excel, read_html
from copy import deepcopy
from selenium import webdriver
from selenium.webdriver.chrome.webdriver import WebDriver
from seleniumwire import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
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
            sleep(3)
            content: str = self.driver.find_element(*self.body).text

            if valor in content:
                break
            
            else:
                if i == 10:
                    raise Exception('Element not interactable')
                
                sleep(2)
                continue

class Orizon(PageElement):
    data_atual = date.today()
    path_anexos_bradesco = r"C:\Users\lucas.paz\Documents\AnexosBradesco"
    path_enviados = r"C:\Users\lucas.paz\Documents\AnexosBradesco\Enviados"
    quantidade_de_nao_enviados = 0
    lista_de_diretorios = [pasta for pasta in listdir(path_anexos_bradesco) if pasta.isdigit()]
    input_usuario = (By.ID, 'username')
    input_senha = (By.ID, 'password')
    button_entrar = (By.ID, 'kc-login')
    terminar = (By.XPATH, '//*[@id="step-0"]/div[3]/button')
    servicos = (By.ID, 'linkTopbarModuloGlosas')
    modulo_de_glosas = (By.ID, 'ctl00_menutab_T2I')
    recurso_de_glosa = (By.ID, 'ctl00_menutab_ctl45_I2i1_T')
    input_data_inicio = (By.ID, 'ctl00_ContentBody_cbpPesquisa_dtePeriodoEnvioRecursoInicio_I')
    input_tipo_identificador = (By.ID, 'ctl00_ContentBody_cbpPesquisa_ddlFiltro_I')
    td_n_lote = (By.ID, 'ctl00_ContentBody_cbpPesquisa_ddlFiltro_DDD_L_LBI2T')
    input_valor_identificador = (By.ID, 'ctl00_ContentBody_cbpPesquisa_txtFiltro_I')
    btn_buscar = (By.ID, 'ctl00_ContentBody_cbpPesquisa_btnPesquisar_CD')
    div_ferramentas_lote = (By.ID, 'ctl00_ContentBody_cbpPesquisa_grdRecursoGlosa_tccell0_0')
    visualizar_guias = (By.ID, 'ctl00_ContentBody_ppmAcoes_DXI1_T')
    opcao_enviar_imagens = (By.ID, 'ctl00_ContentBody_cbpContas_pm_ListaGuias_DXI5_T')
    input_file = (By.ID, 'ctl00_ContentBody_ppcInformativo_FindFile')
    adicionar = (By.ID, 'ctl00_ContentBody_ppcInformativo_AddFile_CD')
    btn_enviar_imagens = (By.ID, 'ctl00_ContentBody_ppcInformativo_UploadFile_CD')
    btn_ok = (By.ID, 'ctl00_ContentBody_MessageBox1_ppcMessage_btnOk_CD')
    table_guias = (By.ID, 'ctl00_ContentBody_cbpContas_grdContas_DXMainTable')

    def __init__(self, driver: WebDriver, url: str, usuario: str, senha: str) -> None:
        super().__init__(driver, url)
        self.usuario = usuario
        self.senha = senha

    def login(self):
        self.driver.find_element(*self.input_usuario).send_keys(self.usuario)
        sleep(1)
        self.driver.find_element(*self.input_senha).send_keys(self.senha)
        sleep(1)
        self.driver.find_element(*self.button_entrar).click()
        sleep(2)

    def caminho(self):
        self.driver.find_element(*self.terminar).click()
        sleep(1)
        self.driver.find_element(*self.servicos).click()
        sleep(2)
        self.driver.switch_to.window(self.driver.window_handles[-1])
        self.driver.find_element(*self.modulo_de_glosas).click()
        sleep(1.5)
        self.driver.find_element(*self.recurso_de_glosa).click()
        sleep(1)

    def get_len_tabela(self) -> int:
        tabela = read_html(self.driver.find_element(*self.table_guias).get_attribute('outerHTML'), header=0)[0]
        valores: list = tabela.values.tolist()
        cabecalho = tabela.columns.values.tolist()
        array_textos = ['Ações', 'Guia Origem', 'Guia Operadora', 'Imagens', 'Senha', 'Valor Total Recursado', 'Qtde. Itens Recursado da Guia', 'Tipo de Recurso']
        dados = [lista for lista in valores if lista[0] not in array_textos]
        return len(dados)
    
    def get_dados_processo(self, path) -> tuple:
        return [
            {
                'guia': arquivo.replace('.pdf', ''),
                'path_arquivo': path + '\\' + arquivo
            }
            for arquivo in listdir(path)
        ]
    
    def img_in_element(self, element):
        self.driver.implicitly_wait(3)
        try:
            element.find_element(By.TAG_NAME, 'img')
            self.driver.implicitly_wait(30)
            return True
        except:
            self.driver.implicitly_wait(30)
            return False
    
    def acessar_enviar_imagem(self, quantidade_guias, guia, path, path_enviado) -> bool:
        for i in range(2, quantidade_guias + 2):
            td_guia = (By.XPATH, f'/html/body/form/div[3]/table[1]/tbody/tr[3]/td[3]/div[2]/div[4]/div[1]/table/tbody/tr/td/table[1]/tbody/tr[{i}]/td[3]')
            td_ferramentas = (By.XPATH, f'/html/body/form/div[3]/table[1]/tbody/tr[3]/td[3]/div[2]/div[4]/div[1]/table/tbody/tr/td/table[1]/tbody/tr[{i}]/td[1]')
            td_anexos = (By.XPATH, f'/html/body/form/div[3]/table[1]/tbody/tr[3]/td[3]/div[2]/div[4]/div[1]/table/tbody/tr/td/table[1]/tbody/tr[{i}]/td[4]')
            guia_portal = self.driver.find_element(*td_guia).text

            if not guia_portal == guia:
                continue

            element_td_imagem = self.driver.find_element(*td_anexos)

            if self.img_in_element(element_td_imagem):
                self.renomear_arquivo(path, path_enviado + '_enviado_anteriormente.pdf')
                self.quantidade_de_nao_enviados += 1
                return False

            self.driver.find_element(*td_ferramentas).click()
            sleep(1.5)
            self.driver.find_element(*self.opcao_enviar_imagens).click()
            sleep(1.5)
            return True
        
        self.renomear_arquivo(path, path_enviado + '_nao_encontrado.pdf')
        self.quantidade_de_nao_enviados += 1
        return False
    
    @staticmethod
    def renomear_arquivo(path, path_novo) -> None:
        rename(path, path_novo)

    def executar_anexo_de_guias(self):
        self.driver.implicitly_wait(30)
        self.open()
        self.login()
        self.caminho()

        for diretorio in self.lista_de_diretorios:
            path_diretorio = self.path_anexos_bradesco + '\\' + diretorio
            path_diretorio_enviados = self.path_enviados + '\\' + diretorio
            data_seis_meses_atras: str = (self.data_atual - timedelta(days=180)).strftime('%d/%m/%Y')
            self.driver.find_element(*self.input_data_inicio).clear()
            self.driver.find_element(*self.input_data_inicio).send_keys(data_seis_meses_atras)
            sleep(1.5)
            self.driver.find_element(*self.input_tipo_identificador).click()
            sleep(1.5)
            self.driver.find_element(*self.td_n_lote).click()
            sleep(1.5)
            self.driver.find_element(*self.input_valor_identificador).clear()
            self.driver.find_element(*self.input_valor_identificador).send_keys(diretorio)
            sleep(1.5)
            self.driver.find_element(*self.btn_buscar).click()
            
            while 'Carregando' in self.driver.find_element(*self.body).text:
                sleep(1)

            if "Não existem recursos para este critério de seleção." in self.driver.find_element(*self.body).text:
                self.renomear_arquivo(path_diretorio, path_diretorio_enviados + "_nao_encontrado")
                continue

            self.driver.find_element(*self.div_ferramentas_lote).click()
            sleep(2)
            self.driver.find_element(*self.visualizar_guias).click()
            sleep(2)
            quantidade_guias = self.get_len_tabela()

            dados_processo = self.get_dados_processo(path_diretorio)

            for dado in dados_processo:
                guia = dado['guia']
                path_arquivo = dado['path_arquivo']
                path_arquivo_enviado = path_diretorio + '\\' + guia

                #TODO quantidade de arquivos não enviados
                if not self.acessar_enviar_imagem(quantidade_guias, guia, path_arquivo, path_arquivo_enviado):
                    continue

                self.driver.find_element(*self.input_file).send_keys(path_arquivo)
                sleep(1)
                self.driver.find_element(*self.adicionar).click()
                sleep(1)
                self.driver.find_element(*self.btn_enviar_imagens).click()
                sleep(2)
                self.renomear_arquivo(path_arquivo, path_arquivo_enviado + '_enviado.pdf')
                self.driver.find_element(*self.btn_ok).click()
        
        self.quantidade_de_enviados = 0
        self.driver.find_element(*self.recurso_de_glosa).click()
        sleep(2)
        self.renomear_arquivo(path_diretorio, path_diretorio_enviados)

def recursar(user, password):
    url = 'https://sso-fature.orizon.com.br/auth/realms/orizon-dativa/protocol/openid-connect/auth?client_id=fature_client&response_type=code&scope=openid&redirect_uri=https://sso-auth-codeflow-fature-apicast-production.api.ocppr.orizon.com.br/sso/token?user_key=32efd36b405a07b8c0e6c6cb9c582047'

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

    usuario = "Integraliscotta"
    senha = "Cotta2024*"

    connect_med = Orizon(driver, url, usuario, senha)
    connect_med.executar_anexo_de_guias()

recursar('lucas.paz', 'WheySesc2024*')