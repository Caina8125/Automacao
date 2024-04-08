from abc import ABC
import math
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
    input_n_lote_recurso = (By.ID, 'txLote')
    btn_consultar = (By.ID, 'botaoConsultar')
    btn_consultar_recurso = (By.ID, 'botaoConsultarLotes')
    visualizar_demonst_pagamento = (By.XPATH, '/html/body/div[2]/div[2]/div/div[2]/form/table[2]/tbody/tr/td[10]/a[1]/img')
    visualizar_lote = (By.XPATH, '/html/body/div[2]/div[2]/div/form/table/tbody/tr/td[10]/a[2]/img')
    clique_aqui = (By.LINK_TEXT, 'clique aqui')
    select_tipo_guia = (By.ID, 'codigoTipoGuia')
    adicionar_guias = (By.XPATH, '/html/body/div[2]/div[2]/div/form/div[1]/fieldset/table/tbody/tr/td/fieldset[2]/table/tbody/tr/td/input')
    table_guias = (By.ID, 'guia')
    btn_adicionar = (By.XPATH, '/html/body/div[2]/div[2]/div/form/div[2]/div[11]/div/button[2]/span')
    btn_salvar = (By.ID, 'botaoSalvar')
    input_numero_guia = (By.ID, 'numeroGuia')
    alterar_guia = (By.XPATH, '/html/body/div[2]/div[2]/div/form/div/fieldset/fieldset[3]/table[2]/tbody/tr/td[11]/a[1]/img')
    fieldset_procedimentos = (By.XPATH, '/html/body/div[2]/div[2]/div/form/div/fieldset/fieldset[7]')
    fieldset_opms = (By.XPATH, '/html/body/div[2]/div[2]/div/form/div/fieldset/fieldset[8]')
    fieldset_outras_despesas = (By.XPATH, '/html/body/div[2]/div[2]/div/form/div/fieldset/fieldset[9]')
    input_campo_motivo = (By.ID, 'idCampoMotivoTemp')
    btn_ok_motivo = (By.XPATH, '/html/body/div[4]/div[11]/div/button')
    botao_salvar_guia = (By.ID, 'botaoSalvarAjax')
    btn_nao_adiciona = (By.XPATH, '/html/body/div[8]/div[10]/div/button[1]')
    btn_ok_salvar = (By.XPATH, '/html/body/div[6]/div[10]/div/button')

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
            lote_tst = f"{df_processo['Lote'][0]}".replace('.0', '')

            if self.is_nan(lote_tst):
                lista_de_guias = set(df_processo['Controle Inicial'].astype(str).values.tolist())
                xpath_tipo_guia = self.pegar_xpath_tipo_guia(int(df_processo['Tipo Guia'][0]))
                self.driver.find_element(*self.input_numero_lote).clear()
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
                self.driver.find_element(By.XPATH, xpath_tipo_guia).click()
                sleep(1.5)
                self.driver.find_element(*self.adicionar_guias).click()
                sleep(1.5)

                while 'Aguarde, consultando...' in self.driver.find_element(*self.body).text:
                    sleep(2)
                
                df_tabela_guia = read_html(self.driver.find_element(*self.table_guias).get_attribute('outerHTML'), header=0)[0]

                for ind, l in df_tabela_guia.iterrows():
                    n_guia_prestador = f"{l['N° da Guia do Prestador']}".replace('.0', '')

                    if n_guia_prestador in lista_de_guias:
                        self.driver.find_element(By.XPATH, f'/html/body/div[2]/div[2]/div/form/div[2]/div[2]/fieldset/div/div/table/tbody/tr[{ind+1}]/td[1]/input[2]').click()
                        sleep(1)

                self.driver.find_element(*self.btn_adicionar).click()
                sleep(2)
                count = 1

                while 'Lote de Guias incluído(a) com sucesso.' not in self.driver.find_element(*self.body).text:
                    self.driver.find_element(*self.input_numero_lote).clear()
                    self.driver.find_element(*self.input_numero_lote).send_keys(f'{numero_processo}-{count}')
                    sleep(1)
                    self.driver.find_element(*self.btn_salvar).click()

                    while 'Aguarde, processando...' in self.driver.find_element(*self.body).text:
                        sleep(2)
                        
                    count += 1

            else:
                self.driver.get('https://aplicacao7.tst.jus.br/tstsaude/LotesRecursoConsultar.do?load=1')
                self.driver.find_element(*self.input_n_lote_recurso).clear()
                self.driver.find_element(*self.input_n_lote_recurso).send_keys(lote_tst)
                sleep(1)
                self.driver.find_element(*self.btn_consultar_recurso).click()
                sleep(1.5)
                self.driver.find_element(*self.visualizar_lote).click()
        
            for index, linha in df_processo.iterrows():
                controle_inicial = f"{linha['Controle Inicial']}".replace('.0', '')
                procedimento = f"{linha['Procedimento']}".replace('.0', '')
                valor_recurso = float(f"{linha['Valor Recursado']}".replace(',','.'))
                justificativa = linha['Recurso Glosa']

                self.driver.find_element(*self.input_numero_guia).send_keys(controle_inicial)
                sleep(1)
                self.driver.find_element(*self.alterar_guia).click()
                sleep(2)
                
                fieldset_proc_element = self.driver.find_element(*self.fieldset_procedimentos)
                fieldset_opms_element = self.driver.find_element(*self.fieldset_opms)
                fieldset_outras_dis_element = self.driver.find_element(*self.fieldset_outras_despesas)

                xpath_procedimento = self.xpath_proc(procedimento, valor_recurso, justificativa, fieldset_proc_element, fieldset_opms_element, fieldset_outras_dis_element)

    def xpath_proc(self, proc, valor_recurso, justificativa, *args):
        for fieldset_element in args:
            if not self.table_in_element(fieldset_element):
                continue

            df_tabela = read_html(fieldset_element.find_element(*self.table).get_attribute('outerHTML'), header=0)[0]
            nome_quarta_coluna = df_tabela.columns.values.tolist()[4]
            num = 1

            for i, linha in df_tabela.iterrows():
                if i % 2 != 0:
                    continue

                numero_proc_portal = self.remove_zeroes(f'{linha[nome_quarta_coluna]}'.replace('.0', ''))

                if numero_proc_portal == proc:
                    checkbox = By.XPATH, f'/html/body/div[2]/div[2]/div/form/div/fieldset/fieldset[9]/table/tbody/tr[{num}]/td[1]/input[5]'
                    checkbox_checked = self.driver.find_element(*checkbox).get_attribute('checked')

                    if not checkbox_checked == 'true':
                        self.driver.find_element(*checkbox).click()

                    sleep(2)
                    input_v_recurso = (By.XPATH, f'/html/body/div[2]/div[2]/div/form/div/fieldset/fieldset[9]/table/tbody/tr[{num+1}]/td[10]/input')
                    motivo_do_recurso = (By.XPATH, f'/html/body/div[2]/div[2]/div/form/div/fieldset/fieldset[9]/table/tbody/tr[{num+1}]/td[13]/input')
                    valor_recurso_default = float(self.driver.find_element(*input_v_recurso).get_attribute('value').replace(',', '.'))
                    class_motivo_do_recurso = self.driver.find_element(*motivo_do_recurso).get_attribute('class')

                    if self.guia_recursada(class_motivo_do_recurso):
                        self.driver.find_element(*checkbox).click()
                        print('Já recursado')
                        return

                    input_qtd = (By.XPATH, f'/html/body/div[2]/div[2]/div/form/div/fieldset/fieldset[9]/table/tbody/tr[{num+1}]/td[7]/input')
                    input_porcentagem = (By.XPATH, f'/html/body/div[2]/div[2]/div/form/div/fieldset/fieldset[9]/table/tbody/tr[{num+1}]/td[9]/div/input')


                    self.ajustar_valor(valor_recurso,  valor_recurso_default, input_qtd, input_porcentagem)
                    self.driver.find_element(*motivo_do_recurso).click()
                    sleep(1.5)
                    self.driver.find_element(*self.input_campo_motivo).send_keys(justificativa)
                    sleep(1)
                    self.driver.find_element(*self.btn_ok_motivo).click()
                    sleep(1)
                    self.driver.find_element(*self.botao_salvar_guia).click()

                    while 'Aguarde, processando...' in self.driver.find_element(*self.body).text:
                        sleep(2)

                    self.driver.find_element(*self.btn_nao_adiciona).click()
                    sleep(1)
                    self.driver.find_element(*self.btn_ok_salvar).click()
                    sleep(1)
                    self.driver.find_element(By.XPATH, f'/html/body/div[2]/div[2]/div/form/div/fieldset/fieldset[9]/table/tbody/tr[{num}]/td[1]/input[5]').click()
                    return

                num += 2

    def table_in_element(self, element):
        self.driver.implicitly_wait(2)
        try:
            element.find_element(*self.table)
            self.driver.implicitly_wait(30)
            return True
        except:
            self.driver.implicitly_wait(30)
            return False
            
    def ajustar_valor(self, valor_recurso, valor_recurso_default, input_qtd, input_porcentagem):
        if valor_recurso > valor_recurso_default:
            quantidade = self.pegar_quantidade(valor_recurso, valor_recurso_default)
            self.driver.find_element(*input_qtd).clear()
            self.driver.find_element(*input_qtd).send_keys(quantidade)
            sleep(1)

        elif valor_recurso < valor_recurso_default:
            porcentagem = self.pegar_porcentagem(valor_recurso, valor_recurso_default)
            self.driver.find_element(*input_porcentagem).clear()
            self.driver.find_element(*input_porcentagem).send_keys(porcentagem)
            sleep(1)

    def desmarcar_checkboxes(self, tamanho):
        ...

    @staticmethod
    def pegar_xpath_tipo_guia(codigo_tipo_guia):
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
            
    @staticmethod
    def guia_recursada(class_motivo_do_recurso):
        match class_motivo_do_recurso:
            case 'btnMotivo':
                return True
            case 'btnSemMotivo':
                return False
            case _:
                return False

    @staticmethod
    def remove_zeroes(procedimento: str) -> str:
        for k, _ in enumerate(procedimento):
            if procedimento[0] != "0":
                return procedimento
            if procedimento[k] != "0":
                procedimento = str(procedimento[k:])
                return procedimento

    @staticmethod    
    def pegar_quantidade(valor: float, valor_default_portal: float):
        return str(1 + ((valor - valor_default_portal) / valor_default_portal)).replace('.', ',')
    
    @staticmethod
    def pegar_porcentagem(valor: float, valor_default_portal: float):
        return str(1 - ((valor_default_portal - valor) / valor_default_portal)).replace('.', ',')
    
    @staticmethod
    def is_nan(valor):
        try:
            return math.isnan(valor)
        except:
            return False

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