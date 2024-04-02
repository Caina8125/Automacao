from datetime import date, timedelta
from os import listdir, rename
from tkinter.filedialog import askdirectory
from openpyxl import load_workbook
from pandas import DataFrame, ExcelWriter, concat, read_excel, read_html
import pyautogui
from selenium.webdriver.remote.webelement import WebElement
from selenium import webdriver
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import time
import requests
from tkinter.messagebox import showerror, showinfo
from page_element import PageElement

class ConnectMed(PageElement):
    data_atual = date.today()
    input_usuario = (By.ID, 'username')
    input_senha = (By.ID, 'password')
    button_entrar = (By.ID, 'submitPrestador')
    extrato = (By.LINK_TEXT, 'Extrato')
    visualizar = (By.LINK_TEXT, 'Visualizar')
    abrir_filtro_extrato = (By.ID, 'abrir-fechar')
    opt_60_dias = (By.XPATH, '//*[@id="cmbPeriodo"]/option[3]')
    opcao_gama = (By.XPATH, '/html/body/div[6]/div[2]/div/div[1]/div/div/a/img')
    btn_consultar = (By.ID, 'btnConsultarExtratoPeriodo')
    button_detalhar_extrato = (By.ID, 'linkDetalharExtrato')
    option_glosados = (By.XPATH, '//*[@id="dadosLotesRecursoAberto_statusRecurso"]/option[2]')
    input_conta_prestador = (By.ID, 'dadosLotesRecursoAberto_numeroContaPrestador')
    input_buscar = (By.ID, 'dadosLotesRecursoAberto_btnBuscar')
    table_guias = (By.ID, 'dadosLotesRecursoAberto-contas_grid')
    td_guia = (By.XPATH, '/html/body/div[2]/div[1]/div[2]/div[2]/div[2]/div[2]/fieldset/div[1]/div/div[1]/div[1]/div/div/div[3]/div[3]/div/table/tbody/tr[2]/td[4]')
    table_proc = (By.ID, 'dadosLotesRecursoAberto-procedimentos_grid')
    text_area_resposta = ...
    input_file = ...
    input_codigo_usuario = (By.ID, 'dadosLotesRecursoAberto_codigoUsuario')
    input_codigo_sistema = (By.ID, 'dadosLotesRecursoAberto_codigoSistema')
    input_codigo_operadora = (By.ID, 'dadosLotesRecursoAberto_codigoOperadoraConnectmed')
    input_codigo_op_processys = (By.ID, 'dadosLotesRecursoAberto_codigoOperadoraProcessys')
    input_refer_id = (By.ID, 'dadosLotesRecursoAberto_referID')
    input_contrato_prestador = (By.ID, 'dadosLotesRecursoAberto_contratoPrestador')
    input_numero_capa_lote = (By.ID, 'dadosLotesRecursoAberto_numeroCapaLote')
    input_numero_lote = (By.ID, 'dadosLotesRecursoAberto_numeroLote')
    input_codigo_empresa = (By.ID, 'dadosLotesRecursoAberto_codigoEmpresa')
    input_last_selected = (By.ID, 'dadosLotesRecursoAberto_lastSelectedRow')
    input_data_pagamento = (By.ID, 'dadosLotesRecursoAberto_dataPagamento')
    div_recurso_agrupado = (By.XPATH, '/html/body/div[4]')
    table_procedimentos_agrupados = (By.ID, 'dadosLotesRecursoAberto-procedimentos_aglutinado_grid')
    fechar_procedimentos_agrupados = (By.XPATH, '/html/body/div[4]/div[1]/button')
    btn_ok = ...
    ok_sucesso = ...
    div_contas_medicas = (By.ID, 'dadosLotesRecursoAberto_divResultado-contas')
    fechar_comunicado = (By.XPATH, '/html/body/div[6]/div[2]/div/div/div[2]/div/a')

    def __init__(self, driver: WebDriver, url: str, usuario: str, senha: str, proxies: dict, diretorio: str='', diretorio_anexos: str='', convenio: str='') -> None:
        super().__init__(driver, url)
        self.diretorio: str = diretorio
        self.diretorio_anexos: str = diretorio_anexos
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
        self.dados_anexos: list[dict] = [
            {
                'caminho': f'{diretorio_anexos}\\{arquivo}',
                'numero_guia': arquivo.split(' ')[-1].replace('.pdf', '')
            } 
            for arquivo in listdir(diretorio_anexos)
            if arquivo.endswith('.pdf')
            ]
        self.proxies = proxies
        self.convenio = convenio

    def login(self) -> None:
        self.driver.find_element(*self.input_usuario).send_keys(self.usuario)
        time.sleep(2)
        self.driver.find_element(*self.input_senha).send_keys(self.senha)
        time.sleep(2)
        self.driver.find_element(*self.button_entrar).click()

        if self.convenio == 'Gama':
            self.driver.find_element(*self.opcao_gama).click()
            time.sleep(2)

    def acessar_extrato(self) -> None:
        if 'COMUNICADO' in self.driver.find_element(*self.body).text:
            self.driver.find_element(*self.fechar_comunicado).click()
        
        try:
            self.driver.find_element(*self.extrato).click()
        except:
            if 'COMUNICADO' in self.driver.find_element(*self.body).text:
                self.driver.find_element(*self.fechar_comunicado).click()
                time.sleep(2)
                self.driver.find_element(*self.extrato).click()
        # time.sleep(2)
        self.driver.find_element(*self.visualizar).click()
        time.sleep(2)

    def get_extrato_df(self) -> DataFrame:
        df_extrato: DataFrame = read_html(self.driver.find_element(*self.table).get_attribute('outerHTML'))[0]
        return df_extrato
    
    def get_lote_no_extrato(self) -> list:
        tables: list[WebElement] = self.driver.find_elements(*self.table)
        lista_df: list[DataFrame] = []

        for table in tables:
            table_df = read_html(table.get_attribute('outerHTML'))[0]
            lista_df.append(table_df)
        
        return concat(lista_df)['Capa de Lote'].astype(str).values.tolist()
    
    def get_planilhas_dos_protocolos(self) -> list[str]:
        protocolos: str = self.get_lote_no_extrato()
        return [dado for dado in self.dados_planilhas if dado['lote'] in protocolos]
    
    def acrescenta_zeros(self, numero: str) -> str:
        while len(numero) < 20:
            numero = '0' + numero
        return numero
    
    def filtrar_guia(self, numero_guia: str, numero_controle: str) -> None:
        if self.convenio == 'Petrobras':
            numero_guia = self.acrescenta_zeros(numero_guia)
        self.driver.find_element(*self.option_glosados).click()
        time.sleep(2)
        self.driver.find_element(*self.input_conta_prestador).clear()
        self.driver.find_element(*self.input_conta_prestador).send_keys(numero_guia)
        time.sleep(2)
        self.driver.find_element(*self.input_buscar).click()
        time.sleep(2)

        if 'Nenhum recurso para visualizar!' in self.driver.find_element(*self.div_contas_medicas).text:
            self.driver.find_element(*self.option_glosados).click()
            time.sleep(2)
            self.driver.find_element(*self.input_conta_prestador).clear()
            self.driver.find_element(*self.input_conta_prestador).send_keys(numero_controle)
            time.sleep(2)
            self.driver.find_element(*self.input_buscar).click()
            time.sleep(2)

    def converter_numero_para_string(self, valor: int | float | str) -> str:
        if isinstance(valor, float) or isinstance(valor, int):
            return "{:.2f}".format(valor).replace('.', ',')

        else:
            return "{:.2f}".format(float(valor.replace('.', '').replace(',', '.')))
        
    def img_in_element(self, element: WebElement) -> bool:
        self.driver.implicitly_wait(3)
        try:
            element.find_element(By.TAG_NAME, 'img')
            self.driver.implicitly_wait(30)
            return True
        except:
            self.driver.implicitly_wait(30)
            return False
        
    def get_query_params_str(self, td_controle: WebElement, tr_proc: WebElement) -> str:
        codigo_usuario: str = self.driver.find_element(*self.input_codigo_usuario).get_attribute('value')
        codigo_sistema: str = self.driver.find_element(*self.input_codigo_sistema).get_attribute('value')
        codigo_operadora: str = self.driver.find_element(*self.input_codigo_operadora).get_attribute('value')
        codigo_op_processys: str = self.driver.find_element(*self.input_codigo_op_processys).get_attribute('value')
        refer_id: str = self.driver.find_element(*self.input_refer_id).get_attribute('value')
        contrato_prestador: str = self.driver.find_element(*self.input_contrato_prestador).get_attribute('value')
        numero_capa_lote: str = self.driver.find_element(*self.input_numero_capa_lote).get_attribute('value')
        numero_lote: str = self.driver.find_element(*self.input_numero_lote).get_attribute('value')
        codigo_empresa: str = self.driver.find_element(*self.input_codigo_empresa).get_attribute('value')
        last_selected: str = self.driver.find_element(*self.input_last_selected).get_attribute('value')
        controle: str = td_controle.get_attribute('title')
        data_pagamento: str = self.driver.find_element(*self.input_data_pagamento).get_attribute('value')
        numero_item: str = tr_proc.get_attribute('id')
        dia: str = data_pagamento.split('/')[0]
        mes: str = data_pagamento.split('/')[1]
        ano: str = data_pagamento.split('/')[2]

        return f'https://wwwt.connectmed.com.br/Connectmed/extratoItensRecursoGlosaLote/listarItensOutrasDespesas/?codigoUsuario={codigo_usuario}&codigoSistema={codigo_sistema}&codigoOperadoraConnect={codigo_operadora}&codigoOperadora={codigo_op_processys}&referID={refer_id}&contratoPrestador={contrato_prestador}&numeroCapaLote={numero_capa_lote}&numeroLote={numero_lote}&codigoEmpresa={codigo_empresa}&propertiesGetPathSaveFileDefault=recursoglosa.path.file.upload.glosa&_search=false&nd=&rows=2000&page=1&sidx=DAT_REALIZACAO&sord=asc&flagDenyItem=&numeroConta={last_selected}&numeroItem={numero_item}&controle={controle}&numeroItemPrincipal={numero_item}&dataCronogramaLote={dia}%2F{mes}%2F{ano}&classAglutinador=-aglutinador&motivoGlosa='
    
    def get_request_response(self, url: str):
        response = requests.get(url=url, proxies=self.proxies, verify=False)
        return response.json()
    
    @staticmethod
    def remove_zeroes(procedimento: str) -> str:
        for k, v in enumerate(procedimento):
            if procedimento[0] != "0":
                return procedimento
            if procedimento[k] != "0":
                procedimento = str(procedimento[k:])
                return procedimento
            
    def inicializar_atributos_enviar_recurso(self, tipo: str) -> None:
        if tipo == 'Aberto':
            self.text_area_resposta = (By.ID, 'dadosLotesRecursoAberto_mensagem')
            self.input_file = (By.XPATH, '/html/body/div[2]/div[1]/div[2]/div[2]/div[2]/div[3]/div[1]/fieldset/div[1]/div/div[2]/fieldset/div/div/div/div/form/div/div[4]/input')
            self.salvar_recurso = (By.XPATH, '/html/body/div[2]/div[1]/div[2]/div[2]/div[2]/div[3]/div[1]/fieldset/div[2]/div[2]/input')
            self.btn_ok = (By.XPATH, '/html/body/div[4]/div[3]/div/button[1]/span')
            self.ok_sucesso = (By.XPATH, '/html/body/div[4]/div[3]/div/button/span')
        else:
            self.text_area_resposta = (By.ID, 'dialogRecursoGlosaAglutinados_mensagem')
            self.input_file = (By.XPATH, '/html/body/div[4]/div[2]/div/div/div[2]/div[2]/div[2]/div/fieldset/div/div/div[2]/fieldset/div/div/div/div/form/div/div[4]/input')
            self.salvar_recurso = (By.XPATH, '/html/body/div[4]/div[2]/div/div/div[2]/div[2]/div[2]/div/fieldset/div/div/div[3]/div[2]/input')
            self.btn_ok = (By.XPATH, '/html/body/div[6]/div[3]/div/button[1]/span')
            self.ok_sucesso = (By.XPATH, '/html/body/div[6]/div[3]/div/button/span')

    def procedimento_in_agrupado(self, lista_de_dados: list[dict], procedimento: str) -> bool:
        if len(lista_de_dados) == 0:
            return False
        
        for dado in lista_de_dados:
            procedimento_agrupado = self.remove_zeroes(dado['codigoItem'])
            valor_recursado = dado['valorRecursado']
            if procedimento == procedimento_agrupado and valor_recursado == None:
                return True
        
        return False
    
    def send_values(self, checkbox: tuple, valor_recurso: WebElement, input_valor_recursar: tuple, vl_recurso: str, justificativa: str, anexo: str) -> None:
        self.get_element_visible(element=checkbox)
        time.sleep(1.5)
        valor_recurso.click()
        time.sleep(1.5)
        self.driver.find_element(*input_valor_recursar).clear()
        self.driver.find_element(*input_valor_recursar).send_keys(vl_recurso)
        time.sleep(1.5)
        self.driver.find_element(*self.text_area_resposta).send_keys(justificativa)
        time.sleep(1.5)

        if anexo != None:
            self.driver.find_element(*self.input_file).send_keys(anexo)
            time.sleep(1.5)

        self.get_element_visible(element=self.salvar_recurso)
        time.sleep(1.5)
        self.driver.find_element(*self.btn_ok).click()
        time.sleep(1.5)
        while 'Sucesso' not in self.driver.find_element(*self.body).text:
            time.sleep(1)
        self.driver.find_element(*self.ok_sucesso).click()
        time.sleep(2)
    
    def recursar_proc_aberto(self, num: str, vl_recurso: str, justificativa: str, anexo: str):
        self.inicializar_atributos_enviar_recurso('Aberto')
        checkbox = (By.XPATH, 
        f'/html/body/div[2]/div[1]/div[2]/div[2]/div[2]/div[2]/fieldset/div[1]/div/div[2]/fieldset/div/div/div/div/div/div[3]/div[3]/div/table/tbody/tr[{num}]/td[1]/input')
        valor_recurso = self.driver.find_element(By.XPATH, 
        f'/html/body/div[2]/div[1]/div[2]/div[2]/div[2]/div[2]/fieldset/div[1]/div/div[2]/fieldset/div/div/div/div/div/div[3]/div[3]/div/table/tbody/tr[{num}]/td[41]')
        input_valor_recursar = (By.XPATH,
        f'/html/body/div[2]/div[1]/div[2]/div[2]/div[2]/div[2]/fieldset/div[1]/div/div[2]/fieldset/div/div/div/div/div/div[3]/div[3]/div/table/tbody/tr[{num}]/td[41]/input')

        self.send_values(checkbox, valor_recurso, input_valor_recursar, vl_recurso, justificativa, anexo)

    def recursar_proc_agrupado(self, procedimento: str, valor_glosado: str, justificativa: str, anexo: str):
        self.inicializar_atributos_enviar_recurso('Agrupado')
        quantidade_proc_agrupados = len(read_html(self.driver.find_element(*self.table_procedimentos_agrupados).get_attribute('outerHTML'))[0]) + 1

        for i in range(2, quantidade_proc_agrupados):
            procedimento_agrupado = self.driver.find_element(By.XPATH,
            f'/html/body/div[4]/div[2]/div/div/div[2]/div[2]/div[1]/div/fieldset/div/div[1]/div/div[1]/div/div[3]/div[3]/div/table/tbody/tr[{i}]/td[21]/div').text
            procedimento_agrupado = self.remove_zeroes(procedimento_agrupado)
            td_valor_glosado = self.driver.find_element(By.XPATH,
            f'/html/body/div[4]/div[2]/div/div/div[2]/div[2]/div[1]/div/fieldset/div/div[1]/div/div[1]/div/div[3]/div[3]/div/table/tbody/tr[{i}]/td[39]')
            valor_glosado_portal = td_valor_glosado.text.replace('R$ ', '').replace('.', '')
            td_valor_recurso = self.driver.find_element(By.XPATH,
            f'/html/body/div[4]/div[2]/div/div/div[2]/div[2]/div[1]/div/fieldset/div/div[1]/div/div[1]/div/div[3]/div[3]/div/table/tbody/tr[{i}]/td[41]')
            valor_recurso_portal = td_valor_recurso.text
            checkbox = (By.XPATH,
            f'/html/body/div[4]/div[2]/div/div/div[2]/div[2]/div[1]/div/fieldset/div/div[1]/div/div[1]/div/div[3]/div[3]/div/table/tbody/tr[{i}]/td[1]/input')
            input_valor_recurso = (By.XPATH,
            f'/html/body/div[4]/div[2]/div/div/div[2]/div[2]/div[1]/div/fieldset/div/div[1]/div/div[1]/div/div[3]/div[3]/div/table/tbody/tr[{i}]/td[41]/input')

            if procedimento_agrupado == procedimento and valor_glosado_portal == valor_glosado and valor_recurso_portal == 'R$0.00':
                self.send_values(checkbox, td_valor_recurso, input_valor_recurso, valor_glosado, justificativa, anexo)
                time.sleep(1)
                self.driver.find_element(*self.fechar_procedimentos_agrupados).click()
                return
            
    def salvar_valor_planilha(self, path_planilha: str, valor, coluna: int, linha: int):
        if isinstance(valor, list) or isinstance(valor, tuple):
            dados = {"Recursado no Portal" : valor}
        else:
            dados = {"Recursado no Portal" : [valor]}

        df_dados = DataFrame(dados)
        book = load_workbook(path_planilha)
        writer = ExcelWriter(path_planilha, engine='openpyxl')
        writer.book = book
        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
        df_dados.to_excel(writer, 'Recurso', startrow=linha, startcol=coluna, header=False, index=False)
        writer.save()
        writer.close()

    def lancar_recurso(self, codigo_proc: str, vl_glosa: str, vl_recurso: str, justificativa: str, anexo: str, path: str, row: str) -> None:
        quantidade_proc = len(read_html(self.driver.find_element(*self.table_proc).get_attribute('outerHTML'))[0]) + 1

        for i in range(2, quantidade_proc):
            procedimento_portal = self.driver.find_element(By.XPATH,
            f'/html/body/div[2]/div[1]/div[2]/div[2]/div[2]/div[2]/fieldset/div[1]/div/div[2]/fieldset/div/div/div/div/div/div[3]/div[3]/div/table/tbody/tr[{i}]/td[21]/div').text
            procedimento_portal = self.remove_zeroes(procedimento_portal)

            valor_glosa_portal = self.driver.find_element(By.XPATH,
            f'/html/body/div[2]/div[1]/div[2]/div[2]/div[2]/div[2]/fieldset/div[1]/div/div[2]/fieldset/div/div/div/div/div/div[3]/div[3]/div/table/tbody/tr[{i}]/td[39]').text
            valor_glosa_portal = valor_glosa_portal.replace('R$ ', '').replace('.', '')

            valor_recursado = self.driver.find_element(By.XPATH,
            f'/html/body/div[2]/div[1]/div[2]/div[2]/div[2]/div[2]/fieldset/div[1]/div/div[2]/fieldset/div/div/div/div/div/div[3]/div[3]/div/table/tbody/tr[{i}]/td[41]').text

            if procedimento_portal == codigo_proc and valor_glosa_portal == vl_glosa and valor_recursado == 'R$0.00':
                self.recursar_proc_aberto(i, vl_recurso, justificativa, anexo)
                self.salvar_valor_planilha(
                    path_planilha=path,
                    valor='Sim',
                    coluna=53,
                    linha=row
                )
                return

            if codigo_proc in self.driver.find_element(*self.table_proc).text:
                continue

            td_controle = self.driver.find_element(By.XPATH,
            f'/html/body/div[2]/div[1]/div[2]/div[2]/div[2]/div[2]/fieldset/div[1]/div/div[2]/fieldset/div/div/div/div/div/div[3]/div[3]/div/table/tbody/tr[{i}]/td[12]')

            tr_proc = self.driver.find_element(By.XPATH,
            f'/html/body/div[2]/div[1]/div[2]/div[2]/div[2]/div[2]/fieldset/div[1]/div/div[2]/fieldset/div/div/div/div/div/div[3]/div[3]/div/table/tbody/tr[{i}]')

            if self.img_in_element(tr_proc):
                url = self.get_query_params_str(td_controle, tr_proc)
                data_agrupado = self.get_request_response(url)

                if self.procedimento_in_agrupado(data_agrupado['rows'], codigo_proc):
                    img_lapis = (By.XPATH, 
                    f'/html/body/div[2]/div[1]/div[2]/div[2]/div[2]/div[2]/fieldset/div[1]/div/div[2]/fieldset/div/div/div/div/div/div[3]/div[3]/div/table/tbody/tr[{i}]/td[6]/img')
                    self.get_element_visible(element=img_lapis)
                    time.sleep(2)
                    while 'Procedimento' not in self.driver.find_element(*self.div_recurso_agrupado).text:
                        time.sleep(1)
                    self.recursar_proc_agrupado(codigo_proc, vl_glosa, justificativa, anexo)
                    self.salvar_valor_planilha(
                        path_planilha=path,
                        valor='Sim',
                        coluna=53,
                        linha=row
                    )   
                    return

        self.salvar_valor_planilha(
            path_planilha=path,
            valor='Não',
            coluna=53,
            linha=row
        )

    def renomear_planilha(self, path_planilha: str, msg: str):
        novo_nome: str = path_planilha.replace('.xlsx', '') + f'_{msg}.xlsx'
        rename(path_planilha, novo_nome)

    def encontrar_anexo_guia(self, numero_guia) -> str | None:
        for dado in self.dados_anexos:
            if dado['numero_guia'] == numero_guia:
                return dado['caminho']
                
    def exec_recurso(self):
        self.driver.implicitly_wait(30)
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
            mes_anterior = (self.data_atual - timedelta(days=1)).replace(day=1).month

            if not mes_extrato == mes_anterior:
                continue

            lupa_extrato = (By.XPATH, f'/html/body/div[2]/div/div/div[2]/div[1]/table/tbody/tr[{index+1}]/td[5]/form/a')
            self.driver.find_element(*lupa_extrato).click()
            time.sleep(2)
            self.driver.find_element(*self.button_detalhar_extrato).click()
            time.sleep(2)
            while 'Capa de Lote' not in self.driver.find_element(*self.body).text:
                time.sleep(1)
            lista_de_dados_no_extrato = self.get_planilhas_dos_protocolos()

            for dado in lista_de_dados_no_extrato:
                caminho = dado['caminho']
                lote = dado['lote']

                if 'Enviado' in caminho or 'Não_Enviado' in caminho:
                    continue

                id = f'formularioBuscarLote_{lote}'

                self.get_element_visible(element=(By.ID, id))
                time.sleep(2)

                df_planilha = read_excel(caminho)
                tabela_guias = self.driver.find_element(*self.table_guias).text

                for i, l in df_planilha.iterrows():
                    if l['Recursado no Portal'] == 'Sim' or l['Recursado no Portal'] == 'Não':
                        continue

                    numero_guia = f"{l['Nro. Guia']}".replace('.0', '')
                    numero_ahmptiss = f"{l['Amhptiss']}"
                    numero_controle = f"{l['Controle Inicial']}"
                    codigo_procedimento = f'{l["Procedimento"]}'.replace('.0', '')
                    justificativa = f'{l["Recurso Glosa"]}'.replace('\t', ' ')
                    valor_glosa = self.converter_numero_para_string(l['Valor Glosa']).replace('-', '')
                    valor_recurso = self.converter_numero_para_string(l['Valor Recursado'])
                    anexo = None

                    if 'anex' in justificativa or 'Anex' in justificativa:
                        anexo = self.encontrar_anexo_guia(numero_ahmptiss)
                        if anexo == None:
                            self.salvar_valor_planilha(caminho, 'Anexo da guia não encontrado', coluna=53, linha=l + 1)
                            continue

                    if numero_guia not in tabela_guias and numero_controle not in tabela_guias:
                        continue

                    self.filtrar_guia(numero_guia, numero_controle)
                    time.sleep(2)
                    self.driver.find_element(*self.td_guia).click()
                    time.sleep(2)

                    self.lancar_recurso(codigo_procedimento, valor_glosa, valor_recurso, justificativa, anexo, caminho, i + 1)
                    time.sleep(2)

                self.driver.back()
                self.driver.refresh()
                time.sleep(2)
                self.driver.find_element(*self.button_detalhar_extrato)
                time.sleep(1)
                self.renomear_planilha(caminho, 'Enviado')

            self.acessar_extrato()
            time.sleep(1.5)
            self.driver.find_element(*self.abrir_filtro_extrato).click()
            time.sleep(2)
            self.driver.find_element(*self.opt_60_dias).click()
            time.sleep(2)
            self.driver.find_element(*self.btn_consultar).click()
            
        showinfo('Automação', 'Recurso realizado!')
        self.driver.quit()

def recursar_gama(user: str, password: str) -> None:
    try:
        showinfo('', 'Selecione a pasta com as planilhas.')
        diretorio_planilhas = askdirectory()
        showinfo('', 'Selecione a pasta com os anexos')
        diretorio_anexos = askdirectory()

        url = 'https://wwwt.connectmed.com.br/conectividade/prestador/home.htm'

        proxies = {
                'http': f'http://{user}:{password}@10.0.0.230:3128',
                'https': f'http://{user}:{password}@10.0.0.230:3128'
            }

        options = {
            'proxy' : proxies
        }

        chrome_options = Options()
        chrome_options.add_argument("--start-maximized")
        # chrome_options.add_argument('--ignore-certificate-errors')
        # chrome_options.add_argument('--ignore-ssl-errors')
        try:
            servico = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=servico, options = chrome_options)
        except:
            driver = webdriver.Chrome(options = chrome_options)

        usuario = "glosaamhp"
        senha = "D8nEUqb!"

        connect_med = ConnectMed(driver, url, usuario, senha, proxies, diretorio_planilhas, diretorio_anexos, "Gama")
        connect_med.open()
        pyautogui.write(user.lower())
        pyautogui.press("TAB")
        time.sleep(1)
        pyautogui.write(password)
        pyautogui.press("enter")
        time.sleep(4)
        connect_med.login()
        connect_med.exec_recurso()

    except Exception as e :
        showerror('Automação', f'Ocorreu uma exceção não trata\n{e.__class__.__name__}:\n{e}')
        driver.quit()