from abc import ABC
import math
from os import listdir
import os
from time import sleep
from tkinter import filedialog
from tkinter.messagebox import showerror, showinfo
from openpyxl import load_workbook
from pandas import DataFrame, ExcelWriter, read_excel, read_html
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium import webdriver
from seleniumwire import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By

from Pidgin import financeiroDemo
from page_element import PageElement

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
    #TODO Alguns XPATHS precisam ser ajustados conforme o tipo de guia
    alterar_guia = ()
    fieldset_procedimentos = (By.XPATH, '/html/body/div[2]/div[2]/div/form/div/fieldset/fieldset[7]')
    fieldset_opms = (By.XPATH, '/html/body/div[2]/div[2]/div/form/div/fieldset/fieldset[8]')
    fieldset_outras_despesas = (By.XPATH, '/html/body/div[2]/div[2]/div/form/div/fieldset/fieldset[9]')
    fieldset_proc_consulta = (By.XPATH, '/html/body/div[2]/div[2]/div/form/div/fieldset/fieldset[5]')
    fieldset_proc_honorario = (By.XPATH, '/html/body/div[2]/div[2]/div/div[2]/form/fieldset/fieldset[6]')
    input_campo_motivo = (By.ID, 'idCampoMotivoTemp')
    btn_ok_motivo = (By.XPATH, '/html/body/div[4]/div[11]/div/button')
    botao_salvar_guia = (By.ID, 'botaoSalvarAjax')
    # modal = (By.CSS_SELECTOR, 'body > div:nth-child(11)')
    btn_ok_salvar = (By.XPATH, '/html/body/div[6]/div[10]/div/button')
    btn_voltar = (By.ID, 'botaoVoltar')

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
        self.driver.implicitly_wait(30)
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

    def recursar_tst(self):
        self.open()
        self.login()
        self.driver.get('https://aplicacao7.tst.jus.br/tstsaude/DemonstrativoPagamentoConsultar.do?load=1')

        for _ in range(0, 10):
            try:
                for arquivo in self.lista_de_arquivos:
                    path_planilha = arquivo['path']
                    novo_path = path_planilha.replace('.xlsx', '_enviado.xlsx')

                    if 'enviado' in path_planilha:
                        continue

                    # numero_processo = arquivo['numero_processo']
                    df_processo =  read_excel(path_planilha)
                    numero_processo = f"{df_processo['Fatura Inicial'][0]}".replace('.0', '')
                    lote_tst = f"{df_processo['Lote'][0]}".replace('.0', '')
                    tipo_guia = int(df_processo['Tipo Guia'][0])

                    if lote_tst == 'nan':
                        lista_de_guias = list(set(df_processo['Controle Inicial'].astype(str).values.tolist()))
                        xpath_tipo_guia = self.pegar_xpath_tipo_guia(tipo_guia)
                        self.acessar_lote(numero_processo, xpath_tipo_guia)

                        while 'Aguarde, consultando...' in self.driver.find_element(*self.body).text:
                            sleep(2)
                        
                        self.click_checkbox_guia(lista_de_guias)

                        n_lote = self.gera_novo_lote(numero_processo)

                        df_processo['Lote'] = n_lote
                        self.salvar_valor_planilha(
                            path_planilha=path_planilha,
                            valor=df_processo['Lote'].values.tolist(),
                            coluna=4,
                            linha=1
                        )

                    else:
                        self.driver.get('https://aplicacao7.tst.jus.br/tstsaude/LotesRecursoConsultar.do?load=1')
                        self.driver.find_element(*self.input_n_lote_recurso).clear()
                        self.driver.find_element(*self.input_n_lote_recurso).send_keys(lote_tst)
                        sleep(1)
                        self.driver.find_element(*self.btn_consultar_recurso).click()
                        sleep(1.5)
                        self.driver.find_element(*self.visualizar_lote).click()

                    # self.inicializar_atributos(tipo_guia)          
                    numero_anterior = ''
                    for i, linha in df_processo.iterrows():

                        if linha['Recursado no Portal'] == 'Sim' or linha['Recursado no Portal'] == 'Não encontrado':
                            continue

                        controle_inicial = f"{linha['Controle Inicial']}".replace('.0', '')

                        if controle_inicial == numero_anterior:
                            continue

                        numero_anterior = f"{linha['Controle Inicial']}".replace('.0', '')
                        self.driver.find_element(*self.input_numero_guia).send_keys(controle_inicial)
                        sleep(1)
                        table_guias = self.driver.find_element(*self.table_guias)
                        a = table_guias.find_elements(By.TAG_NAME, 'a')

                        for element in a:
                            self.driver.implicitly_wait(2)
                            if element.get_attribute('title') == "Alterar Guia":
                                try:
                                    element.click()
                                    break
                                except:
                                    continue
                        self.driver.implicitly_wait(30)

                        df_guia = self.df_guia(df_processo, controle_inicial)
                        num = 0

                        for _, l in df_guia.iterrows():

                            if l['Recursado no Portal'] == 'Sim' or l['Recursado no Portal'] == 'Não encontrado':
                                continue

                            linha_na_planilha = i + num + 1
                            procedimento = f"{l['Procedimento']}".replace('.0', '')
                            valor_recurso = float(f"{l['Valor Recursado']}".replace('.', '').replace(',','.'))
                            justificativa = l['Recurso Glosa']
                        
                            if tipo_guia == 2:    
                                dict_procedimento = self.encontra_id_e_table_do_codigo(procedimento)
                            else:
                                dict_procedimento = self.encontra_id_e_table_do_codigo_restante(procedimento, tipo_guia)

                            if not dict_procedimento:
                                self.salvar_valor_planilha(
                                    path_planilha,
                                    'Não encontrado',
                                    23,
                                    linha_na_planilha
                                )
                                df_processo['Recursado no Portal'][linha + num] = 'Não encontrado'
                                continue
                            
                            id_parcial_tag = self.pegar_id_parcial_tag(dict_procedimento['fieldset'], dict_procedimento['id_proc'])
                            id_motivo_glosa = self.pegar_id_motivo_glosa(dict_procedimento['fieldset'], dict_procedimento['id_proc'])
                            self.driver.find_element(By.ID, 'ckb' + id_parcial_tag).click()
                            sleep(2)
                            valor_total = float(self.driver.find_element(By.ID, 'tot' + id_parcial_tag).get_attribute('value').replace('.', '').replace(',', '.'))
                            valor_recurso_portal = float(self.driver.find_element(By.ID, 'rec' + id_parcial_tag).get_attribute('value').replace('.', '').replace(',', '.'))
                            valor_pago = float(self.driver.find_element(By.ID, 'pago' + id_parcial_tag).get_attribute('value').replace('.', '').replace(',', '.'))
                            valor_unitario = float(self.driver.find_element(By.ID, 'valor' + id_parcial_tag).get_attribute('value').replace('.', '').replace(',', '.'))

                            input_qtd = (By.ID, 'qtd' + id_parcial_tag)
                            input_porcentagem = (By.ID, 'perc' + id_parcial_tag)
                            motivo_do_recurso = (By.ID, id_motivo_glosa)

                            self.ajustar_valor(valor_recurso, valor_recurso_portal,  valor_total, valor_unitario, valor_pago, input_porcentagem, input_qtd)
                            self.driver.find_element(*motivo_do_recurso).click()
                            sleep(1.5)
                            self.driver.find_element(*self.input_campo_motivo).send_keys(justificativa)
                            sleep(1)
                            self.driver.find_element(*self.btn_ok_motivo).click()
                            sleep(1)
                            self.driver.find_element(*self.botao_salvar_guia).click()

                            while 'Aguarde, processando...' in self.driver.find_element(*self.body).text:
                                sleep(2)

                            self.salvar_valor_planilha(
                                path_planilha,
                                'Sim',
                                23,
                                linha_na_planilha
                            )
                            df_processo['Recursado no Portal'][i + num] = 'Sim'

                            sleep(1)
                            self.clickar_btn('Não')
                            sleep(1)
                            self.clickar_btn('Ok')
                            sleep(1)
                            num += 1

                        self.driver.find_element(*self.btn_voltar).click()
                    
                    try:
                        os.rename(path_planilha, novo_path)
                    except:
                        print("Erro ao renomear arquivo")
                    self.caminho()
                
                self.driver.quit()
                break
            
            except Exception as e:
                financeiroDemo(e)
                self.driver.get('https://aplicacao7.tst.jus.br/tstsaude/DemonstrativoPagamentoConsultar.do?load=1')

    def table_in_element(self, element):
        self.driver.implicitly_wait(2)
        try:
            element.find_element(*self.table)
            self.driver.implicitly_wait(30)
            return True
        except:
            self.driver.implicitly_wait(30)
            return False
            
    def ajustar_valor(self, valor_recurso, valor_recurso_portal, valor_total, valor_unitario, valor_pago, input_porcentagem, input_qtd):
        if valor_recurso > valor_recurso_portal:
            quantidade = self.pegar_quantidade(valor_recurso, valor_unitario, valor_pago)
            self.driver.find_element(*input_qtd).clear()
            self.driver.find_element(*input_qtd).send_keys(quantidade)
            sleep(1)

        elif valor_recurso < valor_recurso_portal:
            porcentagem = self.pegar_porcentagem(valor_recurso, valor_total, valor_pago)
            self.driver.find_element(*input_porcentagem).clear()
            self.driver.find_element(*input_porcentagem).send_keys(porcentagem)
            sleep(1)

    def acessar_lote(self, numero_processo, xpath_tipo_guia):
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

    def click_checkbox_guia(self, lista_de_guias):
        df_tabela_guia = read_html(self.driver.find_element(*self.table_guias).get_attribute('outerHTML'), header=0)[0]
        for ind, l in df_tabela_guia.iterrows():
            n_guia_prestador = f"{l['N° da Guia do Prestador']}".replace('.0', '')

            if n_guia_prestador in lista_de_guias:
                self.driver.find_element(By.XPATH, f'/html/body/div[2]/div[2]/div/form/div[2]/div[2]/fieldset/div/div/table/tbody/tr[{ind+1}]/td[1]/input[2]').click()
                sleep(1)

        self.driver.find_element(*self.btn_adicionar).click()
        sleep(2)

    def gera_novo_lote(self, numero_processo):
        count = 1
        while 'Lote de Guias incluído(a) com sucesso.' not in self.driver.find_element(*self.body).text:
            numero_lote = f'{numero_processo}-{count}'
            self.driver.find_element(*self.input_numero_lote).clear()
            self.driver.find_element(*self.input_numero_lote).send_keys(numero_lote)
            sleep(1)
            self.driver.find_element(*self.btn_salvar).click()

            while 'Aguarde, processando...' in self.driver.find_element(*self.body).text:
                sleep(2)
            
            count += 1

        return numero_lote
    
    def encontra_id_e_table_do_codigo_restante(self, codigo_proc, tipo_guia):
        match tipo_guia:
            case 1:
                fieldset_element = self.driver.find_element(*self.fieldset_proc_consulta)
            case 4:
                fieldset_element = self.driver.find_element(*self.fieldset_proc_honorario)

        if not self.table_in_element(fieldset_element):
            return {}
        
        nome_fieldset = fieldset_element.find_element(By.TAG_NAME, 'legend').text
        td_elements = fieldset_element.find_elements(By.TAG_NAME, 'td')

        for td_element in td_elements:
            if not td_element.text:
                continue

            text_element = self.remove_zeroes(f'{td_element.text}').replace('.', '').replace('-','')

            if text_element == codigo_proc:
                id_proc = td_element.find_element(By.TAG_NAME, 'input').get_attribute('id').split('_')[1]
                self.driver.implicitly_wait(1)

                try:
                    self.driver.find_element(By.ID, 'bot_motivo_' + id_proc)
                    self.driver.implicitly_wait(30)
                    continue
                except:
                    self.driver.implicitly_wait(30)
                    motivo = self.driver.find_element(By.ID, 'bot_sem_motivo_proc_' + id_proc)
                    if motivo.get_attribute('class') == 'btnMotivo':
                        continue

                    return {
                        'fieldset': nome_fieldset,
                        'id_proc': id_proc,
                    }
                
        return {}
        
    def encontra_id_e_table_do_codigo(self, codigo_proc):
        fieldsets = [
            self.driver.find_element(*self.fieldset_procedimentos),
            self.driver.find_element(*self.fieldset_opms),
            self.driver.find_element(*self.fieldset_outras_despesas)
        ]

        for fieldset_element in fieldsets:
            if not self.table_in_element(fieldset_element):
                continue
            nome_fieldset = fieldset_element.find_element(By.TAG_NAME, 'legend').text
            td_elements = fieldset_element.find_elements(By.TAG_NAME, 'td')
            for td_element in td_elements:
                if not td_element.text:
                    continue

                text_element = self.remove_zeroes(f'{td_element.text}').replace('.', '').replace('-','')

                if text_element == codigo_proc:
                    id_proc = td_element.find_element(By.TAG_NAME, 'input').get_attribute('id').split('_')[1]
                    self.driver.implicitly_wait(1)

                    if self.procedimento_ja_recursado(id_proc, td_elements):
                        continue

                    return {
                        'fieldset': nome_fieldset,
                        'id_proc': id_proc,
                    }
                
        return {}
    
    def procedimento_ja_recursado(self, id_proc, td_elements):
        self.driver.implicitly_wait(0)
        for td in td_elements:
            input_elements = td.find_elements(By.TAG_NAME, 'input')

            for input_element in input_elements:
                att_class = input_element.get_attribute('class')
                att_onclick = input_element.get_attribute('onclick')

                if att_class == 'btnMotivo' and id_proc in att_onclick:
                    return True
                
        self.driver.implicitly_wait(30)
        return False
    
    def pegar_id_parcial_tag(self, nome_fieldset, id_proc):
        match nome_fieldset:
            case 'Dados do Atendimento / Procedimentos Realizados':
                return f'Proc_{id_proc}'
            case 'Procedimentos Realizados':
                return f'Proc_{id_proc}'
            case 'OPMs Realizados':
                return f'Opm_{id_proc}'
            case 'Outras Despesas':
                return f'Desp_{id_proc}'
            
    def pegar_id_motivo_glosa(self, nome_fieldset, id_proc):
        match nome_fieldset:
            case 'Dados do Atendimento / Procedimentos Realizados':
                return f'bot_sem_motivo_proc_{id_proc}'
            case 'Procedimentos Realizados':
                return f'bot_sem_motivo_proc_{id_proc}'
            case 'OPMs Realizados':
                return f'bot_sem_motivo_opm_{id_proc}'
            case 'Outras Despesas':
                return f'bot_sem_motivo_desp_{id_proc}'
            
    def clickar_btn(self, btn_text):
        botoes = self.driver.find_elements(By.TAG_NAME, 'button')

        for botao in botoes:
            if botao.text == btn_text:
                botao.click()
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

    def inicializar_atributos(self, tipo_guia: int):
        match tipo_guia:
            case 1: 
                self.alterar_guia = (By.XPATH, '/html/body/div[2]/div[2]/div/form/div/fieldset/fieldset[3]/table[2]/tbody/tr[2]/td[11]/a[1]/img')
            case 2:
                self.alterar_guia = (By.XPATH, '/html/body/div[2]/div[2]/div/form/div/fieldset/fieldset[3]/table[2]/tbody/tr[4]/td[11]/a[1]/img')
            case 4:
                self.alterar_guia = (By.XPATH, '/html/body/div[2]/div[2]/div/form/div/fieldset/fieldset[3]/table[2]/tbody/tr[1]/td[11]/a[1]/img')
            
    @staticmethod
    def df_guia(df_processo:DataFrame, controle_inicial):
        df_guia = df_processo.loc[(df_processo["Controle Inicial"] == int(controle_inicial))]

        if df_guia.empty:
            df_guia = df_processo.loc[(df_processo["Controle Inicial"] == controle_inicial)]
        
        return df_guia
            
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
    def pegar_quantidade(valor: float, valor_unitario: float, valor_pago):
        return str((valor + valor_pago)/valor_unitario).replace('.', ',')
    
    @staticmethod
    def pegar_porcentagem(valor: float, valor_total: float, valor_pago):
        return str(1 - ((valor_total - (valor + valor_pago)) / valor_total)).replace('.', ',')
    
    @staticmethod
    def is_nan(valor):
        try:
            return math.isnan(valor)
        except:
            return False

def iniciar_recursar_tst(user, password):
    try:
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
        tst.recursar_tst()
        driver.quit()
        showinfo('', 'Recursos finalizados!')
    
    except Exception as e:
        showerror('Automação', f'Ocorreu uma exceção não tratada\n{e.__class__.__name__}:\n{e}')
        driver.quit()