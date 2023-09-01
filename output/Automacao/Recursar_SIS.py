from tkinter import filedialog
import tkinter.messagebox
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium import webdriver
from openpyxl import load_workbook
from abc import ABC
import pandas as pd
import time
import os
from selenium.webdriver.chrome.options import Options
from seleniumwire import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service

class PageElement(ABC):
    def __init__(self, webdriver, url=''):
        self.webdriver = webdriver
        self.url = url
    def open(self):
        self.webdriver.get(self.url)

class Login(PageElement):
    usuario = (By.XPATH, '//*[@id="UserName"]')
    senha = (By.XPATH, '//*[@id="Password"]')
    acessar = (By.XPATH, '//*[@id="LoginButton"]')

    def exe_login(self, usuario, senha):
        self.webdriver.find_element(*self.usuario).send_keys(usuario)
        self.webdriver.find_element(*self.senha).send_keys(senha)
        self.webdriver.find_element(*self.acessar).click()

class Recurso(PageElement):
    recurso_glosa = (By.XPATH, '//*[@id="sidebar_recursoGlosa"]')
    recursar_glosa = (By.XPATH, '//*[@id="ctl00_SidebarMenu"]/li[6]/ul/li[1]/a')
    lista_excel = []
    protocolo = (By.XPATH, '//*[@id="BTN_RECURSAR_record"]')
    exibir_lista = (By.XPATH, '//*[@id="j1_1_anchor"]')
    tabela = (By.XPATH, '//*[@id="formularioProtocolo"]/div[2]/div/div/table')
    pesq_recurso = (By.XPATH, '//*[@id="ctl00_Main_WIDGETID_636389080027653421_FilterControl_GERAL_2__NUMERO"]')
    buscar = (By.XPATH, '//*[@id="ctl00_Main_WIDGETID_636389080027653421_FilterControl_FilterButton"]')
    recursar = (By.XPATH, '//*[@id="botaoRecursar"]')
    valor_a_recursar = (By.XPATH, '/html/body/form/div[3]/div[3]/div[3]/div/div/div[1]/div/div/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div/div[2]/div[1]/div/div[1]/div[2]/div[2]/table/tbody/tr/td[6]/input')
    motivo_recurso = (By.XPATH, '/html/body/form/div[3]/div[3]/div[3]/div/div/div[1]/div/div/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div/div[2]/div[3]/div/select')
    justificativa = (By.XPATH, '/html/body/form/div[3]/div[3]/div[3]/div/div/div[1]/div/div/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div/div[2]/div[4]/div/textarea')
    fechar = (By.XPATH, '//*[@id="modalRecursoConteudo"]/div[3]/button[3]')
    salvar = (By.XPATH, '//*[@id="botaoSalvar"]')
    enviar = (By.XPATH, '//*[@id="botaoConfirmar"]')
    botao_ok = (By.XPATH, '//*[@id="ctl00_Body"]/div[1]/div/div/div[3]/button')
    n_protocolo = (By.XPATH, '/html/body/div[1]/div/div/div[2]/div/div')
    botao_ok_enviar = (By.XPATH, '/html/body/div[1]/div/div/div[3]/button[2]')


    def caminho(self):
        self.webdriver.find_element(*self.recurso_glosa).click()
        time.sleep(2)
        self.webdriver.find_element(*self.recursar_glosa).click()
        time.sleep(2)
    
    def arquivos(self):
        nomesarquivos = os.listdir(pasta)
        for nome in nomesarquivos:
            if ".~lock" in nome:
                continue
            planilhas = os.path.join(pasta, nome)
            self.lista_excel.append(planilhas)

    def entrar_protocolo(self, planilha):
        df_protocolo = pd.read_excel(planilha)
        for index, linha in df_protocolo.iterrows():
            protocolo = f"{linha['Protocolo Aceite']}".replace('.0', '')
            self.webdriver.find_element(*self.pesq_recurso).send_keys(protocolo)
            self.webdriver.find_element(*self.buscar).click()
            time.sleep(5)
            try:
                self.webdriver.find_element(*self.protocolo).click()
            except:
                time.sleep(3)
                self.webdriver.find_element(*self.protocolo).click()
            time.sleep(5)
            break

    def tabela_site(self):
        WebDriverWait(webdriver, 120).until(EC.presence_of_element_located((self.exibir_lista))).click()
        tabela = self.webdriver.find_element(*self.tabela)
        tabela_html = tabela.get_attribute('outerHTML')
        global df_fatura
        df_fatura = pd.read_html(tabela_html)[0]
        df_fatura["Recursado"] = 'Não'

    def tabela_guia(self, i):
        if quantidade_fatura > 1:
            mostrar_tabela = webdriver.find_element(By.XPATH, f'/html/body/form/div[3]/div[3]/div[3]/div/div/div[1]/div/div/div[2]/div/div[2]/div[2]/div/div/div[1]/div[2]/div[1]/div/div/div/ul/li/ul/li[{i}]/a').click()
        if quantidade_fatura == 1:
            mostrar_tabela = webdriver.find_element(By.XPATH, '/html/body/form/div[3]/div[3]/div[3]/div/div/div[1]/div/div/div[2]/div/div[2]/div[2]/div/div/div[1]/div[2]/div[1]/div/div/div/ul/li/ul/li/a').click()
        tabela_guia = webdriver.find_element(By.XPATH, '//*[@id="formularioGuia"]/div[3]/div/div/table')
        tabela_guia_html = tabela_guia.get_attribute('outerHTML')
        df_guia = pd.read_html(tabela_guia_html)[0]
        global quant_proc
        quant_proc = len(df_guia)
    
    def comparar_proc(self, i, j, linha):
        if quantidade_fatura > 1:
            for num in range(0,1000):
                try:
                    slot_proc = webdriver.find_element(By.XPATH, f'/html/body/form/div[3]/div[3]/div[3]/div/div/div[1]/div/div/div[2]/div/div[2]/div[2]/div/div/div[1]/div[2]/div[1]/div/div/div/ul/li/ul/li[{i}]/ul/li[{j}]')
                    time.sleep(0.2)
                    slot_proc = webdriver.find_element(By.XPATH, f'/html/body/form/div[3]/div[3]/div[3]/div/div/div[1]/div/div/div[2]/div/div[2]/div[2]/div/div/div[1]/div[2]/div[1]/div/div/div/ul/li/ul/li[{i}]/ul/li[{j}]').text
                    break
                except:
                    pass
                
        if quantidade_fatura == 1 and quant_proc > 1:
            for num in range(0,1000):
                try:
                    slot_proc = webdriver.find_element(By.XPATH, f'/html/body/form/div[3]/div[3]/div[3]/div/div/div[1]/div/div/div[2]/div/div[2]/div[2]/div/div/div[1]/div[2]/div[1]/div/div/div/ul/li/ul/li/ul/li[{j}]')
                    time.sleep(0.2)
                    slot_proc = webdriver.find_element(By.XPATH, f'/html/body/form/div[3]/div[3]/div[3]/div/div/div[1]/div/div/div[2]/div/div[2]/div[2]/div/div/div[1]/div[2]/div[1]/div/div/div/ul/li/ul/li/ul/li[{j}]').text
                    break
                except:
                    pass

        if quantidade_fatura == 1 and quant_proc == 1:
            for num in range(0,1000):
                try:
                    slot_proc = webdriver.find_element(By.XPATH, f'/html/body/form/div[3]/div[3]/div[3]/div/div/div[1]/div/div/div[2]/div/div[2]/div[2]/div/div/div[1]/div[2]/div[1]/div/div/div/ul/li/ul/li/ul')
                    time.sleep(0.2)
                    slot_proc = webdriver.find_element(By.XPATH, f'/html/body/form/div[3]/div[3]/div[3]/div/div/div[1]/div/div/div[2]/div/div[2]/div[2]/div/div/div[1]/div[2]/div[1]/div/div/div/ul/li/ul/li/ul').text
                    break
                except:
                   pass
        primeiro_caractere = slot_proc[0]
        if primeiro_caractere == ' ':
            slot_proc = slot_proc.replace(primeiro_caractere, '')
        slot_proc = slot_proc.replace(".", "")
        for k, v in enumerate(slot_proc):
            if slot_proc[0] != "0":
                break
            if slot_proc[k] != "0":
                slot_proc = str(slot_proc[k:])
                break
        global procedimento_plan
        procedimento_plan = f"{linha['Procedimento']}"
        global comparacao
        comparacao = procedimento_plan in slot_proc

    def conferencia_proc(self, i, j):
        if quantidade_fatura > 1:
            visualizar_proc = WebDriverWait(webdriver, 60).until(EC.presence_of_element_located((By.XPATH, f'/html/body/form/div[3]/div[3]/div[3]/div/div/div[1]/div/div/div[2]/div/div[2]/div[2]/div/div/div[1]/div[2]/div[1]/div/div/div/ul/li/ul/li[{i}]/ul/li[{j}]/a'))).click()
        if quantidade_fatura == 1 and quant_proc > 1:
            visualizar_proc = WebDriverWait(webdriver, 60).until(EC.presence_of_element_located((By.XPATH, f'/html/body/form/div[3]/div[3]/div[3]/div/div/div[1]/div/div/div[2]/div/div[2]/div[2]/div/div/div[1]/div[2]/div[1]/div/div/div/ul/li/ul/li/ul/li[{j}]/a'))).click()
        if quantidade_fatura == 1 and quant_proc == 1:
            visualizar_proc = WebDriverWait(webdriver, 60).until(EC.presence_of_element_located((By.XPATH, f'/html/body/form/div[3]/div[3]/div[3]/div/div/div[1]/div/div/div[2]/div/div[2]/div[2]/div/div/div[1]/div[2]/div[1]/div/div/div/ul/li/ul/li/ul/li/a'))).click()
        time.sleep(1)
        global numero_guia
        numero_guia = webdriver.find_element(By.XPATH, '//*[@id="formEventoGuia"]').text
        numero_guia = str(numero_guia)
        print(numero_guia)
        global procedimento
        procedimento = webdriver.find_element(By.XPATH, '/html/body/form/div[3]/div[3]/div[3]/div/div/div[1]/div/div/div[2]/div/div[2]/div[2]/div/div/div[1]/div[2]/div[2]/div/div[1]/div/span').text
        procedimento = procedimento.replace(".", "")
        procedimento = procedimento.replace("PROCEDIMENTO ", "")
        for k, v in enumerate(procedimento):
            if procedimento[0] != "0":
                break
            if procedimento[k] != "0":
                procedimento = str(procedimento[k:])
                break
        print(procedimento)
        global valor_glosado
        valor_glosado = webdriver.find_element(By.XPATH, '/html/body/form/div[3]/div[3]/div[3]/div/div/div[1]/div/div/div[2]/div/div[2]/div[2]/div/div/div[1]/div[2]/div[2]/div/div[4]/div[2]/div[1]/p').text

    def injetar_dados(self, i, j, linha2):
        checkbox = webdriver.find_element(By.XPATH, f'/html/body/form/div[3]/div[3]/div[3]/div/div/div[1]/div/div/div[2]/div/div[2]/div[2]/div/div/div[1]/div[2]/div[1]/div/div/div/ul/li/ul/li[{i}]/ul/li[{j}]/a/i[1]').click()
        page_up = webdriver.find_element(By.XPATH, f'/html/body/form/div[3]/div[3]/div[3]/div/div/div[1]/div/div/div[2]/div/div[2]/div[2]/div/div/div[1]/div[2]/div[1]/div/div/div/ul/li/ul/li[{i}]/a').send_keys(Keys.CONTROL + Keys.HOME)
        time.sleep(0.5)
        self.webdriver.find_element(*self.recursar).click()
        time.sleep(0.5)
        self.webdriver.find_element(*self.valor_a_recursar).clear()
        self.webdriver.find_element(*self.valor_a_recursar).send_keys(f"{linha2['Valor Recursado']}")
        time.sleep(1)
        self.webdriver.find_element(*self.motivo_recurso).send_keys(f"{linha2['Motivo Glosa']}")
        self.webdriver.find_element(*self.justificativa).clear()
        self.webdriver.find_element(*self.justificativa).send_keys(f"{linha2['Recurso Glosa']}")
        time.sleep(2)
        try:
            self.webdriver.find_element(*self.salvar).click()
        except:
            self.webdriver.find_element(*self.enviar).click()
            time.sleep(1)
            numero_remessa = self.webdriver.find_element(*self.n_protocolo).text
            webdriver.save_screenshot(f"{numero_remessa}.png")
            time.sleep(1)
            self.webdriver.find_element(*self.botao_ok_enviar).click()


        time.sleep(3)
        self.webdriver.find_element(*self.botao_ok).click()
        time.sleep(1)

    def fazer_recurso(self):
        #try:
        self.caminho()
        self.arquivos()
        renomear = False
        for planilha in self.lista_excel:
            if "Enviado" in planilha:
                print("PEG já enviado")
                continue
            df_recurso = pd.read_excel(planilha)
            self.entrar_protocolo(planilha)     
            time.sleep(1)
            count = 0
            for index, linha in df_recurso.iterrows():
                count += 1
                if (f"{linha['Recursado no Portal']}") == "Sim":
                    continue
                self.tabela_site()
                global quantidade_fatura
                quantidade_fatura = len(df_fatura)
                self.webdriver.find_element(*self.mostrar_guias).click()
                controle = (f"{linha['Controle Inicial']}").replace(".0", "")
                recursado = False
                for i in range(1, quantidade_fatura + 1):
                    for slot in range(0, 100):
                        try:
                            slot_guia = webdriver.find_element(By.XPATH, f'/html/body/form/div[3]/div[3]/div[3]/div/div/div[1]/div/div/div[2]/div/div[2]/div[2]/div/div/div[1]/div[2]/div[1]/div/div/div/ul/li/ul/li[{i}]')
                            time.sleep(0.2)
                            slot_guia = webdriver.find_element(By.XPATH, f'/html/body/form/div[3]/div[3]/div[3]/div/div/div[1]/div/div/div[2]/div/div[2]/div[2]/div/div/div[1]/div[2]/div[1]/div/div/div/ul/li/ul/li[{i}]').text
                            break
                        except:
                            pass
                    if controle in slot_guia:
                        mostrar_proc = webdriver.find_element(By.XPATH, f'/html/body/form/div[3]/div[3]/div[3]/div/div/div[1]/div/div/div[2]/div/div[2]/div[2]/div/div/div[1]/div[2]/div[1]/div/div/div/ul/li/ul/li[{i}]/i').click()
                        self.tabela_guia(i)
                        for j in range(1, quant_proc + 1):
                            self.comparar_proc(i, j, linha)
                            if comparacao == False:
                                continue
                            self.conferencia_proc(i, j)
                            controle = (f"{linha['Controle Inicial']}")
                            valor_planilha = (f"{linha['Valor Glosa']}").replace(',', '')
                            if (f"{linha['Recursado no Portal']}") == "Sim":
                                continue
                            if controle == numero_guia and valor_planilha == valor_glosado and procedimento_plan in procedimento:
                                self.injetar_dados(i, j, linha)
                                dados = {"Recursado no Portal" : ['Sim']}
                                df = pd.DataFrame(dados)
                                book = load_workbook(planilha)
                                writer = pd.ExcelWriter(planilha, engine='openpyxl')
                                writer.book = book
                                writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
                                df.to_excel(writer, 'Recurso', startrow= count, startcol=21, header=False, index=False)
                                writer.save()
                                recursado = True
                                break

                    if recursado == True:
                        break
                if recursado == False:
                    webdriver.refresh()
                    page_up = webdriver.find_element(By.XPATH, '/html/body/form/div[3]/div[3]/div[3]/div/div/div[1]/div/div/div[2]/div/div[2]/div[2]/div/div/div[1]/div[1]/div/div/a[1]').send_keys(Keys.CONTROL + Keys.HOME)

            try:
                writer.close()
                book.close()
            except:
                pass
            planilha_anterior = planilha
            sem_extensao = planilha.replace('.xlsx', '')
            novo_nome = sem_extensao + '_Enviado.xlsx'
            try:
                time.sleep(2)
                os.rename(planilha_anterior, novo_nome)
            except PermissionError as err:
                print(err)

            self.webdriver.get(r'https://intra4p.senado.leg.br/WebAppPortal/saude/a/portal/prestador/recursosglosaprestador.aspx?i=PORTAL_RECURSARGLOSA&m=MENU_RECURSODEGLOSA_AGRUPADO')
            self.webdriver.find_element(*self.pesq_recurso).clear()
        webdriver.quit()
        # except:
        #     tkinter.messagebox.showerror('Erro Automação' , 'Ocorreu um erro enquanto o Robô trabalhava')
#---------------------------------------------------------------------------------------------------------------------------------
def recursar_sis():
    global pasta
    pasta = filedialog.askdirectory()

    url = r'https://intra4p.senado.leg.br/WebAppPortal/Login?ReturnUrl=%2fWebAppPortal%2fdefault.aspx%3f'

    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument('--ignore-certificate-errors')
    chrome_options.add_argument('--ignore-ssl-errors')

    options = {
    'proxy': {
            'http': 'http://lucas.paz:Gsw2022&@10.0.0.230:3128',
            'https': 'http://lucas.paz:Gsw2022&@10.0.0.230:3128'
        }
    }
    global webdriver

    try:
        servico = Service(ChromeDriverManager().install())
        webdriver = webdriver.Chrome(service=servico, seleniumwire_options=options, options=chrome_options)
    except:
        webdriver = webdriver.Chrome(seleniumwire_options=options, options=chrome_options)

    login_page = Login(webdriver, url)
    login_page.open()

    login_page.exe_login(
        usuario = "01047886154",
        senha = "Amhp0073%"
    )

    Recurso(webdriver, url).fazer_recurso()