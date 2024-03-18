from time import sleep
import pandas as pd
from selenium import webdriver
from seleniumwire import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from tkinter.filedialog import askopenfilename
from bacen_protocolo import BuscarProtocolo
from abc import ABC
import pandas as pd
from seleniumwire import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from tkinter import messagebox
from page_element import PageElement

class LoginLayoutAntigo(PageElement):
    usuario_input = (By.XPATH, '/html/body/table/tbody/tr/td/div/div/div/div/div/div/div/div/div/div/div[2]/table/tbody/tr[2]/td/form/table/tbody/tr[1]/td[2]/input')
    senha_input = (By.XPATH, '/html/body/table/tbody/tr/td/div/div/div/div/div/div/div/div/div/div/div[2]/table/tbody/tr[2]/td/form/table/tbody/tr[2]/td[2]/input')
    login_button = (By.XPATH, '//*[@id="rp1_edt"]/table/tbody/tr[2]/td/form/table/tbody/tr[3]/td/div/input')

    def login(self, usuario, senha):
        self.driver.implicitly_wait(30)
        self.driver.find_element(*self.usuario_input).send_keys(usuario)
        sleep(2)
        self.driver.find_element(*self.senha_input).send_keys(senha)
        sleep(2)
        self.driver.find_element(*self.login_button).click()
        sleep(2)

class ConferirFatura(PageElement):
    body = (By.XPATH, '/html/body')
    faturamento = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/table/tbody/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr[2]/td/nobr/a')
    aguardando_fisico = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr/td[1]/div[2]/div[4]/a')
    input_pesquisar = (By.XPATH, '//*[@id="tsk_toolbar"]/div/div/div/div/div/div/div/div[2]/div/table/tbody/tr/td/form/div[1]/input[1]')
    lupa_pesquisa = (By.XPATH, '//*[@id="tsk_toolbar"]/div/div/div/div/div/div/div/div[2]/div/table/tbody/tr/td/form/div[1]/input[2]')
    lupa_ver_fatura = (By.XPATH, '//*[@id="FormMain"]/table/tbody/tr[1]/td/div/div/div/div/div/div/div/div/div[2]/table/tbody/tr[2]/td[1]/a/img')
    tbody_com_anexo = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr[1]/td/div/div/div/div/div/div/div/div/div[2]/table/tbody')
    tbody_sem_anexo = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[4]/td/table/tbody/tr[1]/td/div/div/div/div/div/div/div/div/div[2]/table/tbody')
    detalhes = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr[1]/td/div/div/div/div/div/div/div/div/div[2]/table/tbody/tr[12]/td/div/div[2]/a')
    detalhes_sem_anexo = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[4]/td/table/tbody/tr[1]/td/div/div/div/div/div/div/div/div/div[2]/table/tbody/tr[12]/td/div/div[2]/a')
    cem_guias = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr/td/form/table/tbody/tr[1]/td/div/div/div/div/div/div/div/div/div[2]/table/tbody/tr[22]/td/div/div[2]/a[2]')
    quarenta_guias = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr/td/form/table/tbody/tr[1]/td/div/div/div/div/div/div/div/div/div[2]/table/tbody/tr[22]/td/div/div[2]/a[1]')
    tabela = (By.XPATH, '//*[@id="FormMain"]/table/tbody/tr[1]/td/div/div/div/div/div/div/div/div/div[2]/table')
    protocolo = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr/td/div/div/div/div/div/div/div/div/div[1]/a[2]')

    def remover_pontos(self, valor):
        return valor.replace('.0', '')
    
    def write_file(self, numero_fatura, texto):
        with open(f"Bacen\\{numero_fatura}.txt", "w") as arquivo:
            arquivo.write(texto)
            arquivo.close()

    def fazer_conferencia(self, planilha):
        self.driver.implicitly_wait(30)
        self.driver.execute_script("window.open('');")
        self.driver.switch_to.window(self.driver.window_handles[-1])
        url = 'https://www3.bcb.gov.br/portalbcsaude/Login'
        busca_de_protocolo = BuscarProtocolo(driver=driver, url=url)
        busca_de_protocolo.open()
        busca_de_protocolo.login_layout_novo(usuario = "00735860000173", senha = "Amhpdf!2023")
        busca_de_protocolo.caminho()
        arquivo = pd.ExcelFile(planilha)
        sheet_names = arquivo.sheet_names
        quantidade = len(sheet_names)
        count = 0

        for i in range(count, quantidade):
            lista_de_nao_encontradas = []
            df = pd.read_excel(planilha, sheet_name=count)
            numero_fatura = df.loc[13, 'Unnamed: 13']
            numero_protocolo = busca_de_protocolo.buscar_protocolo(numero_fatura)

            if not numero_protocolo.isdigit():
                self.write_file(numero_fatura, "Número do protocolo não encontrado.")
                count += 1
                continue

            sleep(2)
            self.driver.switch_to.window(self.driver.window_handles[0])
            self.driver.find_element(*self.faturamento).click()
            sleep(2)
            self.driver.find_element(*self.aguardando_fisico).click()
            sleep(2)
            self.driver.find_element(*self.input_pesquisar).send_keys(numero_protocolo)
            sleep(2)
            self.driver.find_element(*self.lupa_pesquisa).click()
            sleep(2)
            body = self.driver.find_element(*self.body).text

            if "Nenhum registro foi encontrado." in body:
                self.write_file(numero_fatura, "Não há nenhum registro dessa fatura em Aguardando o Físico.")
                count += 1
                self.driver.switch_to.window(self.driver.window_handles[-1])
                continue

            self.driver.find_element(*self.lupa_ver_fatura).click()
            sleep(2)
            conteudo_tabela_com_anexo = self.driver.find_element(*self.tbody_com_anexo).text
            sleep(2)

            if not "Nenhum registro cadastrado." in conteudo_tabela_com_anexo:
                self.driver.find_element(*self.detalhes).click()
                sleep(2)
                try:
                    self.driver.implicitly_wait(4)
                    self.driver.find_element(*self.cem_guias).click()
                    sleep(2)
                except:
                    try:
                        self.driver.implicitly_wait(4)
                        self.driver.find_element(*self.quarenta_guias).click()
                        sleep(2)
                    except:
                        pass
                self.driver.implicitly_wait(30)
                tabela = self.driver.find_element(*self.tabela)
                tabela_html = tabela.get_attribute('outerHTML')
                df_fatura = pd.read_html(tabela_html, header=0)[0]
                df_fatura['Guia Prestador'] = (df_fatura['Guia Prestador'].astype(str)).apply(self.remover_pontos)
                lista_de_numeros_portal = df_fatura['Guia Prestador'].values.tolist()
                df = pd.read_excel(planilha, sheet_name=count, header=18)
                df = df.iloc[:-3]

                for index, linha in df.iterrows():
                    numero_guia = f"{linha['Nº Guia']}".replace('.0', '')
                    if numero_guia not in lista_de_numeros_portal:
                        lista_de_nao_encontradas.append(numero_guia)

                self.driver.find_element(*self.protocolo).click()
                sleep(2)
            
            conteudo_tabela_sem_anexo = self.driver.find_element(*self.tbody_sem_anexo).text

            if not "Nenhum registro cadastrado." in conteudo_tabela_sem_anexo:
                self.driver.find_element(*self.detalhes_sem_anexo).click()
                sleep(2)
                try:
                    self.driver.implicitly_wait(4)
                    self.driver.find_element(*self.cem_guias).click()
                    sleep(2)
                except:
                    try:
                        self.driver.implicitly_wait(4)
                        self.driver.find_element(*self.quarenta_guias).click()
                        sleep(2)
                    except:
                        pass
                self.driver.implicitly_wait(30)
                tabela = self.driver.find_element(*self.tabela)
                tabela_html = tabela.get_attribute('outerHTML')
                df_fatura = pd.read_html(tabela_html, header=0)[0]
                df_fatura['Guia Prestador'] = (df_fatura['Guia Prestador'].astype(str)).apply(self.remover_pontos)
                lista_de_numeros_portal = df_fatura['Guia Prestador'].values.tolist()
                df = pd.read_excel(planilha, sheet_name=count, header=18)
                df = df.iloc[:-3]

                for index, linha in df.iterrows():
                    numero_guia = f"{linha['Nº Guia']}".replace('.0', '')
                    if numero_guia not in lista_de_numeros_portal:
                        lista_de_nao_encontradas.append(numero_guia)

                self.driver.find_element(*self.protocolo).click()

            if len(lista_de_nao_encontradas) == 0:
                self.write_file(numero_fatura, "Todas as guias do relatório desta fatura se encontram no portal.")
                count += 1
                self.driver.switch_to.window(self.driver.window_handles[-1])
                continue

            self.write_file(numero_fatura, "\n".join(lista_de_nao_encontradas))
            count += 1
            self.driver.switch_to.window(self.driver.window_handles[-1])

def conferir_bacen(user, password):
    planilha = askopenfilename()
    url = 'https://www3.bcb.gov.br/pasbcmapa/login.aspx'
    global driver

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

    try:
        login_page = LoginLayoutAntigo(driver, url)
        login_page.open()

        login_page.login(
            usuario = "00735860000173",
            senha = "2024)Amhpdf!"
        )
        ConferirFatura(driver, url).fazer_conferencia(planilha)
        driver.quit()
        messagebox.showinfo( 'Automação Bacen' , 'Todas as pesquisas foram concluídas.' )
    except Exception as e:
        messagebox.showerror( 'Erro Automação' , f'Ocorreu uma excessão não tratada \n {e.__class__.__name__}: {e}' )
        driver.quit()