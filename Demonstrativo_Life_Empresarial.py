import os
import pandas as pd
from selenium import webdriver
from seleniumwire import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import time
import tkinter
import Pidgin
import tkinter.messagebox
import json
from tkinter import filedialog
import datetime
import shutil
from page_element import PageElement

class Login(PageElement):
    prestador = (By.XPATH, '/html/body/table[2]/tbody/tr/td/form/table/tbody/tr/td[2]/table/tbody/tr[3]/td[2]/select/option[3]')
    usuario = (By.ID, 'nmUsuario')
    senha = (By.ID, 'dsSenha')
    entrar = (By.XPATH, '/html/body/table[2]/tbody/tr/td/form/table/tbody/tr/td[2]/table/tbody/tr[16]/td/button[1]')

    def exe_login(self, usuario, senha):
        self.driver.implicitly_wait(30)
        iframe = self.driver.find_elements(By.TAG_NAME, "iframe")[0]
        self.driver.switch_to.frame(iframe)
        self.driver.switch_to.frame('principal')
        self.driver.find_element(*self.prestador).click()
        time.sleep(2)
        self.driver.find_element(*self.usuario).send_keys(usuario)
        time.sleep(2)
        self.driver.find_element(*self.senha).send_keys(senha)
        time.sleep(2)
        self.driver.find_element(*self.entrar).click()
        time.sleep(5)

class BaixarDemonstrativo(PageElement):
    body = (By.XPATH, '/html/body')
    frame1 = (By.XPATH, '/html/body/table/tbody/tr/td/iframe')
    frame2 = (By.XPATH, '/html/frameset/frame[1]')
    conta_medicas = (By.XPATH, '/html/body/div[4]')
    dem_retorno = (By.XPATH, '/html/body/div[5]/font')
    consultas = (By.XPATH, '/html/body/div[1]/div[1]/div[1]')
    dem_analise = (By.XPATH, '/html/body/div[1]/div[2]/center/div[2]/span')
    mes_inicio = (By.XPATH, '/html/body/div[1]/center/fieldset/table/tbody/tr[2]/td[2]/input')
    input_lote = (By.XPATH, '/html/body/div[1]/center/fieldset/table/tbody/tr[2]/td[2]/span[3]/input')
    consultar = (By.XPATH, '/html/body/div[1]/center/fieldset/table/tbody/tr[3]/td[3]/span/button')
    download_xml = (By.XPATH, '/html/body/div[1]/table/tbody/tr[2]/td[1]/a[1]/img')
    download_pdf = (By.XPATH, '/html/body/div[1]/table/tbody/tr[2]/td[1]/a[2]/img')
    numero_protocolo = (By.XPATH, '/html/body/div[1]/table/tbody/tr[2]/td[3]')

    def exe_caminho(self):
        self.driver.implicitly_wait(30)
        frame1 = self.driver.find_element(*self.frame1)
        self.driver.switch_to.frame(frame1)
        frame2 = self.driver.find_element(*self.frame2)
        self.driver.switch_to.frame(frame2)
        time.sleep(2)
        self.driver.find_element(*self.conta_medicas).click()
        time.sleep(2)
        self.driver.find_element(*self.dem_retorno).click()
        time.sleep(2)
        self.driver.switch_to.default_content()
        iframe = self.driver.find_elements(By.TAG_NAME, "iframe")[0]
        self.driver.switch_to.frame(iframe)
        self.driver.switch_to.frame('principal')
        self.driver.switch_to.frame(frame1)
        self.driver.switch_to.frame('paginaPrincipal')
        self.driver.find_element(*self.consultas).click()
        time.sleep(2)
        self.driver.find_element(*self.dem_analise).click()
        time.sleep(2)

    def baixar_demonstrativo(self, planilha):
        df = pd.read_excel(planilha, header=5)
        df = df.iloc[:-1]
        df = df.dropna()
        endereco = r"\\10.0.0.239\automacao_financeiro\LIFE EMPRESARIAL"
        mes_atual = (datetime.date.today()).replace(day=1)  
        um_meses_atras = (mes_atual - datetime.timedelta(days=1)).replace(day=1)
        dois_meses_atras = (um_meses_atras - datetime.timedelta(days=1)).replace(day=1)
        dois_meses_atras = dois_meses_atras.strftime('%d/%m/%Y').replace('/', '')
        self.driver.find_element(*self.mes_inicio).clear()
        self.driver.find_element(*self.mes_inicio).send_keys(dois_meses_atras)
        time.sleep(2)

        for tentativa in range(1, 6):
            print(f'Tentativa {tentativa}')
            try:
                for index, linha in df.iterrows():
                    lista_de_arquivos = [arquivo for arquivo in os.listdir(endereco) if "ANALISE_CONTA_MEDICA" in arquivo]
                    for arquivo in lista_de_arquivos:
                        shutil.move(f"{endereco}\\{arquivo}", r"\\10.0.0.239\automacao_financeiro\LIFE EMPRESARIAL\Não Renomeados")
                    numero_fatura = f"{linha['Nº Fatura']}".replace(".0", "")
                    print(numero_fatura)
                    self.driver.find_element(*self.input_lote).clear()
                    self.driver.find_element(*self.input_lote).send_keys(numero_fatura)
                    time.sleep(2)
                    self.driver.find_element(*self.consultar).click()
                    time.sleep(2)
                    self.driver.find_element(*self.download_xml).click()
                    time.sleep(2)
                    content = self.driver.find_element(*self.body).text
                    while "Aguarde" in content:
                        time.sleep(2)
                        content = self.driver.find_element(*self.body).text
                    print("Download do XML concluído")
                    self.driver.find_element(*self.input_lote).click()
                    time.sleep(1)
                    self.driver.find_element(*self.download_pdf).click()
                    time.sleep(4)
                    novo_nome = f"{endereco}\\{numero_fatura}.pdf"
                    contador = 0

                    while contador < 20:
                        try:
                            nome_antigo = [nome for nome in os.listdir(endereco) if "ANALISE_CONTA_MEDICA" in nome][0]
                            os.rename(f"{endereco}\\{nome_antigo}", novo_nome)
                            print("Download do PDF concluído")
                            break
                        
                        except:
                            print("Download ainda não foi feito")
                            contador += 1

                            if contador == 18:
                                time.sleep(10)
                            
                            else:
                                time.sleep(2)

            except Exception as e:
                print(e)

            break

def demonstrativo_life(user, password):
    try:
        planilha = filedialog.askopenfilename()
        url = 'https://portal.lifeempresarial.com.br/PlanodeSaude/'
        settings = {
        "recentDestinations": [{
                "id": "Save as PDF",
                "origin": "local",
                "account": "",
            }],
            "selectedDestinationId": "Save as PDF",
            "version": 2
        }

        options = {
            'proxy' : {
                'http': f'http://{user}:{password}@10.0.0.230:3128',
                'https': f'http://{user}:{password}@10.0.0.230:3128'
            }
        }

        chrome_options = Options()
        chrome_options.add_experimental_option('prefs', {
            "printing.print_to_pdf": True,
            "download.default_directory": r"\\10.0.0.239\automacao_financeiro\LIFE EMPRESARIAL",
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": 'false',
            "safebrowsing.disable_download_protection,": True,
            "safebrowsing_for_trusted_sources_enabled": False,
            "plugins.always_open_pdf_externally": True,
            "printing.print_preview_sticky_settings.appState": json.dumps(settings),
            "savefile.default_directory": r"\\10.0.0.239\automacao_financeiro\LIFE EMPRESARIAL"
    })
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument('--ignore-certificate-errors')
        chrome_options.add_argument('--ignore-ssl-errors')
        chrome_options.add_argument('--kiosk-printing')

        try:
            servico = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=servico, seleniumwire_options= options, options = chrome_options)

        except:
            driver = webdriver.Chrome(seleniumwire_options= options, options = chrome_options)

        usuario = '00735860000173'
        senha = 'Amhp@0073'

        login_page = Login(driver, url)
        login_page.open()
        login_page.exe_login(usuario, senha)
        BaixarDemonstrativo(driver, url).exe_caminho()
        BaixarDemonstrativo(driver, url).baixar_demonstrativo(planilha)
        driver.quit()
        tkinter.messagebox.showinfo('Automação', 'Downloads concluídos com sucesso!')

    except Exception as err:
        tkinter.messagebox.showerror("Automação", f"Ocorreu uma exceção não tratada. \n {err.__class__.__name__} - {err}")
        Pidgin.financeiroDemo(f"Ocorreu uma exceção não tratada. \n {err.__class__.__name__} - {err}")
        driver.quit()