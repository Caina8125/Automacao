from tkinter import filedialog
import pandas as pd
from selenium import webdriver
from seleniumwire import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import time
import os
import tkinter
import json
import Pidgin
from page_element import PageElement

class Login(PageElement):
    prestador = (By.XPATH, '/html/body/div[1]/div[10]/div[1]/ul/li[2]/ul/li[2]/h2')
    cnpj = (By.XPATH, '//*[@id="login-prestador"]/label[2]/input')
    usuario = (By.XPATH, '//*[@id="pssTpField"]')
    senha = (By.XPATH, '/html/body/div[1]/div[10]/div[1]/ul/li[2]/ul/li[2]/ul/div/form/div[4]/input') 
    entrar = (By.XPATH, '/html/body/div[1]/div[10]/div[1]/ul/li[2]/ul/li[2]/ul/div/form/input[4]')

    def exe_login(self, usuario, senha):
        self.driver.implicitly_wait(30)
        self.driver.find_element(*self.prestador).click()
        time.sleep(2)
        self.driver.find_element(*self.cnpj).click()
        time.sleep(2)
        self.driver.find_element(*self.usuario).send_keys(usuario)
        time.sleep(2)
        self.driver.find_element(*self.senha).send_keys(senha)
        time.sleep(2)
        self.driver.find_element(*self.entrar).click()
        time.sleep(2)

class Caminho(PageElement):
    tiss = (By.XPATH, '//*[@id="portal_salutis"]/div[1]/div[10]/div[2]/ul/li[3]/a/ul/li[2]/h2')
    demonstrativo_de_analise = (By.XPATH, '//*[@id="portal_salutis"]/div[1]/div[11]/div[7]/article/div[1]/a[1]')
    
    def exe_caminho(self):
        self.driver.implicitly_wait(30)
        self.driver.find_element(*self.tiss).click()
        time.sleep(2)
        self.driver.find_element(*self.demonstrativo_de_analise).click()
        time.sleep(2)

class BaixarDemonstrativo(PageElement):
    botao_numero_do_lote = (By.XPATH, '//*[@id="_up_ui_form"]/label[2]/input')
    botao_numero_do_protocolo = (By.XPATH, '//*[@id="_up_ui_form"]/label[1]/input')
    inserir_numero_do_lote = (By.XPATH, '//*[@id="lote"]') 
    inserir_numero_do_protocolo = (By.XPATH, '//*[@id="protocolo"]') 
    enviar = (By.XPATH, '//*[@id="_up_ui_form"]/input[2]')
    imprimir = (By.XPATH, '/html/body/div/input')
    corpo_do_html = (By.XPATH, '/html/body')

    def baixar_demonstrativo(self, planilha):
        time.sleep(2)
        df = pd.read_excel(planilha, header=5)
        df = df.iloc[:-1]
        df = df.dropna()
        df['Concluído'] = ''
        quantidade_de_faturas = len(df)
        lista_diretorio = os.listdir(r"\\10.0.0.239\automacao_financeiro\CODEVASF")
        lista_de_nomes_sem_extensao = [
            nome.replace('.pdf', '') for nome in lista_diretorio
            ]
        count = 0
        erro_portal = False

        for i in range(1, 6):
            print(f"Tentativa {i}")

            try:
                for index, linha in df.iterrows():
                    numero_fatura = str(linha['Nº Fatura']).replace('.0', '')
                    numero_protocolo = str(linha['Nº do Protocolo']).replace('.0', '')

                    if df['Concluído'][index] == "Sim":
                            continue
                    
                    if numero_fatura in lista_de_nomes_sem_extensao:
                        count += 1
                        continue
                    
                    self.driver.implicitly_wait(30)
                    self.driver.find_element(*self.botao_numero_do_lote).click()
                    time.sleep(2)
                    self.driver.find_element(*self.inserir_numero_do_lote).clear()
                    self.driver.find_element(*self.inserir_numero_do_lote).send_keys(numero_fatura)
                    time.sleep(2)
                    self.driver.find_element(*self.enviar).click()
                    time.sleep(1)
                    self.driver.switch_to.window(self.driver.window_handles[-1])
                    time.sleep(2)
                    demonstrativo = self.driver.find_element(*self.corpo_do_html).text

                    if "Não foi possível obter o demonstrativo." in demonstrativo:
                        self.driver.close()
                        self.driver.switch_to.window(self.driver.window_handles[-1])
                        time.sleep(1)
                        self.driver.find_element(*self.botao_numero_do_protocolo).click()
                        time.sleep(2)
                        self.driver.find_element(*self.inserir_numero_do_protocolo).clear()
                        self.driver.find_element(*self.inserir_numero_do_protocolo).send_keys(numero_protocolo)
                        time.sleep(1.5)
                        self.driver.find_element(*self.enviar).click()
                        time.sleep(1)
                        self.driver.switch_to.window(self.driver.window_handles[-1])

                    self.driver.find_element(*self.imprimir).click()
                    time.sleep(2)
                    novo_nome = r"\\10.0.0.239\automacao_financeiro\CODEVASF" + f"\\{numero_fatura}.pdf"
                    lista_faturas_com_erro = []
                    download_feito = False
                    endereco = r"\\10.0.0.239\automacao_financeiro\CODEVASF"

                    for j in range(10):

                        try:
                            os.rename(f"{endereco}\\Imprimir.pdf", novo_nome)
                            download_feito = True
                            break

                        except:
                            print("Download ainda não foi feito")

                            if j == 9:
                                erro_portal = True
                                self.driver.quit()

                            time.sleep(2)

                    if download_feito == True:
                        count += 1
                        print(f"Download da fatura {numero_fatura} concluído com sucesso")

                    else:
                        print(f"Download da fatura {numero_fatura} não foi feito ou não foi possível renomear.")
                        lista_faturas_com_erro.append(numero_fatura)

                    df.loc[index, 'Concluído'] = 'Sim'

                    self.driver.close()
                    self.driver.switch_to.window(self.driver.window_handles[-1])


                if count == quantidade_de_faturas:
                    tkinter.messagebox.showinfo( 'Demonstrativos Codevasf' , f"Downloads na pasta: {count} de {quantidade_de_faturas}." )

                else:
                    tkinter.messagebox.showinfo( 'Demonstrativos Codevasf' , f"Downloads na pasta: {count} de {quantidade_de_faturas}. Conferir fatura(s): {','.join(lista_faturas_com_erro) } na pasta download." )

                self.driver.quit()

                break

            except Exception as err:
                print(err)

                if erro_portal == True:
                    print("Portal sem resposta, tente novamente mais tarde")
                    break
                
                self.driver.get('https://portal.salutis.com.br/index.asp?operadora=codevasf&pag=solicita_demo_tiss&type=arel')

#_________________________________________________________________________________________________________

def demonstrativo_codevasf(user, password):
    try:
        planilha = filedialog.askopenfilename()
        url = 'https://portal.salutis.com.br/index.asp?operadora=codevasf'

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
            "download.default_directory": r"\\10.0.0.239\automacao_financeiro\CODEVASF",
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": 'false',
            "safebrowsing.disable_download_protection,": True,
            "safebrowsing_for_trusted_sources_enabled": False,
            "plugins.always_open_pdf_externally": True,
            "printing.print_preview_sticky_settings.appState": json.dumps(settings),
            "savefile.default_directory": r"\\10.0.0.239\automacao_financeiro\CODEVASF"
    })
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument('--ignore-certificate-errors')
        chrome_options.add_argument('--ignore-ssl-errors')
        chrome_options.add_argument('--kiosk-printing')
        
        try:
            servico = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=servico, seleniumwire_options= options, options = chrome_options)

        except:
            print('except')
            driver = webdriver.Chrome(seleniumwire_options= options, options = chrome_options)

        global usuario, senha
        usuario = "00735860000173"
        senha = '132631514' #int(input('Digite a senha: '))

        global login_page
        login_page = Login(driver, url)
        login_page.open()
        login_page.exe_login(usuario, senha)
        Caminho(driver, url).exe_caminho()
        BaixarDemonstrativo(driver, url).baixar_demonstrativo(planilha)
    
    except FileNotFoundError as err:
        tkinter.messagebox.showerror('Automação', f'Nenhuma planilha foi selecionada!')
    
    except Exception as err:
        tkinter.messagebox.showerror("Automação", f"Ocorreu uma exceção não tratada. \n {err.__class__.__name__} - {err}")
        Pidgin.financeiroDemo(f"Ocorreu uma exceção não tratada. \n {err.__class__.__name__} - {err}")