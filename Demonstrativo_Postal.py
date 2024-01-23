import os
import time
import shutil
import Pidgin
import tkinter
import pandas as pd
from abc import ABC
import tkinter.messagebox
from selenium import webdriver
from tkinter import filedialog
from seleniumwire import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


class PageElement(ABC):
    def __init__(self, driver, url=''):
        self.driver = driver
        self.url = url
    def open(self):
        self.driver.get(self.url)


class Login(PageElement):
    email = (By.XPATH, '//*[@id="Email"]')
    senha = (By.XPATH, '//*[@id="Senha"]')
    logar = (By.XPATH, '//*[@id="btnLogin"]')

    def exe_login(self, email, senha):
        # caminho().Alert()
        self.driver.find_element(*self.email).send_keys(email)
        self.driver.find_element(*self.senha).send_keys(senha)
        self.driver.find_element(*self.logar).click()

class caminho(PageElement):
    demonstrativo        = (By.XPATH, '//*[@id="sidebar-menu"]/li[24]/a/span[1]')
    analise_conta        = (By.XPATH, '//*[@id="sidebar-menu"]/li[24]/ul/li[3]/a/span')
    selecionar_convenio  = (By.XPATH, '//*[@id="s2id_OperadorasCredenciadas_HandleOperadoraSelected"]/a/span[2]/b')
    opcao_postal         = (By.XPATH, '/html/body/div[14]/ul/li[5]/div')
    inserir_protocolo    = (By.XPATH, '//*[@id="Protocolo"]')
    baixar_demonstrativo = (By.XPATH, '//*[@id="btn-Baixar_Demonstrativo"]')
    baixar_xml           = (By.XPATH, '//*[@id="btn-Baixar_XML"]')
    fechar_botao         = (By.XPATH, '//*[@id="bcInformativosModal"]/div/div/div[3]/button[2]')
    fechar_alerta        = (By.XPATH, '/html/body/bc-modal-evolution/div/div/div/div[3]/button[3]')

    def exe_caminho(self):
        time.sleep(1)
        self.driver.find_element(*self.demonstrativo).click()
        time.sleep(1)
        self.driver.find_element(*self.analise_conta).click()
        time.sleep(2)
        caminho(driver, url).Alert()
        self.driver.find_element(*self.selecionar_convenio).click()
        time.sleep(2)
        self.driver.find_element(*self.opcao_postal).click()
        time.sleep(1)

    def buscar_demonstrativo(self):
        df = pd.read_excel(planilha, header=5)
        df = df.iloc[:-1]
        df = df.dropna()
        global count, quantidade_de_faturas, faturas_com_erro
        count = 0
        quantidade_de_faturas = len(df)
        faturas_com_erro = []     

        for index, linha in df.iterrows():

            erro = False

            global fatura
            try:
                protocolo =  f"{linha['Nº do Protocolo']}".replace(".0","")
            except:
                protocolo =  f"{linha['Nº do Protocolo']}"
            try:
                fatura    =  f"{linha['Nº Fatura']}".replace(".0","")
            except:
                fatura    =  f"{linha['Nº Fatura']}"

            self.driver.find_element(*self.inserir_protocolo).send_keys(protocolo)
            time.sleep(1)
            endereco = r"\\10.0.0.239\automacao_financeiro\POSTAL\Renomear"
            arquivo_na_pasta = os.listdir(f"{endereco}")

            for arquivo in arquivo_na_pasta:
                if '.pdf' in arquivo:
                    endereco_arquivo = f'{endereco}\\{arquivo}'
                    shutil.move(endereco_arquivo, r"\\10.0.0.239\automacao_financeiro\POSTAL\Não Renomeados")

            self.driver.find_element(*self.baixar_demonstrativo).click()
            time.sleep(8)
            self.driver.find_element(*self.baixar_xml).click()
            time.sleep(6)

            for i in range(10):
                pasta = r"\\10.0.0.239\automacao_financeiro\POSTAL\Renomear"
                nomes_arquivos = os.listdir(pasta)
                if len(nomes_arquivos) == 0:
                    break
                # time.sleep(2)

                for nome in nomes_arquivos:
                    if '.pdf' in nome:
                        nomepdf = os.path.join(pasta, nome)

                        renomear = r"\\10.0.0.239\automacao_financeiro\POSTAL\Renomear" +f"\\{fatura}"  +  ".pdf"
                        arqDest = r"\\10.0.0.239\automacao_financeiro\POSTAL" + f"\\{fatura}"  +  ".pdf"

                        try:
                            os.rename(nomepdf,renomear)
                            shutil.move(renomear,arqDest)
                            time.sleep(2)
                            print("Arquivo renomeado e guardado com sucesso")
                            break

                        except Exception as e:
                            print(e)
                            print("Download ainda não foi feito/Arquivo não renomeado")
                            time.sleep(2)
                    else:
                        arqDest_xml = r"\\10.0.0.239\automacao_financeiro\POSTAL" + f"\\{nome}"
                        nomexml     = os.path.join(pasta, nome)
                        try:
                            shutil.move(nomexml,arqDest_xml)
                            time.sleep(2)
                            print("Arquivo renomeado e guardado com sucesso")
                            break

                        except Exception as e:
                            print(e)
                            print("Download ainda não foi feito/Arquivo não renomeado")
                            time.sleep(2)

            if i == 9:
                faturas_com_erro.append(fatura)
                erro = True

            time.sleep(2)
            self.driver.find_element(*self.inserir_protocolo).clear()
            time.sleep(1)

            if erro == False:
                count += 1

        if count == quantidade_de_faturas:
            tkinter.messagebox.showinfo( 'Demonstrativos Postal' , f"Downloads concluídos: {count} de {quantidade_de_faturas}." )

        else:
            tkinter.messagebox.showinfo( 'Demonstrativos Postal' , f"Downloads concluídos: {count} de {quantidade_de_faturas}. Conferir fatura(s): {', '.join(faturas_com_erro) }." )




    def Alert(self):
        try:
            modal = WebDriverWait(self.driver, 3.0).until(EC.presence_of_element_located((By.XPATH, '//*[@id="bcInformativosModal"]/div/div')))
            self.driver.find_element(*self.fechar_botao).click()

            while True:
                try:
                    proximo_botao = WebDriverWait(self.driver, 0.2).until(EC.presence_of_element_located((By.XPATH, '//*[@id="bcInformativosModal"]/div/div/div[3]/button[2]')))
                    proximo_botao.click()
                except:
                    break

            try:
                fechar_botao = WebDriverWait(self.driver, 0.2).until(EC.presence_of_element_located((By.XPATH, '/html/body/bc-modal-evolution/div/div/div/div[3]/button[3]')))
                fechar_botao.click()
            except:
                print("Não foi possível encontrar o botão de fechar.")
                pass

        except:
            print("Não teve Modal")
            pass

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
def demonstrativo_postal(user, password):

    try:
        global driver
        global url
        global planilha

        chrome_options = Options()

        chrome_options.add_experimental_option('prefs', { "download.default_directory": r"\\10.0.0.239\automacao_financeiro\POSTAL\Renomear",
                                                "download.prompt_for_download": False,
                                                "download.directory_upgrade": True,
                                                "plugins.always_open_pdf_externally": True,
                                                "safebrowsing.enabled": 'false',
                                                "safebrowsing.disable_download_protection,": True,
                                                "safebrowsing_for_trusted_sources_enabled": False,
                                                })
        
        # chrome_options.add_argument('--disable-features=VizDisplayCompositor')
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument('--ignore-certificate-errors')
        chrome_options.add_argument('--ignore-ssl-errors')
        chrome_options.add_argument('--kiosk-printing')

        url = 'https://portalconectasaude.com.br/Account/Login'
        
        planilha = filedialog.askopenfilename()

        proxy = {
        'proxy': {
                'http': f'http://{user}:{password}@10.0.0.230:3128',
                'https': f'http://{user}:{password}@10.0.0.230:3128'
            }
        }

        driver = webdriver.Chrome(seleniumwire_options=proxy, options=chrome_options)

        login_page = Login(driver, url)
        login_page.open()
        login_page.exe_login(
            email="negociacao.gerencia@amhp.com.br",
            senha="Amhp@0073"
        )

        print('Pegar Alerta Acionado!')
        caminho(driver, url).Alert()

        
        caminho(driver, url).exe_caminho()
        caminho(driver, url).buscar_demonstrativo()

    except FileNotFoundError as err:
        tkinter.messagebox.showerror('Automação', f'Nenhuma planilha foi selecionada!')
    
    except Exception as err:
        tkinter.messagebox.showerror("Automação", f"Ocorreu uma exceção não tratada. \n {err.__class__.__name__} - {err}")
        Pidgin.financeiroDemo(f"Ocorreu uma exceção não tratada. \n {err.__class__.__name__} - {err}")
    driver.quit()