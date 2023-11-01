from tkinter import filedialog
import pandas as pd
from selenium import webdriver
from seleniumwire import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from abc import ABC
import time
import datetime
import os
import tkinter
import Pidgin

class PageElement(ABC):
    def __init__(self, driver, url='') -> None:
        self.driver = driver
        self.url = url

    def open(self):
        self.driver.get(self.url)

class Login(PageElement):
    usuario = (By.XPATH, '//*[@id="rp1_edt"]/table/tbody/tr[2]/td/form/table/tbody/tr[1]/td[2]/input')
    senha = (By.XPATH, '//*[@id="rp1_edt"]/table/tbody/tr[2]/td/form/table/tbody/tr[2]/td[2]/input')
    entrar = (By.XPATH, '//*[@id="rp1_edt"]/table/tbody/tr[2]/td/form/table/tbody/tr[3]/td/div/input')

    def exe_login(self, usuario, senha):
        self.driver.find_element(*self.usuario).send_keys(usuario)
        self.driver.find_element(*self.senha).send_keys(senha)
        time.sleep(1)
        self.driver.find_element(*self.entrar).click()

class Caminho(PageElement):
    consultas_e_relatorios = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/table/tbody/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td[3]/table/tbody/tr[2]/td/nobr/a')
    relatorios = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr/td[1]/div[2]/div[2]/a')
    demonstrativo_de_glosas = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td[1]/div[2]/div[1]/a')

    def exe_caminho(self):
        self.driver.implicitly_wait(10)
        self.driver.find_element(*self.consultas_e_relatorios).click()
        time.sleep(1)
        self.driver.find_element(*self.relatorios).click()
        time.sleep(1)
        self.driver.find_element(*self.demonstrativo_de_glosas).click()

class BaixarDemonstrativo(PageElement):
    competencia = (By.XPATH, '//*[@id="FormMain"]/table/tbody/tr[1]/td[2]/table/tbody/tr/td[1]/input[1]')
    filiais = (By.XPATH, '//*[@id="FormMain"]/table/tbody/tr[1]/td[4]/table/tbody/tr/td[1]/input[1]')
    peg = (By.XPATH, '//*[@id="FormMain"]/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input[1]')
    botao_ok = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[1]/td/div/div[2]/div/div/div/div[1]/a')
    pdf_viewer = (By.XPATH, '/html/body/pdf-viewer')
    demonstrativo_de_glosas = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td[1]/div[2]/div[1]/a')
    texto_peg = (By.XPATH, '//*[@id="FormMain"]/table/tbody/tr[3]/td[2]/div')


    def baixar_demonstrativo(self, planilha):
        df = pd.read_excel(planilha, header=5)
        df = df.iloc[:-1]
        df = df.dropna()
        df['Concluído'] = ''
        count = 0
        quantidade_de_faturas = len(df)
        lista_diretorio = os.listdir(r"\\10.0.0.239\automacao_financeiro\MPU")
        lista_de_nomes_sem_extensao = [
            nome.replace('.pdf', '') for nome in lista_diretorio
            ]
        erro_portal = False
        ano = datetime.date.today().year
        lista_faturas_com_erro = []

        for tentativa in range(1, 6):
            print(f'Tentativa {tentativa}')
            
            try:
                for index, linha in df.iterrows():
                    numero_fatura = f"{linha['Nº Fatura']}".replace(".0", "")

                    if df['Concluído'][index] == "Sim":
                        continue

                    if numero_fatura in lista_de_nomes_sem_extensao:
                        count += 1
                        continue

                    self.driver.implicitly_wait(30)
                    self.driver.find_element(*self.competencia).send_keys(ano)
                    time.sleep(0.5)
                    self.driver.find_element(*self.filiais).send_keys("MPU")
                    time.sleep(0.5)
                    self.driver.find_element(*self.peg).send_keys(int(linha['Nº do Protocolo']))
                    time.sleep(0.5)
                    self.driver.find_element(*self.botao_ok).click()
                    time.sleep(2)
                    self.driver.switch_to.window(self.driver.window_handles[-1])
                    time.sleep(5)
                    novo_nome = r"\\10.0.0.239\automacao_financeiro\MPU" + f"\\{numero_fatura}.pdf"
                    download_feito = False
                    endereco = r"\\10.0.0.239\automacao_financeiro\MPU"

                    for i in range(10):
                        try:
                            os.rename(f"{endereco}\\report.pdf", novo_nome)
                            self.driver.close()
                            self.driver.switch_to.window(self.driver.window_handles[-1])
                            download_feito = True
                            break

                        except Exception as e:
                            print(e)

                            try:
                                self.driver.implicitly_wait(2)
                                self.driver.find_element(*self.texto_peg)
                                texto = self.driver.find_element(*self.texto_peg).text

                                if texto == "Registro não encontrado.":
                                    df.loc[index, 'Concluído'] = 'Sim'
                                    lista_faturas_com_erro.append(numero_fatura)
                                    break
                            
                            except:

                                print("Download ainda não foi feito")

                                if i == 9:
                                    erro_portal = True
                                    self.driver.quit()

                            time.sleep(3)

                    if download_feito == True:
                        count += 1
                        print(f"Download da fatura {numero_fatura} concluído com sucesso")

                    else:
                        print(f"Download da fatura {numero_fatura} não foi feito.")

                    df.loc[index, 'Concluído'] = 'Sim'
                    self.driver.implicitly_wait(30)
                    self.driver.find_element(*self.demonstrativo_de_glosas).click()

                if count == quantidade_de_faturas:
                    tkinter.messagebox.showinfo( 'Demonstrativos MPU' , f"Downloads concluídos: {count} de {quantidade_de_faturas}." )

                else:
                    tkinter.messagebox.showinfo( 'Demonstrativos MPU' , f"Downloads concluídos: {count} de {quantidade_de_faturas}. Conferir fatura(s): {', '.join(lista_faturas_com_erro) }." )

                self.driver.quit()

                break

            except Exception as err:
                print(err)

                if erro_portal == True:
                    print("Portal sem resposta, tente novamente mais tarde")
                    break

                self.driver.get('https://sistema.planassiste.mpu.mp.br/portaltiss/pagemain.aspx?g2p=.k.as0iKTI0.cSgKI0.c.aPC00.a.k.aTS4h_RL1p.lgts1h__S2p.ln_S0iKTR0.cSgKI0.c.aPT_0.ani0.c.aPTS0.a.k.aTS3h___2p.ln_S0iKOA0.cSQjY.a&ptp=.k1.kTFT_F0.brd.bmsavdGs.be7eNo.avp7.cr0h.a.cs__LOP06d.cdnrn.l.0.bKVIRM_1fmiDottoelawr.cc.cd0r.a5u.a3n-c.aTFT_F0.c.ac.astl-.l__LOP06oetenri.m.mosp.ar.ame.ce13s6.cd1vKVIRM_1i1mieoyA4k.a')
                
#_________________________________________________________________________________________

def demonstrativo_mpu():
    try:
        url = 'https://sistema.planassiste.mpu.mp.br/portaltiss/login.aspx'
        planilha = filedialog.askopenfilename()

        options = {
            'proxy' : {
                'http': 'http://lucas.paz:Gsw2022&@10.0.0.230:3128',
                'https': 'http://lucas.paz:Gsw2022&@10.0.0.230:3128'
            }
        }

        chrome_options = Options()
        chrome_options.add_experimental_option('prefs', {
            "download.default_directory": r"\\10.0.0.239\automacao_financeiro\MPU",
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True
    })
        chrome_options.add_argument("--start-maximized")
        servico = Service(ChromeDriverManager().install())

        driver = webdriver.Chrome(service=servico, seleniumwire_options= options, options = chrome_options)

        usuario = "00735860000173"
        senha = "a0r85qbt"

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
    driver.quit()