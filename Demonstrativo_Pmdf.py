from tkinter import filedialog
import pandas as pd
import pyautogui
from selenium import webdriver
from seleniumwire import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import time
import os
import tkinter
import Pidgin
from page_element import PageElement

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
    pesquisa = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/table/tbody/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td[4]/table/tbody/tr[2]/td/nobr/a')
    relatorios = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr/td[1]/div[2]/div[2]/a')
    dem_glosa_analitico = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td[1]/div[2]/div[1]/a')

    def exe_caminho(self):
        self.driver.implicitly_wait(30)
        self.driver.find_element(*self.pesquisa).click()
        time.sleep(1)
        self.driver.find_element(*self.relatorios).click()
        time.sleep(1)
        self.driver.find_element(*self.dem_glosa_analitico).click()

class BaixarDemonstrativo(PageElement):
    fatura = (By.XPATH, '//*[@id="FormMain"]/table/tbody/tr/td[2]/table/tbody/tr/td[1]/input[1]')
    botao_ok = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[1]/td/div/div[2]/div/div/div/div[1]/a')
    dem_glosa_analitico = (By.XPATH, '/html/body/table/tbody/tr[1]/td/div/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td[1]/div[2]/div[1]/a')
    texto_peg = (By.XPATH, '//*[@id="FormMain"]/table/tbody/tr/td[2]/div')

    def baixar_demonstrativo(self, planilha):
        df = pd.read_excel(planilha, header=5)
        df = df.iloc[:-1]
        df = df.dropna()
        df['Concluído'] = ''
        quantidade_de_faturas = len(df)
        lista_diretorio = os.listdir(r"\\10.0.0.239\automacao_financeiro\PMDF")
        lista_de_nomes_sem_extensao = [
            nome.replace('.pdf', '') for nome in lista_diretorio
            ]
        lista_faturas_com_erro = []
        count = 0
        erro_portal = False

        for tentativa in range(1, 6):
            print(f'Tentativa {tentativa}')

            try:
                for index, linha in df.iterrows():
                    numero_fatura = str(linha['Nº Fatura']).replace('.0', '')

                    if df['Concluído'][index] == 'Sim':
                        continue

                    if numero_fatura in lista_de_nomes_sem_extensao:
                        count += 1
                        continue

                    self.driver.implicitly_wait(30)
                    self.driver.find_element(*self.fatura).send_keys(numero_fatura)
                    time.sleep(1)
                    self.driver.find_element(*self.botao_ok).click()
                    time.sleep(2)
                    self.driver.switch_to.window(self.driver.window_handles[-1])
                    time.sleep(5)
                    novo_nome = r"\\10.0.0.239\automacao_financeiro\PMDF" + f"\\{numero_fatura}.pdf"
                    download_feito = False
                    endereco = r"\\10.0.0.239\automacao_financeiro\PMDF"

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
                    self.driver.find_element(*self.dem_glosa_analitico).click()
                
                if count == quantidade_de_faturas:
                    tkinter.messagebox.showinfo( 'Demonstrativos PMDF' , f"Downloads concluídos: {count} de {quantidade_de_faturas}." )

                else:
                    tkinter.messagebox.showinfo( 'Demonstrativos PMDF' , f"Downloads concluídos: {count} de {quantidade_de_faturas}. Conferir fatura(s): {', '.join(lista_faturas_com_erro)}." )

                self.driver.quit()

                break

            except Exception as err:
                print(err)

                if erro_portal == True:
                    print("Portal sem resposta, tente novamente mais tarde")
                    break
                
                self.driver.get('http://saude.pm.df.gov.br/tiss/pagemain.aspx?g2p=.k.as0iKeua.a.k.ats3h_0.a.k.ats0iK1p.lgts1h_si.cSgKi0.c.a0.cSgKi_4h_0.ani0.c.aPqsp.ln_s0iK2p.ln_s0.c.a0.cSJm8.a&ptp=.k1.kTFT_T0.brd.bm.mo.macai.be7eNo.avp4u.a4h.a.cs__LOI1Rd.cdnrn.l.0.bKVIRAE4fmiD.0gsal3dcwr.cc.cd0r.a0s12n-c.aTFT_T0.c.ac.astl-.l__LOI1Roete0lan.f.ftop.ar.ame.ce9.cr2.cd1vKVIRAE4i0mieoycI4.a')

#--------------------------------------------------------------------------------------------------------------------------------------------------------

def demonstrativo_pmdf(user, password):
    try:

        url = 'http://saude.pm.df.gov.br/tiss/pagemain.aspx'
        planilha = filedialog.askopenfilename() 

        # options = {
        #     'proxy' : {
        #         'http': f'http://{user}:{password}@10.0.0.230:3128',
        #         'https': f'http://{user}:{password}@10.0.0.230:3128'
        #     }
        # }

        chrome_options = Options()
        chrome_options.add_experimental_option('prefs', {
            "download.default_directory": r"\\10.0.0.239\automacao_financeiro\PMDF",
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": 'false',
            "safebrowsing.disable_download_protection,": True,
            "safebrowsing_for_trusted_sources_enabled": False,
            "plugins.always_open_pdf_externally": True
    })
        chrome_options.add_argument("--start-maximized")
        
        servico = Service(ChromeDriverManager().install())
        try:
            driver = webdriver.Chrome(service=servico, options = chrome_options)
        except:
            driver = webdriver.Chrome(options = chrome_options)

        usuario = "00735860000173"
        senha = "00735860000173"

        login_page = Login(driver, url)
        login_page.open()
        pyautogui.write(user.lower())
        pyautogui.press("TAB")
        time.sleep(1)
        pyautogui.write(password)
        pyautogui.press("enter")
        time.sleep(4)
        login_page.exe_login(usuario, senha)
        Caminho(driver, url).exe_caminho()
        BaixarDemonstrativo(driver, url).baixar_demonstrativo(planilha)

    except FileNotFoundError as err:
        tkinter.messagebox.showerror('Automação', f'Nenhuma planilha foi selecionada!')
    
    except Exception as err:
        tkinter.messagebox.showerror("Automação", f"Ocorreu uma exceção não tratada. \n {err.__class__.__name__} - {err}")
        Pidgin.financeiroDemow(f"Ocorreu uma exceção não tratada. \n {err.__class__.__name__} - {err}")
    driver.quit()