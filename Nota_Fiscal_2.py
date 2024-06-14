import time
import pyautogui
import time
import Pidgin
from abc import ABC
import pandas as pd
from datetime import datetime
from selenium import webdriver
from tkinter import filedialog
from openpyxl import load_workbook
from selenium.webdriver.common.by import By
from page_element import PageElement


class Login(PageElement):
    logarCertificado = (By.XPATH, '//*[@id="btnAcionaCertificado"]')
    def exe_login(self):
        self.driver.find_element(*self.logarCertificado).click()

class Caminho(PageElement):
    declararServico = (By.XPATH, '//*[@id="Menu1_MenuPrincipal"]/ul/li[3]/div/span[3]')
    incluir         = (By.XPATH, '//*[@id="Menu1_MenuPrincipal"]/ul/li[3]/ul/li[1]/div')
    fecharModal     = (By.XPATH, '//*[@id="base-modal"]/div/div/div[1]/button')
    botaoMenu       = (By.XPATH, '//*[@id="menu-toggle"]')
    

    def exe_caminho(self):
        self.driver.implicitly_wait(180)
        # time.sleep(2)
        # try:
        #     try:
        #         self.driver.implicitly_wait(3)
        #         self.driver.find_element(*self.declararServico).click()
        #         time.sleep(2)
        #         self.driver.find_element(*self.incluir).click()
        #         time.sleep(3)
        #     except:
        #         self.driver.find_element(*self.botaoMenu).click()
        #         time.sleep(1)
        #         try:
        #             self.driver.find_element(*self.declararServico).click()
        #             time.sleep(2)
        #             self.driver.find_element(*self.incluir).click()
        #         except:
        #             time.sleep(2)
        #             self.driver.find_element(*self.incluir).click()
        # except:
        #     try:
        #         self.driver.implicitly_wait(3)
        #         self.driver.find_element(*self.declararServico).click()
        #         time.sleep(2)
        #         self.driver.find_element(*self.incluir).click()
        #         time.sleep(3)
        #     except:
        #         self.driver.find_element(*self.botaoMenu).click()
        #         time.sleep(1)
        #         try:
        #             self.driver.find_element(*self.declararServico).click()
        #             time.sleep(2)
        #             self.driver.find_element(*self.incluir).click()
        #         except:
        #             time.sleep(2)
        #             self.driver.find_element(*self.incluir).click()

        self.driver.switch_to.frame('iframe')
        time.sleep(2)
        self.driver.find_element(*self.fecharModal).click()
        time.sleep(2)
        self.driver.switch_to.default_content()



class Nf(PageElement):
    inserirCNPJ1    = (By.XPATH, f'//*[@id="dgDeclContratados__ctl2_txtCpfCnpjPrestador"]')
    inserirNumDoc   = (By.XPATH, f'//*[@id="dgDeclContratados__ctl2_txtNumDoc"]')
    botaoGravar     = (By.XPATH, '//*[@id="btnGravar"]')
    campoVlDoc      = (By.XPATH, '//*[@id="dgDeclContratados__ctl2_txtValorDoc"]')
    fecharModalErro = (By.XPATH, '//*[@id="base-modal"]/div/div/div[1]/button')
    botaoCancelar   = (By.XPATH, '//*[@id="Button4"]')
    selectMes       = (By.XPATH, '//*[@id="ddlMes"]')
    janeiro         = (By.XPATH, '//*[@id="ddlMes"]/option[2]')
    fevereiro       = (By.XPATH, '//*[@id="ddlMes"]/option[3]')
    marco           = (By.XPATH, '//*[@id="ddlMes"]/option[4]')
    abril           = (By.XPATH, '//*[@id="ddlMes"]/option[5]')
    maio            = (By.XPATH, '//*[@id="ddlMes"]/option[6]')
    junho           = (By.XPATH, '//*[@id="ddlMes"]/option[7]')
    julho           = (By.XPATH, '//*[@id="ddlMes"]/option[8]')
    agosto          = (By.XPATH, '//*[@id="ddlMes"]/option[9]')
    setembro        = (By.XPATH, '//*[@id="ddlMes"]/option[10]')
    outubro         = (By.XPATH, '//*[@id="ddlMes"]/option[11]')
    novembro        = (By.XPATH, '//*[@id="ddlMes"]/option[12]')
    dezembro        = (By.XPATH, '//*[@id="ddlMes"]/option[13]')
    confirmarMes    = (By.XPATH, '//*[@id="btnAlterarCompetencia"]')
    botaoNfGravado  = (By.XPATH, '/html/body/div[1]/div/div/div[3]/div/button')
    botaoOKErro     = (By.XPATH, '//*[@id="base-modal"]/div/div/div[1]/button')
    incluir         = (By.XPATH, '//*[@id="Menu1_MenuPrincipal"]/ul/li[4]/ul/li[1]/div')
    fecharModal     = (By.XPATH, '//*[@id="base-modal"]/div/div/div[1]/button')

    def ordenarNFPorData(self,dados):
        df = pd.read_excel(dados)

        df['NFECOMPETENCIA'] = pd.to_datetime(df['NFECOMPETENCIA'])
        df_ordenado = df.sort_values(by='NFECOMPETENCIA')

        df_ordenado.to_excel(planilha,index=False,sheet_name="Compromisso ISS")
        return df_ordenado
    

    def alterarCompetencia(self,mes):
        self.driver.implicitly_wait(3)
        driver.get("https://df.issnetonline.com.br/online/Default/alterar_competencia.aspx")
        time.sleep(1)
        self.driver.find_element(*self.selectMes).click()
        time.sleep(1)
        self.verificarMesCompetente(mes)
        time.sleep(1)
        self.driver.find_element(*self.confirmarMes).click()
        time.sleep(1)
        Caminho(driver,url).exe_caminho()

    # Função para verificar se a DataCompetencia da NF é a mesma Data do Portal, 
    # Se não for essa função vai alterar a data do Portal
    def verificarMesCompetente(self,mesPlanilha): 
        self.driver.implicitly_wait(3)
        match mesPlanilha:
            case "Janeiro":
                self.driver.find_element(*self.janeiro).click()
            case "Fevereiro":
                self.driver.find_element(*self.fevereiro).click()
            case "Março":
                self.driver.find_element(*self.marco).click()
            case "Abril":
                self.driver.find_element(*self.abril).click()
            case "Maio":
                self.driver.find_element(*self.maio).click()
            case "Junho":
                self.driver.find_element(*self.junho).click()
            case "Julho":
                self.driver.find_element(*self.julho).click()
            case "Agosto":
                self.driver.find_element(*self.agosto).click()
            case "Setembro":
                self.driver.find_element(*self.setembro).click()
            case "Outubro":
                self.driver.find_element(*self.outubro).click()
            case "Novembro":
                self.driver.find_element(*self.novembro).click()
            case "Dezembro":
                self.driver.find_element(*self.dezembro).click()

    def inserirDezEmDez(self,count,i,cnpj,nf):
        self.driver.implicitly_wait(3)
        self.driver.switch_to.frame('iframe')
        self.inserirModeloDocumento(i)
        self.driver.find_element(By.XPATH, f'//*[@id="dgDeclContratados__ctl{i}_txtCpfCnpjPrestador"]').clear()
        self.driver.find_element(By.XPATH, f'//*[@id="dgDeclContratados__ctl{i}_txtCpfCnpjPrestador"]').send_keys(cnpj)
        self.driver.find_element(*self.inserirNumDoc).click()
        time.sleep(2)
        try:
            self.driver.find_element(By.XPATH, f'//*[@id="dgDeclContratados__ctl{i}_txtNumDoc"]').clear()
            self.driver.find_element(By.XPATH, f'//*[@id="dgDeclContratados__ctl{i}_txtNumDoc"]').send_keys(nf)
        except:
            self.driver.find_element(By.XPATH, f'//*[@id="dgDeclContratados__ctl{i}_txtNumDoc"]').clear()
            self.driver.find_element(By.XPATH, f'//*[@id="dgDeclContratados__ctl{i}_txtNumDoc"]').send_keys(nf)
        self.driver.find_element(*self.campoVlDoc).click()
        self.salvarResultadoExcel("Enviado no Portal",count)

    def inserirNF(self,count,cnpj,nf):
        self.driver.implicitly_wait(3)
        self.driver.switch_to.frame('iframe')
        self.driver.find_element(By.XPATH, f'//*[@id="dgContratados__ctl2_txtCPF_CNPJ"]').clear()
        self.driver.find_element(By.XPATH, f'//*[@id="dgContratados__ctl2_txtCPF_CNPJ"]').send_keys(cnpj)
        self.driver.find_element(*self.inserirNumDoc).click()
        time.sleep(2)
        try:
            self.driver.find_element(By.XPATH, f'//*[@id="dgContratados__ctl2_txtNum_Doc"]').clear()
            self.driver.find_element(By.XPATH, f'//*[@id="dgContratados__ctl2_txtNum_Doc"]').send_keys(nf)
        except:
            self.driver.find_element(By.XPATH, f'//*[@id="dgContratados__ctl2_txtNum_Doc"]').clear()
            self.driver.find_element(By.XPATH, f'//*[@id="dgContratados__ctl2_txtNum_Doc"]').send_keys(nf)
        self.driver.find_element(*self.campoVlDoc).click()
        time.sleep(2)

    def apagarInseridos(self):
        self.driver.implicitly_wait(3)
        self.driver.switch_to.frame('iframe')
        self.driver.find_element(By.XPATH, f'//*[@id="dgContratados__ctl2_txtCPF_CNPJ"]').clear()
        # self.driver.find_element(*self.inserirNumDoc).click()
        time.sleep(2)
        try:
            self.driver.find_element(By.XPATH, f'//*[@id="dgContratados__ctl2_txtNum_Doc"]').clear()
        except:
            self.driver.find_element(By.XPATH, f'//*[@id="dgContratados__ctl2_txtNum_Doc"]').clear()
        # self.driver.find_element(*self.campoVlDoc).click()
        time.sleep(2)

    # Função para salvar o resultado da NF no Excel
    def salvarResultadoExcel(self,resultado,count_linha):
        result        = [resultado]
        df            = pd.DataFrame(result)
        book          = load_workbook(planilha)
        writer        = pd.ExcelWriter(planilha, engine='openpyxl')
        writer.book   = book
        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
        print(count_linha)
        df.to_excel(writer, "Compromisso ISS", startrow=count_linha, startcol=25, header=False, index=False)
        writer.save()
        print('anexado na planilha')
        

    def mesCompetencia(self,mes): # Função para escrever o mês por extenso da DataCompetencia
        match mes:
            case 1:
                return "Janeiro"
            case 2:
                return "Fevereiro"
            case 3:
                return "Março"
            case 4:
                return "Abril"
            case 5:
                return "Maio"
            case 6:
                return "Junho"
            case 7:
                return "Julho"
            case 8:
                return "Agosto"
            case 9:
                return "Setembro"
            case 10:
                return "Outubro"
            case 11:
                return "Novembro"
            case 12:
                return "Dezembro"

    def converterData(self,data):
        # Converter dataCompetencia para DateTime
        convertData     = datetime.strptime(data, '%Y-%m-%d').date()
        return convertData

    def obterMes(self,data):
        convertMes      = data.month
        mesPorExtenso   = self.mesCompetencia(convertMes)
        return mesPorExtenso
    
    def verificarMesVigentePortal(self,mes,i):
        # Pegar o mês vigente do portal
        try:
            self.driver.switch_to.default_content()
            campoMesPortal = driver.find_element(By.XPATH, '//*[@id="lblCompetencia"]').text
        except:
            try:
                self.driver.switch_to.default_content()
                campoMesPortal = driver.find_element(By.XPATH, '/html/body/form/nav/div/div[2]/ul[2]/li[1]/a/span[1]').text
            except:
                campoMesPortal = driver.find_element(By.XPATH, '/html/body/form/nav/div/div[2]/ul[2]/li[1]/a/span[1]').text

        # Verificar se o mês vigente do portal é o mesmo da NF,
        # Se não for, altera no portal com a função alterarCompetencia()
        if mes in campoMesPortal:
            print("Mês vigente da NF")
        else:
            if i > 2:
                self.gravarNF()
                self.verificarNFGravada()
            self.alterarCompetencia(mes)

    def verificaCnpj(self, linha):
        # Alterar cnpj e nf para string
        cnpj = str(f"{linha['CNPJCPF']}")

        # Inserir zero a esquerda quando o CNPJ não vir com 14 caracteres
        while not len(cnpj) == 14:
            cnpj = "0" + cnpj

        return cnpj
    
    def gravarNF(self):
        try:
            self.driver.switch_to.frame('iframe')
            self.driver.find_element(*self.botaoGravar).click()
        except:
            self.driver.find_element(*self.botaoGravar).click()

        

    def verificarNFGravada(self):
        try:
            try:
                status = driver.find_element(By.XPATH, '//*[@id="base-modal"]/div/div/div[1]/h4').text

                if status == "Erros":

                    self.driver.find_element(*self.botaoOKErro).click()
                    # time.sleep(1)
                    # self.salvarResultadoExcel("Enviado no Portal",count)
                    time.sleep(1)
                    self.driver.switch_to.default_content()
                    self.reset()
            except:
                self.driver.find_element(*self.botaoNfGravado).click()
                # self.salvarResultadoExcel("Enviado no Portal",count)
                self.driver.switch_to.default_content()

        except:
            try:
                # self.driver.switch_to.default_content()
                self.driver.find_element(By.XPATH, f'//*[@id="dgContratados__ctl2_txtCPF_CNPJ"]').click()
            except:
                print("Excesão não tratada")


    def reset(self):
        self.driver.find_element(*self.incluir).click()
        self.driver.switch_to.frame('iframe')
        time.sleep(2)
        self.driver.find_element(*self.fecharModal).click()
        time.sleep(2)
        self.driver.switch_to.default_content()
    def inserirModeloDocumento(self,i):
        campoClick = f'//*[@id="dgDeclContratados__ctl{i}_ddlModeloDoc"]'

        self.driver.find_element(By.XPATH,f'//*[@id="dgDeclContratados__ctl{i}_ddlModeloDoc"]/option[2]').click()
            
    def inserirDadosNf(self,dadosNF):
        global i
        self.driver.implicitly_wait(3)
        count_linha = 1
        i = 2

        for index, linha in dadosNF.iterrows():
            # Se a NF ja estiver sido enviada, pular para próxima linha
            if 'Enviado no Portal' in f"{linha['VERIFICAR']}":
                count_linha+=1
                print(count_linha)
                continue

            #Tratando Data Copetencia
            dataCompetencia = f"{linha['NFECOMPETENCIA']}".replace(" 00:00:00","")


            # Verificando se Data copetencia está vazio, se estiver pular para próxima
            if not '-' in dataCompetencia:
                count_linha+=1
                print(count_linha)
                continue

            # Converter string para DateTime
            data = self.converterData(dataCompetencia)
            #Obter o Mês por Extenso
            mes = self.obterMes(data)
            #Verifica mes do portal e trocar caso não seja o mesmo da NF
            self.verificarMesVigentePortal(mes,i)
            #Converter cnpj em string e verificar tamanho
            cnpj = self.verificaCnpj(linha)
            #Converter NF em string
            nf = str(f"{linha['NFENUMERO']}")
            # Inserir dados da NF no portal
            self.inserirDezEmDez(count_linha,i,cnpj,nf)
            
            # Condição para inserir 10 NFs no de uma vez no portal
            if i < 11:
                i+=1
                count_linha+=1
                continue

            time.sleep(1)
            self.gravarNF()
            time.sleep(1)
            self.verificarNFGravada()

            count_linha += 1
            i = 2
            print(count_linha)

        self.gravarNF()
        time.sleep(1)
        # self.verificarNFGravada()

#---------------------------------------------------------------------------------------------------------------------------------------------

def subirNF2(user, password):
    try:
        global planilha
        global driver
        global url
        
        planilha = filedialog.askopenfilename()

        url = 'https://www2.geap.org.br/auth/prestadorVue.asp'
        refresh = "https://df.issnetonline.com.br/online/Login/Login.aspx#"
        driver = webdriver.Chrome()

        dados = Nf(driver,url).ordenarNFPorData(planilha)

        login_page = Login(driver, url)

        driver.maximize_window()
        driver.get(url)

        pyautogui.write(user)
        pyautogui.press("TAB")
        pyautogui.write(password)
        pyautogui.press("enter")

        driver.get(refresh)
        time.sleep(1)
        login_page.exe_login()
        time.sleep(2)
        pyautogui.press("enter")
        time.sleep(2)

        Caminho(driver,url).exe_caminho()

        time.sleep(1)
        Nf(driver,url).inserirDadosNf(dados)
        Pidgin.notaFiscal("Todas as Notas Concluídas")
        time.sleep(1)
    except Exception as e:
        Pidgin.notaFiscal(f"Ocorreu um erro sem tratamento: {e.__class__.__name__}: {e}")

        driver.quit()
        pass
    
    driver.quit()