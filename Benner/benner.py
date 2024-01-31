import os
from time import sleep
from tkinter.messagebox import showerror, showinfo, showwarning
import pandas as pd
from seleniumwire.webdriver import Chrome
from page_element import PageElement
from selenium.webdriver.common.by import By
from datetime import date

class Benner(PageElement):
    arquivos_anexados: list = []
    arquivos_invalidos: list = []
    body: tuple = (By.XPATH, '/html/body')
    botao_fechar: tuple = (By.XPATH, '/html/body/bc-modal-evolution/div/div/div/div[3]/button[3]')
    check_box: tuple = (By.XPATH, '/html/body/div[3]/div[2]/div/div[2]/div/div[2]/div/div/div[2]/div/bc-smart-table/div[2]/table/tbody[1]/tr[1]/td[2]/input')
    confirm_fechar_lote: tuple = (By.XPATH, '/html/body/div[8]/div/div/div[3]/button[2]')
    contador_fatura: int = 0
    detalhes_do_lote: tuple = (By.XPATH, '/html/body/div[3]/div[2]/div/div[2]/div/div[2]/div/div/div[2]/div/bc-smart-table/div[2]/table/thead/tr[1]/th/div[1]/div/div/a[3]/span')
    email_input: tuple = (By.XPATH, '//*[@id="Email"]')
    enviar_lote: tuple = (By.XPATH, '/html/body/div[3]/div[2]/div/div[2]/div/bc-detalhes-lote/div[3]/div/div[1]/div[2]/a[3]/span')
    exames_complementares_opt: tuple = (By.XPATH, '/html/body/div[10]/div/div/div[2]/bc-anexo-dropzone/div/div/div/div[2]/div[7]/select/option[12]')
    fechar_lote: tuple = (By.XPATH, '/html/body/div[3]/div[2]/div/div[2]/div/bc-detalhes-lote/div[3]/div/div[1]/div[2]/a[1]/span')
    incluir: tuple = (By.XPATH, '/html/body/div[10]/div/div/div[3]/button[2]')
    incluir_anexo: tuple = (By.XPATH, '//*[@id="incluir-anexo"]/span')
    incluir_xml: tuple = (By.XPATH, '/html/body/div[3]/div[2]/div/div[2]/div/div/div[2]/div/div/div[2]/div/bc-incluir-arquivo/div/p/button')
    input_convenio: tuple = (By.XPATH, '/html/body/div[3]/div[2]/div/div[2]/div/bc-detalhes-lote/div[3]/div/div[2]/div/div/div[2]/div[11]/div/div/input[1]')
    input_file: tuple = (By.XPATH, '/html/body/input')
    input_n_lote: tuple = (By.XPATH, '/html/body/div[3]/div[2]/div/div[2]/div/div[1]/div[2]/form/div[1]/div/bc-pesquisa-numero-lote/div[2]/div[3]/div/input')
    input_xml: tuple = (By.XPATH, '')
    logar: tuple = (By.XPATH, '//*[@id="btnLogin"]')
    lote_de_pagamento: tuple = (By.XPATH, '/html/body/div[3]/div[1]/div/ul/li[23]/a/span[1]')
    modal: tuple = (By.XPATH, '//*[@id="bcInformativosModal"]/div/div')
    nome_arquivo_reduzido: str = ''
    pesquisar: tuple = (By.XPATH, '/html/body/div[3]/div[2]/div/div[2]/div/div[1]/div[2]/form/div[2]/button/span')
    pesquisar_lotes_li: tuple = (By.XPATH, '/html/body/div[3]/div[1]/div/ul/li[23]/ul/li[3]/a/span[1]')
    proximo_botao: tuple = (By.XPATH, '//*[@id="bcInformativosModal"]/div/div/div[3]/button[2]')
    registros_lotes: tuple = (By.XPATH, '/html/body/div[3]/div[2]/div/div[2]/div/div[2]/div/div/div[2]/div/bc-smart-table/div[2]/table/tbody')
    senha_input: tuple = (By.XPATH, '//*[@id="Senha"]')
    table_anexos = (By.XPATH, '//*[@id="AnexosDataGrid"]')
    upload_novo: tuple = (By.XPATH, '/html/body/div[3]/div[1]/div/ul/li[36]/a')

    def __init__(self, driver: Chrome, url: str, email: str, senha: str) -> None:
        super().__init__(driver=driver, url=url)
        self.email: str = email
        self.senha: str = senha

    def get_arquivos_pdf(cls, diretorio: str) -> list:
        LISTA_DE_ARQUIVOS: list[str] = [
            f"{diretorio}//{arquivo}" 
            for arquivo in os.listdir(diretorio)
            if arquivo.endswith(".xls") or arquivo.endswith(".xlsx")
            ]
        return LISTA_DE_ARQUIVOS
    
    def get_arquivos_xml(cls, diretorio: str) -> list:
        LISTA_DE_ARQUIVOS: list[str] = [
            f"{diretorio}//{arquivo}" 
            for arquivo in os.listdir(diretorio)
            if arquivo.endswith(".xml")
            ]
        return LISTA_DE_ARQUIVOS

    def exe_login(cls) -> None:
        cls.driver.find_element(*cls.email_input).send_keys(cls.email)
        sleep(2)
        cls.driver.find_element(*cls.senha_input).send_keys(cls.senha)
        sleep(2)
        cls.driver.find_element(*cls.logar).click()
        sleep(2)

    def pegar_alerta(cls) -> None:
        cls.driver.implicitly_wait(3)
        while True:
            try:
                cls.driver.find_element(*cls.proximo_botao).click()
                sleep(0.5)
            except:
                break
        try:
            cls.driver.find_element(*cls.botao_fechar).click()
        except:
            return


    def caminho_lote_pagamento(cls) -> None:
        cls.driver.find_element(*cls.lote_de_pagamento).click()
        sleep(2)
        cls.driver.find_element(*cls.pesquisar_lotes_li).click()
        sleep(2)

    def lote_existe(cls, numero_fatura: str) -> bool:
        cls.driver.find_element(*cls.input_n_lote).send_keys(numero_fatura)
        sleep(2)
        cls.get_click(cls.pesquisar)
        sleep(2)
        registros_content: str = cls.driver.find_element(*cls.registros_lotes).text

        if "Não existem registros." in registros_content:
            return False
        
        return True

    def fechar_e_enviar(cls) -> None:
        cls.driver.find_element(*cls.fechar_lote).click()
        sleep(2)
        cls.driver.find_element(*cls.confirm_fechar_lote).click()
        sleep(2)
        cls.driver.find_element(*cls.enviar_lote).click()
        sleep(2)

    def adicionar_anexo(cls, caminho: str) -> None:
        cls.get_click(cls.incluir_anexo)
        sleep(2)
        cls.driver.find_element(*cls.input_file).send_keys(caminho)
        sleep(2)
        cls.driver.find_element(*cls.exames_complementares_opt).click()
        sleep(2)
        cls.driver.find_element(*cls.incluir).click()
        sleep(2)

    def get_click(self, element: tuple) -> bool:
        for i in range(10):
            try:
                self.driver.find_element(*element).click()
                return True
            
            except:
                if i == 10:
                    return False
                
                self.driver.execute_script('scrollBy(0,100)')
                continue

    def get_df(cls, file_path: str) -> pd.DataFrame:
        DF = pd.read_excel(file_path, header=23).iloc[:-6]
        return DF
    
    def is_menor(self, caminho: str, convenio: str) -> bool:
        SIZE: float = (os.path.getsize(caminho) / 1024) / 1024

        match convenio:
            case "POSTAL SAÚDE | 419133":
                if SIZE > 15:
                    return False
            case "CÂMARA DOS DEPUTADOS | 888888":
                if SIZE > 20:
                    return False
        
        return True

    def fazer_envio_postal(cls, dia_atual: int) -> None:
        if dia_atual >= 1 and dia_atual <= 5:
            cls.fechar_e_enviar()
        
        else:
            cls.driver.find_element(*cls.fechar_lote).click()
            sleep(2)
            cls.driver.find_element(*cls.confirm_fechar_lote).click()
            sleep(2)
            cls.driver.get('https://portalconectasaude.com.br/Pagamentos/PesquisaLote/Index')

    def show_message(self):
        if len(self.arquivos_invalidos) > 0:
            showwarning('Automação', f"{len(self.arquivos_invalidos)} arquivos não foram anexados por ultrapassarem o limite de MB:\n{', '.join(self.arquivos_invalidos)}")

        if len(self.arquivos_anexados) > 0:
            showwarning('Automação', f"{len(self.arquivos_anexados)} arquivos não foram anexados por já ter um anexo no portal:\n{', '.join(self.arquivos_anexados)}")

        showinfo("Automação", "Concluído!")

    def exec_anexo_guias(self, diretorio: str, dia_atual: int) -> None:
        try:
            LISTA_DE_ARQUIVOS: list[str] = self.get_arquivos_pdf(diretorio)
            self.open()
            self.exe_login()
            sleep(2)
            self.pegar_alerta()
            self.caminho_lote_pagamento()
            sleep(2)

            for arquivo in LISTA_DE_ARQUIVOS:
                self.nome_arquivo_reduzido = arquivo.replace(diretorio, '')
                self.contador_fatura = 0
                self.pegar_alerta()
                df: pd.DataFrame = self.get_df(arquivo)

                for index, linha in df.iterrows():
                    numero_fatura: str = f"{linha['Nº Fatura']}"
                    caminho: str = f"{linha['Observações']}".replace('"', '')

                    if not self.lote_existe(numero_fatura):
                        continue
                    
                    self.get_click(self.check_box)
                    
                    sleep(2)
                    self.driver.find_element(*self.detalhes_do_lote).click()
                    sleep(2)
                    self.pegar_alerta()
                    convenio: str = self.driver.find_element(*self.input_convenio).get_attribute('value')

                    if not self.is_menor(caminho, convenio):
                        self.driver.get('https://portalconectasaude.com.br/Pagamentos/PesquisaLote/Index')
                        self.arquivos_invalidos.append(numero_fatura)
                        continue

                    table_content: str = self.driver.find_element(*self.table_anexos).text
                    
                    if not "Não existem registros." in table_content and convenio == "CÂMARA DOS DEPUTADOS | 888888":
                        print("Já existe um anexo neste lote!")
                        self.arquivos_anexados.append(numero_fatura)
                        continue
                    
                    self.adicionar_anexo(caminho)
                    page_content: str = self.driver.find_element(*self.body).text

                    while "Aguarde" in page_content:
                        sleep(2)
                        page_content: str = self.driver.find_element(*self.body).text

                    self.contador_fatura += 1
                    sleep(2)
                    self.pegar_alerta()

                    match convenio:
                        case "POSTAL SAÚDE | 419133":
                            self.fazer_envio_postal(dia_atual)

                        case "CÂMARA DOS DEPUTADOS | 888888":
                            self.driver.get('https://portalconectasaude.com.br/Pagamentos/PesquisaLote/Index')

        except Exception as err:
            if self.contador_fatura:
                showerror("Automação", f"Ocorreu uma excessão não tratada. Arquivo: {self.nome_arquivo_reduzido}\nQuantidade de arquivos enviados:{self.contador_fatura}\n{err.__class__.__name__}:\n{err}")

            else:
                showerror("Automação", f"Ocorreu uma excessão não tratada\n{err.__class__.__name__}:\n{err}")

    def exec_envio_xml(self, diretorio: str) -> None:
        self.open()
        self.exe_login()
        sleep(2)
        self.pegar_alerta()
        self.driver.find_element(*self.upload_novo).click()
        sleep(4)
        self.pegar_alerta()
        LISTA_DE_ARQUIVOS: list[str] = self.get_arquivos_xml(diretorio)

        for arquivo in LISTA_DE_ARQUIVOS:
            self.driver.find_element(*self.input_file).send_keys(arquivo)
            sleep(1)

        self.driver.find_element(*self.incluir_xml).click()