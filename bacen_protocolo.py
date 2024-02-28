from time import sleep
from selenium.webdriver.common.by import By
from page_element import PageElement


class BuscarProtocolo(PageElement):
    body = (By.XPATH, '/html/body')
    usuario_input = (By.XPATH, '//*[@id="UserName"]')
    senha_input = (By.XPATH, '//*[@id="Password"]')
    login_button = (By.XPATH, '//*[@id="LoginButton"]')
    envio_de_arquivo_xml = (By.XPATH, '//*[@id="sidebar_envioXML"]/span[1]')
    consultar_arquivo = (By.XPATH, '/html/body/form/div[3]/div[3]/div[1]/div/ul/li[10]/ul/li[2]/a')
    input_numero_lote = (By.XPATH, '/html/body/form/div[3]/div[3]/div[3]/div/div/div[1]/div/div/div[2]/div/div[2]/span/div/div/div/div[3]/div/div/div[2]/div[1]/span/input')
    lupa_pesquisar = (By.XPATH, '/html/body/form/div[3]/div[3]/div[3]/div/div/div[1]/div/div/div[2]/div/div[2]/span/div/div/div/div[2]/a[2]')
    a_numero_protocolo = (By.XPATH, '/html/body/form/div[3]/div[3]/div[3]/div/div/div[1]/div/div/div[2]/div/div[2]/div[4]/table/tbody/tr/td[4]/a')
    
    def login_layout_novo(self, usuario, senha):
        self.driver.implicitly_wait(30)
        self.driver.find_element(*self.usuario_input).send_keys(usuario)
        sleep(2)
        self.driver.find_element(*self.senha_input).send_keys(senha)
        sleep(2)
        self.driver.find_element(*self.login_button).click()
        sleep(2)
    
    def caminho(self):
        self.driver.implicitly_wait(30)
        self.driver.find_element(*self.envio_de_arquivo_xml).click()
        sleep(2)
        self.driver.find_element(*self.consultar_arquivo).click()
        sleep(2)

    def buscar_protocolo(self, numero_fatura):
        self.driver.implicitly_wait(30)
        self.driver.find_element(*self.input_numero_lote).send_keys(numero_fatura)
        sleep(2)
        self.driver.find_element(*self.lupa_pesquisar).click()
        sleep(2)
        body = self.driver.find_element(*self.body).text

        if "Nenhum registro encontrado." in body:
            self.driver.find_element(*self.input_numero_lote).clear()
            return 'Nenhum registro encontrado.'
        
        else:
            numero_protocolo = self.driver.find_element(*self.a_numero_protocolo).text
            sleep(1)
            self.driver.find_element(*self.input_numero_lote).clear()
            return numero_protocolo