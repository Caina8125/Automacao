from asyncio import sleep
import requests
import pandas as pd
from seleniumwire import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.webdriver import WebDriver
from page_element import PageElement
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import os
    
class ExtrairDados():
    url_revisao = 'https://apicore.geap.com.br/auditoriadigital/v1/Guias/prestadores/guias/0?PageSize=1000000&PageNumber=1&Parameters=IdeSituacaoGuia%3A3%3BDtaInicioImportacao%3A%3BDtaFimImportacao%3A%3BNroTpoGsp%3A%3BNomeBeneficiario%3A%3B&Situacao=3'
    url_pendentes = 'https://apicore.geap.com.br/auditoriadigital/v1/Guias/prestadores/guias/0?PageSize=1000000&PageNumber=1&Parameters=IdeSituacaoGuia%3A1%3BDtaInicioImportacao%3A%3BDtaFimImportacao%3A%3BNroTpoGsp%3A%3BNomeBeneficiario%3A%3B&Situacao=1'
    proxy = {
    'http': '10.0.0.230:3128',
    'https': 'lucas.paz:WheySesc2024*@10.0.0.230:3128'
    }

    def acessar_revisao_prestador(self):
        global token, authorization
        token = input('Cole aqui o token: ')
        authorization = {
        "Authorization": f'Bearer {token}'
        }
        response = requests.get(url=self.url_revisao, headers=authorization, proxies=self.proxy, verify=False)
        self.data_revisao = response.json()
        self.quant_guia = len(self.data_revisao["resultData"]["items"])

    def acessar_pendentes(self):
        global token, authorization
        token = input('Cole aqui o token: ')
        authorization = {
        "Authorization": f'Bearer {token}'
        }
        response = requests.get(url=self.url_pendentes, headers=authorization, proxies=self.proxy, verify=False)
        self.data_revisao = response.json()
        self.quant_guia = len(self.data_revisao["resultData"]["items"])

    def pegar_id(self):
        if __name__ == '__main__':
            updater = ImportarGoogleSheets()
            creds = updater.authenticate()
            updater.get_spreadsheet(creds)

        lista_url = []
        for i in range(0, self.quant_guia):
            id = str(self.data_revisao["resultData"]["items"][i]["id"])
            id_encontrado = False

            for lista in values:
                link_plan = lista[0]

                if id in link_plan:
                    id_encontrado = True
                    break
                
                else:
                    continue
            if id_encontrado == False:
                lista_url.append(f"https://apicore.geap.com.br/auditoriadigital/v1/Guias/guias/{id}")

        return lista_url
    
    def pegar_id_pendentes(self):
        lista_url = []
        for i in range(0, self.quant_guia):
            id = str(self.data_revisao["resultData"]["items"][i]["id"])
            lista_url.append(f"https://apicore.geap.com.br/auditoriadigital/v1/Guias/guias/{id}")

        return lista_url
        
    def extrair_dados(self):
        self.acessar_revisao_prestador()
        lista_url = self.pegar_id()

        for link_guia in lista_url:
            
            response = requests.get(url=link_guia, headers=authorization, proxies=self.proxy, verify=False)
            print(response)
            data = response.json()
            n_amhp = data['resultData']["numeroGuiaPrestador"]
            nome_paciente = data['resultData']["nomeBeneficiario"]
            carteirinha = data["resultData"]["numeroCartao"]
            link_aba = link_guia + '/abas'
            response_aba = requests.get(url=link_aba, headers=authorization, proxies=self.proxy, verify=False)
            data_aba = response_aba.json()
            quant_procedimento = len(data_aba["resultData"][6]["itemOutputModel"]['resultData']["items"])
            id = "https://sisgeap.geap.com.br/auditoriadigital/guia/" + link_guia.replace("https://apicore.geap.com.br/auditoriadigital/v1/Guias/guias/", "")
            motivo_glosa = []

            for i in range(0, quant_procedimento):
                id_procedimento = data_aba["resultData"][6]["itemOutputModel"]['resultData']["items"][i]["id"]
                link_procedimento = f'https://apicore.geap.com.br/auditoriadigital/v1/Itens/{id_procedimento}/historico-revisoes?tipoItem=Procedimentos'
                response_procedimento = requests.get(url=link_procedimento, headers=authorization, proxies=self.proxy, verify=False)
                data_procedimento = response_procedimento.json()
                ultimo_motivo = len(data_procedimento["resultData"]["items"]) - 1
                codigo_procedimento = data_procedimento["resultData"]["items"][ultimo_motivo]["codigo"]

                try:
                    justificativa = data_procedimento["resultData"]["items"][ultimo_motivo]["justificativa"]
                    motivo_glosa.append(f'{codigo_procedimento} - {justificativa}')

                except:
                    pass

            motivos = "/".join(motivo_glosa)    
            df = pd.DataFrame({'Endereço': [id], 'GUIA': [n_amhp], 'PACIENTE': [nome_paciente], 'CARTEIRINHA': [carteirinha], 'MOTIVO DE GLOSA': [motivos], 'SITUAÇÃO': [""],
                          'RESPONSAVEL': [""], 'REVISÃO TATIANE': [""], 'OBSERVAÇÃO': [""], 'PARECER TÉCNICO - DR. RICARDO': [""]})
            global lista_df
            lista_df = df.values.tolist()
            if __name__ == '__main__':
                ImportarGoogleSheets().main()

    def atualizar_situacao(self):
        updater = ImportarGoogleSheets()
        creds = updater.authenticate()
        updater.get_spreadsheet(creds)
        cabecalho = values.pop(0)
        # cabecalho.append('1')
        df = pd.DataFrame(values)
        df.columns = cabecalho
        token = input('Cole aqui o token: ')
        authorization = {
    "Authorization": f'Bearer {token}'
        }

        for index, linha in df.iterrows():
            if linha['SITUAÇÃO'] == 'Revisão Geap':
                id = str(linha['ENDEREÇO']).replace('https://sisgeap.geap.com.br/auditoriadigital/guia/', '')
                url = f"https://apicore.geap.com.br/auditoriadigital/v1/Guias/guias/{id}"
                response = requests.get(url=url, headers=authorization, proxies=self.proxy, verify=False)
                data = response.json()
                situacao = data["resultData"]["situacao"]

                match situacao:
                    case 'RevisaoGeap':
                        continue
                    case 'RevisaoPrestador':
                        valor = [['Revisão Prestador ']]
                        ImportarGoogleSheets().update_situacao(str(index + 2), valor)
                    case 'Consensuada':
                        valor = [['Consensuada']]
                        ImportarGoogleSheets().update_situacao(str(index + 2), valor)
                    case 'Cancelada':
                        valor = [['Guia cancelada ']]
                        ImportarGoogleSheets().update_situacao(str(index + 2), valor)

    def pegar_dados_pendentes(self):
        self.acessar_pendentes()
        lista_id = self.pegar_id_pendentes()
        lista = []
        token = input('Cole aqui o token: ')
        authorization = {
    "Authorization": f'Bearer {token}'
        }
        for link_guia in lista_id:    
            response = requests.get(url=link_guia, headers=authorization, proxies=self.proxy, verify=False)
            data = response.json()
            n_amhp = data['resultData']["numeroGuiaPrestador"]
            nome_paciente = data['resultData']["nomeBeneficiario"]
            carteirinha = data["resultData"]["numeroCartao"]
            data_atendimento = data["resultData"]["dataHoraAtendimento"]
            vet_data_atendimento = data_atendimento.split('T')
            vet_data_atendimento = vet_data_atendimento[0].split('-')
            data_atendimento = f"{vet_data_atendimento[2]}/{vet_data_atendimento[1]}/{vet_data_atendimento[0]}"
            data_importacao = data["resultData"]["dataAtualizacao"]
            vet_data_importacao = data_importacao.split('T')
            vet_data_importacao = vet_data_importacao[0].split('-')
            data_importacao = f"{vet_data_importacao[2]}/{vet_data_importacao[1]}/{vet_data_importacao[0]}"
            link_aba = link_guia + '/abas'
            response_aba = requests.get(url=link_aba, headers=authorization, proxies=self.proxy, verify=False)
            data_aba = response_aba.json()
            quant_procedimento = len(data_aba["resultData"][6]["itemOutputModel"]["resultData"]["items"])
            valor_total = 0.0

            for i in range(0, quant_procedimento):
                valor_proc = float(data_aba["resultData"][6]["itemOutputModel"]["resultData"]["items"][i]["valorTotal"])
                valor_total += valor_proc
            
            lista_guia = [n_amhp, nome_paciente, carteirinha, data_atendimento, data_importacao, valor_total]
            lista.append(lista_guia)
        
        cabecalho = ["Nº AMHPTISS", "Nome do Paciente", "Matrícula", "Data Atendimento", "Data Importação", "Valor Total"]
        df = pd.DataFrame(lista)
        df.columns = cabecalho
        df.to_excel(r'C:\Users\lucas.paz\Documents\Planilhas\Geap.xlsx')

class SisGeap(PageElement):
    input_usuario = '/html/body/div[1]/div[2]/div/div/div[1]/div/div/div/div/div[2]/div/div[1]/div/label[1]/div/div[1]/div/input'
    input_senha = '/html/body/div[1]/div[2]/div/div/div[1]/div/div/div/div/div[2]/div/div[1]/div/label[2]/div/div[1]/div[1]/input'
    btn_entrar = '/html/body/div[1]/div[2]/div/div/div[1]/div/div/div/div/div[2]/div/div[2]/button'

    def __init__(self, driver: WebDriver, url: str, usuario, senha) -> None:
        super().__init__(driver, url)
        self.usuario = usuario
        self.senha = senha

    def login(self):
        self.driver.find_element(*self.input_usuario).send_keys(self.usuario)
        sleep(2)
        self.driver.find_element(*self.input_senha).send_keys(self.senha)
        sleep(2)
        self.driver.find_element(*self.btn_entrar).click()
        sleep(2)

    def get_token(self):
        self.open()
        self.login()
        local_storage = self.driver.execute_script('localstorage;')

class ImportarGoogleSheets(ExtrairDados):
    def __init__(self)-> None:
        # Autenticação no proxy
        os.environ['HTTP_PROXY'] = 'http://10.0.0.230:3128'
        os.environ['HTTPS_PROXY'] = 'http://10.0.0.230:3128'

        # Se modificar esse scopo exclua o token.json.
        self.SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

        # ID é o código para rastrear a planilha.
        self.SAMPLE_SPREADSHEET_ID = '1gEGk8OUD9fvuVrIdEgFrPJvtn1OLrOX1i1CQMxezTs4'
        self.SAMPLE_RANGE_NAME = 'A1:Z1000'

    def authenticate(self)-> None:
        creds = None

        # O arquivo token.json armazena os tokens de acesso e atualização do usuário e é
        # criado automaticamente quando o fluxo de autorização é concluído pela primeira vez
        if os.path.exists('token.json'):
            print('usa token existente')
            creds = Credentials.from_authorized_user_file('token.json', self.SCOPES)

        # Se não houver credenciais (válidas) disponíveis, deixe o usuário fazer login.
        if not creds or not creds.valid:
            print('novo token')

            if creds and creds.expired and creds.refresh_token:
                print('refresh token')
                creds.refresh(Request())
            else:
                print('define parametros servidor local')
                flow = InstalledAppFlow.from_client_secrets_file("credentials.json", self.SCOPES)
                print('inicia servidor local')
                creds = flow.run_local_server(port=0)
                print('servidor local iniciado')

            # Salve as credenciais para a próxima execução
            with open('token.json', 'w') as token:
                print('write token')
                token.write(creds.to_json())

        return creds

    def get_spreadsheet(self, creds)-> None:
        try:
            global service
            service = build('sheets', 'v4', credentials=creds)

            # chamar a planilha
            sheet = service.spreadsheets()
            result = sheet.values().get(spreadsheetId=self.SAMPLE_SPREADSHEET_ID,
                                        range=self.SAMPLE_RANGE_NAME).execute()
            global values
            values = result.get('values', [])
        except HttpError as err:
            print(err)

    def injetar_dados(self):        
        try:
            sheet_size = str(len(values) + 1)

            # Injeta os dados na planilha do Google Sheets.
            sheet = service.spreadsheets()
            result = sheet.values().update(spreadsheetId=self.SAMPLE_SPREADSHEET_ID, range='A' + sheet_size,
                                           valueInputOption="USER_ENTERED", body={'values': lista_df}).execute()


        except HttpError as err:
            print(err)

    def update_situacao(self, num, valor):
        try:
            sheet_size = str(len(values) + 1)

            # Injeta os dados na planilha do Google Sheets.
            sheet = service.spreadsheets()
            result = sheet.values().update(spreadsheetId=self.SAMPLE_SPREADSHEET_ID, range='F' + num,
                                           valueInputOption="USER_ENTERED", body={'values': valor}).execute()


        except HttpError as err:
            print(err)

    def main(self):
        updater = ImportarGoogleSheets()
        creds = updater.authenticate()
        updater.get_spreadsheet(creds)
        updater.injetar_dados()

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------
url = r'https://login.geap.com.br/account/signin?p=%2Fconnect%2Fauthorize%2Fcallback%3Fclient_id%3Dgeap.auditoria.digital%26redirect_uri%3Dhttps%253A%252F%252Fsisgeap.geap.com.br%252Fauditoriadigital%252Flogin-callback%26response_type%3Dcode%26scope%3Dopenid%2520profile%252017_E%252037_E%25201104_S%25201104_U%25201104_I%25201104_D%252016_E%26state%3Df2199211ba7b494fa63da011dda15601%26code_challenge%3DrY3xZIoRaQrVs7uAVW1oq8HOFGVeLVGs_0RIytfgPJo%26code_challenge_method%3DS256%26response_mode%3Dquery'
options = {
    'proxy' : {
        'http': 'http://lucas.paz:WheySesc2024*@10.0.0.230:3128',
        'https': 'http://lucas.paz:WheySesc2024*@10.0.0.230:3128'
    },
    'verify_ssl': False
}

chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument('--ignore-certificate-errors')
try:
    servico = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=servico, seleniumwire_options= options, options = chrome_options)
except Exception as err:
    print(err)
    driver = webdriver.Chrome(seleniumwire_options= options, options = chrome_options)

usuario = "66661692120"
senha = "Amhp2023"
SisGeap(driver, url, usuario, senha).get_token()
ExtrairDados().atualizar_situacao()
ExtrairDados().extrair_dados()