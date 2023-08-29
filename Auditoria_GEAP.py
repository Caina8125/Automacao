import requests
import pandas as pd
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import os

class LogarGeap():
    proxies = {
    'http': '10.0.0.230:3128',
    'https': 'lucas.paz:Gsw2022&@10.0.0.230:3128'
    }
    login = {
    "username": "23003723",
    "password": "amhpdf0073",
    "nrotpousuario": "1",
    "grant_type": "password",
    "CpfMultiusuario": "66661692120"
}
    data = []
    token = ''
    headers = {}
    url = 'https://wwwapi.geap.com.br/authentication/api/Token'
    
    def logar(self):
        response_login = requests.post(url=self.url, data=self.login, proxies=self.proxies)
        self.data = response_login.json()

    def gerar_token(self):
        self.token = self.data["access_token"]
        self.headers = {'Authorization': f'Bearer {self.token}'}
    
class ExtrairDados(LogarGeap):
    url_revisao = 'https://wwwapi.geap.com.br/AuditoriaDigital/api/v1/prestadores/undefined/guias?PageSize=10000000&PageNumber=1&Situacao=3&Filtro=&Ordenacao=Crescente'
    lista_id = []
    data_revisao = []
    quant_guia = 0
    lista_df = []

    def acessar_revisao_prestador(self):
        response = requests.get(url=self.url_revisao, headers=self.headers, proxies=self.proxies)
        self.data_revisao = response.json()
        self.quant_guia = len(self.data_revisao["ResultData"]["Items"])

    def pegar_id(self):
        updater = ImportarGoogleSheets()
        creds = updater.authenticate()
        updater.get_spreadsheet(creds)

        for i in range(0, self.quant_guia):
            id = str(self.data_revisao["ResultData"]["Items"][i]["Id"])
            id_encontrado = False
            for lista in values:
                link_plan = lista[0]
                if id in link_plan:
                    id_encontrado = True
                    break
                else:
                    continue
            if id_encontrado == False:
                self.lista_id.append(f"https://wwwapi.geap.com.br/AuditoriaDigital/api/v1/guias/{id}")
        
    def extrair_dados(self):
        self.logar()
        self.gerar_token()
        self.acessar_revisao_prestador()
        self.pegar_id()

        for link_guia in self.lista_id:
            
            response = requests.get(url=link_guia, headers=self.headers, proxies=self.proxies)
            print(response)
            data = response.json()
            n_amhp = data['ResultData']["NumeroGuiaPrestador"]
            nome_paciente = data['ResultData']["NomeBeneficiario"]
            carteirinha = data["ResultData"]["NumeroCartao"]
            link_aba = link_guia + '/abas'
            response_aba = requests.get(url=link_aba, headers=self.headers, proxies=self.proxies)
            data_aba = response_aba.json()
            quant_procedimento = len(data_aba["ResultData"][6]["ResultData"]["Items"])
            id = "https://www2.geap.com.br/AuditoriaDigital/guia/" + link_guia.replace("https://wwwapi.geap.com.br/AuditoriaDigital/api/v1/guias/", "")
            motivo_glosa = []
            for i in range(0, quant_procedimento):
                id_procedimento = data_aba["ResultData"][6]["ResultData"]["Items"][i]["Id"]
                link_procedimento = f'https://wwwapi.geap.com.br/AuditoriaDigital/api/v1/itens/{id_procedimento}/historico-revisoes?tipoItem=Procedimentos'
                response_procedimento = requests.get(url=link_procedimento, headers=self.headers, proxies=self.proxies)
                data_procedimento = response_procedimento.json()
                ultimo_motivo = len(data_procedimento["ResultData"]["Items"]) - 1
                codigo_procedimento = data_procedimento["ResultData"]["Items"][ultimo_motivo]["Codigo"]
                try:
                    justificativa = data_procedimento["ResultData"]["Items"][ultimo_motivo]["Justificativa"]
                    motivo_glosa.append(f'{codigo_procedimento} - {justificativa}')

                except:
                    pass
            motivos = "/".join(motivo_glosa)    
            df = pd.DataFrame({'Endereço': [id], 'GUIA': [n_amhp], 'PACIENTE': [nome_paciente], 'CARTEIRINHA': [carteirinha], 'MOTIVO DE GLOSA': [motivos], 'SITUAÇÃO': [""],
                          'RESPONSAVEL': [""], 'REVISÃO TATIANE': [""], 'OBSERVAÇÃO': [""], 'PARECER TÉCNICO - DR. RICARDO': [""]})
            global lista_df
            lista_df = df.values.tolist()
            ImportarGoogleSheets().main()

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

    def main(self):
        updater = ImportarGoogleSheets()
        creds = updater.authenticate()
        updater.get_spreadsheet(creds)
        updater.injetar_dados()
