import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


class GoogleSheets:
    def __init__(self, spreadsheet_id)-> None:
        # Autenticação no proxy
        # os.environ['HTTP_PROXY'] = 'http://10.0.0.230:3128'
        # os.environ['HTTPS_PROXY'] = 'http://10.0.0.230:3128'

        # Se modificar esse scopo exclua o token.json.
        self.SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

        # ID é o código para rastrear a planilha.
        self.SAMPLE_SPREADSHEET_ID = spreadsheet_id
        self.SAMPLE_RANGE_NAME = 'A2:Z1000'

    def authenticate(self):
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
    
    def sheet_values(self, creds):
        try:
            service = build('sheets', 'v4', credentials=creds)

            # chamar a planilha
            sheet = service.spreadsheets()
            result = sheet.values().get(spreadsheetId=self.SAMPLE_SPREADSHEET_ID,
                                        range=self.SAMPLE_RANGE_NAME).execute()
            return result.get('values', [])

        except HttpError as err:
            print(err)

    def sheet_columns(self, creds):
        try:
            service = build('sheets', 'v4', credentials=creds)

            # chamar a planilha
            sheet = service.spreadsheets()
            result = sheet.values().get(spreadsheetId=self.SAMPLE_SPREADSHEET_ID,
                                        range='A1:Z1').execute()
            if isinstance(result.get('values', []), list):
                return result.get('values', []).pop(0)

        except HttpError as err:
            print(err)

    def send_values(self, creds, intervalo, valores):
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        result = sheet.values().update(spreadsheetId=self.SAMPLE_SPREADSHEET_ID, range=intervalo,
                                       valueInputOption="USER_ENTERED", body={'values': valores}).execute()

    # def update_spreadsheet(self, creds)-> None:
    #     try:
    #         service = build('sheets', 'v4', credentials=creds)

    #         # chamar a planilha
    #         sheet = service.spreadsheets()
    #         result = sheet.values().get(spreadsheetId=self.SAMPLE_SPREADSHEET_ID,
    #                                     range=self.SAMPLE_RANGE_NAME).execute()
    #         values = result.get('values', [])
    #         sheet_size = str(len(values) + 1)

    #         # Transforma as linhas da planilha em listas
    #         df = pd.read_excel(planilha, sheet_name="Plan2")
    #         lista = df.values.tolist()

    #         # Injeta os dados na planilha do Google Sheets.
    #         sheet = service.spreadsheets()
    #         result = sheet.values().update(spreadsheetId=self.SAMPLE_SPREADSHEET_ID, range='A' + sheet_size,
    #                                        valueInputOption="USER_ENTERED", body={'values': lista}).execute()


    #     except HttpError as err:
    #         print(err)

#---------------------------------------------------------------------------------