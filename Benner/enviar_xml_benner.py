from selenium.webdriver.chrome.options import Options
from Benner.benner import Benner
from seleniumwire.webdriver import Chrome
from tkinter.filedialog import askdirectory

def enviar_xml_benner(user, password) -> None:
    URL = r'https://portalconectasaude.com.br/Account/Login?ReturnUrl=%2FHome%2FIndex'

    DIRETORIO = askdirectory()

    PROXY: dict = {
    'proxy': {
            'http': f'http://{user}:{password}@10.0.0.230:3128',
            'https': f'http://{user}:{password}@10.0.0.230:3128'
        }
    }

    chrome_options: Options = Options()          
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument('--ignore-certificate-errors')
    chrome_options.add_argument('--ignore-ssl-errors')
    chrome_options.add_argument('--kiosk-printing')


    EMAIL: str = 'negociacao.gerencia@amhp.com.br'
    SENHA: str = 'Amhp@0073'

    DRIVER: Chrome = Chrome(seleniumwire_options=PROXY, options=chrome_options)

    envio_xml_benner = Benner(DRIVER, URL, EMAIL, SENHA)
    envio_xml_benner.exec_envio_xml(DIRETORIO)