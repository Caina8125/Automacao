from selenium.webdriver.chrome.options import Options
from seleniumwire.webdriver import Chrome
from Benner.benner import Benner
from tkinter.filedialog import askdirectory
from datetime import date

def enviar_pdf_benner(user: str, password: str) -> None:
    chrome_options: Options = Options()          
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument('--ignore-certificate-errors')
    chrome_options.add_argument('--ignore-ssl-errors')
    chrome_options.add_argument('--kiosk-printing')

    URL = 'https://portalconectasaude.com.br/Account/Login'
    

    PROXY: dict = {
    'proxy': {
            'http': f'http://{user}:{password}@10.0.0.230:3128',
            'https': f'http://{user}:{password}@10.0.0.230:3128'
        }
    }

    DIRETORIO: str = askdirectory()

    DRIVER = Chrome(seleniumwire_options=PROXY, options=chrome_options)

    EMAIL: str = 'negociacao.gerencia@amhp.com.br'
    SENHA: str = 'Amhp@0073'
    DIA_ATUAL: int = date.today().day

    enviar_pdf: Benner = Benner(DRIVER, URL, EMAIL, SENHA)
    enviar_pdf.exec_anexo_guias(DIRETORIO, DIA_ATUAL)
    DRIVER.quit()