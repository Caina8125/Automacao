from Benner.benner import Benner
from selenium.webdriver.chrome.options import Options
from seleniumwire.webdriver import Chrome
from tkinter.filedialog import askdirectory

URL = r'https://portalconectasaude.com.br/Account/Login?ReturnUrl=%2FHome%2FIndex'

DIRETORIO = askdirectory()

chrome_options: Options = Options()          
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument('--ignore-certificate-errors')
chrome_options.add_argument('--ignore-ssl-errors')
chrome_options.add_argument('--kiosk-printing')


EMAIL: str = 'negociacao.gerencia@amhp.com.br'
SENHA: str = 'Amhp@0073'

DRIVER: Chrome = Chrome()

envio_xml_benner = Benner(DRIVER, URL, EMAIL, SENHA)
envio_xml_benner.exec_envio_xml(DIRETORIO)