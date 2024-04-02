import time
from tkinter.filedialog import askdirectory
from tkinter.messagebox import showinfo, showerror
import pyautogui
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver

from classes.connect_med import ConnectMed


def recursar_petrobras(user: str, password: str) -> None:
    try:
        showinfo('', 'Selecione a pasta com as planilhas.')
        diretorio_planilhas = askdirectory()
        showinfo('', 'Selecione a pasta com os anexos')
        diretorio_anexos = askdirectory()

        url = 'https://portaltiss.saudepetrobras.com.br/saudeweb/seguranca/login'

        proxies = {
                'http': f'http://{user}:{password}@10.0.0.230:3128',
                'https': f'http://{user}:{password}@10.0.0.230:3128'
            }

        options = {
            'proxy' : proxies
        }

        chrome_options = Options()
        chrome_options.add_argument("--start-maximized")
        # chrome_options.add_argument('--ignore-certificate-errors')
        # chrome_options.add_argument('--ignore-ssl-errors')
        try:
            servico = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=servico, options = chrome_options)
        except:
            driver = webdriver.Chrome(options = chrome_options)

        usuario = "glosaamhp"
         
        senha = "REC0073"

        connect_med = ConnectMed(driver, url, usuario, senha, proxies, diretorio_planilhas, diretorio_anexos, "Petrobras")
        connect_med.open()
        pyautogui.write(user.lower())
        pyautogui.press("TAB")
        time.sleep(1)
        pyautogui.write(password)
        pyautogui.press("enter")
        time.sleep(4)
        connect_med.login()
        connect_med.exec_recurso()

    except Exception as e :
        showerror('Automação', f'Ocorreu uma exceção não trata\n{e.__class__.__name__}:\n{e}')