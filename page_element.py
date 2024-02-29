from abc import ABC
from selenium.webdriver import Chrome, Firefox, Edge
from selenium.webdriver.common.by import By

class PageElement(ABC):
    body: tuple = (By.XPATH, '/html/body')
    a: tuple = (By.TAG_NAME, 'a')
    p: tuple = (By.TAG_NAME, 'p')
    h1: tuple = (By.TAG_NAME, 'h1')
    h2: tuple = (By.TAG_NAME, 'h2')
    h3: tuple = (By.TAG_NAME, 'h3')
    h4: tuple = (By.TAG_NAME, 'h4')
    h5: tuple = (By.TAG_NAME, 'h5')
    h6: tuple = (By.TAG_NAME, 'h6')

    def __init__(self, driver: Chrome | Firefox | Edge, url: str = '') -> None:
        self.driver: Chrome = driver
        self.url:str = url

    def open(self) -> None:
        self.driver.get(self.url)