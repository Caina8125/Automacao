from abc import ABC
from selenium.webdriver import Chrome

class PageElement(ABC):
    def __init__(self, driver: Chrome, url: str = ''):
        self.driver: Chrome = driver
        self.url:str = url
    def open(self):
        self.driver.get(self.url)