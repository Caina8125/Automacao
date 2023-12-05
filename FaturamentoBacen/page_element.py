from abc import ABC

class PageElement(ABC):
    def __init__(self, driver, url, by_xpath)-> None:
        self.driver = driver
        self.url = url
        self.by_xpath = by_xpath
        
    def open(self)-> None:
        self.driver.get(self.url)