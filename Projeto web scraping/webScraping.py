# TODO Criar uma classe que receba uma função 
# TODO Fazer a função abrir um site e pesquisar sobre o slogan de uma empresa

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver import Keys, ActionChains
from selenium.webdriver.common.by import By
from time import sleep
import dotenv
import os

class Nomus:
    
    def __init__(self, username, password):
        
        self.url = 'https://metalservtech.nomus.com.br/metalservtech/Login.do?metodo=LogOff'
        self.username = username
        self.password = password
        self.chromedriverPath = r'C:\Users\jogol\AppData\Local\Programs\chromedriver-win64'
        self.service = Service(executable_path=self.chromedriverPath)
        self.driver = webdriver.Chrome()

    def loginNomus(self):

        self.driver.get(self.url)

        text_input = self.driver.find_element(By.NAME, "login")
        ActionChains(self.driver)\
            .send_keys_to_element(text_input, self.username)\
            .perform()

        text_input = self.driver.find_element(By.NAME, "senha")
        ActionChains(self.driver)\
            .send_keys_to_element(text_input, self.password)\
            .perform()
        
        click = self.driver.find_element(By.NAME, "metodo")
        ActionChains(self.driver)\
            .click(click)\
            .perform()
        sleep(2)

    def acessarPagina(self):

        click = self.driver.find_element(By.XPATH, './/span[@title="Engenharia"]')
        ActionChains(self.driver)\
            .click(click)\
            .perform()

        click = self.driver.find_element(By.XPATH, './/span[@title="Produtos"]')
        ActionChains(self.driver)\
            .click(click)\
            .perform()
        sleep(2)

    def criarProduto(self):
        
        click = self.driver.find_element(By.ID, "botao_criar_produto")
        ActionChains(self.driver)\
            .click(click)\
            .perform()
        sleep(20)
        self.driver.quit()

dotenv.load_dotenv()
loginOS = os.getenv("LOGIN")
password = os.getenv("PASSWORD")

nomus = Nomus(loginOS, password)
sleep(5)
nomus.loginNomus()
nomus.acessarPagina()
nomus.criarProduto()

sleep(10)