# TODO Criar uma classe que receba uma função 
# TODO Fazer a função abrir um site e pesquisar sobre o slogan de uma empresa

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
import dotenv
import os
tamanho = 4.75

class Nomus:
    
    def __init__(self, username, password):
        
        self.url = 'https://metalservtech.nomus.com.br/metalservtech/Login.do?metodo=LogOff'
        self.username = username
        self.password = password
        self.chromedriverPath = r'C:\Users\Thalles\AppData\Local\Programs\chromedriver-win64'
        self.service = Service(executable_path=self.chromedriverPath)
        self.driver = webdriver.Chrome()

    def login_nomus(self):

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
        sleep(1)

    def acessar_pagina(self):

        click = self.driver.find_element(By.XPATH, './/span[@title="Engenharia"]')
        ActionChains(self.driver)\
            .click(click)\
            .perform()

        click = self.driver.find_element(By.XPATH, './/span[@title="Produtos"]')
        ActionChains(self.driver)\
            .click(click)\
            .perform()
        sleep(1)

    def criar_produto(self):
        
        click = self.driver.find_element(By.ID, "botao_criar_produto")
        ActionChains(self.driver)\
            .click(click)\
            .perform()

    def preencher_campos_producao_propria(self):
        
        # Descrição do produto
        text_input = self.driver.find_element(By.NAME, "descricao")
        ActionChains(self.driver)\
            .send_keys_to_element(text_input, "Teste")\
            .perform()
        
        # Unidade de medida
        select_element = self.driver.find_element(By.NAME, "idUnidadeMedida")
        select = Select(select_element)
        select.select_by_value("4")

        # Tipo de produto
        select_element = self.driver.find_element(By.NAME, "idTipoProduto")
        select = Select(select_element)
        select.select_by_value("8")

        # Clicando no Fiscal
        click = self.driver.find_element(By.ID, "ui-id-29")
        ActionChains(self.driver)\
            .click(click)\
            .perform()

        # Digitando NCM
        text_input = self.driver.find_element(By.NAME, "nomeNcm")

        if tamanho <= 3.1:
            ActionChains(self.driver)\
                .send_keys_to_element(text_input, "72085400")\
                .perform()
        elif tamanho > 3.2 <= 4.75:
            ActionChains(self.driver)\
                .send_keys_to_element(text_input, "72085300")\
                .perform()
        elif tamanho > 4.76 <= 10:
            ActionChains(self.driver)\
                .send_keys_to_element(text_input, "72085200")\
                .perform()
        elif tamanho > 10.1:
            ActionChains(self.driver)\
                .send_keys_to_element(text_input, "72085100")\
                .perform()

        WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located((By.ID, "ui-id-43")))
        
        # Selecionando o NCM
        item = self.driver.find_element(By.CLASS_NAME, "ui-menu-item")
        ActionChains(self.driver)\
            .click(item)\
            .perform()
        
        # Seleciona a origem do produto
        origem = self.driver.find_element(By.NAME, "origemProdutoPadrao")
        select = Select(origem)
        select.select_by_value("0")

        # Seleciona o tipo de venda
        tipo_de_venda = self.driver.find_element(By.NAME, "idTipoMovimentacaoPadraoVenda")
        select = Select(tipo_de_venda)
        select.select_by_value("53")

        # Clica no PCP
        click = self.driver.find_element(By.ID, "ui-id-34")
        ActionChains(self.driver)\
            .click(click)\
            .perform()
        
        # Seleciona tipo de ordem de produção
        tipo_de_ordem = self.driver.find_element(By.NAME, "idTipoOrdemPadrao")
        select = Select(tipo_de_ordem)
        select.select_by_value("1")

        # Seleciona tipo de movimentação do reporte da produção
        tipo_de_reporte = self.driver.find_element(By.NAME, "idTipoMovimentacaoPadraoReporte")
        select = Select(tipo_de_reporte)
        select.select_by_value("6")

        # Seleciona setor de estoque
        estoque = self.driver.find_element(By.NAME, "idSetorEstoqueEntradaPadraoReporte")
        select = Select(estoque)
        select.select_by_value("4")

        # Cria o produto
        click = self.driver.find_element(By.ID, "botao_salvar")
        ActionChains(self.driver)\
            .click(click)\
            .perform()
        
dotenv.load_dotenv()
loginOS = os.getenv("LOGIN")
password = os.getenv("PASSWORD")

nomus = Nomus(loginOS, password)
sleep(5)
nomus.login_nomus()
nomus.acessar_pagina()
nomus.criar_produto()
nomus.preencher_campos_producao_propria()

sleep(200)