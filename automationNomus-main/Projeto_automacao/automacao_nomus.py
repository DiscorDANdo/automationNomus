from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver import Keys, ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
import dotenv
import os
import pdfplumber
import re

class Nomus:
    
    def __init__(self, username, password, caminho_pdf):
        
        self.url = 'https://metalservtech.nomus.com.br/metalservtech/Login.do?metodo=LogOff'
        self.username = username
        self.password = password
        self.chromedriverPath = r'C:\Users\Thalles\AppData\Local\Programs\chromedriver-win64'
        self.service = Service(executable_path=self.chromedriverPath)
        self.driver = webdriver.Chrome()
        self.leitura = LeituraPDF(caminho_pdf)
        self.paginas = LeituraPDF.extrair_paginas(self.leitura)
        self.itens_verificados = []

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

    def verificar_pecas(self, pagina):

        # Digitando nome do produto
        try:
            text_input = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.NAME, "descricaoPesquisa"))
            )
            ActionChains(self.driver)\
                .send_keys_to_element(text_input, self.leitura.encontrar_codigo(pagina))\
                .perform()
        except Exception as e:
            print(f"Erro: {e}")
        
        # Pesquisa o produto
        try:
            click = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.ID, "botao_pesquisar"))
            )
            ActionChains(self.driver)\
                .click(click)\
                .perform()
        except Exception as e:
            print(f"Erro: {e}")
        
        # Verificando se a peça já exite no sistema
        try:
            WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.CLASS_NAME, "tela-vazia"))
            )
            self.itens_verificados.append(self.leitura.encontrar_codigo(pagina))
        except Exception as e:  
            print(f"Erro: {e}")

        try:
            campo_pesquisa = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.NAME, "descricaoPesquisa")))
            ActionChains(self.driver)\
                .click(campo_pesquisa)\
                .perform()
            
            ActionChains(self.driver)\
                .key_down(Keys.CONTROL)\
                .send_keys("a")\
                .key_up(Keys.CONTROL)\
                .perform()
            
            ActionChains(self.driver)\
                .send_keys(Keys.DELETE)\
                .perform()
        except Exception as e:
            print(f"Erro: {e}")
        

    def criar_produto(self):
        
        click = self.driver.find_element(By.ID, "botao_criar_produto")
        ActionChains(self.driver)\
            .click(click)\
            .perform()
        sleep(2)

    def preencher_campos(self, pagina):
        
        if self.leitura.encontrar_codigo(pagina) in self.itens_verificados:
            # Descrição do produto
            text_input = self.driver.find_element(By.NAME, "descricao")
            ActionChains(self.driver)\
                .send_keys_to_element(text_input, self.leitura.encontrar_codigo(pagina))\
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

            if self.leitura.encontrar_cliente(pagina) in "ADR":
                ActionChains(self.driver)\
                    .send_keys_to_element(text_input, "87169090")\
                    .perform()
            else:
                if self.leitura.encontrar_espessura(pagina) <= 3:
                    ActionChains(self.driver)\
                        .send_keys_to_element(text_input, "72085400")\
                        .perform()
                if 3.1 <= self.leitura.encontrar_espessura(pagina) <= 4.75:
                    ActionChains(self.driver)\
                        .send_keys_to_element(text_input, "72085300")\
                        .perform()
                if 4.76 < self.leitura.encontrar_espessura(pagina) <= 10:
                    ActionChains(self.driver)\
                        .send_keys_to_element(text_input, "72085200")\
                        .perform()
                if 10.1 < self.leitura.encontrar_espessura(pagina) <= 25:
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
        else:
            pass
        
class LeituraPDF:

    def __init__(self, caminho_arquivo):
        self.caminho_arquivo = caminho_arquivo
        self.texto = self.extrair_texto()
        # self.paginas = self.extrair_texto()

    def extrair_texto(self):
        paginas_texto = []

        with pdfplumber.open(self.caminho_arquivo) as pdf:
            for page in pdf.pages:
                extraido = page.extract_text()
                if extraido:
                    paginas_texto.append(extraido)
        return paginas_texto
    
    def extrair_paginas(self):
        paginas = 0

        with pdfplumber.open(self.caminho_arquivo) as pdf:
            for page in pdf.pages:
                extraido = page.extract_text()
                if extraido:
                    paginas += 1
        return paginas
    
    def encontrar_codigo(self, pagina=0):
        padrao = r"Código:\s*(.*?)\s*(?=FIM)"
        match = re.search(padrao, self.texto[pagina])
        
        if match:
            nome_peca = match.group(1).strip()
        else:
            nome_peca = "Peça não encontrada."
        
        return nome_peca
    
    def encontrar_espessura(self, pagina=0):
        padrao = r"Espessura:\s*(.*?)\s*(?=FIM)"
        match = re.search(padrao, self.texto[pagina])

        if match:
            espessura_peca = match.group(1)
            espessura_peca_convertido = float(espessura_peca.replace(",", "."))
        else:
            espessura_peca = "Espessura não encontrada."
        
        return espessura_peca_convertido
    
    def encontrar_material(self, pagina=0):
        padrao = r"Material:\s*(.*?)\s*(?=FIM)"
        match = re.search(padrao, self.texto[pagina])

        if match:
            material_peca = match.group(1).strip()
        else:
            material_peca = "Material não encontrado."

        return material_peca

    def encontrar_peso(self, pagina=0):
        padrao = r"Peso:\s*(.*?)\s*(?=FIM)"
        match = re.search(padrao, self.texto[pagina])
        
        if match:
            peso_peca = float(match.group(1).strip())
        else:
            peso_peca = "Peso não encontrado."
        
        return peso_peca
    
    def encontrar_dobra(self, pagina=0):
        padrao = r"Dobras:\s*(.*?)\s*(?=FIM)"
        match = re.search(padrao, self.texto[pagina])

        if match:
            dobra_peca = match.group(1).strip()
        else:
            dobra_peca = "Dobras não encontrada."
        
        return dobra_peca
    
    def encontrar_cliente(self, pagina=0):
        padrao = r"Cliente:\s*(.*?)\s*(?=FIM)"
        match = re.search(padrao, self.texto[pagina])

        if match:
            cliente = match.group(1).strip()
        else:
            cliente = "Cliente não encontrado."
        
        return cliente
    
    def encontrar_quantidade(self, pagina=0):
        padrao = r"Quantidade:\s*(.*?)\s*(?=FIM)"
        match = re.search(padrao, self.texto[pagina])

        if match:
            quantidade_peca = match.group(1).strip()
        else:
            quantidade_peca = "Quantidade não encontrada."

        return quantidade_peca

    def exibe_informacoes(self):
        print(f"""
    Código: {LeituraPDF.encontrar_codigo(self)}
    Espessura: {LeituraPDF.encontrar_espessura(self)}
    Material: {LeituraPDF.encontrar_material(self)}
    Peso: {LeituraPDF.encontrar_peso(self)}
    Dobras: {LeituraPDF.encontrar_dobra(self)}
    Cliente: {LeituraPDF.encontrar_cliente(self)}
    Quantidade: {LeituraPDF.encontrar_quantidade(self)}
    """)

dotenv.load_dotenv()
loginOS = os.getenv("LOGIN")
senhaOS = os.getenv("PASSWORD")
caminho_pdf = r"C:\Users\Thalles\Desktop\Listas\ListaAtualizada10.pdf"
automacao = Nomus(loginOS, senhaOS, caminho_pdf)

automacao.login_nomus()
automacao.acessar_pagina()

# Verificar se as peças já estão criadas
for pagina in range(automacao.paginas):
    automacao.verificar_pecas(pagina)
print(automacao.itens_verificados)

# Criar produto
for pagina in range(automacao.paginas):
    automacao.criar_produto()
    automacao.preencher_campos(pagina)