from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver import Keys, ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException
from leitura import leitura
from orcamento import orcamento
from icecream import ic
from time import sleep
from datetime import timedelta, date
import tkinter_class
import traceback
import dotenv
import os

#TODO: Finalizar a função de criar orçamento
#TODO: Adicionar as funções do orcamento.py no main()

class Nomus:
    
    def __init__(self, username, password, caminho_pdf):
        
        # Selenium
        self.url = 'https://metalservtech.nomus.com.br/metalservtech/'
        self.username = username
        self.password = password
        self.chromedriverPath = r'C:\Users\Programação\Documents\chromedriver\chromedriver-win64'
        self.service = Service(executable_path=self.chromedriverPath)
        self.driver = webdriver.Chrome()
        self.itens_verificados = []
        self.clientes_terceirizacao = ["ADR", "JLS", "FACTORY NAUTICA"]

        # leitura
        self.caminho_pdf = caminho_pdf
        self.leitura = leitura(self.caminho_pdf)
        self.dados = self.leitura.extrair_dados()
        self.peca = self.leitura.peca
        self.extracao_pecas = self.leitura.extrair_pecas()
        self.pecas = self.leitura.extrair_tamanho()

    def login_nomus(self):

        self.driver.get(self.url + 'Login.do?metodo=LogOff')

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

        try:
            self.driver.get(self.url + 'Produto.do?metodo=Pesquisar')
            
        except Exception as e:
            ic(f"Error acessing page: {e}")


    def verificar_pecas(self, peca: dict):

        while not self.driver.current_url.startswith(self.url + 'Produto.do?metodo=Pesquisar'):
            self.acessar_pagina()

        # Digitando nome do produto
        try:
            text_input = WebDriverWait(self.driver, 5).until(
                EC.visibility_of_element_located((By.NAME, "descricaoPesquisa"))
            )
            ActionChains(self.driver)\
                .send_keys_to_element(text_input, self.leitura.lista_pecas[peca]["Código"])\
                .perform()
        except Exception as e:
            ic(f"Error inserting product name: {e}")
        
        # Pesquisa o produto
        try:
            click = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.ID, "botao_pesquisar"))
            )
            ActionChains(self.driver)\
                .click(click)\
                .perform()
        
        
            # Verificando se a peça já exite no sistema
            peca_existente = True

            nome = self.leitura.lista_pecas[peca]["Código"].upper()
            # xpath = f"//span[contains(text(), '{nome}')]"
            xpath = f"//span[text()='{nome}']"

            try:
                WebDriverWait(self.driver, 2).until(
                    EC.visibility_of_element_located((By.XPATH, xpath))
                )
            except TimeoutException:
                ic("Product don't exists")
                peca_existente = False

                return peca_existente
        
            if peca_existente:
                pass
        except Exception as e:
            ic(f"Error verifying product: {e}")
        
        finally:
            if not peca_existente:
                self.itens_verificados.append(self.leitura.lista_pecas[peca].copy())
                ic("Item inserted in list successfully!")

            # Apaga informações do campo de pesquisa
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
    

    def criar_produto(self):
        
        self.driver.get(self.url + 'Produto.do?metodo=Criar_produto')

    def preencher_campos(self, peca : dict):
        
        while not self.driver.current_url.startswith(self.url + 'Produto.do?metodo=Criar_produto'):
            self.criar_produto()

        cod = self.leitura.lista_pecas[peca]["Código"]
        ic(f"Verificated items: {self.itens_verificados}")

        if any(item["Código"] == cod for item in self.itens_verificados):
            # Descrição do produto
            descricao_produto_criacaoPA = WebDriverWait(self.driver, 20).until(
                EC.visibility_of_element_located((By.NAME, "descricao"))
            )
            ActionChains(self.driver)\
                .send_keys_to_element(descricao_produto_criacaoPA, self.leitura.lista_pecas[peca]["Código"].upper())\
                .perform()
            
            # Unidade de medida
            unidade_medida_criacaoPA = WebDriverWait(self.driver, 20).until(
                EC.visibility_of_element_located((By.NAME, "idUnidadeMedida"))
            )
            select = Select(unidade_medida_criacaoPA)
            select.select_by_value("4")

            # Tipo de produto
            tipo_produto_criacaoPA = WebDriverWait(self.driver, 20).until(
                EC.visibility_of_element_located((By.NAME, "idTipoProduto"))
            )
            select = Select(tipo_produto_criacaoPA)
            select.select_by_value("8")

            # Grupo de produto
            grupo_produto_criacaoPA = WebDriverWait(self.driver, 20).until(
                EC.visibility_of_element_located((By.NAME, "idGrupoProduto"))
            )
            if "CLIENTE" in self.leitura.lista_pecas[peca]["Material"].upper():
                select = Select(grupo_produto_criacaoPA)
                select.select_by_value("21") # Industrialização
            else:
                select = Select(grupo_produto_criacaoPA)
                select.select_by_value("20") # Produção própria

            # Clicando no Fiscal
            click_fiscal_criacaoPA = WebDriverWait(self.driver, 20).until(
                EC.visibility_of_element_located((By.ID, "ui-id-29"))
            )
            ActionChains(self.driver)\
                .click(click_fiscal_criacaoPA)\
                .perform()

            # Digitando NCM
            ncm_criacaoPA = WebDriverWait(self.driver, 20).until(
                EC.visibility_of_element_located((By.NAME, "nomeNcm"))
            )

            if "ADR" in self.leitura.lista_pecas[peca]["Cliente"].upper():
                ActionChains(self.driver)\
                    .send_keys_to_element(ncm_criacaoPA, "87169090")\
                    .perform()
            else:
                if self.leitura.lista_pecas[peca]["Espessura"] <= 3:
                    ActionChains(self.driver)\
                    .send_keys_to_element(ncm_criacaoPA, "72085400")\
                        .perform()
                if 3.1 <= self.leitura.lista_pecas[peca]["Espessura"] <= 4.75:
                    ActionChains(self.driver)\
                        .send_keys_to_element(ncm_criacaoPA, "72085300")\
                        .perform()
                if 4.76 < self.leitura.lista_pecas[peca]["Espessura"] <= 10:
                    ActionChains(self.driver)\
                        .send_keys_to_element(ncm_criacaoPA, "72085200")\
                        .perform()
                if 10.1 < self.leitura.lista_pecas[peca]["Espessura"] <= 25:
                    ActionChains(self.driver)\
                        .send_keys_to_element(ncm_criacaoPA, "72085100")\
                        .perform()
                    
            # Espera abrir o menu de NCM
            WebDriverWait(self.driver, 20).until(EC.visibility_of_element_located((By.ID, "ui-id-43")))
            
            # Selecionando o NCM
            click_ncm_criacaoPA = WebDriverWait(self.driver, 20).until(
                EC.visibility_of_element_located((By.CLASS_NAME, "ui-menu-item"))
            )
            ActionChains(self.driver)\
                .click(click_ncm_criacaoPA)\
                .perform()
            
            # Seleciona a origem do produto
            select_origem_criacaoPA = WebDriverWait(self.driver, 20).until(
                EC.visibility_of_element_located((By.NAME, "origemProdutoPadrao"))
            )
            select = Select(select_origem_criacaoPA)
            select.select_by_value("0")

            # Seleciona o tipo de venda
            select_tipo_venda_criacaoPA = WebDriverWait(self.driver, 20).until(
                EC.visibility_of_element_located((By.NAME, "idTipoMovimentacaoPadraoVenda"))
            )
            if "Cliente" in self.leitura.lista_pecas[peca]["Material"]:
                select = Select(select_tipo_venda_criacaoPA)
                select.select_by_value("45") # Industrialização
            else:
                select = Select(select_tipo_venda_criacaoPA)
                select.select_by_value("53") # Produção própria

            # Clica no PCP
            click_pcp = WebDriverWait(self.driver, 20).until(
                EC.visibility_of_element_located((By.ID, "ui-id-34"))
            )
            ActionChains(self.driver)\
                .click(click_pcp)\
                .perform()
            
            # Seleciona tipo de ordem de produção
            seleciona_tipo_ordem_criacaoPA = WebDriverWait(self.driver, 20).until(
                EC.visibility_of_element_located((By.NAME, "idTipoOrdemPadrao"))
            )
            if self.leitura.lista_pecas[peca]["Material"] in "Carbono Cliente" or "Inox Cliente":
                select = Select(seleciona_tipo_ordem_criacaoPA)
                select.select_by_value("2") # Industrialização
            else:
                select = Select(seleciona_tipo_ordem_criacaoPA)
                select.select_by_value("1") # Produção própria

            # Seleciona tipo de movimentação do reporte da produção
            select_tipo_reporte_criacaoPA = WebDriverWait(self.driver, 20).until(
                EC.visibility_of_element_located((By.NAME, "idTipoMovimentacaoPadraoReporte"))
            )
            select = Select(select_tipo_reporte_criacaoPA)
            select.select_by_value("6")

            # Seleciona setor de estoque
            seleciona_setor_estoque_criacaoPA = WebDriverWait(self.driver, 20).until(
                EC.visibility_of_element_located((By.NAME, "idSetorEstoqueEntradaPadraoReporte"))
            )
            select = Select(seleciona_setor_estoque_criacaoPA)
            select.select_by_value("4")

            # Cria o produto
            click_botao_salvar_criacaoPA = WebDriverWait(self.driver, 20).until(
                EC.visibility_of_element_located((By.ID, "botao_salvar"))
            )
            ActionChains(self.driver)\
                .click(click_botao_salvar_criacaoPA)\
                .perform()

    def obter_mp(self, espessura, peca : dict):
        cliente = peca["Cliente"].upper()
        material = peca["Material"].upper()

        mp_metal = [
            ((1, 1.5), "MP 00002"),
            ((1.9, 2), "MP 00001"),
            ((3, 3.1), "MP 00019"),
            ((4.25, 4.25), "MP 00129"),
            ((4.75, 5), "MP 00003"),
            ((6, 6.3), "MP 00006"),
            ((7, 7.5), "MP 00017"),
            ((7.9, 8), "MP 00123"),
            ((9, 10.00), "MP 00032"),
            ((12, 12.7), "MP 00011"),
            ((15.8, 16), "MP 00036"),
            ((19, 20), "MP 00001"), # SEM CHAPA NO SISTEMA
            ((25, 25), "MP 00001"), # SEM CHAPA NO SISTEMA
        ]
        mp_adr = [
            ((4, 5), "MP 00026"),
            ((6, 6.3), "MP 00016"),
            ((7.9, 8), "MP 00014"),
            ((9, 9.5), "MP 00015"),
            ((12, 12.7), "MP 00009"),
            ((15, 16), "MP 00046"),
            ((19, 20), "MP 00035"),
            ((25, 25), "MP 00025"),
        ]
        mp_jls_carbono = [
            ((1, 1.2), "MP 00115"),
            ((1.4, 1.5), "MP 00076"),
            ((1.9, 2), "MP 00087"),
            ((2.5, 2.66), "MP 00112"),
            ((3, 3.1), "MP 00079"),
            ((3.7, 3.75), "MP 00080"),
            ((4, 4.76), "MP 00081"),
            ((6, 6.5), "MP 00089"),
            ((7.9, 8), "MP 00082"),
            ((9, 9.53), "MP 00083"),
            ((12, 12.7), "MP 00084"),
        ]
        mp_jls_inox = [
            ((1, 1.5), "MP 00090"),
            ((1.9, 2), "MP 00086"),
            ((3, 3.1), "MP 00085"),
            ((3.9, 4), "MP 00116"),
            ((4.75, 5), "MP 00094"),
            ((7.9, 8), "MP 00114"),
        ]

        if "ADR" in cliente:
            for (min_esp, max_esp), codigo in mp_adr:
                if min_esp <= espessura <= max_esp:
                    return codigo
                else:
                    ic("Material not found")
                    pass
        elif "JLS" in cliente:
            if "CARBONO" in material:
                for (min_esp, max_esp), codigo, in mp_jls_carbono:
                    if min_esp <= espessura <= max_esp:
                        return codigo
                    else:
                        ic("Material not found")
                        pass
            if "INOX" in material:
                for (min_esp, max_esp), codigo in mp_jls_inox:
                    if min_esp <= espessura <= max_esp:
                        return codigo
                    else:
                        ic("Material not found")
                        pass
        else:
            for (min_esp, max_esp), codigo in mp_metal:
                if min_esp <= espessura <= max_esp:
                    return codigo
                else:
                    ic("Material not found")
                    pass

    def criar_lista_materiais(self):
        for index, item in enumerate(self.leitura.lista_pecas):
            espessura = float(item["Espessura"])
            codigo_mp = self.obter_mp(espessura, item)

            # Verifica se está na pagina correta
            while not self.driver.current_url.startswith(self.url + "Produto.do?metodo=Pesquisar"):
                self.acessar_pagina()
            sleep(1)
            
            # Limpa o campo de pesquisa
            campo_pesquisa = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.NAME, "descricaoPesquisa")))
            campo_pesquisa.clear()
            sleep(1)

            # Digita o nome do produto
            digita_produto_pesquisa = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.NAME, "descricaoPesquisa"))
            )
            digita_produto_pesquisa.send_keys(item["Código"])
            
            # Clica no botão para pesquisar
            click_pesquisa_produto = WebDriverWait(self.driver, 10).until(
                    EC.visibility_of_element_located((By.ID, "botao_pesquisar"))
                )
            click_pesquisa_produto.click()
            sleep(1)
            
            # Encontra o produto na pagina e clica no produto
            nome = item["Código"].upper()
            xpath = f"//span[contains(text(), '{nome}')]"

            click_produto = self.driver.find_element(By.XPATH, xpath)
            click_produto.click()
            sleep(1)

            # Clica no submenu de lista de materiais
            click_submenu_lista_materiais = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.ID, "produtoAtivoLiberado_itemSubMenu_acessarListaMateriais"))
            )
            click_submenu_lista_materiais.click()
            sleep(1)

            # Verifica se já existe lista de materiais:
            botao_localizado = False
            try:
                WebDriverWait(self.driver, 5).until(
                    EC.visibility_of_element_located((By.ID, "botao_acessaradicionaritemestrutura"))
                )
                botao_localizado = True
                return botao_localizado
            except TimeoutException:
                continue
            finally:
                if not botao_localizado:
                    # Clica em salvar para criar a lista de materiais
                    click_salvar_lista_materiais = WebDriverWait(self.driver, 10).until(
                        EC.visibility_of_element_located((By.ID, "botao_salvar"))
                    )
                    click_salvar_lista_materiais.click()

                    # Clica no botão para adicionar um item a estrutura
                    click_adicionar_item = WebDriverWait(self.driver, 10).until(
                        EC.visibility_of_element_located((By.ID, "botao_acessaradicionaritemestrutura"))
                    )
                    click_adicionar_item.click()
                    sleep(1)

                    # Clica no geral
                    click_geral = WebDriverWait(self.driver, 10).until(
                        EC.visibility_of_element_located((By.XPATH, '//*[@id="ui-id-3"]'))
                    )
                    click_geral.click()

                    # Digita o código da matéria prima
                    digita_materia_prima = WebDriverWait(self.driver, 10).until(
                        EC.visibility_of_element_located((By.ID, "nome_produto"))
                    )
                    sleep(1)
                    
                    # Verifica se o material é de terceiros
                    if item["Cliente"] in self.clientes_terceirizacao:
                        digita_materia_prima.send_keys(codigo_mp)
                        # Clica no material
                        click_material = WebDriverWait(self.driver, 10).until(
                            EC.visibility_of_element_located((By.CLASS_NAME, "ui-menu-item"))
                        )
                        click_material.click()
                        
                        # Digita o peso
                        digita_peso = WebDriverWait(self.driver, 10).until(
                            EC.visibility_of_element_located((By.NAME, "qtdeNecessariaProdutoFilho"))
                        )
                        digita_peso.send_keys(item["Peso (str)"])
                        
                        # Clica em avançado
                        click_avancado = WebDriverWait(self.driver, 5).until(
                            EC.visibility_of_element_located((By.XPATH, '//*[@id="ui-id-7"]'))
                        )
                        click_avancado.click()
                        
                        # Clica em recebe componentes de terceiros para industrialização sob encomenda 
                        click_industrializacao = WebDriverWait(self.driver, 5).until(
                            EC.visibility_of_element_located((By.ID, "id_recebe_componente_industrializacao"))
                        )
                        click_industrializacao.click()

                    else:    
                        digita_materia_prima.send_keys(codigo_mp)
                        # Clica no material
                        click_material = WebDriverWait(self.driver, 10).until(
                            EC.visibility_of_element_located((By.CLASS_NAME, "ui-menu-item"))
                        )
                        click_material.click()

                        # Digita o peso
                        digita_peso = WebDriverWait(self.driver, 10).until(
                            EC.visibility_of_element_located((By.NAME, "qtdeNecessariaProdutoFilho"))
                        )
                        digita_peso.send_keys(item["Peso (str)"])
                    
                    sleep(1)

                    # Salva a materia prima adicionada
                    click_salvar_materia_prima = WebDriverWait(self.driver, 10).until(
                        EC.visibility_of_element_located((By.ID, "botao_salvar"))
                    )
                    click_salvar_materia_prima.click()
                else:
                    continue

    def criar_roteiro(self, peca):

        # Verifica se está na página correta
        while not self.driver.current_url.startswith(self.url + 'Produto.do?metodo=Pesquisar'):
            self.acessar_pagina()

        try:
            # Apaga informações do campo de pesquisa
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
            ic(f"Error erasing info on descripton field: {e}")

        # Digita o nome do produto
        sleep(1)
        digita_produto_pesquisa = WebDriverWait(self.driver, 10).until(
            EC.visibility_of_element_located((By.NAME, "descricaoPesquisa"))
        )
        ActionChains(self.driver)\
            .send_keys_to_element(digita_produto_pesquisa, self.leitura.lista_pecas[peca]["Código"])\
            .perform()
        
        # Clica no botão para pesquisar
        click_pesquisa_produto = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.ID, "botao_pesquisar"))
            )
        ActionChains(self.driver)\
            .click(click_pesquisa_produto)\
            .perform()

        # Encontra o produto na pagina e clica no produto
        try:
            nome = self.leitura.lista_pecas[peca]["Código"].upper()
            # xpath = f"//span[contains(text(), '{nome}')]"
            xpath = f"//span[text()='{nome}']"

            click_produto = self.driver.find_element(By.XPATH, xpath)
            ActionChains(self.driver)\
                .click(click_produto)\
                .perform()
            
            # Seleciona o roteiro de produção
            click_submenu_roteiro = self.driver.find_element(By.ID, "produtoAtivoLiberado_itemSubMenu_acessarRoteiroProduto")
            ActionChains(self.driver)\
                .click(click_submenu_roteiro)\
                .perform()
        except Exception as e:
            ic(f"Error clicking on product: {e}")
            
        # Verifica se existe roteiro
        botao_localizado = False
        try:
            WebDriverWait(self.driver, 5).until(
                EC.visibility_of_element_located((By.ID, "botao_criar_operacao"))
            )
            botao_localizado = True
        except Exception:
            ic("Product script already exists.")

        if not botao_localizado:
        
            # Digita o nome do roteiro de produção
            digita_nome_roteiro = WebDriverWait(self.driver, 5).until(
                EC.visibility_of_element_located((By.NAME, "roteiroDescricao"))
            )
            ActionChains(self.driver)\
                .send_keys_to_element(digita_nome_roteiro, "PADRÃO")\
                .perform()

            # Salva o roteiro de produção
            click_cria_roteiro = WebDriverWait(self.driver, 5).until(
                EC.visibility_of_element_located((By.ID, "botao_salvar_como_roteiro_adicional"))
            )
            ActionChains(self.driver)\
                .click(click_cria_roteiro)\
                .perform()
            
            # Cria operação do roteiro de produção
            click_cria_operacao = WebDriverWait(self.driver, 5).until(
                EC.visibility_of_element_located((By.ID, "botao_criar_operacao"))
            )
            ActionChains(self.driver)\
                .click(click_cria_operacao)\
                .perform()
            
            # Digita o número da operação
            digita_numero_operacao = WebDriverWait(self.driver, 5).until(
                EC.visibility_of_element_located((By.NAME, "operacaoFase"))
            )
            ActionChains(self.driver)\
                .send_keys_to_element(digita_numero_operacao, "10")\
                .perform()
            
            # Digita o nome da operação
            digita_nome_operacao = WebDriverWait(self.driver, 5).until(
                EC.visibility_of_element_located((By.NAME, "operacaoDescricao"))
            )
            ActionChains(self.driver)\
                .send_keys_to_element(digita_nome_operacao, "CORTE")\
                .perform()
            
            # Seleciona o tipo de operação
            select_tipo_operacao = WebDriverWait(self.driver, 5).until(
                EC.visibility_of_element_located((By.NAME, "idTipoOperacao"))
            )
            select = Select(select_tipo_operacao)
            select.select_by_value("1")

            # Seleciona centro de trabalho
            select_centro_trabalho = WebDriverWait(self.driver, 5).until(
                EC.visibility_of_element_located((By.NAME, "operacaoCentroTrabalhoId"))
            )
            select = Select(select_centro_trabalho)
            select.select_by_value("1")

            # Digita o tempo de setup
            digita_tempo_setup = WebDriverWait(self.driver, 5).until(
                EC.visibility_of_element_located((By.NAME, "tempoSetupCentroTrabalho"))
            )
            ActionChains(self.driver)\
                .send_keys_to_element(digita_tempo_setup, "000000")\
                .perform()
            
            # Digita o tempo de operação
            digita_tempo_operacao = WebDriverWait(self.driver, 5).until(
                EC.visibility_of_element_located((By.NAME, "tempoOperacaoCentroTrabalho"))
            )
            ActionChains(self.driver)\
                .send_keys_to_element(digita_tempo_operacao, self.leitura.lista_pecas[peca]["Tempo"])\
                .perform()
            
            # Seleciona tipo de tempo de operação
            select_tipo_tempo_operacao = WebDriverWait(self.driver, 5).until(
                EC.visibility_of_element_located((By.NAME, "tipoTempoCentroTrabalho"))
            )
            select = Select(select_tipo_tempo_operacao)
            select.select_by_value("2")

            # Digita o tempo de MOD no setup
            digita_tempo_MOD_setup = WebDriverWait(self.driver, 5).until(
                EC.visibility_of_element_located((By.NAME, "tempoMODSetupCentroTrabalho"))
            )
            ActionChains(self.driver)\
                .send_keys_to_element(digita_tempo_MOD_setup, "000000")\
                .perform()
            
            # Digita o tempo de MOD na operação
            digita_tempo_MOD_operacao = WebDriverWait(self.driver, 5).until(
                EC.visibility_of_element_located((By.NAME, "tempoMODOperacaoCentroTrabalho"))
            )
            ActionChains(self.driver)\
                .send_keys_to_element(digita_tempo_MOD_operacao, self.leitura.lista_pecas[peca]["Tempo"])\
                .perform()
            
            # Verifica se existe dobra e converte a quantidade de string para int
            if self.leitura.lista_pecas[peca]["Dobras"] == 0:
                pass
            else:
                quantidade_dobra = int(self.leitura.lista_pecas[peca]["Dobras"])
 
                # Verifica se existe dobra
                if quantidade_dobra >= 1:
                    
                    # Clica em salvar e continuar na operação
                    click_salvar_continuar = WebDriverWait(self.driver, 5).until(
                        EC.visibility_of_element_located((By.XPATH, "//input[@type='button' and @id='botao_salvarepermanecernatela']"))
                    )
                    ActionChains(self.driver)\
                        .click(click_salvar_continuar)\
                        .perform()
                    
                    sleep(2)
                    # Digita o número da operação
                    digita_numero_operacao_dobra = WebDriverWait(self.driver, 5).until(
                        EC.visibility_of_element_located((By.NAME, "operacaoFase"))
                    )
                    ActionChains(self.driver)\
                        .send_keys_to_element(digita_numero_operacao_dobra, "20")\
                        .perform()
                    
                    # Digita o nome da operação
                    digita_nome_operacao_dobra = WebDriverWait(self.driver, 5).until(
                        EC.visibility_of_element_located((By.NAME, "operacaoDescricao"))
                    )
                    ActionChains(self.driver)\
                        .send_keys_to_element(digita_nome_operacao_dobra, "DOBRA")\
                        .perform()
                    
                    # Seleciona o tipo de operação
                    select_tipo_operacao_dobra = WebDriverWait(self.driver, 5).until(
                        EC.visibility_of_element_located((By.NAME, "idTipoOperacao"))
                    )
                    select = Select(select_tipo_operacao_dobra)
                    select.select_by_value("5")

                    # Seleciona centro de trabalho
                    select_centro_trabalho_dobra = WebDriverWait(self.driver, 5).until(
                        EC.visibility_of_element_located((By.NAME, "operacaoCentroTrabalhoId"))
                    )
                    select = Select(select_centro_trabalho_dobra)
                    select.select_by_value("3")

                    # Digita o tempo de setup
                    digita_tempo_setup_dobra = WebDriverWait(self.driver, 5).until(
                        EC.visibility_of_element_located((By.NAME, "tempoSetupCentroTrabalho"))
                    )
                    ActionChains(self.driver)\
                        .send_keys_to_element(digita_tempo_setup_dobra, "000000")\
                        .perform()
                    
                    # Digita o tempo de operação
                    tempo = 10
                    digita_tempo_operacao_dobra = WebDriverWait(self.driver, 5).until(
                        EC.visibility_of_element_located((By.NAME, "tempoOperacaoCentroTrabalho"))
                    )
                    
                    # Verifica a quantidade de dobras para fazer o cálculo do tempo
                    if quantidade_dobra <= 5:
                        for c in range(quantidade_dobra):
                            tempo += 10
                        tempo_formatado = str(timedelta(seconds=tempo)).zfill(8)

                        # Digita o tempo formatado
                        ActionChains(self.driver)\
                            .send_keys_to_element(digita_tempo_operacao_dobra, tempo_formatado)\
                            .perform()
                        
                    if quantidade_dobra >= 6:
                        for c in range(quantidade_dobra):
                            tempo += 50
                        tempo_formatado = str(timedelta(seconds=tempo)).zfill(8)

                        # Digita o tempo formatado
                        ActionChains(self.driver)\
                            .send_keys_to_element(digita_tempo_operacao_dobra, tempo_formatado)\
                            .perform()
                    
                    # Seleciona tipo de tempo de operação
                    select_tipo_tempo_operacao_dobra = WebDriverWait(self.driver, 5).until(
                        EC.visibility_of_element_located((By.NAME, "tipoTempoCentroTrabalho"))
                    )
                    select = Select(select_tipo_tempo_operacao_dobra)
                    select.select_by_value("2")

                    # Digita o tempo de MOD no setup
                    digita_tempo_MOD_setup_dobra = WebDriverWait(self.driver, 5).until(
                        EC.visibility_of_element_located((By.NAME, "tempoMODSetupCentroTrabalho"))
                    )
                    ActionChains(self.driver)\
                        .send_keys_to_element(digita_tempo_MOD_setup_dobra, "000000")\
                        .perform()
                    
                    # Digita o tempo de MOD na operação
                    digita_tempo_MOD_operacao_dobra = WebDriverWait(self.driver, 5).until(
                        EC.visibility_of_element_located((By.NAME, "tempoMODOperacaoCentroTrabalho"))
                    )
                    tempo = 10
                    if quantidade_dobra <= 5:
                        for c in range(quantidade_dobra):
                            tempo += 10
                        tempo_formatado = str(timedelta(seconds=tempo)).zfill(8)

                        # Digita o tempo formatado
                        ActionChains(self.driver)\
                            .send_keys_to_element(digita_tempo_MOD_operacao_dobra, tempo_formatado)\
                            .perform()
                    if quantidade_dobra >= 6:
                        for c in range(quantidade_dobra):
                            tempo += 50
                        tempo_formatado = str(timedelta(seconds=tempo)).zfill(8)

                        # Digita o tempo formatado
                        ActionChains(self.driver)\
                            .send_keys_to_element(digita_tempo_MOD_operacao_dobra, tempo_formatado)\
                            .perform()
                        
            # Clica no botão para salvar
            click_botao_salvar = WebDriverWait(self.driver, 5).until(
                EC.visibility_of_element_located((By.XPATH, ".//input[@type='button' and @id='botao_salvar']"))
            )
            ActionChains(self.driver)\
                .double_click(click_botao_salvar)\
                .perform()
            
            # Log de sucesso
            ic("Production script sucessfully created.")
        else:
            pass
                
        # Acessa a página
        while not self.driver.current_url.startswith(self.url +'Produto.do?metodo=Pesquisar'):
            self.acessar_pagina()

        # Limpa o campo de pesquisa
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

    def criar_orcamento(self, peca : dict, caminho_pdf, caminho_planilha):
        # orcamento
        dados_orcamento = orcamento(caminho_pdf, caminho_planilha)
        dados_orcamento.extrair_dados()
        dados_orcamento.extrair_dados_excel()

        while not self.driver.current_url.startswith(self.url + "Proposta.do?metodo=Criar_proposta"):
            self.driver.get(self.url + "Proposta.do?metodo=Criar_proposta")

        # Preenche campo de cliente
        campo_cliente = WebDriverWait(self.driver, 10).until(
            EC.visibility_of_element_located((By.ID, "nome_cliente"))
        )
        ActionChains(self.driver)\
            .send_keys_to_element(campo_cliente, self.leitura.lista_pecas[peca]["Cliente"])\
            .perform()
        
        # Espera menu de clientes aparecer
        WebDriverWait(self.driver, 10).until(
            EC.visibility_of_element_located((By.ID, "ui-id-30"))
        )

        # Clica no cliente
        cliente = WebDriverWait(self.driver, 10).until(
            EC.visibility_of_element_located((By.CLASS_NAME, "ui-menu-item"))
        )
        ActionChains(self.driver)\
            .click(cliente)\
            .perform()
        sleep(1)

        
        # Seleciona tipo de movimentação
        seleciona_movimentação = WebDriverWait(self.driver, 10).until(
            EC.visibility_of_element_located((By.ID, "id_idTipoMovimentacao"))
        )
        select = Select(seleciona_movimentação)
        if "Cliente" in self.leitura.lista_pecas[peca]["Material"]:
            select.select_by_value("45")
        else:
            select.select_by_value("53")

        # Localiza prazo de validade
        prazo_validade = WebDriverWait(self.driver, 10).until(
            EC.visibility_of_element_located((By.ID, "id_dataValidade"))
        )

        sleep(2)

        # Apaga data antiga
        prazo_validade.clear()
        
        # Pega data atual e futura
        data = date.today()
        data_futura = data + timedelta(days=10)
        data_br = str(data_futura.strftime("%d/%m%Y"))
        ic(data_br)

        sleep(2)

        prazo_validade.send_keys(data_br)
        prazo_validade.send_keys(Keys.TAB)
        
        # Digita prazo de entrega
        prazo_entrega = WebDriverWait(self.driver, 10).until(
            EC.visibility_of_element_located((By.ID, "id_prazoEntregaDias"))
        )

        prazo_entrega.clear()
        prazo_entrega.send_keys("20")
        prazo_entrega.send_keys(Keys.ENTER)
        
        ic(dados_orcamento.dados_excel)

        ic(f"Quantidade de peças extraídas: {len(dados_orcamento.dados_excel)}")
        for i, peca in enumerate(dados_orcamento.dados_excel):
            ic(f"{i}: {peca}")

        # Cria item de proposta
        cria_item = WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((By.ID, "botao_criar_item_de_proposta"))
        )
        ActionChains(self.driver)\
            .click(cria_item)\
            .perform()

        for index, item in enumerate(dados_orcamento.dados_excel):
            try:
                sleep(2)

                # Digita código da peça
                try:
                    campo_produto = WebDriverWait(self.driver, 10).until(
                        EC.visibility_of_element_located((By.ID, "id_nomeProduto"))
                    )
                except StaleElementReferenceException as e:
                    ic(f"Erro ao encontrar o campo de código: {traceback.format_exc(e)}")
                except Exception as e:
                    ic(f"Erro ao criar o produto: {traceback.format_exc(e)}")

                campo_produto.clear()
                campo_produto.send_keys(item["Código"])
                
                # Clica no produto
                click_produto = WebDriverWait(self.driver, 10).until(
                    EC.visibility_of_element_located((By.ID, "ui-id-8"))
                )
                click_produto.click()
                
                # Digita a quantidade
                campo_quantidade = WebDriverWait(self.driver, 10).until(
                    EC.visibility_of_element_located((By.ID, "id_qtdeInformada"))
                )
                campo_quantidade.send_keys(item["Quantidade"])

                # Digita valor
                campo_valor = WebDriverWait(self.driver, 10).until(
                    EC.visibility_of_element_located((By.ID, "id_precoUnitario"))
                )
                campo_valor.send_keys(item["Valor"])
                
                # Verifica quantidade de peças para salvar ou continuar a cadastrar
                if index < len(dados_orcamento.dados_excel) - 1:
                    click_salvar = WebDriverWait(self.driver, 10).until(
                        EC.element_to_be_clickable((By.ID, "botao_salvar_criar_novo_item"))
                    )
                else:
                    click_salvar = WebDriverWait(self.driver, 10).until(
                        EC.element_to_be_clickable((By.ID, "botao_salvar"))
                    )
                click_salvar.click()
            
            except Exception:
                ic(f"Error creating product: {traceback.format_exception}")
            
        # Salva o orçamento
        click_finalizar = WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((By.ID, "botao_salvar"))
        )
        click_finalizar.click()
    
    def alterar_chapa(self):
        for index, item, in enumerate(self.leitura.lista_pecas):
            try:
                espessura = float(item["Espessura"])
                codigo_mp = self.obter_mp(espessura, item)

                # Verifica se está na pagina correta
                while not self.driver.current_url.startswith(self.url + "Produto.do?metodo=Pesquisar"):
                    self.acessar_pagina()
                sleep(1)
                
                # Limpa o campo de pesquisa
                campo_pesquisa = WebDriverWait(self.driver, 10).until(
                    EC.visibility_of_element_located((By.NAME, "descricaoPesquisa")))
                campo_pesquisa.clear()
        
                peca = item["Código"]
                # Verifica se a peça é da JLS
                if "JLS" in item["Cliente"]:
                    if "DS" in item["Código"] or "RM" in item["Código"] or "DU" in item["Código"]:
                        peca = item["Código"][:17]
                    else:
                        peca = item["Código"][:8]

                # Digita o nome do produto
                digita_produto_pesquisa = WebDriverWait(self.driver, 10).until(
                    EC.visibility_of_element_located((By.NAME, "descricaoPesquisa"))
                )
                digita_produto_pesquisa.send_keys(peca)
                
                # Clica no botão para pesquisar
                click_pesquisa_produto = WebDriverWait(self.driver, 10).until(
                        EC.visibility_of_element_located((By.ID, "botao_pesquisar"))
                    )
                click_pesquisa_produto.click()
                
                # Encontra o produto na pagina e clica no produto
                nome = peca.upper()
                xpath = f"//span[contains(text(), '{nome}')]"
                try:
                    click_produto = WebDriverWait(self.driver, 2).until(
                        EC.element_to_be_clickable((By.XPATH, xpath))
                    )                
                    click_produto.click()
                    
                    # Clica no submenu de lista de materiais
                    click_submenu_lista_materiais = WebDriverWait(self.driver, 10).until(
                        EC.visibility_of_element_located((By.ID, "produtoAtivoLiberado_itemSubMenu_acessarListaMateriais"))
                    )
                    click_submenu_lista_materiais.click()
                    sleep(1)
                except (TimeoutException, StaleElementReferenceException):
                    ic(f"Product {nome} not created.")
                    continue

                # Verifica se já existe lista de materiais:
                botao_localizado = False
                try:
                    click_adicionar_item = WebDriverWait(self.driver, 5).until(
                        EC.visibility_of_element_located((By.ID, "botao_acessaradicionaritemestrutura"))
                    )
                    botao_localizado = True
                    return botao_localizado
                except TimeoutException:
                    ic("Material list not found, creating a new one.")
                finally:
                    # Verifica se já existe lista de materiais
                    if not botao_localizado:
                        # Clica em salvar para criar a lista de materiais
                        click_salvar_lista_materiais = WebDriverWait(self.driver, 10).until(
                            EC.visibility_of_element_located((By.ID, "botao_salvar"))
                        )
                        click_salvar_lista_materiais.click()

                        # Clica no botão para adicionar um item a estrutura
                        click_adicionar_item = WebDriverWait(self.driver, 5).until(
                        EC.visibility_of_element_located((By.ID, "botao_acessaradicionaritemestrutura"))
                        )
                        click_adicionar_item.click()
                    elif botao_localizado:                        
                        # Verifica se o código da matéria prima esta correto
                        xpath = f"//span[contains(text(), '{codigo_mp}')]"
                        try:
                            WebDriverWait(self.driver, 2).until(
                                EC.visibility_of_element_located((By.XPATH, xpath))
                            )
                            continue
                        except TimeoutException:
                            try:
                                xpath = "//span[contains(text(), 'MP')]"
                                click_materia_prima = WebDriverWait(self.driver, 2).until(
                                    EC.element_to_be_clickable((By.XPATH, xpath))
                                )
                                click_materia_prima.click()
                                sleep(1)

                                click_editar = WebDriverWait(self.driver, 2).until(
                                    EC.element_to_be_clickable((By.ID, "componentesProduto_itemSubMenu_acessarEditarItemListaMateriais"))
                                )
                                click_editar.click()
                            except TimeoutException:
                                click_adicionar_item = WebDriverWait(self.driver, 5).until(
                                   EC.visibility_of_element_located((By.ID, "botao_acessaradicionaritemestrutura"))
                                )       
                                click_adicionar_item.click()

                    sleep(1)

                    # Clica no geral
                    click_geral = WebDriverWait(self.driver, 10).until(
                        EC.visibility_of_element_located((By.XPATH, '//*[@id="ui-id-3"]'))
                    )
                    click_geral.click()

                    # Digita o código da matéria prima
                    digita_materia_prima = WebDriverWait(self.driver, 10).until(
                        EC.visibility_of_element_located((By.ID, "nome_produto"))
                    )
                    digita_materia_prima.clear()
                    sleep(1)
                    digita_materia_prima.send_keys(codigo_mp)
                    sleep(1)

                    # Clica no material
                    click_material = WebDriverWait(self.driver, 10).until(
                        EC.visibility_of_element_located((By.CLASS_NAME, "ui-menu-item"))
                    )
                    click_material.click()
                    sleep(1)
                    
                    # Digita o peso
                    digita_peso = WebDriverWait(self.driver, 10).until(
                        EC.visibility_of_element_located((By.NAME, "qtdeNecessariaProdutoFilho"))
                    )
                    digita_peso.clear()
                    digita_peso.send_keys(item["Peso (str)"])
                
                    # Verifica se o material é de terceiros
                    if item["Cliente"] in self.clientes_terceirizacao:
                        # Clica em avançado
                        click_avancado = WebDriverWait(self.driver, 5).until(
                            EC.visibility_of_element_located((By.XPATH, '//*[@id="ui-id-7"]'))
                        )
                        click_avancado.click()

                        try:
                            xpath = "//input[@id='id_recebe_componente_industrializacao' and @checked='checked']"
                            WebDriverWait(self.driver, 2).until(
                                EC.visibility_of_element_located((By.XPATH, xpath))
                            )
                        except TimeoutException:
                            ic("Not checked, checking now.")
                            checkmark = WebDriverWait(self.driver, 2).until(
                                EC.element_to_be_clickable((By.ID, "id_recebe_componente_industrializacao"))
                            )
                            checkmark.click()
                    
                    # Salva a materia prima adicionada
                    click_salvar_materia_prima = WebDriverWait(self.driver, 10).until(
                        EC.visibility_of_element_located((By.ID, "botao_salvar"))
                    )
                    click_salvar_materia_prima.click()
                    continue
            except Exception as e:
                ic(f"Error editing material: {traceback.format_exception(e)}")
                ic(f"Item: {item["Código"]}")
                continue
    
    def habilitar_industrializacao(self):
        for index, item in enumerate(self.leitura.lista_pecas):
            try:
                xpath = f"//span[contains(text(), '{item["Código"]})]"
                espessura = float(item["Espessura"])
                codigo_mp = self.obter_mp(espessura, item)

                while not self.driver.current_url.startswith(self.url + "Ordem.do?metodo=pesquisarPaginado"):
                    self.driver.get(self.url + "Ordem.do?metodo=pesquisarPaginado")

                peca = item["Código"]
                # Verifica se a peça é da JLS
                if "JLS" in item["Cliente"] or "PROFENG" in item["Código"]:
                    if "DS" in item["Código"] or "RM" in item["Código"] or "DU" in item["Código"] or "GB" in item["Código"]:
                        peca = item["Código"][:17]
                    else:
                        peca = item["Código"][:8]

                # Digita o nome do produto
                digita_produto = WebDriverWait(self.driver, 5).until(
                    EC.element_to_be_clickable((By.ID, "descricaoProdutoPesquisa"))
                )
                digita_produto.clear()
                digita_produto.send_keys(peca)
                digita_produto.send_keys(Keys.ENTER)
                sleep(1)

                # Clica no produto
                clica_produto = WebDriverWait(self.driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, xpath))
                )
                clica_produto.click()

                # Clica em ver empenhos
                clica_empenhos = WebDriverWait(self.driver, 5).until(
                    EC.element_to_be_clickable((By.ID, "ordemRequisitadaParcialmenteOuOLiberada_itemSubMenu_verEmpenhos"))
                )
                clica_empenhos.click()
                sleep(1)

                # Verifica se existe empenho
                empenho_localizado = False
                xpath = "//span[contains(text(), 'MP')]"

                try:
                    WebDriverWait(self.driver, 2).until(
                        EC.visibility_of_element_located((By.XPATH, xpath))
                    )
                    empenho_localizado = True
                except TimeoutException:
                    ic(f"Empenho not found for {item["Código"]}.")
                
                if empenho_localizado:
                    # Clica no empenho
                    clica_empenho = WebDriverWait(self.driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, xpath))
                    )
                    clica_empenho.click()

                    # Clica em editar
                    clica_editar = WebDriverWait(self.driver, 5).until(
                        EC.element_to_be_clickable((By.ID, "empenhosSubmenu_itemSubMenu_acessarEditarEmpenho"))
                    )
                    clica_editar.click()
                    sleep(1)

                    # Clica no geral
                    click_geral = WebDriverWait(self.driver, 10).until(
                        EC.visibility_of_element_located((By.XPATH, '//*[@id="ui-id-3"]'))
                    )
                    click_geral.click()

                    # Digita o código da matéria prima
                    digita_materia_prima = WebDriverWait(self.driver, 10).until(
                        EC.visibility_of_element_located((By.ID, "nome_produto"))
                    )
                    digita_materia_prima.clear()
                    sleep(1)
                    digita_materia_prima.send_keys(codigo_mp)
                    sleep(1)

                    # Clica no material
                    click_material = WebDriverWait(self.driver, 10).until(
                        EC.visibility_of_element_located((By.CLASS_NAME, "ui-menu-item"))
                    )
                    click_material.click()
                    sleep(1)
                    
                    # Digita o peso
                    digita_peso = WebDriverWait(self.driver, 10).until(
                        EC.visibility_of_element_located((By.NAME, "qtdeNecessariaProdutoFilho"))
                    )
                    digita_peso.clear()
                    digita_peso.send_keys(item["Peso (str)"])

                    # Clica em terceirização
                    clica_terceirizacao = WebDriverWait(self.driver, 5).until(
                        EC.element_to_be_clickable((By.ID, "ui-id-7"))
                    )
                    clica_terceirizacao.click()

                    # Verifica se o checkbox de industrialização está marcado. Caso não esteja, marca ele.
                    try:
                            xpath = "//input[@id='id_recebe_componente_industrializacao' and @checked='checked']"
                            WebDriverWait(self.driver, 2).until(
                                EC.visibility_of_element_located((By.XPATH, xpath))
                            )
                    except TimeoutException:
                        ic("Not checked, checking now.")
                        checkmark = WebDriverWait(self.driver, 2).until(
                            EC.element_to_be_clickable((By.ID, "id_recebe_componente_industrializacao"))
                        )
                        checkmark.click()
                    
                    # Salva o empenho
                    clica_salvar_empenho = WebDriverWait(self.driver, 10).until(
                        EC.element_to_be_clickable((By.ID, "botao_salvar"))
                    )
                    clica_salvar_empenho.click()
                    continue
            except Exception as e:
                ic(f"Error enabling industrialization: {traceback.format_exception(e)}")
                continue

def main():
    dotenv.load_dotenv()
    loginOS = os.getenv("LOGIN")
    senhaOS = os.getenv("PASSWORD")
    tkinter = tkinter_class.tkinter_class
    caminho_pdf = tkinter.escolher_pdf()
    automacao = Nomus(loginOS, senhaOS, caminho_pdf)
    linhas = '-=' * 20

    automacao.login_nomus()
    automacao.acessar_pagina()

    while True:
        print("""
        [1] Criação completa (Verifizcação, produto, lista de materiais, roteiro de produção).
        [2] Verificar e criar produto.
        [3] Criar lista de materiais.
        [4] Criar roteiro de produção.
        [5] Criar orçamento.
        [6] Alterar materia prima.
        [7] Habilitar terceirização.
        [8] Sair.
        """)
        resposta = input("Digite sua resposta: ")

        while True:
            
            try:
                resposta = int(resposta)
                if 1 <= resposta <= 8:
                    break
                else:
                    ic("Invalid answer. Try again!")
                    break
            except ValueError:
                ic("Invalid answer. Try again!")
                break
        
        if resposta == 1:

            # Verificar se as peças já estão criadas
            ic(linhas)
            for pecas in range(automacao.pecas):
                automacao.verificar_pecas(pecas)
            ic(f'List lenght: {len(automacao.itens_verificados)}')

            sleep(2)

            # Criar produto
            ic(linhas)

            tamanho_lista = len(automacao.itens_verificados)
            ic(f"List lenght: {tamanho_lista}")
            if tamanho_lista > 0:
                for pecas in range(automacao.pecas):
                        automacao.criar_produto()
                        automacao.preencher_campos(pecas)

            sleep(2)

            # Criar lista de materiais
            automacao.criar_lista_materiais()
            
            ic("Material list created successfully!")
            ic(linhas)

            # Criar roteiro de produção
            ic(linhas)
            for pecas in range(automacao.pecas):
                automacao.criar_roteiro(pecas)

            ic("Production scripts created successfully!")
            ic(linhas)

        elif resposta == 2:
            # Verificar se as peças já estão criadas
            ic(linhas)
            for pecas in range(automacao.pecas):
                automacao.verificar_pecas(pecas)
            ic(f'Verificated items: \n {automacao.itens_verificados}')
            ic(f'List lenght: {len(automacao.itens_verificados)}')

            sleep(2)

            # Criar produto
            ic(linhas)

            tamanho_lista = len(automacao.itens_verificados)
            ic(f"List lenght: {tamanho_lista}")
            if tamanho_lista > 0:
                for pecas in range(automacao.pecas):
                        automacao.criar_produto()
                        automacao.preencher_campos(pecas)

            sleep(2)
            ic("Products created successfully!")

        elif resposta == 3:
            # Criar lista de materiais
            automacao.criar_lista_materiais()
            
            ic("Material list created successfully!")
            ic(linhas)

        elif resposta == 4:
             # Criar roteiro de produção
            ic(linhas)
            for pecas in range(automacao.pecas):
                automacao.criar_roteiro(pecas)

            ic("Production scripts created successfully!")
            ic(linhas)

        elif resposta == 5:
            # Criar orçamento
            ic("Escolha a planilha desejada.")
            caminho_planilha = tkinter.escolher_planilha()
            automacao.criar_orcamento(0, caminho_pdf, caminho_planilha)

        elif resposta == 6:
            # Alterar materia prima
            automacao.alterar_chapa()
        
        elif resposta == 7:
            # Habilitar industrialização
            automacao.habilitar_industrializacao()
            ic("Industrialization enabled successfully!")

        elif resposta == 8:
            ic("Exiting the program...")
            break

if __name__ == "__main__":
    main()