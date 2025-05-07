from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver import Keys, ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from icecream import ic
from time import sleep
from datetime import timedelta 
import leituraPDF
import tkinter_class
import dotenv
import os

#TODO: Adicionar as funções do orcamento.py no main()

class Nomus:
    
    def __init__(self, username, password, caminho_pdf):
        
        self.url = 'https://metalservtech.nomus.com.br/metalservtech/'
        self.username = username
        self.password = password
        self.chromedriverPath = r'C:\Users\Thalles\AppData\Local\Programs\chromedriver-win64'
        self.service = Service(executable_path=self.chromedriverPath)
        self.driver = webdriver.Chrome()
        self.leitura = leituraPDF.leituraPDF(caminho_pdf)
        self.paginas = leituraPDF.leituraPDF.extrair_paginas(self.leitura)
        self.itens_verificados = []

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


    def verificar_pecas(self, pagina):

        while not self.driver.current_url.startswith(self.url + 'Produto.do?metodo=Pesquisar'):
            self.acessar_pagina()

        # Digitando nome do produto
        try:
            text_input = WebDriverWait(self.driver, 5).until(
                EC.visibility_of_element_located((By.NAME, "descricaoPesquisa"))
            )
            ActionChains(self.driver)\
                .send_keys_to_element(text_input, self.leitura.encontrar_codigo(pagina))\
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
        except Exception as e:
            ic(f"Error searching the product: {e}")
        
        # Verificando se a peça já exite no sistema
        try:
            elemento = WebDriverWait(self.driver, 2).until(
                EC.visibility_of_element_located((By.CLASS_NAME, "tela-vazia"))
            )

            if elemento.is_displayed():
                self.itens_verificados.append(self.leitura.encontrar_codigo(pagina))
            else:
                pass
        except Exception:  
            ic("Product already exists on system.")

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
        

    def criar_produto(self):
        
        self.driver.get(self.url + 'Produto.do?metodo=Criar_produto')

    def preencher_campos(self, pagina):
        
        while not self.driver.current_url.startswith(self.url + 'Produto.do?metodo=Criar_produto'):
            self.criar_produto()

        if self.leitura.encontrar_codigo(pagina) in self.itens_verificados:
            # Descrição do produto
            descricao_produto_criacaoPA = WebDriverWait(self.driver, 20).until(
                EC.visibility_of_element_located((By.NAME, "descricao"))
            )
            ActionChains(self.driver)\
                .send_keys_to_element(descricao_produto_criacaoPA, self.leitura.encontrar_codigo(pagina).upper())\
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
            if self.leitura.encontrar_material(pagina) in "Carbono Cliente" or "Inox Cliente":
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

            if self.leitura.encontrar_cliente(pagina) in "ADR":
                ActionChains(self.driver)\
                    .send_keys_to_element(ncm_criacaoPA, "87169090")\
                    .perform()
            else:
                if self.leitura.encontrar_espessura(pagina) <= 3:
                    ActionChains(self.driver)\
                .send_keys_to_element(ncm_criacaoPA, "72085400")\
                        .perform()
                if 3.1 <= self.leitura.encontrar_espessura(pagina) <= 4.75:
                    ActionChains(self.driver)\
                        .send_keys_to_element(ncm_criacaoPA, "72085300")\
                        .perform()
                if 4.76 < self.leitura.encontrar_espessura(pagina) <= 10:
                    ActionChains(self.driver)\
                        .send_keys_to_element(ncm_criacaoPA, "72085200")\
                        .perform()
                if 10.1 < self.leitura.encontrar_espessura(pagina) <= 25:
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
            if self.leitura.encontrar_material() in "Carbono Cliente" or "Inox Cliente":
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
            if self.leitura.encontrar_material() in "Carbono Cliente" or "Inox Cliente":
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

    def criar_lista_materiais(self, pagina):
        espessura = int(self.leitura.encontrar_espessura(pagina))

        # Verifica se está na pagina correta
        while not self.driver.current_url.startswith(self.url + "Produto.do?metodo=Pesquisar"):
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
        
        sleep(1)

        # Digita o nome do produto
        digita_produto_pesquisa = WebDriverWait(self.driver, 10).until(
            EC.visibility_of_element_located((By.NAME, "descricaoPesquisa"))
        )
        ActionChains(self.driver)\
            .send_keys_to_element(digita_produto_pesquisa, self.leitura.encontrar_codigo(pagina))\
            .perform()
        
        # Clica no botão para pesquisar
        click_pesquisa_produto = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.ID, "botao_pesquisar"))
            )
        ActionChains(self.driver)\
            .click(click_pesquisa_produto)\
            .perform()
        
        # Encontra o produto na pagina e clica no produto
        nome = self.leitura.encontrar_codigo(pagina).upper()
        xpath = f"//span[contains(text(), '{nome}')]"

        click_produto = self.driver.find_element(By.XPATH, xpath)
        ActionChains(self.driver)\
            .click(click_produto)\
            .perform()
        
        # Clica no submenu de lista de materiais
        click_submenu_lista_materiais = WebDriverWait(self.driver, 10).until(
            EC.visibility_of_element_located((By.ID, "produtoAtivoLiberado_itemSubMenu_acessarListaMateriais"))
        )
        ActionChains(self.driver)\
            .click(click_submenu_lista_materiais)\
            .perform()

        # Verifica se já existe lista de materiais:
        botao_localizado = False
        try:
            WebDriverWait(self.driver, 5).until(
                EC.visibility_of_element_located((By.ID, "botao_acessaradicionaritemestrutura"))
            )
            botao_localizado = True
            return botao_localizado
        except Exception:
            pass

        if not botao_localizado:
            # Clica em salvar para criar a lista de materiais
            click_salvar_lista_materiais = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.ID, "botao_salvar"))
            )
            ActionChains(self.driver)\
                .click(click_salvar_lista_materiais)\
                .perform()
            
            # Clica no botão para adicionar um item a estrutura
            click_adicionar_item = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.ID, "botao_acessaradicionaritemestrutura"))
            )
            ActionChains(self.driver)\
                .click(click_adicionar_item)\
                .perform()

            # Clica no geral
            click_geral = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, '//*[@id="ui-id-3"]'))
            )
            ActionChains(self.driver)\
                .click(click_geral)\
                .perform()
            
            # Digita o código da matéria prima
            digita_materia_prima = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.ID, "nome_produto"))
            )
            
            # Verifica se o material é da ADR
            if self.leitura.encontrar_cliente() in "ADR":
                if 4 <= espessura <= 5:
                    ActionChains(self.driver)\
                        .send_keys_to_element(digita_materia_prima, "MP 00026")\
                        .perform()
                elif 6 <= espessura <= 6.3:
                    ActionChains(self.driver)\
                        .send_keys_to_element(digita_materia_prima, "MP 00016")\
                        .perform()
                elif 7.9 <= espessura <= 8:
                    ActionChains(self.driver)\
                        .send_keys_to_element(digita_materia_prima, "MP 00014")\
                        .perform()
                elif 9 <= espessura <= 9.5:
                    ActionChains(self.driver)\
                        .send_keys_to_element(digita_materia_prima, "MP 00015")\
                        .perform()
                elif 12 <= espessura <= 12.7:
                    ActionChains(self.driver)\
                        .send_keys_to_element(digita_materia_prima, "MP 00009")\
                        .perform()
                elif 15 <= espessura <= 16:
                    ActionChains(self.driver)\
                        .send_keys_to_element(digita_materia_prima, "MP 00046")\
                        .perform()
                elif 19 <= espessura < 19.5:
                    ActionChains(self.driver)\
                        .send_keys_to_element(digita_materia_prima, "MP 00035")\
                        .perform()
                elif espessura <= 25:
                    ActionChains(self.driver)\
                        .send_keys_to_element(digita_materia_prima, "MP 00025")\
                        .perform()
                    
                    # Clica no material
                click_material = WebDriverWait(self.driver, 10).until(
                    EC.visibility_of_element_located((By.CLASS_NAME, "ui-menu-item"))
                )
                ActionChains(self.driver)\
                    .click(click_material)\
                    .perform()
                
                # Digita o peso
                digita_peso = WebDriverWait(self.driver, 10).until(
                    EC.visibility_of_element_located((By.NAME, "qtdeNecessariaProdutoFilho"))
                )
                ActionChains(self.driver)\
                    .send_keys_to_element(digita_peso, self.leitura.encontrar_peso())\
                    .perform()
                
                # Clica em avançado
                click_avancado = WebDriverWait(self.driver, 5).until(
                    EC.visibility_of_element_located((By.XPATH, '//*[@id="ui-id-7"]'))
                )
                ActionChains(self.driver)\
                    .click(click_avancado)\
                    .perform()
                
                # Clica em recebe componentes de terceiros para industrialização sob encomenda 
                click_industrializacao = WebDriverWait(self.driver, 5).until(
                    EC.visibility_of_element_located((By.ID, "id_recebe_componente_industrializacao"))
                )
                ActionChains(self.driver)\
                    .click(click_industrializacao)\
                    .perform()

            else:    
                if espessura <= 1.5:
                    ActionChains(self.driver)\
                        .send_keys_to_element(digita_materia_prima, "MP 00002")\
                        .perform()
                elif 1.9 <= espessura <= 2:
                    ActionChains(self.driver)\
                        .send_keys_to_element(digita_materia_prima, "MP 00001")\
                        .perform()
                elif 3 <= espessura <= 3.1:
                    ActionChains(self.driver)\
                        .send_keys_to_element(digita_materia_prima, "MP 00019")\
                        .perform()
                elif 4 <= espessura <= 5:
                    ActionChains(self.driver)\
                        .send_keys_to_element(digita_materia_prima, "MP 00003")\
                        .perform()
                elif 6 <= espessura <= 6.3:
                    ActionChains(self.driver)\
                        .send_keys_to_element(digita_materia_prima, "MP 00006")\
                        .perform()
                elif 7 <= espessura <= 7.5:
                    ActionChains(self.driver)\
                        .send_keys_to_element(digita_materia_prima, "Colocar espessura no banco")\
                        .perform()
                elif 7.9 <= espessura <= 8:
                    ActionChains(self.driver)\
                        .send_keys_to_element(digita_materia_prima, "MP 00018")\
                        .perform()
                elif 9 <= espessura <= 9.5:
                    ActionChains(self.driver)\
                        .send_keys_to_element(digita_materia_prima, "MP 00032")\
                        .perform()
                elif 12 <= espessura <= 12.7:
                    ActionChains(self.driver)\
                        .send_keys_to_element(digita_materia_prima, "MP 00011")\
                        .perform()
                elif 15.88 <= espessura <= 16:
                    ActionChains(self.driver)\
                        .send_keys_to_element(digita_materia_prima, "MP 00036")\
                        .perform()
                elif 19 <= espessura < 19.5:
                    ActionChains(self.driver)\
                        .send_keys_to_element(digita_materia_prima, "Colocar espessura no banco")\
                        .perform()
                elif espessura <= 25:
                    ActionChains(self.driver)\
                        .send_keys_to_element(digita_materia_prima, "Colocar espessura no banco")\
                        .perform()

                # Clica no material
                click_material = WebDriverWait(self.driver, 10).until(
                    EC.visibility_of_element_located((By.CLASS_NAME, "ui-menu-item"))
                )
                ActionChains(self.driver)\
                    .click(click_material)\
                    .perform()
                
                # Digita o peso
                digita_peso = WebDriverWait(self.driver, 10).until(
                    EC.visibility_of_element_located((By.NAME, "qtdeNecessariaProdutoFilho"))
                )
                ActionChains(self.driver)\
                    .send_keys_to_element(digita_peso, self.leitura.encontrar_peso())\
                    .perform()
            
            # Salva a materia prima adicionada
            click_salvar_materia_prima = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.ID, "botao_salvar"))
            )
            ActionChains(self.driver)\
                .click(click_salvar_materia_prima)\
                .perform()
        else:
            pass

    def criar_roteiro(self, pagina):

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
            .send_keys_to_element(digita_produto_pesquisa, self.leitura.encontrar_codigo(pagina))\
            .perform()
        
        # Clica no botão para pesquisar
        click_pesquisa_produto = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.ID, "botao_pesquisar"))
            )
        ActionChains(self.driver)\
            .click(click_pesquisa_produto)\
            .perform()

        # Encontra o produto na pagina e clica no produto
        nome = self.leitura.encontrar_codigo(pagina).upper()
        xpath = f"//span[contains(text(), '{nome}')]"

        click_produto = self.driver.find_element(By.XPATH, xpath)
        ActionChains(self.driver)\
            .click(click_produto)\
            .perform()
        
        # Seleciona o roteiro de produção
        click_submenu_roteiro = self.driver.find_element(By.ID, "produtoAtivoLiberado_itemSubMenu_acessarRoteiroProduto")
        ActionChains(self.driver)\
            .click(click_submenu_roteiro)\
            .perform()
        
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
                .send_keys_to_element(digita_tempo_operacao, self.leitura.encontrar_tempo(pagina))\
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
                .send_keys_to_element(digita_tempo_MOD_operacao, self.leitura.encontrar_tempo(pagina))\
                .perform()
            
            # Verifica se existe dobra e converte a quantidade de string para int
            if self.leitura.encontrar_dobra(pagina) in " " and "FIM" and "Dobras não encontrada.":
                pass
            else:
                quantidade_dobra = int(self.leitura.encontrar_dobra(pagina))
 
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
                    
                # Clica no botão para salvar1
                click_botao_salvar = WebDriverWait(self.driver, 5).until(
                    EC.visibility_of_element_located((By.XPATH, ".//input[@type='button' and @id='botao_salvar']"))
                )
                ActionChains(self.driver)\
                    .double_click(click_botao_salvar)\
                    .perform()
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
        
        # Log de sucesso
        ic("Production script sucessfully created.")

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
        [6] Sair.
        """)
        resposta = input("Digite sua resposta: ")

        while True:
            
            try:
                resposta = int(resposta)
                if 1 <= resposta <= 6:
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
            for pagina in range(automacao.paginas):
                automacao.verificar_pecas(pagina)
            ic(f'Verificated items: \n {automacao.itens_verificados}')
            ic(f'List lenght: {len(automacao.itens_verificados)}')

            sleep(2)

            # Criar produto
            ic(linhas)

            tamanho_lista = len(automacao.itens_verificados)
            ic(f"List lenght: {tamanho_lista}")
            if tamanho_lista > 0:
                for pagina in range(automacao.paginas):
                        automacao.criar_produto()
                        automacao.preencher_campos(pagina)

            sleep(2)

            # Criar lista de materiais
            for pagina in range(automacao.paginas):
                automacao.criar_lista_materiais(pagina)
            
            ic("Material list created successfully!")
            ic(linhas)

            # Criar roteiro de produção
            ic(linhas)
            for pagina in range(automacao.paginas):
                automacao.criar_roteiro(pagina)

            ic("Production scripts created successfully!")
            ic(linhas)

        elif resposta == 2:
            # Verificar se as peças já estão criadas
            ic(linhas)
            for pagina in range(automacao.paginas):
                automacao.verificar_pecas(pagina)
            ic(f'Verificated items: \n {automacao.itens_verificados}')
            ic(f'List lenght: {len(automacao.itens_verificados)}')

            sleep(2)

            # Criar produto
            ic(linhas)

            tamanho_lista = len(automacao.itens_verificados)
            ic(f"List lenght: {tamanho_lista}")
            if tamanho_lista > 0:
                for pagina in range(automacao.paginas):
                        automacao.criar_produto()
                        automacao.preencher_campos(pagina)

            sleep(2)
            ic("Products created successfully!")

        elif resposta == 3:
            # Criar lista de materiais
            for pagina in range(automacao.paginas):
                automacao.criar_lista_materiais(pagina)
            
            ic("Material list created successfully!")
            ic(linhas)

        elif resposta == 4:
             # Criar roteiro de produção
            ic(linhas)
            for pagina in range(automacao.paginas):
                automacao.criar_roteiro(pagina)

            ic("Production scripts created successfully!")
            ic(linhas)

        elif resposta == 6:
            ic("Exiting the program...")
            break

if __name__ == "__main__":
    main()