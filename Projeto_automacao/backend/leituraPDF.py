import pdfplumber
import re

class leituraPDF:

    def __init__(self, caminho_arquivo):
        self.caminho_arquivo = caminho_arquivo
        self.texto = self.extrair_texto()

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
        if pagina < len(self.texto):

            if self.encontrar_cliente().strip() in "JLS":
                padrao_jls = r"Código:\s*(\d{3}-\d{4})\s*-\s*\d+[.,]?\d*\s*MM"
                    
                match = re.search(padrao_jls, self.texto[pagina], re.IGNORECASE)
                if match:
                    nome_peca = match.group(1).strip()
                    return nome_peca

            if self.encontrar_cliente().strip() in "ALIMENTEC":
                padrao_alimentec = r"Código:\s*(NSA-\d+)"

                match = re.search(padrao_alimentec, self.texto[pagina])

                if match:
                    nome_peca = match.group(1).strip()
                    return nome_peca

            padroes = [
                r"Código:\s*(.*?)\s*(?=FIM)",
                r"Código:\s*(\d+)\s*(?=FIM)",
                r"Código:\s*(\d{3}\.\d{1,2}\s-\s.*)",
                r"Código:\s*(\d{3}\s-\s.*)",
                r"Código:\s*(\d{3}\.\d+\s-\s[\w\s-]+-\s[\d\w-]+)",
                r"Código:\s*(\d{3}\s-\s[\w\s-]+-\s[\d\w-]+)"
            ]

            for padrao in padroes:

                match = re.search(padrao, self.texto[pagina], re.IGNORECASE)
                if match:
                    nome_peca = match.group(1).strip()
                    return nome_peca
            
            return "Peça não encontrada"
    
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
    
    def encontrar_largura(self, pagina=0):
        padrao = r"Largura:\s*(.*?)\s*(?=FIM)"
        match = re.search(padrao, self.texto[pagina])

        if match:
            largura_peca = match.group(1).strip()
        else:
            largura_peca = "Largura não encontrada."
        
        return largura_peca, largura_peca
    
    def encontrar_comprimento(self, pagina=0):
        padrao = r"Comprimento:\s*(.*?)\s*(?=FIM)"
        match = re.search(padrao, self.texto[pagina])

        if match:
            comprimento_peca = match.group(1).strip()
        else:
            comprimento_peca = "Comprimento não encontrado."
        
        return comprimento_peca, comprimento_peca

    def encontrar_peso(self, pagina=0):
        padrao = r"Peso:\s*(.*?)\s*(?=FIM)"
        match = re.search(padrao, self.texto[pagina])
        
        if match:
            peso_peca = match.group(1).strip()
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
    
    def encontrar_tempo(self, pagina=0):
        padrao = r"Tempo:\s*(.*?)\s*(?=FIM)"
        match = re.search(padrao, self.texto[pagina])

        if match:
            tempo_de_corte = match.group(1).strip()
        else:
            tempo_de_corte = "Tempo não encontrado."
        
        return tempo_de_corte

    def exibe_informacoes(self):
        print(f"""
    Código: {leituraPDF.encontrar_codigo(self)}
    Espessura: {leituraPDF.encontrar_espessura(self)}
    Material: {leituraPDF.encontrar_material(self)}
    Peso: {leituraPDF.encontrar_peso(self)}
    Dobras: {leituraPDF.encontrar_dobra(self)}
    Cliente: {leituraPDF.encontrar_cliente(self)}
    Quantidade: {leituraPDF.encontrar_quantidade(self)}
    """)