# Precisa alterar o metódo para exibição
# Também precisa alterar o PDF para ficar mais espaçado as informações
import PyPDF2
import re

class leitura_PDF:

    def __init__(self, caminho_arquivo):
        self.caminho_arquivo = caminho_arquivo

    def extrair_texto(self):
        texto_extraido = ""
        
        with open(self.caminho_arquivo, "rb") as file:
            reader = PyPDF2.PdfReader(file)

            texto_extraido = ""
            
            for page in reader.pages:
                texto_extraido += page.extract_text() + "\n"

        return texto_extraido
    
    def encontrar_nome_peca(self):
        texto = self.extrair_texto()
        padrao = r"Código:\s*(.*)"

        match = re.search(padrao, texto)

        if match:
            nome_peca = match.group(1)
            # print(nome_peca)
        else:
            nome_peca = "Peça não encontrada."
            # print(nome_peca)

        return nome_peca

    def encontrar_espessura_peca(self):
        texto = self.extrair_texto()
        padrao = r"Espessura:\s*(.*)"

        match = re.search(padrao, texto)

        if match:
            espessura_peca = match.group(1)
            # print(espessura_peca)
        else:
            espessura_peca = "Espessura não encontrada."
            # print(espessura_peca)

        return espessura_peca

    def encontrar_material_peca(self):
        texto = self.extrair_texto()
        padrao = r"Material:\s*(.*)"

        match = re.search(padrao, texto)

        if match:
            material_peca = match.group(1)
            # print(material_peca)
        else:
            material_peca = "Material não encontrada."
            # print(material_peca)
        
        return material_peca
    
    def encontrar_peso_peca(self):
        texto = self.extrair_texto()
        padrao = r"Peso:\s*(.*)"

        match = re.search(padrao, texto)

        if match:
            peso_peca = match.group(1)
            # print(peso_peca)
        else:
            peso_peca = "Peso não encontrado."
            # print(peso_peca)

        return peso_peca
    
    def encontrar_dobra_peca(self):
        texto = self.extrair_texto()
        padrao = r"Dobras:\s*(.*)"

        match = re.search(padrao, texto)

        if match:
            dobras_peca = match.group(1)
            # print(dobras_peca)
        else:
            dobras_peca = "Dobras não encontradas."
            # print(dobras_peca)
        
        return dobras_peca
    
    def encontrar_quantidade(self):
        texto = self.extrair_texto()
        padrao = r"Quantidade:\s*(.*)"

        match = re.search(padrao, texto)

        if match:
            quantidade_peca = match.group(1)
            # print(quantidade_peca)
        else:
            quantidade_peca = "Quantidade não encontrada."
            # print(quantidade_peca)

        return quantidade_peca
        
    def encontrar_valor_unitario(self):
        texto = self.extrair_texto()
        padrao = r"R$\s*(.*)"

        match = re.search(padrao, texto)

        if match:
            valor_unitario = match.group(1)
            # print(valor_unitario)
        else:
            valor_unitario = "Valor unitário não encontrado."
            # print(valor_unitario)
        
        return valor_unitario
    
    def encontrar_valor_total(self):
        valor_total = leitura_PDF.encontrar_quantidade * leitura_PDF.encontrar_valor_unitario

        return valor_total

    def extrai_informacoes(self):
        print(leitura_PDF.encontrar_nome_peca())
        print(leitura_PDF.encontrar_espessura_peca())
        print(leitura_PDF.encontrar_material_peca())
        print(leitura_PDF.encontrar_peso_peca())
        print(leitura_PDF.encontrar_dobra_peca())
        print(leitura_PDF.encontrar_quantidade())

leitura = leitura_PDF(r'C:\Users\Thalles\Desktop\Listas\teste.pdf')
print(leitura.encontrar_nome_peca())
print(leitura.encontrar_espessura_peca())
print(leitura.encontrar_material_peca())
print(leitura.encontrar_peso_peca())
print(leitura.encontrar_dobra_peca())
print(leitura.encontrar_quantidade())
