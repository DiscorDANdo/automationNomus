from icecream import ic
from tkinter_class import tkinter_class
import pdfplumber
import pandas as pd
import tabula

# TODO: Alterar a função de extrair o cliente.

class leitura:

    def __init__(self, arquivo_pdf):
        self.arquivo_pdf = arquivo_pdf
        self.peca = {
            "Código" : "",
            "Espessura" : "",
            "Material" : "",
            "Largura" : "",
            "Comprimento" : "",
            "Peso" : "",
            "Peso (str)" : "",
            "Tempo" : "",
            "Quantidade" : "",
            "Dobras" : "",
            "Cliente" : ""
        }
        self.lista_pecas = []
        self.dados = self.extrair_dados()

    def extrair_dados(self):
        try:
            df = tabula.read_pdf(self.arquivo_pdf, pages='all', multiple_tables=True, lattice=True)

            dfs_completed = pd.concat(df, ignore_index=True)

            return dfs_completed
        except Exception as e:
            ic(f"Error reading PDF: {e}")

    def extrair_pecas(self):
        try:
            self.peca["Cliente"] = self.extrair_cliente()

            for i, row in self.dados.iterrows():
                self.peca["Código"] = self.corrigir_quebra_texto(row["Peça"])
                self.peca["Espessura"] = self.converter_para_float(row["Espessura"])
                self.peca["Material"] = self.corrigir_quebra_texto(row["Material"])
                self.peca["Largura"] = row["Largura"]
                self.peca["Comprimento"] = row["Comprimento"]
                self.peca["Peso"] = self.calcular_peso(row["Espessura"], row["Largura"], row["Comprimento"])
                self.peca["Peso (str)"] = self.converter_peso_para_str(self.peca["Peso"])
                self.peca["Tempo"] = row["Tempo"]
                self.peca["Quantidade"] = int(row["QNTD"])
                self.peca["Dobras"] = 0 if pd.isna(row["DOBRA"]) else row["DOBRA"]
                self.peca["Cliente"] = self.extrair_cliente()

                self.lista_pecas.append(self.peca.copy())
                ic(f"Product successfully extracted: {self.peca}")
        except Exception as e:
            ic(f"Error extracting product: {e}")
    
    def extrair_tamanho(self):
        return len(self.lista_pecas)
    
    def extrair_cliente(self):
        with pdfplumber.open(self.arquivo_pdf) as pdf:
            primeira_pagina = pdf.pages[0]

            # Coordenadas (x0, y0, x1, y1) em pontos PDF
            # Ajuste os valores abaixo com base na sua área
            # Exemplo: região no topo da página
            bbox = (0, 50, 500, 70)  # esquerda, topo, direita, inferior

            area = primeira_pagina.within_bbox(bbox)
            texto_extraido = area.extract_text()

            return texto_extraido
        
    def calcular_peso(self, espessura, largura, comprimento):
        try:
            espessura_formatada = float(str(espessura).replace(",", "."))
            largura_formatada = float(str(largura).replace(",", "."))
            comprimento_formatado = float(str(comprimento).replace(",", "."))

            peso = (espessura_formatada * largura_formatada * comprimento_formatado * 7850) / 1000000000 
            peso_formatado = (f"{peso:.2f}")
            
            return peso_formatado
        except Exception as e:
            ic(f"Error calculating weight: {e}")

    def corrigir_quebra_texto(self, texto):
        try:
            if isinstance(texto, float):
                texto_bruto = str(int(texto))
            else:
                texto_bruto = str(texto)
                
            texto_formatado = texto_bruto.replace("\r", " ")

            return str(texto_formatado)
        except Exception as e:
            ic(f"Error formatting text: {e}")
    
    def converter_para_float(self, valor):
        try:
            valor = float(str(valor).replace(",", "."))

            return valor
        except Exception as e:
            ic(f"Error converting to float: {e}")
            return 0.0
    
    def converter_peso_para_str(self, valor):
        valor = str(valor).replace(".", ",")

        return valor
    
    def main():
        ic("Starting data extraction...")
        caminho_pdf = tkinter_class.escolher_pdf()
        leitor = leitura(caminho_pdf)
        leitor.extrair_pecas()
        leitor.extrair_tamanho()
        ic("Data extraction completed successfully.")

        ic(leitor.lista_pecas[0]["Espessura"])
        
if __name__ == "__main__":
    leitura.main()
