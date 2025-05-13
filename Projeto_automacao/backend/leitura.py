from icecream import ic
from tkinter_class import tkinter_class
import pandas as pd
import tabula

class leitura:

    def __init__(self, arquivo_pdf):
        self.arquivo_pdf = arquivo_pdf
        self.peca = {
            "Código" : "",
            "Espessura" : "",
            "Material" : "",
            "Largura" : "",
            "Comprimento" : "",
            "Tempo" : "",
            "QNTD" : "",
            "DOBRA" : ""
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
            for i, row in self.dados.iterrows():
                self.peca["Código"] = self.corrige_quebra_texto(row["Peça"])
                self.peca["Espessura"] = row["Espessura"]
                self.peca["Material"] = self.corrige_quebra_texto(row["Material"])
                self.peca["Largura"] = row["Largura"]
                self.peca["Comprimento"] = row["Comprimento"]
                self.peca["Tempo"] = row["Tempo"]
                self.peca["QNTD"] = row["QNTD"]
                self.peca["DOBRA"] = row["DOBRA"]

                self.lista_pecas.append(self.peca.copy())
                ic(f"Product successfully extracted: {self.peca}")
        except Exception as e:
            ic(f"Error extracting product: {e}")
    
    def corrige_quebra_texto(self, texto):
        try:
            texto_bruto = texto
            texto_formatado = texto_bruto.replace("\r", " ")

            return texto_formatado
        except Exception as e:
            ic(f"Error formatting text: {e}")

    def main():
        ic("Starting data extraction...")
        caminho_pdf = tkinter_class.escolher_pdf()
        leitor = leitura(caminho_pdf)
        leitor.extrair_pecas()
        ic("Data extraction completed successfully.")

if __name__ == "__main__":
    leitura.main()
