from openpyxl import load_workbook
import pandas as pd
import pdfplumber
import os

class orcamento:
    
    def __init__(self, orcamento_pdf, orcamento_planilha):
        self.arquivo_pdf = orcamento_pdf
        self.arquivo_planilha = orcamento_planilha
        self.dados = []

    def extrair_dados(self):
        dados = []

        with pdfplumber.open(self.arquivo_pdf) as pdf:
            for pagina in pdf.pages:
                tabela = pagina.extract_table()
                
                if tabela:
                    colunas = tabela[0]
                    for linha in tabela[1:]:
                        dados.append(dict(zip(colunas, linha)))
        
        df = pd.DataFrame(dados)

        return df
    
    def preencher_planilha(self):

        wb = load_workbook(self.arquivo_planilha)
        ws = wb.active

        colunas = {"PEÇA": "A", "ESPESSURA": "B", "LARGURA": "C", "COMPRIMENTO": "D", "TEMPO": "E", "DOBRA": "F", "QTDE": "G"}

        for i, linhas in enumerate(self.extrair_dados().itertuples(index=False), start=2):
            for idx, campo in enumerate(colunas.keys()):
                letra_coluna = colunas[campo]
                ws[f"{letra_coluna}{i}"] = getattr(linhas, idx)

        wb.save(r"C:\Users\Thalles\Desktop\Planilha\Planilha_atualizada.xlsx")
        print(f"salvo em: {os.path.abspath("Planilha_atualizada.xlsx")}")
