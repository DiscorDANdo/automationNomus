from openpyxl import load_workbook
import pandas as pd
import tabula
import tkinter_class


class orcamento:
    
    def __init__(self, caminho_pdf, caminho_planilha):
        self.arquivo_pdf = caminho_pdf
        self.arquivo_planilha = caminho_planilha
        self.dados = []

    def extrair_dados(self):
        dfs = tabula.read_pdf(self.arquivo_pdf, pages='all', multiple_tables=True, lattice=True)
        
        dfs_completed = pd.concat(dfs, ignore_index=True)

        #dfs_completed.to_excel(self.arquivo_planilha, index=False)

        df = pd.DataFrame(dfs_completed)

        book = load_workbook(self.arquivo_planilha)

        with pd.ExcelWriter(self.arquivo_planilha, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            writer.book = book

            df.to_excel(writer, sheet_name='Plan1', startrow=1, startcol=0, index=False, header=False)
        
    
    def atualizar_planilha(self):
        pass

caminho_pdf = tkinter_class.tkinter_class.escolher_pdf()
caminho_planilha = tkinter_class.tkinter_class.escolher_planilha()
orcamento = orcamento(caminho_pdf, caminho_planilha)
print(orcamento.extrair_dados())