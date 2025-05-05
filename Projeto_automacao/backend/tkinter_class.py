from tkinter.filedialog import askopenfilename

class tkinter_class():

    def __init__(self):
        pass

    def escolher_pdf():
        caminho_pdf = askopenfilename(
            title="Selecione o arquivo desejado.",
            filetypes=[("PDF files", "*.pdf")]
        )

        return caminho_pdf
    
    def escolher_planilha():
        caminho_planilha = askopenfilename(
            title="Selecione o arquivo desejado.",
            filetypes=[("Planilhas excel", "*.xlsx")]
        )

        return caminho_planilha