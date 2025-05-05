from tkinter.filedialog import askopenfilename

class tkinter_class():

    def __init__(self):
        pass

    def escolher_pdf():
        caminho_pdf = askopenfilename(
            title="Selecione o PDF",
            filetypes=[("PDF files", "*.pdf")]
        )

        return caminho_pdf
    
    def escolher_planilha():
        caminho_planilha = askopenfilename(
            title="Selecione a planilha.",
            filetypes=[("Planilhas excel", "*.xlsx")]
        )

        return caminho_planilha