from cx_Freeze import setup, Executable

build_exe_options = {
    "packages": ["selenium", "tkinter", "dotenv"],
    "include_files": ["chromedriver.exe", ".env"],
}

setup(
    name="AutomacaoNomus",
    version="1.0",
    description="Automação Nomus com terminal",
    options={"build_exe": build_exe_options},
    executables=[
        Executable("automacao_nomus.py", base=None)  # <- terminal ativado aqui
    ]
)
