import sys
from cx_Freeze import setup, Executable

base = None

if sys.platform == 'win32':
    base = "Win32GUI"

# Lista de includes necessários
includes = [
    'customtkinter',  
    'tkinter', 
    'openpyxl', 
    'pandas', 
    'numpy', 
    'keyboard', 
    'pyautogui', 
    'pyperclip', 
    'pyscreeze',  
    'pytweening', 
    'time', 
    'os', 
    'datetime'
]

# Configuração do cx_Freeze
setup(
    name="Verificador de Contravale",
    version='1.0',
    description="Programa que localiza um número de contravale na planilha do sistema e cola o contravale em uma frente de caixa.",
    options={
        'build_exe': {
            'includes': includes,
            'include_msvcr': True,
            'packages': ['pandas', 'numpy', 'customtkinter']  
        }
    },
    executables=[Executable('Verificador de Contravale.py', base=base)]  
)