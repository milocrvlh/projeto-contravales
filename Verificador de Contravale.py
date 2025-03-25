from customtkinter import * # Versão moderna do Tkinter
from tkinter import messagebox # Habilitar Pop-up's 
import keyboard # Para controle das teclas
import pandas as pd
import os
from datetime import datetime
import pyperclip  # Biblioteca para gerenciar o clipboard
import pyautogui  # Biblioteca para automação de foco e colagem
from time import sleep # Para pausas entre ações

# Caminho do arquivo Excel
ARQUIVO_CONTRAVALES = "contravales.xlsx" # Substitua pelo endereço da planilha

# Nome do aplicativo a ser focado
NOME_APP = "Teste - Bloco de Notas"  # Substitua pelo nome da janela do aplicativo

# Inicializa o arquivo Excel, se não existir
if not os.path.exists(ARQUIVO_CONTRAVALES):
    with pd.ExcelWriter(ARQUIVO_CONTRAVALES) as writer:
        pd.DataFrame(columns=["Nº Contravale"]).to_excel(writer, sheet_name="ContravalesAbertos", index=False)
        pd.DataFrame(columns=["Nº Contravale", "Data e Hora Baixa"]).to_excel(writer, sheet_name="ContravalesBaixados", index=False)
        pd.DataFrame(columns=["Nº Contravale"]).to_excel(writer, sheet_name="BaixaAutomatica", index=False)

# Função para simular a tecla ESC
def pressionar_esc():
    keyboard.press_and_release('esc')

# Função para focar no aplicativo específico
def focar_em_app(nome_app):
    try:
        # Traz a janela do aplicativo para frente
        pyautogui.getWindowsWithTitle(nome_app)[0].activate()
        sleep(0.5)  # Aguarda para garantir que a janela esteja ativa
        return True # Para usar em lógicas que indicam que o app foi encontrado
    except IndexError:
        messagebox.showerror("Erro", f"Não foi possível localizar o aplicativo '{nome_app}'.")
        return False  # Para usar em lógicas que indicam que o app não foi encontrado

# Função para processar um contravale específico
def processar_contravale():
    # Carrega os dados das planilhas
    planilhas = pd.ExcelFile(ARQUIVO_CONTRAVALES)
    contravales_abertos = pd.read_excel(planilhas, sheet_name="ContravalesAbertos")
    contravales_baixados = pd.read_excel(planilhas, sheet_name="ContravalesBaixados")
    
    # Solicita o número do contravale
    numero = CTkInputDialog(title="Contravale", text="Digite o número do contravale:").get_input()
    # Caso o usuário cancele a operação
    if numero is None:
        return
    # Caso o usuário digite algo que não é um número de contravale
    while not numero.isnumeric():
        messagebox.showerror("Erro", "Entrada inválida. Apenas números são permitidos.")
        numero = CTkInputDialog(title="Contravale", text="Digite o número do contravale:").get_input()
        
    # Converte o número para string (para validação)
    numero = str(numero)
    
    # Verifica se o contravale existe
    if numero not in contravales_abertos["Nº Contravale"].astype(str).values:
        messagebox.showerror("Erro", f"O contravale {numero} não foi encontrado!")
        pressionar_esc()  # Simula o pressionamento de ESC
        return
    
    # Localiza o contravale correspondente
    linha = contravales_abertos[contravales_abertos["Nº Contravale"].astype(str) == numero]
    
    # Adiciona a data e hora de baixa
    data_hora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    linha["Data e Hora Baixa"] = data_hora
    
    # Move o contravale para a planilha de baixados
    contravales_abertos = contravales_abertos[contravales_abertos["Nº Contravale"].astype(str) != numero]
    contravales_baixados = pd.concat([contravales_baixados, linha], ignore_index=True)
    
    # Salva no Excel
    with pd.ExcelWriter(ARQUIVO_CONTRAVALES, mode="a", if_sheet_exists="replace") as writer:
        contravales_abertos.to_excel(writer, sheet_name="ContravalesAbertos", index=False)
        contravales_baixados.to_excel(writer, sheet_name="ContravalesBaixados", index=False)
    
    # Copia o número do contravale para o clipboard
    pyperclip.copy(numero)

    # Tenta localizar o aplicativo para colar o número
    if focar_em_app(NOME_APP) == True:
        pyautogui.hotkey("ctrl", "v")  # Cola o número do contravale no aplicativo
        sleep(1)
        pyautogui.press("enter")
        sleep(1)
        messagebox.showinfo("Sucesso", f"Contravale {numero} baixado e colado na frente de caixa.")

    # O aplicativo não foi encontrado, logo é um contravale solitário
    else:
        messagebox.showinfo("Contravale Solitário", f"Contravale {numero} baixado na planilha.")
    
# Função para monitorar teclas
def monitorar_teclas():
    while True:
        if keyboard.is_pressed("f2"):  # Tecla 'F2' para processar contravale único
            processar_contravale()
        

# Inicializa o app
def iniciar_app():
    root = CTk()
    root.withdraw()  # Oculta a janela principal
    monitorar_teclas()

# Rodando o aplicativo
if __name__ == "__main__":
    iniciar_app()