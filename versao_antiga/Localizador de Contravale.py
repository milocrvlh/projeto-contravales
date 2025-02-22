# -*- coding: utf-8 -*-
import Tkinter as tk
import tkSimpleDialog as simpledialog
import tkMessageBox as messagebox
import keyboard
import pandas as pd
import os
from datetime import datetime
import pyperclip  # Biblioteca para gerenciar o clipboard
import pyautogui  # Biblioteca para automação de foco e colagem
import time  # Para pausas entre ações
from pywinauto import Application, findwindows # Para gerenciamento de janelas

# Caminho do arquivo Excel
ARQUIVO_CONTRAVALES = "contravales.xlsx"

# Nome do aplicativo a ser focado
NOME_APP = u"Teste - WordPad"  # Substitua pelo nome da janela do aplicativo

# Inicializa o arquivo Excel, se não existir
if not os.path.exists(ARQUIVO_CONTRAVALES):
    with pd.ExcelWriter(ARQUIVO_CONTRAVALES) as writer:
        pd.DataFrame(columns=[u"Nº Contravale"]).to_excel(writer, sheet_name="ContravalesAbertos", index=False)
        pd.DataFrame(columns=[u"Nº Contravale", "Data e Hora Baixa"]).to_excel(writer, sheet_name="ContravalesBaixados", index=False)
        pd.DataFrame(columns=[u"Nº Contravale"]).to_excel(writer, sheet_name="BaixaAutomatica", index=False)

# Função para simular a tecla ESC
def pressionar_esc():
    keyboard.press_and_release('esc')

# Função para focar no aplicativo específico
def focar_em_app(nome_app):
    try:
        # Encontra a janela do aplicativo pelo título
        janela = findwindows.find_window(title_re=nome_app)
        # Conecta-se ao aplicativo e ativa a janela
        app = Application().connect(handle=janela)
        app.window(handle=janela).set_focus()
        time.sleep(0.5)  # Aguarda para garantir que a janela esteja ativa
        return True
    except:
        messagebox.showerror("Erro", u"Incapaz de localizar o aplicativo '{}'.".format(nome_app))
        return False

# Função para processar um contravale específico
def processar_contravale():
    # Carrega os dados das planilhas
    planilhas = pd.ExcelFile(ARQUIVO_CONTRAVALES)
    contravales_abertos = pd.read_excel(planilhas, sheet_name="ContravalesAbertos", dtype=str)
    contravales_baixados = pd.read_excel(planilhas, sheet_name="ContravalesBaixados", dtype=str)
    
    # Garantir que o nome da coluna está correto
    if u'Nº Contravale' not in contravales_abertos.columns:
        # Tentativa de corrigir o nome da coluna
        contravales_abertos.columns = [col.encode('latin1').decode('utf-8') for col in contravales_abertos.columns]
    
    # Solicita o número do contravale
    try:
        numero = simpledialog.askinteger("Contravale", u"Digite o número do contravale:")
        if numero is None:  # Caso o usuário cancele a operação
            return
    except:
        pressionar_esc()  # Simula o pressionamento de ESC
        return
    
    # Converte o número para string (para validação)
    numero = str(numero)
    
    # Verifica se o contravale existe
    if numero not in contravales_abertos[u"Nº Contravale"].astype(str).values:
        messagebox.showerror("Erro", u"O contravale {} é inexistente ou já foi usado!".format(numero))
        pressionar_esc()  # Simula o pressionamento de ESC
        return
    
    # Localiza o contravale correspondente
    linha = contravales_abertos[contravales_abertos[u"Nº Contravale"].astype(str) == numero]
    
    # Adiciona a data e hora de baixa
    data_hora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    linha["Data e Hora Baixa"] = data_hora
    
    # Move o contravale para a planilha de baixados
    contravales_abertos = contravales_abertos[contravales_abertos[u"Nº Contravale"].astype(str) != numero]
    contravales_baixados = pd.concat([contravales_baixados, linha], ignore_index=True)
    
    # Salva no Excel (sobrescrevendo o arquivo)
    with pd.ExcelWriter(ARQUIVO_CONTRAVALES, mode="w") as writer:  
        contravales_abertos.to_excel(writer, sheet_name="ContravalesAbertos", index=False)
        contravales_baixados.to_excel(writer, sheet_name="ContravalesBaixados", index=False)

    
    # Copia o número do contravale para o clipboard
    pyperclip.copy(numero)
    
    # Focar no aplicativo e cola o número e dá enter, caso tenha encontrado a janela do aplicativo
    if focar_em_app(NOME_APP) == True:
        pyautogui.hotkey("ctrl", "v")  # Cola o número do contravale no aplicativo
        pyautogui.press("enter")
        messagebox.showinfo("Sucesso", u"Contravale {} baixado e colado no aplicativo.".format(numero))

    else:
        messagebox.showinfo("Aviso", u"Contravale {} baixado na planilha, mas incapaz de ser colado no aplicativo.".format(numero))

# Função para monitorar teclas
def monitorar_teclas():
    while True:
        if keyboard.is_pressed("f"):  # Tecla 'f' para processar contravale único
            processar_contravale()

# Inicializa o app
def iniciar_app():
    root = tk.Tk()
    root.withdraw()  # Oculta a janela principal
    monitorar_teclas()

# Rodando o aplicativo
if __name__ == "__main__":
    iniciar_app()
