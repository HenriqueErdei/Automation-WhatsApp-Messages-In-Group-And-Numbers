import pyautogui
import time
import openpyxl
import os
from datetime import date
import subprocess
from termcolor import colored

subprocess.call('',shell=True)

# Função para abrir o WhatsApp e enviar mensagem
def enviar_mensagem(nome_grupo, mensagem):
    # Obter a data atual
    data_atual = date.today().strftime("%d/%m/%Y")
    
    print("Abrindo o WhatsApp...")
    # Abrir a barra de pesquisa do Windows
    pyautogui.hotkey('win', 's')
    time.sleep(1)  # Esperar a barra de pesquisa abrir

    # Digitar "WhatsApp" na barra de pesquisa e pressionar Enter
    pyautogui.typewrite("WhatsApp")
    time.sleep(1)  # Aguardar a digitação
    pyautogui.press('enter')
    print("Aguarde enquanto o WhatsApp é carregado...")
    time.sleep(10)  # Aguardar o WhatsApp carregar

    print("Pesquisando pelo nome do grupo...")
    # Pesquisar pelo nome do grupo
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(2)
    pyautogui.typewrite(nome_grupo, interval=0.1)
    time.sleep(2)

    print("Selecionando o grupo...")
    # Mover para baixo na lista de resultados e pressionar Enter
    pyautogui.press('down')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(5)  # Aguardar o carregamento da conversa

    print("Enviando mensagem...")
    # Digitar a mensagem (incluindo a data atual) e enviar
    pyautogui.typewrite(f"{data_atual}: {mensagem}", interval=0.1)
    pyautogui.press('enter')
    print("Mensagem enviada!")

    print("Minimizando o WhatsApp...")
    # Minimizar a janela
    pyautogui.hotkey('win', 'm')

# Função para ler os dados do arquivo Excel
def ler_dados_excel(nome_arquivo):
    # Verificar se o arquivo existe
    if os.path.exists(nome_arquivo):
        # Carregar o arquivo Excel
        wb = openpyxl.load_workbook(nome_arquivo)
        # Selecionar a primeira planilha
        ws = wb.active
        # Inicializar uma lista para armazenar os dados
        dados = []
        # Ler os dados das duas primeiras colunas de todas as linhas da planilha
        for row in ws.iter_rows(values_only=True):
            if row[0] is not None and row[1] is not None:  # Verificar se os valores das duas primeiras colunas são diferentes de None
                dados.append((row[0], row[1]))
            else:
                print("A linha contém valores nulos nas duas primeiras colunas:", row)
        # Fechar o arquivo Excel
        wb.close()
        # Retornar os dados lidos
        return dados
    else:
        print("O arquivo especificado não existe.")
        return []

# Nome do arquivo Excel
nome_arquivo_excel = "enviar.xlsx"

# Ler os dados do arquivo Excel
dados = ler_dados_excel(nome_arquivo_excel)

# Verificar se os dados foram lidos com sucesso
if dados:
    print("Dados lidos do arquivo Excel:")
    print(dados)
    # Loop infinito para enviar as mensagens sequencialmente
    while True:
        for linha in dados:
            nome_grupo, mensagem = linha
            enviar_mensagem(nome_grupo, mensagem)
            time.sleep(5)  # Aguardar 5 segundos antes de enviar a próxima mensagem
else:
    print("Não foi possível ler os dados do arquivo Excel.")
