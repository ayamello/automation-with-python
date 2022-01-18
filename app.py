import pyautogui
import pyperclip
import time
import pandas as pd

pyautogui.PAUSE = 1

# ENTRAR NO SISTEMA DA EMPRESA ABRINDO UMA NOVA ABA

# abre uma nova aba no navegador
# pyautogui.hotkey("ctrl", "t") 

# copia o link com caractere especial
#pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")

# cola o link no navegador
# pyautogui.hotkey("ctrl", "v")

# pressiona enter para entrar no link colado
# pyautogui.press("enter")

# fazer o código esperar por 5 segundos para completar o carregamento do sistema
# time.sleep(5)


# # ENTRAR NO SISTEMA DA EMPRESA ABRINDO O NAVEGADOR

# abre o menu do windows através da tecla windows
pyautogui.press("win")

# pesquisa o edge
pyautogui.write("edge")

# pressiona enter para abrir o navegador
pyautogui.press("enter")

# copia o link com caractere especial
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")

# cola o link no navegador
pyautogui.hotkey("ctrl", "v")

# pressiona enter para entrar no link colado
pyautogui.press("enter")

# fazer o código esperar por 5 segundos para completar o carregamento do sistema
time.sleep(5)


# NAVEGAR NO SISTEMA

# clicar e abrir na pasta exportar
pyautogui.click(x=356, y=251, clicks=2) # captar a posição da pasta através do pyautogui.position() e posicionando o mouse em cima da paste

time.sleep(3)

# clicar no arquivo
pyautogui.click(x=356, y=251)

# clicar nos 3 pontos
pyautogui.click(x=1162, y=157)

# clicar em fazer download
pyautogui.click(x=932, y=558)

# esperar download
time.sleep(5)


# IMPORTAR A BASE DE VENDAS PRO PYTHON

# Ler arquivo de vendas
table = pd.read_excel(r"C:\Users\ayana\Downloads\Vendas - Dez.xlsx")

faturamento = table["Valor Final"].sum()
qtde_produtos = table["Quantidade"].sum()


# ENVIAR EMAIL

# abre email
pyautogui.hotkey("ctrl", "t") 
pyperclip.copy("https://mail.google.com/mail/u/1/?ogbl#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

time.sleep(3)

# clica no botão escrever
pyautogui.click(x=88, y=166)

time.sleep(5)

# preencher destino
pyautogui.write("dev.ayanamello@gmail.com")
pyautogui.press("tab") # seleciona o email e muda para o próximo campo
pyautogui.press("tab")

# preencher o assunto
pyperclip.copy("Relatório de vendas")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab") # muda para o campo de corpo do email

# preencher o corpo do email
text = f'''
Prezados diretores,

O faturamento de ontem foi de R${faturamento:,.2f}
A quantidade de produtos foi {qtde_produtos:.0f}

At.te,
Ayana Mello.
'''

pyperclip.copy(text)
pyautogui.hotkey("ctrl", "v")


# clica em enviar
pyautogui.hotkey("ctrl", "enter")

time.sleep(4)

pyautogui.click(x=1344, y=12)