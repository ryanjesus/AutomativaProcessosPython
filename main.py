# pip install pyautogui
"""
Automação de processo
"""
import pyautogui
import pyperclip
import time
import pandas as pd
# import openpyxl
# import numpy
from IPython.display import display

pyautogui.PAUSE = 0.5

# Abrindo o navegador
pyautogui.press("win")
pyautogui.write("chrome")
pyautogui.press("enter")
# pyautogui.press("win")

# Passo 1: Entrar no sistema (no nosso caso vai ser entrar no link)
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

# site ta carregando
time.sleep(3)

# Passo 2: Navegar no sistema e encontrar a base de dados (entrar na pasta Exportar)
time.sleep(3)
pyautogui.click(x=436, y=320, clicks=2)
time.sleep(3)

# Passo 3: DowLoad da base de dados
pyautogui.click(x=441, y=441)  # clicou no arquivo
pyautogui.click(x=1009, y=209)  # clicou na opções
pyautogui.click(x=837, y=644)  # clicou em Dowload
time.sleep(3)
pyautogui.press("enter")
time.sleep(3)

# Passo 4: Calcular os indicadores(faturamento, quantidade de produtos)
tabela = pd.read_excel(r"C:\Users\Usuario\Desktop\Intensivo de python\Vendas - Dez.xlsx")
display(tabela)

quantidade = tabela["Quantidade"].sum()  # Soma da coluna quantidade

faturamente = tabela["Valor Final"].sum()  # Soma da coluna valor final

print(quantidade)
print(faturamente)

# Passo 5: Entra no email
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(7)

# Passo 6: Mandar um email para a diretoria com os indicadores

# Clicar no portão +
pyautogui.click(x=119, y=209)
time.sleep(2)

# escreve o destinatário
pyautogui.write("ryanoliveira.vm@gmail.com")
pyautogui.press("tab")
pyautogui.press("tab")

# escreve o assunto
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab")

# escreve o corpo do email
texto = f"""
Prezados, bom dia

O faturamento de ontem foi de: R${faturamente:,.2f}
A quantidade de produtos foi de: {quantidade:,}

Abs
Ryan Oliveira
"""
pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")

# enviar o email
pyautogui.hotkey("ctrl", "enter")
