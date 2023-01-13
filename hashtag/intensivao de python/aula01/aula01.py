import time
import pyautogui as pa
import pandas as pd
import numpy
import openpyxl
import pyperclip


# passo 1: entrar no sistema da empresa (no link do drive, ou no site do programa)
# passo 2: navegar até o local do relatório (entrar na pasta do explorer)
# passo 3: exportar o relatório (fazer download)
# passo 4: calcular os indicadores (faturamento e quantidade de produtos)
# passo 5: enviar um e-mail para diretoria

def pegar_posicao_mouse(segundos):
    time.sleep(segundos)
    print(pa.position())


def clica(x, y):
    pa.click(x, y)
    espere(0.25)


def cliques(x, y, qtd_cliques):
    pa.click(x, y, clicks=qtd_cliques)


def atalho(x, y):
    pa.hotkey(x, y)
    espere(0.25)


def pressione(tecla):
    pa.press(tecla)
    espere(0.25)


def colar():
    atalho("ctrl", "v")


def copiar_colar(copiar):
    pyperclip.copy(copiar)
    colar()
    espere(1)


def espere(segundos):
    time.sleep(segundos)


def escreva(texto):
    pa.write(texto)
    espere(1)


def tab(segundos):
    pressione("tab")
    espere(segundos)


df = pd.read_excel(r'C:\Users\Lucas\Downloads\Exportar-20230109T230421Z-001\Exportar\Vendas-Dez.xlsx')

faturamento = df["Valor Final"].sum()
quantidade = df["Quantidade"].sum()

url_gmail = "https://mail.google.com/mail/u/0/#inbox"
botao_escrever = 145, 198   # coloque aqui os valores que você recebeu da função pegar_posicao_mouse
email = "imparcialista@gmail.com"
assunto = "Relatório de vendas"
corpo = f"""Prezados, bom dia!

O faturamento de ontem foi de: {faturamento:,.2f}
A quantidade de produtos vendidos foi de: {quantidade:,.2f}

Atenciosamente,
Lucas A.
"""


def enviar_relatorio():
    espere(3)
    atalho("alt", "tab")
    # programa iniciado

    atalho("ctrl", "t")  # abre uma nova guia
    espere(1)

    copiar_colar(url_gmail)
    pressione("enter")  # entra no link do gmail
    espere(5)

    clica(145, 198)  # clica em escrever
    espere(1)
    escreva(email)  # escreve o e-mail
    espere(1)
    pressione("tab")  # seleciona o e-mail
    espere(1)
    pressione("tab")  # pula para o campo de assunto
    espere(1)

    # assunto
    copiar_colar(assunto)  # pula para o campo de corpo e espera um segundo
    tab(1)

    # corpo
    copiar_colar(corpo)  # escreve o corpo e espera dois segundos

    # enviar o e-mail
    espere(3)
    atalho("ctrl", "enter")

    espere(3)
    atalho("alt", "tab")
    # programa finalizado


# pegar_posicao_mouse(5)
# precisamos primeiro pegar a posição do mouse no botão escrever, pois, isso muda conforme o monitor
enviar_relatorio()

