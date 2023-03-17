import time
import PyPDF2
import re
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from time import sleep
import urllib.parse
from tkinter import *

def EnviarMensagens():
    # importando a planilha excel

    listacontatos = pd.read_excel(r"mensagem.xlsx", engine='openpyxl')

    # fazendo o chrome abrir o whatsapp web

    s = Service('chromedriver.exe')
    navegador = webdriver.Chrome(service=s)
    navegador.get("https://web.whatsapp.com/")

    # a cada 10 segundos confira se foi logado

    while len(navegador.find_elements(By.XPATH, "//*[@id=\"pane-side\"]")) < 1:
        sleep(10)

    # Quando logado, envie a mensagem da planilha para cada contato dela

    for i, mensagem in enumerate(listacontatos['Mensagem']):
        if listacontatos.loc[i, 'Check'] != 'ok':

            try:
                Telefone = listacontatos.loc[i, "Telefone"]
                Telefone = re.sub(r'[^\w]', ' ', Telefone)
                Telefone = re.sub('', '', Telefone)
                texto = urllib.parse.quote(f"{mensagem}")
                link = f"https://web.whatsapp.com/send?phone=55{Telefone}"
                #&text={texto} - envia texto do Excel
                navegador.get(link)

                # esperar o whatsapp com a msg do contato carregar
                while len(navegador.find_elements(By.XPATH, "//*[@id=\"pane-side\"]")) < 1:
                    time.sleep(7)
                # envia a midia e espera 7s para começar o loop

                # navegador.find_element(By.XPATH, "//*[@id=\"main\"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button").send_keys(Keys.ENTER) - enter para enviar mensagem de texto

                # mídia a ser enviada
                midia = "C:/Users/LTMat/Desktop/DNA_bot/video.mp4"
                navegador.find_element(By.CSS_SELECTOR, "span[data-icon='clip']").click()
                attach = navegador.find_element(By.CSS_SELECTOR, "input[type='file']")
                attach.send_keys(midia)
                time.sleep(7)
                send = navegador.find_element(By.CSS_SELECTOR, "span[data-icon='send']")
                time.sleep(7)
                send.click()

                listacontatos.loc[i, 'Check'] = 'ok'
                time.sleep(7)

            except:
                print("erro")
                listacontatos.loc[i, 'Check'] = 'erro'

    listacontatos.to_excel(r"mensagem.xlsx", index=False)
    print(listacontatos)

# INTERFACE

Window = Tk()
Window.title("Bot")

botao0 = Button(Window, text='Selecionar mídia', bg="green", fg="white", font="Arial")
botao0.grid(column=5, row=36, padx=10, pady=10)

botao1 = Button(Window, text='Selecionar Coluna', bg="green", fg="white", font="Arial")
botao1.grid(column=5, row=37, padx=10, pady=10)

bota2 = Button(Window, text='Enviar Mensagens', command=EnviarMensagens, bg="green", fg="white", font="Arial")
bota2.grid(column=5, row=40, padx=10, pady=290)

botao3 = Button(Window, text='fechar', bg="green", fg="white", font="Arial")
botao3.grid(column=5, row=45, padx=10, pady=300)

# andamento = Label(janela, text='0 de ')
# andamento.grid(column=0, row=4, padx=10, pady=10)

Window.geometry("650x500+200+200")
Window['bg'] = "#403B3B"
Window.resizable(False, False)
Window.mainloop()

# sempre colocar isso no final pra manter a janela aberta