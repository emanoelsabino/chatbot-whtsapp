""" Funções Uteis da aplicação """
# imports
from inputimeout import inputimeout, TimeoutOccurred
from datetime import datetime, timedelta
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
import time
import urllib
from openpyxl import Workbook, load_workbook
import pandas as pd


# define o nome da planilha inicial (padrão)
def nome_planilha_inicial():
    return 'planilhas/controle_guia_fs.xlsx'

# Função que aguarda um novo nome de planilha, ou mantem a mesma caso não exista novo input
def input_nome_planilha(nome_planilha):
    try:
        nome_planilha = inputimeout(prompt=f"Atualize o caminho da planilha, ou aguarde para manter a atual [{nome_planilha}]: ", timeout=20)
        print(f'Voce digitou -> {nome_planilha}')
    except TimeoutOccurred:
        print(f'Você não digitou nada, input continua sendo -> {nome_planilha}')
    finally:
        return nome_planilha

# Função que aguarda um nova opção no final [executar novamente / sair / aguardar novo envio],
# ou mantem a mesma (aguardar novo envio) caso não exista novo input
def input_opcao(op):
    print("""
          \n\nOpções:
          [1] - Executar Novamente.
          [2] - Sair do Chatbot.
          [3] - Aguarde para o próximo envio.
          [4] - Alterar planilha atual.
          """)
    try:
        op = inputimeout(prompt=f"Digite uma opção: [ou aguarde para o próximo envio] -> ", timeout=1800)
        print(f'Voce digitou -> {op}')
    except TimeoutOccurred:
        print(f'Você não digitou nada, input continua sendo -> {op}')
    finally:
        return op

# Condição verifica opção para novo envio de mensagens
def condicao_novo_envio(hoje, nome_planilha, op):
    amanha = hoje + timedelta(days=1)
    while hoje.strftime('%d-%m-%y-%H') != amanha.strftime('%d-%m-%y-%H'):
        print(f'\n\nMensagens do dia {hoje.strftime("%d/%m/%Y")} enviadas com sucesso! \nPróximo envio: {amanha.strftime("%d/%m/%Y às %H")} horas.')
        # Recebe a opção do Modulo Utils.py
        op = input_opcao(op)
        if op == '1':
            print('op 1')
            return op, nome_planilha
        elif op == '2':
            op = 'sair'
            return op, nome_planilha
        elif op == '4':
            # Recebo o nome da planilha do Modulo Utils.py
            op = '3'
            nome_planilha = input_nome_planilha(nome_planilha)
        else:
            print('op 3')
            hoje = datetime.now()

# Verifica se WhatsApp Web está totalmente carregado
def verifica_conexao_whatsapp(navegador):
    print('Aguardando conectar WhatsApp.', end='')
    while len(navegador.find_elements(By.ID, "side")) < 1:
        print('.', end='')
        time.sleep(2)

def envia_msg_whatsapp(navegador):
    print('numero valido, enviando mensagem...')
    navegador.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div/p/span').send_keys(Keys.ENTER)
