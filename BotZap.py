import os
import random
import re
import threading
import time
import tkinter as tk
import urllib
from tkinter import filedialog
import phonenumbers
import numpy as np
import openpyxl
import pandas as pd
import requests
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager


################################################################################################
def contagem_regressiva(segundos, mensagem):
    for i in range(segundos, -1, -1):
        countdown = f"{i:2d}"
        texto = f"Aguardando {countdown} segundos para {mensagem}"
        print(f"\r{texto}", end="", flush=True)
        time.sleep(1)
    print("")

print('Ao abrir o explorer informe o caminho da planilha com os contatos')
contagem_regressiva(2, "segundos para começar")

def ler_banco_de_dados():
    root = tk.Tk()
    root.withdraw()
    caminho = filedialog.askopenfilename()
    if not caminho:
        print("Nenhum arquivo selecionado!")
        return None
    banco_de_dados = pd.read_excel(caminho)
    banco_de_dados = banco_de_dados.astype(str)
    return banco_de_dados

DataProducao = ler_banco_de_dados()
print(DataProducao)

while True:
    declarou = input("Deseja enviar mensagem apenas para quem NÃO declarou a última campanha? (0 para sim | 1 para não 'TODOS' | 3 Para continuar de onde parou)\n")
    if declarou == "0" or declarou == "1" or declarou == "3":
        break
    else:
        print("Opção inválida. Tente novamente.\n")

if declarou == "3":
    DataProducao = DataProducao
elif declarou == "0":
    print('Excluindo fichas não declaradas')
    DataProducao.drop(labels=DataProducao[DataProducao['Dec. Rebanho'] == "1"].index, axis=0, inplace=True)
    DataProducao = DataProducao.drop(columns=['Dec. Rebanho'])
else:
    declarou == "1"
    print('Mantido Todos os registros')

    print('Excluindo colunas excedentes')
    colunas_a_manter = ['Nome do Titular da Ficha de bovideos', 'Nome da Propriedade', 'Endereço da Prop.', 'Telefone 1', 'Telefone 2', 'Celular']
    DataProducao = DataProducao[colunas_a_manter]

    print('Criando uma coluna para o Status')
    if 'Status' not in DataProducao.columns:
        DataProducao['Status'] = 0
        DataProducao = DataProducao.reindex(columns=['Status'] + list(DataProducao.columns[:-1]))

    print('Agrupando as colunas Nome, endereço e propriedade')
    def concatenar_informacoes(row):
        nome = row["Nome do Titular da Ficha de bovideos"]
        propriedade = row["Nome da Propriedade"]
        endereco = row["Endereço da Prop."]
        return "{} - {} - {}".format(nome, propriedade, endereco)

    # Aplica a função ao dataframe
    DataProducao["Nome do Titular da Ficha de bovideos"] = DataProducao.apply(concatenar_informacoes, axis=1)

    # Remove as colunas excedentes
    DataProducao = DataProducao.drop(columns=["Nome da Propriedade", "Endereço da Prop."])
    # Renomeia as duas colunas
    DataProducao = DataProducao.rename(columns={"Nome do Titular da Ficha de bovideos": "nome", "Telefone 1": "telefone"})

def atualizar_planilha(df, nome_arquivo='BancoProducao.xlsx'):
    diretorio_atual = os.path.dirname(os.path.abspath(__file__))
    caminho_arquivo = os.path.join(diretorio_atual, nome_arquivo)
    df.to_excel(caminho_arquivo, index=False)

    print('Agrupando as três colunas de telefone em uma só')
    # Agrupa os telefones Telefone 2 e Celular todos na coluna telefone 
    DataProducao = DataProducao.melt(id_vars=["nome", "Status"], value_vars=["telefone", "Telefone 2", "Celular"], var_name="tipo_telefone", value_name="telefone")
    DataProducao = DataProducao.loc[:, ["Status", "nome", "telefone"]]

    print('Removendo linhas com valores nulos ou apenas caracteres especiais na coluna "telefone"')
# Remover linhas com valores nulos ou apenas caracteres especiais na coluna "telefone"
    DataProducao = DataProducao.dropna(subset=["telefone"])
    DataProducao = DataProducao.drop(DataProducao[(DataProducao["telefone"].isnull()) | (DataProducao["telefone"].str.contains(r'^[\(\)\s-]+$'))].index)

    print('Substituindo nulos e inválidos na coluna telefone')
    # Substitui valores nulos ou inválidos na coluna 'telefone'
    DataProducao['telefone'] = DataProducao['telefone'].apply(lambda x: 'aaaa' if (isinstance(x, str) and not x[-4:].isdigit()) else x)
    DataProducao['telefone'] = DataProducao['telefone'].fillna('aaaa')
    DataProducao = DataProducao.loc[DataProducao['telefone'] != 'aaaa']

    print('Removendo caracteres inválidos')
    # Remove caracteres não numéricos do telefone
    DataProducao['telefone'] = DataProducao['telefone'].apply(lambda x: re.sub('[^0-9]', '', x))

    print('Adicionando o nono dígito quando ele não estiver presente')
    # Adiciona o 9 na frente quando estiver ausente
    def adicionar_nono_digito(df):
     df['telefone'] = df['telefone'].astype(str)
     df['telefone'] = df['telefone'].apply(lambda x: x[:2] + '9' + x[2:] if len(x) == 10 else x)
    return df
    DataProducao = adicionar_nono_digito(DataProducao)

def corrigir_e_formatar_numero(numero):
    # Parse do número de telefone
    numero_telefone = phonenumbers.parse(numero, "BR")

    # Verificar se o número é válido
    if phonenumbers.is_valid_number(numero_telefone):
        # Formatação do número de telefone
        telefone_formatado = phonenumbers.format_number(numero_telefone, phonenumbers.PhoneNumberFormat.INTERNATIONAL)
    else:
        # Tentar corrigir o número de telefone
        numero_corrigido = phonenumbers.parse(numero, "BR")
        telefone_formatado = phonenumbers.format_number(numero_corrigido, phonenumbers.PhoneNumberFormat.INTERNATIONAL)

    return telefone_formatado

print('Formatando os telefones para o padrão nacional/internacional')
DataProducao['telefone'] = DataProducao['telefone'].apply(corrigir_e_formatar_numero)

print('Agrupando as linhas pelo número do telefone')
# Agrupa as linhas pelo número de telefone e concatena os nomes
DataProducao = DataProducao.groupby(["Status", "telefone"])["nome"].apply(lambda x: " || ".join(x)).reset_index()

# Renomeia as colunas
DataProducao = DataProducao.rename(columns={"nome": "Nomes"})

print('Removendo as duplicatas, se houver')
# Remove duplicatas de nome e telefone
DataProducao.drop_duplicates(subset=['Nomes', 'telefone'], keep='first', inplace=True)

atualizar_planilha(DataProducao, 'BancoProd.xlsx')
print('Exibindo o produto final')
print(DataProducao)

# ################################################################################################
# ################Iniciar logica de envio e leitura de mensagens##################################
# ################################################################################################
print('Abrindo o whatzapp')

agent = {"User-Agent": 'Mozilla/5.0 (Windows NT 6.3; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/59.0.3071.115 Safari/537.36'}
api = requests.get("https://editacodigo.com.br/index/api-whatsapp/xgLNUFtZsAbhZZaxkRh5ofM6Z0YIXwwv" ,  headers=agent)
time.sleep(1)
api = api.text
api = api.split(".n.")
bolinha_notificacao = api[3].strip()
contato_cliente = api[4].strip()
caixa_msg = api[5].strip()
msg_cliente = api[6].strip()
dir_path = os.getcwd()
chrome_options2 = Options()
chrome_options2.add_argument(r"user-data-dir=" + dir_path + "/pasta/sessao")
Servico = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=Servico)
driver.get('https://web.whatsapp.com')
while len(driver.find_elements(By.ID, 'side')) < 1:
       time.sleep(1)
time.sleep(2)

# ################################################################################################
# esperar a tela do whatsapp carregar -> espera um elemento que só existe na tela já carregada aparecer
# -> lista for vazia -> que o elemento não existe ainda
while len(driver.find_elements(By.ID, 'side')) < 1:
       time.sleep(1)
time.sleep(2)  # só uma garantia

# # ################################################################################################
def CompletarComNonoDig(numero_telefone):
    numero_telefone = str(numero_telefone)
    if len(numero_telefone) == 16:
        numero_telefone = numero_telefone[:7] + '9' + numero_telefone[7:]
    return numero_telefone

def NovaMensagem():
 
    try:      
        # PEGA A BOLINHA VERDE
        bolinha = driver.find_element(By.CLASS_NAME, bolinha_notificacao)
        bolinha = driver.find_elements(By.CLASS_NAME, bolinha_notificacao)
        clica_bolinha = bolinha[-1]
        acao_bolinha = webdriver.common.action_chains.ActionChains(driver)
        acao_bolinha.move_to_element_with_offset(clica_bolinha, 0, -20)
        acao_bolinha.click()
        acao_bolinha.perform()
        acao_bolinha.click()
        acao_bolinha.perform()

        # PEGA O TELEFONE DO CLIENTE
        telefone_cliente = driver.find_element(By.XPATH, contato_cliente)
        telefone_final = telefone_cliente.text
        print(telefone_final)

        # PEGA A MENSAGEM DO CLIENTE
        todas_as_msg = driver.find_elements(By.CLASS_NAME, msg_cliente)
        todas_as_msg_texto = [e.text for e in todas_as_msg]
        msg = todas_as_msg_texto[-1]
        print(msg)

        # DEFINIR A RESPOSTA
        print('Respondendo... ')
        print(telefone_final)
        telefone_format = CompletarComNonoDig(telefone_final)
        numero_status = DataProducao.loc[DataProducao['telefone'] == telefone_format, 'Status'].values[0]

        nome = DataProducao.loc[DataProducao['telefone'] == telefone_format, 'Nomes'].values
        if len(nome) > 0:
           nome = nome[0]
        Respondeu_sim = 'Prezado produtor estamos nos ultimos dias da campanha e não consta em nosso banco de dados a sua declaração obrigatória semestral de rebanho, por favor procure a agência IDARON o mais breve possivel e faça sua declaração evitando transtornos e aborrecimentos, caso tenha senha cadastrada pode fazer sua declaração tambem pelo Site: http://www.idaron.ro.gov.br. Para maiores informações pode entrar em contato com nosso numero de whatsapp (69)9245-2646, Estamos aguardando, obrigado. '
        Respondeu_nao = 'Obrigado por responder, vamos providenciar para que seu numero seja retirado de nossa base de contatos'
        Desculpe = f"Desculpe, não entendi sua resposta. Vamos tentar novamente?\n\nEste número ({telefone_final}) está cadastrado na *IDARON* para contato com o produtor - ({nome}). Você é ele(a) ou responde por ele(a)? RESPONDA  *Sim* ou  *Não*\n\nSim\nNão"
        
        if numero_status == 'Env1':
           if msg == 'Sim':
            #if msg.lower() in RespostasValidas_SIM:
             RESPOSTA = Respondeu_sim
             DataProducao.loc[DataProducao['telefone'] == telefone_format, 'Status'] = 'Mensagem completa'
             atualizar_planilha(DataProducao, 'BancoProd.xlsx')
     
           elif msg == 'Não':
            #if msg.lower() in RespostasValidas_NAO:
             DataProducao.loc[DataProducao['telefone'] == telefone_format, 'Status'] = 'N Resp. pelo contato'
             atualizar_planilha(DataProducao, 'BancoProd.xlsx')
             RESPOSTA = Respondeu_nao
         
           else:
             RESPOSTA = Desculpe
        
        else:
            RESPOSTA = 'Desculpe, esse contato só opera envio de mensagens automaticas, para atendimento pode entrar em contato pelo numero (69)9245-2646'    
        # RESPONDER A MENSAGEM
        campo_de_texto = driver.find_element(By.XPATH, caixa_msg)
        campo_de_texto.click()
        # resposta = requests.get("http://localhost/bot/index.php", params={'msg': msg, 'telefone': telefone_final})
        # bot_resposta = resposta.text
        time.sleep(3)
        campo_de_texto.send_keys(RESPOSTA, Keys.ENTER)
        linhas_alt= DataProducao.loc[DataProducao['telefone'] == telefone_format]
        print(linhas_alt)
        # FECHA O CONTATO
        webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    except:
        print('buscando novas mensagens')
        time.sleep(3)

####################################################################
####################################################################

def criar_link_whatsapp(numero, mensagem):
    mensagem_codificada = urllib.parse.quote(mensagem.format(telefone=numero))
    link = f"https://web.whatsapp.com/send?phone={numero}&text={mensagem_codificada}"
    return link

# # ################################################################################################
# # Estrutura pra enviar a mensagem
def disparar_mensagem(link_whatsapp, arquivo="N"):
    driver.get(link_whatsapp)
    
    # Esperar até que o elemento 'side' esteja presente
    while len(driver.find_elements(By.ID, 'side')) < 1:
        time.sleep(1)
    time.sleep(2)
    
    # Verificar se o número é inválido
    if len(driver.find_elements(By.XPATH, '//*[@id="app"]/div/span[2]/div/span/div/div/div/div/div/div[1]')) < 1:
        # Enviar a mensagem
        driver.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span').click()
        
        if arquivo != "N":
            caminho_completo = os.path.abspath(f"arquivos/{arquivo}")
            driver.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/div/span').click()
            driver.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/span/div/div/ul/li[4]/button/input').send_keys(caminho_completo)
            time.sleep(2)
            driver.find_element(By.XPATH, '//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/div/div[2]/div[2]/div/div').click()
            
        webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        return True
    else:
        return False

def EnviarMensagem():
    contagem_regressiva(11, "Segundos para começar os disparos")
    
    try:
        filtro = DataProducao['Status'] == 0
        if not DataProducao.loc[filtro, 'telefone'].empty:
            numero = DataProducao.loc[filtro, 'telefone'].iloc[0]
            NomeContato = DataProducao.loc[filtro, 'Nomes'].iloc[0]
           
            mensa = f'Olá tudo bem?😊\nEste número ({numero}) está cadastrado na *IDARON* para contato com o produtor - {NomeContato}. Você é ele(a) ou responde por ele(a)? RESPONDA *Sim* ou *Não*'
            MensagemPergunta = mensa + "\n\nSim\nNão"
            link = criar_link_whatsapp(numero, MensagemPergunta)
            print(f'Enviando mensagem para ({numero}).')
            # Se a mensagem foi enviada
            if disparar_mensagem(link, "N"):
                DataProducao.loc[DataProducao['telefone'] == numero, 'Status'] = 'Env1'
                atualizar_planilha(DataProducao, 'BancoProd.xlsx')
                linhas_alteradas = DataProducao.loc[DataProducao['telefone'] == numero]
                print(linhas_alteradas)
               
            # Se o número for inválido
            else:
                DataProducao.loc[DataProducao['telefone'] == numero, 'Status'] = 'Invalido'
                atualizar_planilha(DataProducao, 'BancoProd.xlsx')
                print("Número inválido!")
                NovaMensagem()

        else:
            print("A lista de envios foi completada, não há mais contatos para enviar.")
        
        time.sleep(2)  # Aguarda 2 segundos
        
    except Exception as e:
        # print(f"Erro ao enviar mensagem: {e}")
        pass

while True:
    has_new_message = NovaMensagem()
    if not has_new_message:
        EnviarMensagem()