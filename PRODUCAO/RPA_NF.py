#
#
# Este é o exemplo do codigo do meu projeto de RPA, está versão não está funcional pois falta informações essenciais para rodar
#
#

#importando bibliotecas e funcoes necessarias 
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from time import sleep 
from pyautogui import press
import pyautogui
from datetime import datetime
from googleapiclient.discovery import build
from google.oauth2 import service_account
import os
import fitz
import re
import mysql.connector
import tkinter as tk
from tkinter import messagebox

#GUI para interação com o usuario 
janela = tk.Tk()
janela.title("Coleta de Mês para Notas")

rotulo_mes = tk.Label(janela, text="Mês (1 a 12):")
rotulo_mes.grid(row=0, column=0, padx=5, pady=5)

entrada_mes = tk.Entry(janela)
entrada_mes.grid(row=0, column=1, padx=5, pady=5)

def validar_mes():
    try:
        global data
        data = int(entrada_mes.get())
        if 1 <= data <= 12:
            # Mês válido, prosseguir com a lógica do programa
            print(f"Mês selecionado: {data}")
            janela.destroy()  # Fechar a janela após a validação
        else:
            messagebox.showerror("Mês Inválido!", "Por favor, digite um mês entre 1 e 12.")
    except ValueError:
        messagebox.showerror("Erro de Digitação!", "Digite um número inteiro para o mês.")

botao_validar = tk.Button(janela, text="Validar", command=validar_mes)
botao_validar.grid(row=2, column=0, columnspan=2, padx=5, pady=5)

janela.protocol("WM_DELETE_WINDOW", validar_mes)
janela.mainloop()


#conexão com banco de dados
db = mysql.connector.connect(
    host=###,
    user=###,
    password=###, 
    database=###)

cursor = db.cursor()

nome_tabela = 'Mes' +  '_' + str(data)

#executando comando para criar a tabela do mês no MySQL
cursor.execute(f"""
    CREATE TABLE IF NOT EXISTS {nome_tabela} (
        id INT PRIMARY KEY AUTO_INCREMENT,
        conta INT NOT NULL UNIQUE,
        nome_nota VARCHAR(30) NOT NULL,
        valor_total VARCHAR(15) NOT NULL,
        juros VARCHAR(15),
        data_vencimento VARCHAR (10) NOT NULL,
        ref VARCHAR (2) NOT NULL,
        data_emissao VARCHAR (10) NOT NULL,
        id_drive VARCHAR(50) UNIQUE NOT NULL
    );
    """)

db.commit()


SCOPES = ["https://www.googleapis.com/auth/drive"] #escopo de permissão que o codigo tem sobre o drive
SERVICE_ACCOUNT_FILE = 'CHAVES/token.json' #token para acesso drive
SHARED_DRIVE_ID = #ID da pasta no drive para armazenar as NFs PRECISA SER TROCADO A CADA ANO!!!!


#dicionario contendo XPATHs utilizados
Portal_map = {
    "buttons": {
        "download":{
            "xpath": "/html/body/center/form/table/tbody/tr[6]/td/input"
        }
    }
}



pasta_downloads = os.path.join(os.path.expanduser("~"), "Downloads") #declarando automaticamente o caminho até a pasta de downloads de qualquer maquina 


#configurando o navegador
chrome_options = Options()
chrome_options.add_experimental_option("detach", True)
#servico = Service(ChromeDriverManager().install()) #instalando o driver atual do chrome de forma automatica
navegador = webdriver.Chrome(options=chrome_options)

def reinciando(): #funçao para reiniciar o processo
    navegador.close() #fechando a aba de download
    navegador.switch_to.window(navegador.window_handles[0]) #voltando a controlar a tela de login do portal
    navegador.find_element(By.XPATH, '/html/body/center/a').click() #clicando no botao "reiniciar" para o processo de login

def verificando_conta(conta): #funçao para selecionar as contas 
    combobox_contas = navegador.find_element(By.XPATH, "/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr/td[6]/form/select")
    combobox_contas.click()
    options_contas = combobox_contas.find_elements(By.TAG_NAME, "option")
    sleep(2)
    
    for option in options_contas:
        if option.text == conta:
            option.click()
            break

def verificando_data(): #funçao para verificar as datas das opçoes disponiveis
    global fatura #declarando a variavel fatura (ela que diz se a fatura do mes atual está disponivel ou não)
    combobox = navegador.find_element(By.XPATH, "/html/body/center/form/table/tbody/tr[4]/td/select") #armazenando a lista de datas em uma variavel
    combobox.click() #abrindo a lista
    options = combobox.find_elements(By.TAG_NAME, "option") #passando por todas as opçoes dentro da lista da variavel de datas disponiveis
    sleep(2)
    fatura = False #recebendo false para que a cada conta logada/fatura ele tenha que verificar novamente

    #lendo todas opções disponiveis
    for option in options:
        final_option = option.text.split("/")[2] #pegando a parte final da opção (o ano e referencia estão nela)
        ref_option = final_option.split(".")[1] #pegando a referenia da opção
        year_option = final_option.split("-")[0] #pegando o ano da opção
        
        #transformando em inteiros para comparar as datas com as atuais
        ref_option = int(ref_option)
        ano_option = int(year_option)
        
        #pegando ao ano atual
        ano_atual = datetime.now().year
        
        #se encontrar uma fatura com o mes que o usuario informou e ano atual vai selecionar ela para download
        if ref_option == data and ano_option == ano_atual :
            fatura = True
            option.click()
            break

def autenticar():  #autenticando as credenciais da API do drive
    creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    return creds

def check_file_exists(file_name): #função para checar se a NF já está na pasta do drive
    creds = autenticar()
    service = build('drive', 'v3', credentials=creds)

    # Busca por arquivos com nome correspondente na pasta pai
    response = service.files().list(
        supportsAllDrives=True, #parametro para incluir drivers compartilhados
        q=f"name = '{file_name}' and parents = '{SHARED_DRIVE_ID}'"
    ).execute()

    # Verifica se algum arquivo foi encontrado
    if response['files']:
        return True  # O arquivo existe no Drive
    else:
        return False  # O arquivo não existe no Drive

def arquivo_recente(): #função para pegar o arquivo mais recente na pasta downloads (ultima NF baixada)
    lista_arquivos = os.listdir(pasta_downloads) #listando o diretorio de downloads
    lista_datas = [] #criando outra lista pata armezenar o tempo de modificação de cada arquivo em downloads

    for arquivo in lista_arquivos:
    # descobrir a data desse arquivo
        if ".pdf" in arquivo:
            data = os.path.getmtime(f"{pasta_downloads}/{arquivo}") #pegando a data de modificaçãp de cada pdf 
            lista_datas.append((data, arquivo)) #armazenadno o tempo de modificação e o nome do arquivo em formato de matriz

    #ordenando em ordem crescente e pegando o nome do primerio arquivo 
    lista_datas.sort(reverse=True)     
    global ultimo_arquivo
    ultimo_arquivo = lista_datas[0]

def extração(): #função para extrair as informações (Total a pagar, Juros, Data de Emissão) das NFs

    #abrindo o arquivo e passando as paginas
    pdf = fitz.open(f'{pasta_downloads}/{ultimo_arquivo[1]}')
    first_page = pdf[0]
    second_page = pdf[1]
    three_page = pdf[2]
    
    #obtendo os textos das paginas
    text = first_page.get_text() 
    text2 = second_page.get_text()
    text3 = three_page.get_text()
    
    global data_emissao
    data_emissao = None
        # Expressão regular para encontrar "Data de Emissão"
    match = re.search(r"Data de emissão:\s+(.+)", text2)
    if match:
        data_emissao = match.group(1)

    if data_emissao:
        print(f"data de emissão: {data_emissao}")
    else:
        match = re.search(r"Data de emissão:\s+(.+)", text3)
        if match:
            data_emissao = match.group(1)
            print(f"data de emissão: {data_emissao}")


    global valor_total
    valor_total = None
        # Expressão regular para encontrar "Total a pagar" 
    match = re.search(r"Total a pagar\s+R\$ ([\d.,]+)", text)
    if match:
        valor_total = match.group(1)

    if valor_total:
        print(f"Valor total a pagar: R$ {valor_total}")
    else:
        print("Valor total não encontrado na fatura.")

    global juros
    juros = None
        # Expressão regular para encontrar "Itens Adicionais" 
    match = re.search(r"Itens Adicionais\s+R\$ ([\d.,]+)", text)
    if match:
        juros = match.group(1)

    if juros:
        print(f"Valor total de juros: R$ {juros}")
    else:
        print("Valor total de juros não encontrado.")

def upload_drive(file_path): #função para realizar upload no drive
    creds = autenticar()
    service = build('drive', 'v3', credentials=creds)

    #lista de contas com seus respectivos estados para nomear a NF
    estados = [ (#, "SP"), (, "PA"), (, "BA"), (, "ES"), (, "SP"), 
                (#, "RJ"), (, "MA"), (, "AM"), (, "PI"), (, "PE"), 
                (#, "MG"), (, "CE") ]

    #numero de conta e vencimento para nomear a NF
    global file_name
    global n_conta
    global vencimento
    n_conta = ultimo_arquivo[1].split("_")[0]
    vencimento = ultimo_arquivo[1].split("_")[1]
    referencia = ultimo_arquivo[1].split("_")[2]

    #iterando sobre a lista de estados para ver qual é o estado da conta atual
    for e in estados:
        if e[0] == login:
            estado = str(e[1])
            
    #nessa conta são NFs de serviços, mans nas outra sera seguranca
    if login == :   
        file_name = os.path.basename(f"{n_conta}_Claro GR Servico {estado}" + "_" + f"{vencimento}" + "_" + f"ref-{referencia}.pdf") #padrao para NF servico
    else:       
        file_name = os.path.basename(f"{n_conta}_Claro GR Seguranca {estado}" + "_" + f"{vencimento}" + "_" + f"ref-{referencia}.pdf") #padrao para NF seguranca
    
    #chamando a função para checar a pasta no drive e evitar duplicação
    if check_file_exists(file_name) == True:
        print(f"O arquivo '{file_name}' já existe no Drive. Pulando envio.")
        return

    else:
        #nome do arquivo e para que pasta vai
        file_metadata = {
            'name' : (f"{file_name}"),
            'parents' : [SHARED_DRIVE_ID]
        }
        
        response = service.files().create(
            supportsAllDrives=True, #linha para permitir o codigo acessar pastas compartilhadas no drive
            body=file_metadata, #corpo do arquivo no drive
            media_body=file_path #informações do arquivo no drive
            ).execute()
        
        #coletando o id do arquivo no drive para armazenar no banco de dados
        global file_id
        file_id = response['id']

def insert_MySQL():

    #dando insert na tabela criado no MySQL
    try:
        cursor.execute(f"""
            INSERT INTO {nome_tabela} (
                conta, nome_nota, valor_total, juros, data_vencimento, ref,  data_emissao, id_drive
            ) VALUES (
                '{n_conta}', '{file_name}', '{valor_total}', '{juros}', '{vencimento}', '{data}', '{data_emissao}', '{file_id}'
            )
        """)
        db.commit()
    except:
        print(f"Nota já está no banco de dados")

def remover_arquivo(): #funçao para remover o pdf da nota da pasta downloads
    path = (f'{pasta_downloads}/{ultimo_arquivo[1]}')
    if os.path.exists(path):
        os.remove(path)

#abrindo o portal da claro 
navegador.get("https://contaonline.claro.com.br/webbow/login/initPJ_oqe.do ")

#abrindo o arquivo com os logins
with open("PRODUCAO/claro.txt", 'r') as arquivo:
    for linha in arquivo:
        login = int(linha) #armezando a linha atual em uma variavel int para verificar se é um login com mais de uma fatura para baixar

        #logando no portal e indo para a aba de downloads
        navegador.find_element(By.XPATH, '/html/body/form/table/tbody/tr[2]/td[1]/input').send_keys(linha) #preenchendo o campo de login
        press('enter') #solucao para o pop up que aparece na tela de login
        sleep(1)
        navegador.find_element(By.XPATH, '/html/body/form/table/tbody/tr[2]/td[2]/input').send_keys() #preenchendo a senha
        navegador.find_element(By.XPATH, '/html/body/form/table/tbody/tr[2]/td[3]/input').click() #dando ok para login
        sleep(3)
        pyautogui.hotkey("ctrl", "w") #comando para fechar guia pop-up do site
        navegador.switch_to.window(navegador.window_handles[1]) #mudando a aba para ser controlada
        sleep(3)
        navegador.find_element(By.XPATH, '/html/body/table[1]/tbody/tr/td[1]/ul/table/tbody/tr/td[5]/li/a/img').click() #indo para gerenciamento
        (sleep(2))
        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td[1]/ul/table/tbody/tr/td[5]/li/ul/li[3]/a').click() #indo para a aba de downloads de ultimas faturas


        #caso especifico de login com mais de uma fatura
        if login == :
            verificando_conta() #passsando o numero da conta para selecionar ela dentro do login
            sleep(2)
            verificando_data()
            if fatura == True: #se a primeira fatura não estiver disponivel ele vai selecionar a lista de faturas e ir para a segunda opção
                navegador.find_element(By.XPATH, Portal_map['buttons']["download"]["xpath"]).click() #clicando em ok para download
                sleep(10)
                arquivo_recente()
                extração()
                upload_drive(f"{pasta_downloads}/{ultimo_arquivo[1]}")
                insert_MySQL()
                remover_arquivo()
            verificando_conta()
            sleep(1)
            verificando_data()
            if fatura == True: #se a segunda fatura não esiver disponivel ele vai selecionar a lista de faturas e ir para a primeira opção
                navegador.find_element(By.XPATH, Portal_map['buttons']["download"]["xpath"]).click() #clicando em ok para download
                sleep(10)
                arquivo_recente()
                extração()
                upload_drive(f"{pasta_downloads}/{ultimo_arquivo[1]}")
                insert_MySQL()
                remover_arquivo()
                verificando_conta()
                sleep(2)
            reinciando()

        #caso especifico de login com mais de uma fatura
        elif login == :
            verificando_conta()
            sleep(1)
            verificando_data()
            if fatura == True:
                navegador.find_element(By.XPATH, Portal_map['buttons']["download"]["xpath"]).click() #clicando em ok para download
                sleep(10)
                arquivo_recente()
                extração()
                upload_drive(f"{pasta_downloads}/{ultimo_arquivo[1]}")
                insert_MySQL()
                remover_arquivo()
                sleep(2)
            verificando_conta()
            sleep(1)
            verificando_data()
            if fatura == True:
                navegador.find_element(By.XPATH, Portal_map['buttons']["download"]["xpath"]).click() #clicando em ok para download
                sleep(10)
                arquivo_recente()
                extração()
                upload_drive(f"{pasta_downloads}/{ultimo_arquivo[1]}")
                insert_MySQL()
                remover_arquivo()
                sleep(2)
            verificando_conta()
            sleep(1)
            verificando_data()
            if fatura == True:
                navegador.find_element(By.XPATH, Portal_map['buttons']["download"]["xpath"]).click() #clicando em ok para download
                sleep(10)
                arquivo_recente()
                extração()
                upload_drive(f"{pasta_downloads}/{ultimo_arquivo[1]}")
                insert_MySQL()
                remover_arquivo()
                verificando_conta()
                sleep(2)
            reinciando()

        #caso especifico de login com mais de uma fatura
        elif login == :
            verificando_conta()
            sleep(2)
            verificando_data()
            if fatura == True:
                navegador.find_element(By.XPATH, Portal_map['buttons']["download"]["xpath"]).click() #clicando em ok para download
                sleep(10)
                arquivo_recente()
                extração()
                upload_drive(f"{pasta_downloads}/{ultimo_arquivo[1]}")
                insert_MySQL()
                remover_arquivo()
                sleep(2)
            verificando_conta()
            sleep(2)
            verificando_data()
            if fatura == True:
                navegador.find_element(By.XPATH, Portal_map['buttons']["download"]["xpath"]).click() #clicando em ok para download
                sleep(10)
                arquivo_recente()
                extração()
                upload_drive(f"{pasta_downloads}/{ultimo_arquivo[1]}")
                insert_MySQL()
                remover_arquivo()
            verificando_conta()
            sleep(2)
            reinciando()

        #caso padrao de apenas uma fatura no login
        else:
            verificando_data() #chamando a função para verificar a data das faturas disponiveis
            if fatura == False: #se não estiver nenhuma do mes e ano atual disponivel vai voltar a controlar a aba de login e pular essa conta
                reinciando()
            else: #se tiver uma disponivel ele confirma cliando em ok e depois volta para a aba de login
                try:
                    navegador.find_element(By.XPATH, Portal_map['buttons']["download"]["xpath"]).click() #clicando em ok para download
                    sleep(10)
                    arquivo_recente()
                    extração()
                    upload_drive(f"{pasta_downloads}/{ultimo_arquivo[1]}")
                    insert_MySQL()
                    remover_arquivo()
                    reinciando()
                except:
                    reinciando()

navegador.quit() #fechando o navegador depois de verificar todos os logins
db.close() #fechando o banco de dados 