import re #procura e compilação de padrão
import os
import smtplib
import openpyxl
import email
from email.message import EmailMessage
from time import sleep
import pandas as pd
import pywhatkit as kit

class ToDo:
   
    def iniciar(self):
        # self.lista_tarefa = []
        # self.lista_datas = []
        # self.email_destino()
        # self.menu()
        # self.visualizar()
        # self.criar_planilha()
        # sleep(1)
        # self.enviar_email()
        self.enviar_whatsapp()
    
    def email_destino(self):
         while True:
            self.email = str(input('Email de destino: ')).lower()
            padrao_email = re.search(
            "^[a-z0-9._]+@[a-z0-9]+.[a-z](.[a-z]+)?$", self.email #texto/numero._@ texto/numero . texto (.texto)
            )

            if padrao_email: # false ' ' ou none
                print('Email Valido')
                break
            else:
                print("Email inválido, tente outro...")
    
    def menu(self):
        while True:
            menu_principal = int(input('''
            Menu Principal
            [1] CADASTRAR
            [2] VIZUALIZAR
            [3] SAIR
            Opção: '''))
            match menu_principal:
                case 1:
                    self.cadastrar()
                case 2:
                    self.visualizar()
                case 3:
                    break
                case _:
                    print('\nOpção inválida')
    
    def cadastrar(self):
        while True:
            self.tarefa = str(input('Digite uma letra ou [s] para sair: ')).capitalize()
            
            if self.tarefa == 'S': 
                break
            else:
                self.data = str(input('Data: '))
                self.lista_tarefa.append(self.tarefa)
                self.lista_datas.append(self.data)
                try:
                    with open('./src/Tarefas/historico_tarefas.txt', 'a', encoding='utf8') as arquivo:
                        arquivo.write(f'{self.tarefa}\n')
                except FileNotFoundError as e:
                    print(f'Erro: {e}')
    
    def visualizar(self):
        try:
            with open('./src/Tarefas/historico_tarefas.txt', 'r', encoding='utf8') as arquivo:
                print(arquivo.read())
        except FileExistsError as e:
            print(f'Erro: {e}')

    def criar_planilha(self):
        if len(self.lista_tarefa) > 0:
            try:
                df = pd.DataFrame({
                    "Tarefas": self.lista_tarefa,
                    'data': self.lista_datas
                    })
                
                self.nome_arquivo = str(input('Nome do arquivo: ')).lower()

                if self.nome_arquivo[-5:] == '.xlsx':
                    df.to_excel(f'./src/Tarefas/{self.nome_arquivo}', index=False) #isso aqui ele tira os numeros sequenciais
                else:
                    df.to_excel(f'./src/Tarefas/{self.nome_arquivo}.xlsx', index=False)
                print('\nPlanilha criada com sucesso!')
            
            except Exception as e:
                print(f'Erro: {e}')
        else:
            print('\nLista de tarefas vazia.')

    def enviar_email(self):
        endereço = '...@gmail.com'

        with open('./src/Senha/Senha.txt', 'r', encoding='utf8') as arquivo:
            s = arquivo.readlines()
        senha = s[0]

        msg = EmailMessage()
        msg['From'] = endereço
        msg['To'] = self.email
        msg['Subject'] = 'Oooo sé, chegou a planilha!'
        msg.set_content(
            'Planilha em anexo.'
        )
        arquivos = [f'./src/Tarefas/{self.nome_arquivo}.xlsx']

        for arquivo in arquivos:
            with open(arquivo, 'rb') as arq: #rb leitura programação (read binario)
                dados = arq.read()
                nome_arquivo = arq.name

            msg.add_attachment(
                dados,
                maintype = 'application',
                subtype = 'octet-stream',
                filename = nome_arquivo
            )
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465) #protocolo de internet para acessar e-mail (dentro dos parentes tem o padrão para o gmail)   
        server.login(endereço, senha, initial_response_ok=True)
        server.send_message(msg)
        print("Email enviado com sucesso.")
            
    def enviar_whatsapp(self):
        try:
            numero_destino = '+5511...'
            mensagem = 'Ooooou mandei a planilha lá!!!'
            kit.sendwhatmsg_instantly(numero_destino, mensagem, wait_time=60)
            print('\nWhats enviado')
        except Exception as e:
            print(f'Erro: {e}')













































































start = ToDo()
start.iniciar()
