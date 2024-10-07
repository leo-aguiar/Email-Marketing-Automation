import os
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from dotenv import load_dotenv

# Carregar variáveis de ambiente de um arquivo .env
load_dotenv()

SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587
EMAIL_USER = os.getenv('EMAIL_USER')
EMAIL_PASS = os.getenv('EMAIL_PASS')

def enviar_email(destinatario, nome_destinatario):
    try:
        msg = MIMEMultipart()
        msg['Subject'] = 'Novidades da nossa Equipe :)'
        msg['From'] = EMAIL_USER
        msg['To'] = destinatario

        # Corpo do email
        mensagem = f'''Olá, {nome_destinatario}!
        
        Esta é uma mensagem de marketing automática. Confira nossas novidades =)

        Atenciosamente,
        Leonardo'''
        
        msg.attach(MIMEText(mensagem, 'plain'))

        # Enviar email
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_USER, EMAIL_PASS)
            server.sendmail(msg['From'], msg['To'], msg.as_string())
            print(f'Email enviado com sucesso para {nome_destinatario} ({destinatario})')

    except Exception as e:
        print(f'Erro ao enviar email para {nome_destinatario} ({destinatario}): {e}')

def enviar_emails_marketing():
    # Carregar a lista de clientes
    try:
        clientes = pd.read_excel('./clientes.xlsx')
    except FileNotFoundError:
        print('Erro: Arquivo "clientes.xlsx" não encontrado.')
        return
    except Exception as e:
        print(f'Erro ao carregar arquivo: {e}')
        return

    # Enviar emails para todos os clientes
    for index, cliente in clientes.iterrows():
        enviar_email(cliente['email'], cliente['nome'])

if __name__ == "__main__":
    enviar_emails_marketing()
