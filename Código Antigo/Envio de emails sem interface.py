import pandas as pd
import smtplib
import time
import os
import random
from pathlib import Path
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime


def carregar_corpo_email():
    with open("corpo_email.html", "r") as file:
        return file.read()


def escrever_envio(mensagem):
    arquivo_enviados = Path('enviados.txt')
    # Cria o arquivo se ele não existir
    if not arquivo_enviados.exists():
        arquivo_enviados.touch()
        
    # Escreve os logs no arquivo
    with open("enviados.txt", "a") as log_file:
        log_file.write(f"{mensagem}\n")
            
            
def escrever_erro(mensagem):
    # Cria o arquivo se ele não existir
    arquivo_erros = Path('erros.txt')
    if not arquivo_erros.exists():
        arquivo_erros.touch()
            
    # Escreve os logs no arquivo
    with open("erros.txt", "a") as log_file:
        log_file.write(f"{mensagem}\n")
        
        
def limpar_tela():
    # Verifica o sistema operacional para determinar o comando apropriado para limpar a tela, espera 3 segundos antes disso
    time.sleep(3)
    os.system('cls' if os.name == "nt" else "clear")


def ler_emails(planilha_path):
        # Lê os e-mails da planilha Excel
        df_emails = pd.read_excel(planilha_path)
        if 'E-mail' in df_emails.columns:
            tamanho_antes = len(df_emails)

            # Remove duplicatas mantendo apenas a primeira ocorrência, considerando as colunas 'E-mail', 'Nome' e 'Produto'
            df_sem_duplicatas = df_emails.drop_duplicates(
                subset=['E-mail', 'Nome', 'Produto'], keep='first')

            tamanho_depois = len(df_sem_duplicatas)

            if tamanho_antes != tamanho_depois:
                print(f"\nAlgumas entradas duplicadas foram removidas! Agora são: {tamanho_depois}\n")
            else:
                print('\nNão há entradas duplicadas na planilha!\n')

            # Salva a planilha sem duplicatas
            df_sem_duplicatas.to_excel('Enviar E-mails.xlsx', index=False)

            # Retorna a lista de e-mails e o dataframe
            return df_sem_duplicatas['E-mail'].tolist(), df_sem_duplicatas
        else:
            print('\nA planilha não contém a coluna "E-mail"!\n')
            return [], pd.DataFrame()
        

def obter_nome_produto(df_emails, email_destino):
    # Verifica se o e-mail de destino existe na planilha
    if email_destino in df_emails['E-mail'].values:
        # Encontre o nome e o produto correspondentes ao e-mail de destino na planilha
        nome = df_emails.loc[df_emails['E-mail'] == email_destino, 'Nome'].iloc[0]
        produto = df_emails.loc[df_emails['E-mail'] == email_destino, 'Produto'].iloc[0]
        return nome, produto
    else:
        # Retorna None se o e-mail de destino não for encontrado na planilha
        return None, None
    

def atualizar_mensagem(df_emails, email_destino, titulo_html, mensagem_html, numero_whatsapp):
    # Obtém o nome e o produto correspondentes ao e-mail de destino
    nome, produto = obter_nome_produto(df_emails,email_destino)
   
    # Verifica se nome e produto foram encontrados
    if nome is not None:
        mensagem_html = mensagem_html.replace('{nome}', str(nome))
        titulo_html = titulo_html.replace('{nome}', str(nome))
    if produto is not None:
        mensagem_html = mensagem_html.replace('{produto}', str(produto))
        titulo_html = titulo_html.replace('{produto}', str(produto))
        
    # Substitui o marcador de posição do número do WhatsApp na mensagem HTML
    mensagem_html = mensagem_html.replace('{numero}', str(numero_whatsapp))
        
    # Substitui o marcador de posição do número do WhatsApp no link de redirecionamento
    link_whatsapp = f"https://api.whatsapp.com/send?phone={numero_whatsapp}&text=Ol%C3%A1!%20Vim%20pelo%20site%20e%20gostaria%20de%20tirar%20algumas%20d%C3%BAvidas"
            
    # Retorna a mensagem original se nome ou produto não forem encontrados
    return titulo_html, mensagem_html, link_whatsapp
        

def enviar_email(destinatarios, planilha_path):
    # Configurações para o servidor SMTP do Gmail
    servidor_smtp = 'smtp.gmail.com'
    porta_smtp = 587
    
    # Configurando email do remetente e senha
    email_remetente = 'fin.envios.3@gmail.com'
    senha = 'zljt edek hejb ohjv'
    
    titulo_html = 'teste {nome}'
    assunto = 'testando amigo'
    mensagem_html = 'teste opa {nome}'
    redirecionar_whatsapp = '553291376630'

    corpo_email_original = carregar_corpo_email()
    
    df_emails = pd.read_excel(planilha_path)
    status_coluna = 'STATUS'

    for i, email_destino in enumerate(destinatarios):
        corpo_email = corpo_email_original
            
        # Atualiza a mensagem de e-mail com o nome e o produto correspondentes ao destinatário atual
        novo_titulo, nova_mensagem, link_whatsapp = atualizar_mensagem(df_emails, email_destino, titulo_html, mensagem_html, redirecionar_whatsapp)
            
        # Substitui os marcadores de posição no corpo do e-mail pelos valores correspondentes
        corpo_email = corpo_email.replace("{titulo_html}", novo_titulo)
        corpo_email = corpo_email.replace("{mensagem_html}", nova_mensagem)
            
        # Substitui o link de redirecionamento no corpo do e-mail
        corpo_email = corpo_email.replace("{link_whatsapp}", link_whatsapp)
        
        corpo_mensagem = MIMEText(corpo_email, 'html')

        # Configuração do e-mail
        msg = MIMEMultipart()

        # Configura o destinatário atual
        msg.attach(corpo_mensagem)
        msg['To'] = email_destino
        msg['Subject'] = assunto  # Assunto do email
        msg['From'] = email_remetente # Substitua pelo seu endereço de e-mail do Outlook
        password = senha  # Substitua pela sua senha ou pela senha de aplicativo

        hora_atual = datetime.now()
        
        # Randomiza os intervalos de tempo
        intervalo_envio = random.randint(30, 60)

        try:
            # Inicia a conexão SMTP e envia o e-mail
            with smtplib.SMTP(servidor_smtp, porta_smtp) as s:
                s.starttls()
                s.login(msg['From'], password)
                s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
            
            # Atualiza o status para 'SUCESSO'    
            df_emails.loc[df_emails['E-mail'] == email_destino, status_coluna] = 'SUCESSO'
            df_emails.to_excel('Enviar E-mails.xlsx', index=False)
            
            mensagem_de_envio = f'\nE-mail enviado para {email_destino} dia {hora_atual.strftime("%d")} às {hora_atual.strftime("%H:%M:%S")}, com {intervalo_envio}s de intervalo.\n'
            escrever_envio(mensagem_de_envio)
            
            limpar_tela()
            
            print(mensagem_de_envio + f'\nEnviados: {i+1}')
        
            
        except Exception as e:
            # Se ocorrer um erro, imprima a mensagem de erro e atualize o status para 'ERRO'
            df_emails.loc[df_emails['E-mail'] == email_destino, status_coluna] = 'ERRO'
            df_emails.to_excel('Enviar E-mails.xlsx', index=False)
            
            mensagem_de_erro = f'\nErro ao enviar e-mail para {email_destino} dia {hora_atual.strftime("%d")} ás {hora_atual.strftime("%H:%M:%S")}: {e}\n'
            escrever_erro(mensagem_de_erro)
            
            limpar_tela()
            
            print(mensagem_de_erro)
            
        time.sleep(intervalo_envio)


limpar_tela()

# Caminho da planilha com os e-mails
planilha_path = Path('Enviar E-mails.xlsx')

# Lê os e-mails da planilha
destinatarios, _ = ler_emails(planilha_path)

print('Iniciando envio de e-mails!')

# Envia e-mails para os destinatários
enviar_email(destinatarios, planilha_path)

print('\nTodos os e-mails enviados com sucesso!\n')