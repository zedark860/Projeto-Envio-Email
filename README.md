# Projeto Envio Email

Bem-vindo ao Projeto Envio Email! Este é um projeto de automação desenvolvido em Python para facilitar o envio de emails.

## Índice

- [Introdução](#introdução)
- [Funcionalidades](#funcionalidades)
- [Tecnologias Utilizadas](#tecnologias-utilizadas)
- [Instalação](#instalação)
- [Como Usar](#como-usar)
- [Exemplo de Uso](#exemplo-de-uso)
- [Licença](#licença)

## Introdução

Este projeto foi criado para automatizar o envio de emails para múltiplos destinatários com a possibilidade de personalização de mensagens. Utiliza Python e bibliotecas específicas para gerenciar o envio de emails de maneira eficiente e segura.

## Funcionalidades

- Envio de emails para uma lista de destinatários.
- Personalização do assunto e do corpo do email.
- Suporte para anexos.
- Registro de histórico de envio.

## Tecnologias Utilizadas

- **Python**: Linguagem de programação principal.
- **Biblioteca de Envio de Emails**: [smtplib](https://docs.python.org/3/library/smtplib.html).
- **Outras dependências**: Pandas, Email, Webbrowser, Json, Atexit, Random, Time, Os, Datetime, Pyqt5

## Instalação

1. Clone o repositório:
    ```bash
    git clone https://github.com/seu-usuario/projeto-envio-email.git
    cd projeto-envio-email
    ```

2. Crie e ative um ambiente virtual (recomendado):
    ```bash
    python -m venv venv
    source .venv/bin/activate  # No Windows, use `.venv\Scripts\activate`
    ```

3. Instale as dependências listas na parte de tecnologias utilizadas:

## Como Usar

1. Configuração:
   - Configure as credenciais do servidor SMTP no arquivo `config.py`.
   - Configure os destinatários e o conteúdo do email no arquivo `emails.py`.

2. Execução:
   - Execute o script principal do projeto com interface:
     ```bash
     python Envio de emails com interface.py
     ```
   -Ou execute o script principal do projeto sem interface:
     ```bash
     python Envio de emails sem interface.py
     ```

## Exemplo de Uso

Aqui está um exemplo básico de como usar este projeto para enviar um email:

```python
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from config import EMAIL_ADDRESS, EMAIL_PASSWORD

def enviar_email(destinatario, assunto, corpo):
    msg = MIMEMultipart()
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = destinatario
    msg['Subject'] = assunto

    msg.attach(MIMEText(corpo, 'plain'))

    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.starttls()
        server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        server.send_message(msg)
        print(f'Email enviado para {destinatario}')

# Exemplo de uso
destinatarios = ['exemplo1@dominio.com', 'exemplo2@dominio.com']
assunto = "Teste de Envio de Email"
corpo = "Este é um email de teste."

for destinatario in destinatarios:
    enviar_email(destinatario, assunto, corpo)
```

## Licença

Este projeto está licenciado sob a licença MIT. Veja o arquivo [LICENSE](link) para mais detalhes.

Agredeço por sua visita!

