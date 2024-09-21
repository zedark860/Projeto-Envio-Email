import random
import time
import smtplib
import os
import pandas as pd
from PyQt5.QtCore import pyqtBoundSignal
from datetime import datetime
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from src.controller.DataUserController import DataUserController
from src.model.SenderEmailModel import SenderEmailModel
from src.log.Log import Log

class SenderEmailController(SenderEmailModel):
    
    @staticmethod
    def load_attachments(list_attachments: list[str]) -> list[MIMEApplication] | None:
        attachments_in_memory: list[MIMEApplication] = []
        
        if not list_attachments:
            return attachments_in_memory
        
        try:
            for attachment_name in list_attachments:
                attachment_path: str = os.getcwd() + f"\\anexos\\{attachment_name}"
                with open(attachment_path, 'rb') as file:
                        part: MIMEApplication = MIMEApplication(file.read(), Name=os.path.basename(attachment_name))
                        part['Content-Disposition'] = f'attachment; filename="{os.path.basename(attachment_name)}"'
                        attachments_in_memory.append(part)
                        
            return attachments_in_memory
        except FileNotFoundError as fnfe:
            raise FileNotFoundError("Arquivo não encontrado.") from fnfe
        except PermissionError as pe:
            raise PermissionError("Permissão negada para abrir o arquivo. Verifique se o arquivo está em uso.") from pe
        except IsADirectoryError as ide:
            raise IsADirectoryError("O caminho é um diretório, não um arquivo.") from ide
        except OSError as ose:
            raise OSError(f"Ocorreu um erro ao tentar abrir o arquivo: {ose.strerror}") from ose
    
     
    @staticmethod
    def load_html(corpo_email_path: str):
        with open(corpo_email_path, "r", encoding="utf-8") as file:
            return file.read()
        
    
    @staticmethod
    def wait_random_send_email(send_interval: int) -> int:
        return random.randint(send_interval - 5, send_interval + 5)
    
    
    def __init__(self, data_user: DataUserController, spreadsheet_path: str, html_path: str, email_send_quantity: int, send_interval: int, email_subject: str, email_title: str, email_message: str, whatsapp_redirect_number: str):
        super().__init__(
            data_user=data_user,
            spreadsheet_path=spreadsheet_path,
            html_path=html_path,
            email_send_quantity=email_send_quantity, 
            send_interval=send_interval, 
            email_subject=email_subject, 
            email_title=email_title, 
            email_message=email_message, 
            whatsapp_redirect_number=whatsapp_redirect_number
        )

    
    def read_emails(self) -> tuple[list[str], pd.DataFrame]:
        df_emails: pd.DataFrame = pd.read_excel(self.spreadsheet_path)
        
        required_columns: list[str] = ['E-mail', 'Nome', 'Protocolo', 'Produto', 'Nome_Arquivos_Anexo']
        
        try:
            df_no_duplicates: pd.DataFrame = df_emails.drop_duplicates(
                subset=['E-mail', 'Nome', 'Protocolo', 'Produto', 'Nome_Arquivos_Anexo'], keep='first')

            df_no_duplicates.to_excel(self.spreadsheet_path, index=False)

            return df_no_duplicates['E-mail'].tolist(), df_no_duplicates
        except KeyError as error:
            missing_columns: list[str] = list(set(required_columns) - set(df_emails.columns))
            raise KeyError(f"Colunas ausentes na planilha: {', '.join(missing_columns)}") from error
    
    
    def get_name_product(self, df_emails: pd.DataFrame, destination_email: str) -> tuple[str, str, str]:
        if destination_email in df_emails['E-mail'].values:
            name: str = df_emails.loc[df_emails['E-mail'] == destination_email, 'Nome'].iloc[0]
            product: str = df_emails.loc[df_emails['E-mail'] == destination_email, 'Produto'].iloc[0]
            protocol: str = df_emails.loc[df_emails['E-mail'] == destination_email, 'Protocolo'].iloc[0]
            
            return name, product, protocol

        return '', '', ''
    
    
    def get_name_attachment(self, df_emails: pd.DataFrame, destination_email: str) -> list[str] | list:
        if destination_email in df_emails['E-mail'].values:
            attachments: str = df_emails.loc[df_emails['E-mail'] == destination_email, 'Nome_Arquivos_Anexo'].iloc[0]
            if pd.isna(attachments) or not attachments:
                return list()

            attachments = attachments.replace(' ', '')
            arquives_name_list: list[str] = attachments.split(';')
            
            return arquives_name_list
        return list()
    

    def update_message(self, name: str, product: str, protocol: str) -> str:
        aux_dict: dict[str, str] = {
            "{nome}": str(name),
            "{produto}": str(product),
            "{protocolo}": str(protocol),
        }
        
        if all(value is not None for value in aux_dict.values()):
            for key, value in aux_dict.items():
                self.email_message = self.email_message.replace(key, value)
                self.email_title = self.email_title.replace(key, value)

        return f"https://api.whatsapp.com/send?phone={self.whatsapp_redirect_number}&text=Ol%C3%A1!%20Vim%20pelo%20email%20e%20gostaria%20de%20tirar%20algumas%20d%C3%BAvidas"
    
    
    def type_verify_email(self) -> str:
        email_and_server: dict[str, str] = {
            "gmail.com": "smtp.gmail.com",
            "webcertificados.com.br": "stmp.gmail.com",
            "outlook.com": "smtp.office365.com",
            "hotmail.com": "smtp.live.com",
            "aol.com": "smtp.aol.com",
            "icloud.com": "smtp.mail.me.com",
        }
        
        domain: str = self.data_user.email.split('@')[-1]

        return email_and_server.get(domain, "Servidor não encontrado")    
    
    def config_send_email(self, destination_email: str, copy_body_html: str, attachments: list[MIMEApplication] | None) -> None:
        server_smtp: str = self.type_verify_email()
        port_smtp: int = 587
        
        msg = MIMEMultipart('alternative')
        msg['Subject'] = self.email_subject
        msg['From'] = self.data_user.email
        msg['To'] = destination_email
        msg.attach(MIMEText(copy_body_html, 'html'))
        
        if attachments:
            for attachment in attachments:
                msg.attach(attachment)
        
        with smtplib.SMTP(server_smtp, port_smtp) as server:
            server.starttls()
            server.login(msg['From'], self.data_user.app_password)
            server.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
            

    def replace_email_html(self, body_html_original: str, link_whatsapp: str) -> str:
        copy_body_html: str = body_html_original
        copy_body_html = copy_body_html.replace('{titulo_html}', self.email_title)
        copy_body_html = copy_body_html.replace('{mensagem_html}', self.email_message)
        copy_body_html = copy_body_html.replace('{link_whatsapp}', link_whatsapp)
        
        return copy_body_html
    
    
    def write_on_console_and_txt(self, destination_email: str, destination_emails: list[str], validator: bool, i: int, error=None, interval_send_email=None) -> str:
        message: str = ""
        if not validator:
            message = f"Erro ao enviar email para {destination_email} às {datetime.now().strftime('%H:%M:%S')}. Erro: {error}\n"
            Log.write_error(self.data_user.email, message)
            return message
        
        message = f"Email enviado para {destination_email} às {datetime.now().strftime('%H:%M:%S')}\nTotal de emails enviados: {i + 1} de {len(destination_emails)} " + f"Intervalo de envio: {interval_send_email} segundos\n"
        Log.write_success(self.data_user.email, message)
        return message
    
    
    def send_emails(self, log_signal: pyqtBoundSignal) -> None:
        body_html_original: str = SenderEmailController.load_html(self.html_path)
        destination_emails, df_emails = self.read_emails()
        
        
        for i, destination_email in enumerate(destination_emails):
            if i + 1 > self.email_send_quantity:
                return
            
            name, product, protocol = self.get_name_product(df_emails, destination_email)
            attachments_list_in_memory: list[MIMEApplication] | None = SenderEmailController.load_attachments(
                                    self.get_name_attachment(df_emails, destination_email))
            link_whatsapp: str = self.update_message(name, product, protocol)
            copy_body_html: str = self.replace_email_html(body_html_original, link_whatsapp)
            self.config_send_email(destination_email, copy_body_html, attachments_list_in_memory)
            random_interval: int = SenderEmailController.wait_random_send_email(self.send_interval)
            
            try:
                df_emails.loc[df_emails['E-mail'] == destination_email, 'STATUS'] = 'ENVIADO'
                df_emails.to_excel(self.spreadsheet_path, index=False)
                message_success: str = self.write_on_console_and_txt(destination_email, destination_emails, True, i, None, random_interval)
                log_signal.emit(message_success)
                time.sleep(random_interval)   
            except Exception as error:
                df_emails.loc[df_emails['E-mail'] == destination_email, 'STATUS'] = 'ERRO'
                df_emails.to_excel(self.spreadsheet_path, index=False)
                message_error: str = self.write_on_console_and_txt(destination_email, destination_emails, False, i, error, None)
                log_signal.emit(message_error)