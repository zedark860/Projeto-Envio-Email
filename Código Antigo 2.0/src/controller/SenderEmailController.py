import random
import time
import smtplib
import os
import sys
import pandas as pd
from PyQt5.QtCore import pyqtBoundSignal
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from src.controller.DataUserController import DataUserController
from src.model.SenderEmailModel import SenderEmailModel
from src.log.Log import Log

class SenderEmailController(SenderEmailModel):
    
    # Não está sendo utilizado
    @staticmethod
    def resource_path(relative_path):
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(
            os.path.abspath(__file__)))
    
        return os.path.join(base_path, relative_path)
    
    
    @staticmethod
    def clean_text(text: str) -> str:
        return text.replace('\xa0', ' ').encode('ascii', 'ignore').decode('ascii')
    
    
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
            whatsapp_redirect_number=whatsapp_redirect_number,
            list_columns=['E-mail', 'Nome', 'Protocolo', 'Produto', 'STATUS']
        )

    
    def read_emails(self) -> tuple[list[str], pd.DataFrame]:
        df_emails = pd.read_excel(self.spreadsheet_path)
        
        required_columns: list[str] = self.list_columns
        
        try:
            df_emails.columns = required_columns
            
            df_no_duplicates: pd.DataFrame = df_emails.drop_duplicates(
                subset=self.list_columns, keep='first')

            df_no_duplicates.to_excel(self.spreadsheet_path, index=False)

            return df_no_duplicates[self.list_columns[0]].tolist(), df_no_duplicates
        except KeyError as error:
            missing_columns: list[str] = list(set(required_columns) - set(df_emails.columns))
            raise KeyError(f"Colunas ausentes na planilha: {', '.join(missing_columns)}") from error
    
    
    def get_name_product(self, df_emails: pd.DataFrame, destination_email: str) -> tuple[str, str, str]:
        if destination_email in df_emails[self.list_columns[0]].values:
            name: str = df_emails.loc[df_emails[self.list_columns[0]] == destination_email, self.list_columns[1]].iloc[0]
            product: str = df_emails.loc[df_emails[self.list_columns[0]] == destination_email, self.list_columns[2]].iloc[0]
            protocol: str = df_emails.loc[df_emails[self.list_columns[0]] == destination_email, self.list_columns[3]].iloc[0]
            
            return name, product, protocol

        name, product, protocol = '', '', ''
        return name, product, protocol
    

    def update_message(self, name: str, product: str, protocol: str) -> tuple[str, str, str]:
        aux_dict: dict[str, str] = {
            "{nome}": str(name),
            "{produto}": str(product),
            "{protocolo}": str(protocol),
        }
        
        message_copy: str = self.email_message
        title_copy: str = self.email_title
        
        if name is not None or product is not None or protocol is not None:
            for key, value in aux_dict.items():
                message_copy = message_copy.replace(key, value)
                title_copy = title_copy.replace(key, value)
                
            link_whatsapp_copy: str = f"https://api.whatsapp.com/send?phone={self.whatsapp_redirect_number}&text=Ol%C3%A1!%20Vim%20pelo%20email%20e%20gostaria%20de%20tirar%20algumas%20d%C3%BAvidas"

            return message_copy, title_copy, link_whatsapp_copy
    
    
    def type_verify_email(self) -> str:
        email_and_server: dict[str, str] = {
            "gmail.com": "smtp.gmail.com",
            "webcertificados.com.br": "smtp.gmail.com",
            "webcertificado.com.br": "smtp.gmail.com",
            "outlook.com": "smtp.office365.com",
            "hotmail.com": "smtp.live.com",
            "aol.com": "smtp.aol.com",
            "icloud.com": "smtp.mail.me.com",
        }
        
        domain: str = self.data_user.email.split('@')[-1]

        return email_and_server.get(domain, "smtp.gmail.com")  
    
    
    def config_send_email(self, destination_email: str, copy_body_html: str) -> None:
        server_smtp = self.type_verify_email()
        port_smtp = 587
        
        msg = MIMEMultipart('alternative')
        msg['Subject'] = self.clean_text(self.email_subject)
        msg['From'] = self.clean_text(self.data_user.email)
        msg['To'] = self.clean_text(destination_email)
        msg.attach(MIMEText(self.clean_text(copy_body_html), 'html', 'utf-8'))
        
        with smtplib.SMTP(server_smtp, port_smtp) as server:
            server.starttls()
            server.login(msg['From'], self.data_user.app_password)
            server.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
            

    def replace_email_html(self, body_html_original: str, message_copy: str, title_copy: str, link_whatsapp_copy: str) -> str:
        copy_body_html: str = body_html_original
        copy_body_html = copy_body_html.replace('{titulo_html}', title_copy)
        copy_body_html = copy_body_html.replace('{mensagem_html}', message_copy)
        copy_body_html = copy_body_html.replace('{link_whatsapp}', link_whatsapp_copy)
        
        return copy_body_html
    
    
    def write_on_console_and_txt(self, destination_email: str, destination_emails: list[str], validator: bool, i: int, start_index: int=None, error=None, interval_send_email=None) -> str:
        if not validator:
            message: str = f"Erro ao enviar email para {destination_email} às {datetime.now().strftime('%d/%m/%Y-%H:%M:%S')}. Erro: {error}\n"
            Log.write_error(self.data_user.email, message)
            return message
        
        message: str = f"Email enviado para {destination_email} às {datetime.now().strftime('%d/%m/%Y-%H:%M:%S')}\nTotal de emails enviados: {i + 1} de {len(destination_emails[start_index:])} " + f"Intervalo de envio: {interval_send_email} segundos\n"
        Log.write_success(self.data_user.email, message)
        return message
    
    
    def send_emails(self, log_signal: pyqtBoundSignal) -> None:
        body_html_original: str = SenderEmailController.load_html(self.html_path)
        destination_emails, df_emails = self.read_emails()
        
        last_sent_index: int = df_emails[df_emails[self.list_columns[4]].isin(['ENVIADO', 'ERRO'])].index.max()
        start_index: int = (last_sent_index + 1) if pd.notna(last_sent_index) else 0
        
        for i, destination_email in enumerate(destination_emails[start_index:]):
            if (i + 1) > self.email_send_quantity:
                break
            
            try:
                name, product, protocol = self.get_name_product(df_emails, destination_email)
                message_copy, title_copy, link_whatsapp_copy = self.update_message(name, product, protocol)
                copy_body_html: str = self.replace_email_html(body_html_original, message_copy, title_copy, link_whatsapp_copy)
                self.config_send_email(destination_email, copy_body_html)
                random_interval: int = SenderEmailController.wait_random_send_email(self.send_interval)
            
                df_emails.loc[i, self.list_columns[4]] = 'ENVIADO'
                df_emails.to_excel(self.spreadsheet_path, index=False)
                message_success: str = self.write_on_console_and_txt(destination_email, destination_emails, True, i, start_index, None, random_interval)
                log_signal.emit(message_success)
                time.sleep(random_interval)
            except Exception as error:
                df_emails.loc[i, self.list_columns[4]] = 'ERRO'
                df_emails.to_excel(self.spreadsheet_path, index=False)
                message_error: str = self.write_on_console_and_txt(destination_email, destination_emails, False, i, None, error, None)
                log_signal.emit(message_error)
                continue