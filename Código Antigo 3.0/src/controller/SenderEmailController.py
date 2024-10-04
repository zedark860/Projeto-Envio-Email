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
    
    @staticmethod
    def resource_path(relative_path):
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(
            os.path.abspath(__file__)))
    
        return os.path.join(base_path, relative_path)
    
    @staticmethod
    def load_html():
        with open(SenderEmailController.resource_path("corpo_email.html"), "r", encoding="utf-8") as file:
            return file.read()
        
    
    @staticmethod
    def wait_random_send_email(send_interval: int) -> int:
        return random.randint(send_interval - 5, send_interval + 5)
    
    
    def __init__(self, data_user: DataUserController, spreadsheet_path: str, email_send_quantity: int, send_interval: int, email_subject: str, email_title: str, email_message: str, whatsapp_redirect_number: str):
        super().__init__(
            data_user=data_user,
            spreadsheet_path=spreadsheet_path, 
            email_send_quantity=email_send_quantity, 
            send_interval=send_interval, 
            email_subject=email_subject, 
            email_title=email_title, 
            email_message=email_message, 
            whatsapp_redirect_number=whatsapp_redirect_number
        )

    
    def read_emails(self) -> tuple[list[str], pd.DataFrame]:
        df_emails = pd.read_excel(self.spreadsheet_path)
        
        if 'E-mail' in df_emails.columns:
            df_no_duplicates: pd.DataFrame = df_emails.drop_duplicates(
                subset=['E-mail', 'Nome', 'Protocolo', 'Produto'], keep='first')

            df_no_duplicates.to_excel(self.spreadsheet_path, index=False)

            return df_no_duplicates['E-mail'].tolist(), df_no_duplicates

        return [], pd.DataFrame()
    
    
    def get_name_product(self, df_emails: pd.DataFrame, destination_email: str) -> tuple[str, str, str]:
        if destination_email in df_emails['E-mail'].values:
            name: str = df_emails.loc[df_emails['E-mail'] == destination_email, 'Nome'].iloc[0]
            product: str = df_emails.loc[df_emails['E-mail'] == destination_email, 'Produto'].iloc[0]
            protocol: str = df_emails.loc[df_emails['E-mail'] == destination_email, 'Protocolo'].iloc[0]
            
            return name, product, protocol

        return '', '', ''
    

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
    
    def config_send_email(self, destination_email: str, copy_body_html: str) -> None:
        server_smtp = 'smtp.gmail.com'
        port_smtp = 587
        
        msg = MIMEMultipart('alternative')
        msg['Subject'] = self.email_subject
        msg['From'] = self.data_user.email
        msg['To'] = destination_email
        msg.attach(MIMEText(copy_body_html, 'html'))
        
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
    
    
    def write_on_console_and_txt(self, destination_email: str, destination_emails: list[str], validator: bool, i: int, error=None, interval_send_email=None) -> str:
        if not validator:
            message = f"Erro ao enviar email para {destination_email} às {datetime.now().strftime('%H:%M:%S')}. Erro: {error}\n"
            Log.write_error(self.data_user.email, message)
            return message
        
        message = f"Email enviado para {destination_email} às {datetime.now().strftime('%H:%M:%S')}\nTotal de emails enviados: {i + 1} de {len(destination_emails)} " + f"Intervalo de envio: {interval_send_email} segundos\n"
        Log.write_success(self.data_user.email, message)
        return message
    
    
    def send_emails(self, log_signal: pyqtBoundSignal) -> None:
        body_html_original: str = SenderEmailController.load_html()
        destination_emails, df_emails = self.read_emails()
        
        for i, destination_email in enumerate(destination_emails):
            if i + 1 > self.email_send_quantity:
                return
            
            name, product, protocol = self.get_name_product(df_emails, destination_email)
            message_copy, title_copy, link_whatsapp_copy = self.update_message(name, product, protocol)
            copy_body_html: str = self.replace_email_html(body_html_original, message_copy, title_copy, link_whatsapp_copy)
            self.config_send_email(destination_email, copy_body_html)
            
            try:
                df_emails.loc[df_emails['E-mail'] == destination_email, 'STATUS'] = 'ENVIADO'
                df_emails.to_excel(self.spreadsheet_path, index=False)
                message_success: str = self.write_on_console_and_txt(destination_email, destination_emails, True, i, None, self.send_interval)
                log_signal.emit(message_success)
                time.sleep(self.send_interval)   
            except Exception as error:
                df_emails.loc[df_emails['E-mail'] == destination_email, 'STATUS'] = 'ERRO'
                df_emails.to_excel(self.spreadsheet_path, index=False)
                message_error: str = self.write_on_console_and_txt(destination_email, destination_emails, False, i, error, None)
                log_signal.emit(message_error)
                time.sleep(self.send_interval)
