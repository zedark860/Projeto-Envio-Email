import time
from src.controller.DataUserController import DataUserController
from src.controller.SenderEmailController import SenderEmailController
from PyQt5.QtCore import QThread, pyqtSignal, pyqtBoundSignal

class EmailSenderThread(QThread):
    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal()

    def __init__(self, data: dict, parent=None):
        super().__init__(parent)
        self.data = data
        self.running = True

    def run(self):
        try:
            while self.running:
                self.requisicao_iniciar_envios()
                time.sleep(self.data["send_interval"])   
                self.finished_signal.emit()
            return
        except Exception as e:
            self.log_signal.emit(f'Erro ao tentar obter logs: {str(e)}')
            self.finished_signal.emit()
            return
        
        
    def requisicao_iniciar_envios(self) -> None:    
        try:
            dataUserController: DataUserController = DataUserController(
                email = self.data["email"],
                app_password = self.data["app_password"]
            )
            
            validator: bool = dataUserController.check_data_user()
            
            if validator:
                self.log_signal.emit('Envio de e-mails iniciado com sucesso!\n')
                
                SenderEmailController(
                    data_user = dataUserController, 
                    spreadsheet_path = self.data["spreadsheet_path"],
                    html_path = self.data["html_path"],
                    email_send_quantity = self.data["email_send_quantity"],
                    send_interval = self.data["send_interval"],
                    email_subject = self.data["email_subject"],
                    email_title = self.data["email_title"],
                    email_message = self.data["email_message"],
                    whatsapp_redirect_number = self.data["whatsapp_redirect_number"]   
                ).send_emails(self.log_signal)
                
            self.stop()
        except Exception as e:
            self.log_signal.emit(f'Erro ao iniciar o envio de e-mails: {str(e)}')
            
            
    def stop(self):
        self.running = False
            
    
    # def requisicao_logs(self) -> None:
    #     last_message_success: str = Log.read_last_success(self.data["email"])
    #     last_message_error: str = Log.read_last_error(self.data["email"])
        
    #     response: str = ""
        
    #     if not last_message_success and not last_message_error:
    #         response: str = 'Nenhum Email enviado ainda.'
    #         self.log_signal.emit(response)
        
    #     if last_message_success:
    #         response: str = last_message_success
        
    #     if last_message_error:
    #         response: str = last_message_error

    #     self.log_signal.emit(str(response))

