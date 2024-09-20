from src.controller.DataUserController import DataUserController
from pydantic import BaseModel, Field

class SenderEmailModel(BaseModel):
    data_user: DataUserController
    spreadsheet_path: str = Field(..., description="Caminho da planilha")
    email_send_quantity: int = Field(..., description="Quantidade de envio de email")
    send_interval: int = Field(..., description="Intervalo de envio")
    email_subject: str = Field(..., description="Assunto do email")
    email_title: str = Field(..., description="Título do email")
    email_message: str = Field(..., description="Mensagem do email")
    whatsapp_redirect_number: str = Field(..., description="Número do whatsapp para redirecionamento")
    
    class Config:
        from_attributes = True
    