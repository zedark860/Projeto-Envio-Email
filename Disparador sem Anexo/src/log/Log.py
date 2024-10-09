import os
import sys

class Log:
    
    @staticmethod
    def resource_path(relative_path):
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(
            os.path.abspath(__file__)))
    
        return os.path.join(base_path, relative_path)


    @staticmethod
    def write_success(user_email: str, message: str) -> None:
        archive_name: str = fr"send_success\{user_email}_sucessos.txt"
        if not Log.resource_path(archive_name):
            with open(Log.resource_path(archive_name), "w", encoding="utf-8") as log_file:
                log_file.write(f"{message}\n")
            return
            
        with open(Log.resource_path(archive_name), "a", encoding="utf-8") as log_file:
            log_file.write(f"{message}\n")
            
            
    @staticmethod
    def write_error(user_email: str, message: str) -> None:
        archive_name: str = fr"send_error\{user_email}_erros.txt"
        if not Log.resource_path(archive_name):
            with open(Log.resource_path(archive_name), "w", encoding="utf-8") as log_file:
                log_file.write(f"{message}\n")
                return
                
        with open(Log.resource_path(archive_name), "a", encoding="utf-8") as log_file:
            log_file.write(f"{message}\n")
            
    
    # @staticmethod
    # def read_last_success(user_email: str) -> str:
    #     archive_name: str = fr"send_success\{user_email}_sucessos.txt"
        
    #     if not Log.resource_path(archive_name):
    #         with open(Log.resource_path(archive_name), "w", encoding="utf-8"):
    #             pass
        
    #     with open(Log.resource_path(archive_name), 'r', encoding="utf-8") as file:
    #         lines = file.readlines()
    #         if lines:
    #             return lines[0] + lines[1]
    #         return ""

            
    # @staticmethod
    # def read_last_error(user_email: str) -> str:
    #     archive_name: str = fr"send_error\{user_email}_erros.txt"
    #     if not Log.resource_path(archive_name):
    #         with open(Log.resource_path(archive_name), "w", encoding="utf-8"):
    #             pass
        
    #     with open(Log.resource_path(archive_name), 'r', encoding="utf-8") as file:
    #         lines = file.readlines()
    #         if lines:
    #             return lines[0] + lines[1]
    #         return ""
            
    
