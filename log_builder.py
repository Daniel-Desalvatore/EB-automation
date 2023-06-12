import logging

class MyLogger:
    def __init__(self,log_file='EBiennial.log') -> None:
        self.log_file = log_file
        self.setup_logger()
    
    def setup_logger(self):
        logging.basicConfig(
            level=logging.DEBUG,
            format = '%(asctime)s - %(levelname)s - %(message)s',
            filename=self.log_file,
            filemode= 'w'
        )
    
    def debug(self,message,data=None):
        log_message = f'{message} - {data}' if data else message
        logging.debug(log_message)
    
    def info(self,message,data=None):
        log_message = f'{message} - {data}' if data else message
        logging.info(log_message)
    
    def warning(self,message,data=None):
        log_message = f'{message} - {data}' if data else message
        logging.warning(log_message)
    
    def error(self,message,data=None):
        log_message = f'{message} - {data}' if data else message
        logging.error(log_message)
    
    def critical(self,message,data=None):
        log_message = f'{message} - {data}' if data else message
        logging.critical(log_message)
