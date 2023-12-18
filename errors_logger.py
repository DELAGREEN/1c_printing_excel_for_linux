import logging


class ErrorsLogger():


    #def __init__(self, message = None) -> None:
    #    self._message = message

    
    def _log_config(func) -> object:
        def printer(*args, **kwgs) -> object:
            logging.basicConfig(level=logging.INFO,
                            filename="/tmp/excel_printer.log",
                            filemode="a+",
                            format="%(asctime)s %(levelname)s %(message)s",
                            datefmt="%d-%m-%Y %H:%M:%S")
            return func(*args, **kwgs)
        return printer
    

    @_log_config
    def print_warning(_message:str):
        return logging.warning(f'{_message}')
    

    @_log_config
    def print_info(_message:str):
        return logging.info(f'{_message}')
    
    
    @_log_config
    def print_critical(_message:str):
        return logging.critical(f'{_message}')
    

    @_log_config
    def print_debug(_message:str):
        return logging.debug(f'{_message}')


    @_log_config
    def print_error(_message:str):
        return logging.error(f'{_message}')
    

'''
#############
###Example###
#############
logging.debug("A DEBUG Message")
logging.info("An INFO")
logging.warning("A WARNING")
logging.error("An ERROR")
logging.critical("A message of CRITICAL severity")
'''