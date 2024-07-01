import logging
from logging.handlers import RotatingFileHandler
import traceback
from tkinter import messagebox


def log_exceptions(func):
    def wrapper(*args, **kwargs):
        logger = logging.getLogger("Rotating Log")
        logger.setLevel(logging.ERROR)
        handler = RotatingFileHandler("log.txt", maxBytes=10000, backupCount=5)
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        handler.setFormatter(formatter)
        logger.addHandler(handler)

        try:
            return func(*args, **kwargs)
        except Exception as e:
            logger.error(str(e))
            logger.error(traceback.format_exc())
            messagebox.showerror("Error", "Unexpected error, check logs for more info")
            raise

    return wrapper
