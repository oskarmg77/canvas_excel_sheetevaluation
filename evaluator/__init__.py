# evaluator/__init__.py

import logging
from logging.handlers import RotatingFileHandler

LOG_FILENAME = 'app.log'


def setup_logging():
    """Configura el logger para guardar en un archivo y mostrar en consola."""
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)

    if logger.hasHandlers():
        logger.handlers.clear()

    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

    fh = RotatingFileHandler(LOG_FILENAME, maxBytes=1 * 1024 * 1024, backupCount=1)
    fh.setFormatter(formatter)
    logger.addHandler(fh)

    ch = logging.StreamHandler()
    ch.setFormatter(formatter)
    logger.addHandler(ch)