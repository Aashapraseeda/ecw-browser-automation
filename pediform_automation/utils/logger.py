"""
utils/logger.py
---------------
Single-call logger factory used across all modules.
Writes to logs/automation.log and to stdout.
"""

import logging
import os
from logging.handlers import RotatingFileHandler

_LOG_DIR  = os.path.join(os.path.dirname(os.path.dirname(__file__)), "logs")
_LOG_FILE = os.path.join(_LOG_DIR, "automation.log")

os.makedirs(_LOG_DIR, exist_ok=True)

_fmt = logging.Formatter(
    fmt="%(asctime)s  %(levelname)-8s  %(name)s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)

_file_handler = RotatingFileHandler(
    _LOG_FILE, maxBytes=5 * 1024 * 1024, backupCount=3, encoding="utf-8"
)
_file_handler.setFormatter(_fmt)

_console_handler = logging.StreamHandler()
_console_handler.setFormatter(_fmt)


def get_logger(name: str) -> logging.Logger:
    logger = logging.getLogger(name)
    if not logger.handlers:
        logger.setLevel(logging.DEBUG)
        logger.addHandler(_file_handler)
        logger.addHandler(_console_handler)
        logger.propagate = False
    return logger
