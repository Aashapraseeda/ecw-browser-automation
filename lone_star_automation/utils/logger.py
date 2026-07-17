"""
utils/logger.py
----------------
Shared logger factory: writes to logs/automation.log (rotating) and stdout.
"""

import os
import logging
from logging.handlers import RotatingFileHandler

from config import settings

_LOG_FILE = os.path.join(settings.LOG_DIR, "automation.log")


def get_logger(name):
    os.makedirs(settings.LOG_DIR, exist_ok=True)

    logger = logging.getLogger(name)
    if logger.handlers:
        return logger  # already configured

    logger.setLevel(logging.INFO)
    formatter = logging.Formatter("%(message)s")

    file_handler = RotatingFileHandler(
        _LOG_FILE, maxBytes=5 * 1024 * 1024, backupCount=3
    )
    file_handler.setFormatter(formatter)

    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    logger.addHandler(stream_handler)
    logger.propagate = False

    return logger


if __name__ == "__main__":
    log = get_logger("smoke_test")
    log.info("logger smoke-test: writing to %s", _LOG_FILE)
    print("Done — check logs/automation.log")
