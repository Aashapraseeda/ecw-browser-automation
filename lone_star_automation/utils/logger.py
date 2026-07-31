"""
utils/logger.py
----------------
Shared logger factory: writes to logs/automation.log (rotating) and stdout.

(fix 2026-07-30) Every module in this project calls get_logger(__name__)
with a DIFFERENT name (ecw.login, pcarelink.messenger, patient_forms_now.*,
...). The previous implementation attached a BRAND NEW RotatingFileHandler
to each of those differently-named loggers - so a single production run
ends up with roughly a dozen independent handler objects, each holding its
own open file handle, all pointed at the exact same automation.log. That's
harmless for plain appending, but the moment the file actually crosses the
5MB rotation threshold, more than one of those handlers will try to
rename/reopen the same file at once - and on Windows specifically,
os.rename() refuses to touch a file that another handle still has open,
which is exactly this situation. Confirmed live: 11 separate
RotatingFileHandler instances were found attached across this project's
own module loggers, all targeting the same path.

Fixed by attaching the file/stream handlers ONCE, to the root logger.
Every module's own named logger (returned below) has no handlers of its
own and simply propagates up to root (Python's default) - so there is
only ever ONE handle managing automation.log, no matter how many modules
log to it. Visible output/format is unchanged.
"""

import os
import logging
from logging.handlers import RotatingFileHandler

from config import settings

_LOG_FILE = os.path.join(settings.LOG_DIR, "automation.log")


def _configure_root_once():
    root = logging.getLogger()
    if root.handlers:
        return  # already configured this process

    os.makedirs(settings.LOG_DIR, exist_ok=True)
    root.setLevel(logging.INFO)
    formatter = logging.Formatter("%(message)s")

    file_handler = RotatingFileHandler(
        _LOG_FILE, maxBytes=5 * 1024 * 1024, backupCount=3
    )
    file_handler.setFormatter(formatter)

    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(formatter)

    root.addHandler(file_handler)
    root.addHandler(stream_handler)


def get_logger(name):
    _configure_root_once()
    return logging.getLogger(name)


if __name__ == "__main__":
    log = get_logger("smoke_test")
    log.info("logger smoke-test: writing to %s", _LOG_FILE)
    print("Done — check logs/automation.log")
