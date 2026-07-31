"""
logging_setup.py
-----------------
(added 2026-07-30) This project (unlike lone_star_automation) has NEVER
had any file-based logging - every step only ever used print(), which
only ever reaches the terminal it was run in. Confirmed by inspection:
no `import logging` anywhere in main.py/main_1.py, and no logs/ directory
existed on disk before this file was added. That means the "early steps
scrolled out of view" symptom was NOT just a terminal scrollback limit for
this project - there was no file anywhere holding that output, so once
it's gone from the terminal buffer, it's gone for good.

This module fixes that WITHOUT touching any of the ~400 existing print()
calls across main.py/main_1.py: it transparently duplicates everything
written to stdout into a real, rotating log file, the same way
lone_star_automation's utils/logger.py already does for its own output.
Call enable_full_run_logging() once, near the top of the
`if __name__ == "__main__":` block, before asyncio.run(...).
"""

import os
import sys
import logging
from logging.handlers import RotatingFileHandler

LOG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
_LOG_FILE = os.path.join(LOG_DIR, "automation.log")

_file_logger = logging.getLogger("ecw_automation_full_run")


class _TeeStdout:
    """Writes everything to the real stdout AND to the rotating file
    logger, so the terminal looks identical to before while every line is
    also persisted. print()'s content and its trailing newline arrive as
    two separate .write() calls - splitlines() + skipping empty results
    means neither call produces a spurious blank log line."""

    def __init__(self, real_stream, file_logger):
        self._real = real_stream
        self._file_logger = file_logger

    def write(self, data):
        self._real.write(data)
        for line in data.splitlines():
            if line:
                self._file_logger.info(line)

    def flush(self):
        self._real.flush()

    def isatty(self):
        return getattr(self._real, "isatty", lambda: False)()


def enable_full_run_logging():
    """Idempotent - safe to call more than once (e.g. if a future script
    imports this and also calls it directly)."""
    if not _file_logger.handlers:
        os.makedirs(LOG_DIR, exist_ok=True)
        _file_logger.setLevel(logging.INFO)
        handler = RotatingFileHandler(_LOG_FILE, maxBytes=5 * 1024 * 1024, backupCount=3)
        handler.setFormatter(logging.Formatter("%(message)s"))
        _file_logger.addHandler(handler)
        _file_logger.propagate = False

    if not isinstance(sys.stdout, _TeeStdout):
        sys.stdout = _TeeStdout(sys.stdout, _file_logger)
    if not isinstance(sys.stderr, _TeeStdout):
        sys.stderr = _TeeStdout(sys.stderr, _file_logger)

    return _LOG_FILE
