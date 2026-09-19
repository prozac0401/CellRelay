"""Application logging configuration."""

from __future__ import annotations

import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path


def configure_logging(log_dir: Path) -> Path:
    """Configure console and rotating-file logging and return the log path."""
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / "cellrelay.log"
    formatter = logging.Formatter(
        "%(asctime)s %(levelname)s [pid=%(process)d thread=%(threadName)s] "
        "%(name)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    file_handler = RotatingFileHandler(
        log_path,
        maxBytes=5 * 1024 * 1024,
        backupCount=3,
        encoding="utf-8",
    )
    file_handler.setFormatter(formatter)
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logging.basicConfig(
        level=logging.INFO,
        handlers=[file_handler, console_handler],
        force=True,
    )
    return log_path
