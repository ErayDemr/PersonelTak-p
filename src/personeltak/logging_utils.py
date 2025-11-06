"""Centralized logging configuration for PersonelTak."""
from __future__ import annotations

import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path

from .config import AppConfig


def setup_logging(config: AppConfig) -> logging.Logger:
    """Configure application wide logging and return the root logger."""

    log_path = Path(config.log_path)
    log_path.parent.mkdir(parents=True, exist_ok=True)

    logger = logging.getLogger("personeltak")
    logger.setLevel(getattr(logging, config.log_level.upper(), logging.INFO))
    logger.propagate = False

    # avoid duplicate handlers when called multiple times
    if not any(isinstance(handler, RotatingFileHandler) for handler in logger.handlers):
        file_handler = RotatingFileHandler(
            log_path,
            maxBytes=2 * 1024 * 1024,
            backupCount=5,
            encoding="utf-8",
        )
        formatter = logging.Formatter(
            fmt="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

    if not any(isinstance(handler, logging.StreamHandler) for handler in logger.handlers):
        stream_handler = logging.StreamHandler()
        stream_handler.setFormatter(
            logging.Formatter("%(levelname)s: %(message)s")
        )
        logger.addHandler(stream_handler)

    logger.debug("Logging initialized at %s level", config.log_level.upper())
    return logger


__all__ = ["setup_logging"]
