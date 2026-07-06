from __future__ import annotations

import logging
from pathlib import Path

from .config import LOG_DIR, LOG_FILE


def configure_logging(project_root: Path) -> Path:
    log_dir = project_root / LOG_DIR
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / LOG_FILE

    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)
    root_logger.handlers.clear()

    formatter = logging.Formatter(
        fmt="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setFormatter(formatter)
    root_logger.addHandler(file_handler)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    root_logger.addHandler(console_handler)

    return log_path
